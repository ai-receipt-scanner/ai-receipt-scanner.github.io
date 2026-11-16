/* Receipt to CSV converter. Everything runs in the browser: the image is read by
   Tesseract.js locally and is never uploaded. */
(function () {
    'use strict';

    // ---- Parsing (kept free of DOM access so it can be tested on its own) ----

    /** Turns "1.234,56", "1,234.56" or "12,50" into a number. */
    function parseAmount(token) {
        var commas = (token.match(/,/g) || []).length;
        var dots = (token.match(/\./g) || []).length;
        var s = token;
        if (commas && dots) {
            var decimal = token.lastIndexOf(',') > token.lastIndexOf('.') ? ',' : '.';
            var thousands = decimal === ',' ? '.' : ',';
            s = token.split(thousands).join('').replace(decimal, '.');
        } else if (commas === 1) {
            var after = token.length - token.indexOf(',') - 1;
            s = after === 3 ? token.replace(',', '') : token.replace(',', '.');
        } else if (commas > 1) {
            s = token.split(',').join('');
        } else if (dots > 1) {
            s = token.split('.').join('');
        }
        var n = parseFloat(s);
        return isNaN(n) ? null : n;
    }

    var AMOUNT = /\d{1,3}(?:[.,]\d{3})*[.,]\d{2}(?!\d)|\d+[.,]\d{2}(?!\d)/g;
    var TOTAL_WORDS = /(grand total|total|amount due|balance due|summe|gesamt|zu zahlen|net [àa] payer|[àa] payer|importe|totale|totaal|te betalen|valor total)/i;
    var SUBTOTAL_WORDS = /(sub\s*-?\s*total|zwischensumme|sous[\s-]*total|subtotaal|subtotale|total tax|tax total)/i;
    var TAX_WORDS = /\b(sales tax|tax|vat|gst|hst|mwst|ust|tva|iva|btw)\b/i;
    var MONTHS = 'jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|mär|mai|okt|dez|janv|févr|avr|juin|juil|août|déc|ene|abr|ago|dic|gen|mag|giu|lug|set|ott';
    var DATE_PATTERNS = [
        /\b\d{4}[-\/.]\d{1,2}[-\/.]\d{1,2}\b/,
        /\b\d{1,2}[-\/.]\d{1,2}[-\/.]\d{2,4}\b/,
        new RegExp('\\b\\d{1,2}\\.?\\s+(?:' + MONTHS + ')[a-zéûä]*\\.?,?\\s+\\d{2,4}\\b', 'i'),
        new RegExp('\\b(?:' + MONTHS + ')[a-zéûä]*\\.?\\s+\\d{1,2},?\\s+\\d{4}\\b', 'i')
    ];

    function lastAmount(line) {
        var found = line.match(AMOUNT);
        return found ? parseAmount(found[found.length - 1]) : null;
    }

    function detectCurrency(text) {
        if (/€|\bEUR\b/.test(text)) return 'EUR';
        if (/£|\bGBP\b/.test(text)) return 'GBP';
        if (/₪|\bILS\b|\bNIS\b/.test(text)) return 'ILS';
        if (/¥|\bJPY\b/.test(text)) return 'JPY';
        if (/\bCAD\b/.test(text)) return 'CAD';
        if (/\bAUD\b/.test(text)) return 'AUD';
        if (/\$|\bUSD\b/.test(text)) return 'USD';
        return '';
    }

    /** Extracts vendor, date, total, tax and currency from recognised receipt text. */
    function parseReceipt(text) {
        var lines = text.split(/\r?\n/).map(function (l) { return l.trim(); }).filter(Boolean);
        var result = { vendor: '', date: '', total: '', tax: '', currency: detectCurrency(text) };

        for (var i = 0; i < lines.length; i++) {
            var letters = lines[i].replace(/[^A-Za-zÀ-ÿ]/g, '');
            if (letters.length >= 3 && !lines[i].match(AMOUNT)) { result.vendor = lines[i]; break; }
        }

        for (var d = 0; d < lines.length && !result.date; d++) {
            for (var p = 0; p < DATE_PATTERNS.length; p++) {
                var m = lines[d].match(DATE_PATTERNS[p]);
                if (m) { result.date = m[0]; break; }
            }
        }

        var total = null, tax = null, largest = null;
        lines.forEach(function (line) {
            var value = lastAmount(line);
            if (value === null) return;
            if (largest === null || value > largest) largest = value;
            if (SUBTOTAL_WORDS.test(line)) {
                if (TAX_WORDS.test(line)) tax = value;
                return;
            }
            if (TOTAL_WORDS.test(line)) { total = value; return; }
            if (TAX_WORDS.test(line)) tax = value;
        });
        if (total === null) total = largest;
        if (total !== null) result.total = total.toFixed(2);
        if (tax !== null && (total === null || tax < total)) result.tax = tax.toFixed(2);
        return result;
    }

    function csvCell(value) {
        var s = String(value == null ? '' : value);
        if (/^[=+\-@]/.test(s)) s = "'" + s;           // stop spreadsheet formula injection
        return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    }

    function toCSV(rows) {
        var header = ['Vendor', 'Date', 'Total', 'Tax', 'Currency', 'Source file'];
        var out = [header.join(',')];
        rows.forEach(function (r) {
            out.push([r.vendor, r.date, r.total, r.tax, r.currency, r.file].map(csvCell).join(','));
        });
        return '﻿' + out.join('\r\n') + '\r\n';     // BOM so Excel reads UTF-8
    }

    window.ReceiptTool = { parseAmount: parseAmount, parseReceipt: parseReceipt, toCSV: toCSV };

    // ---- Page wiring ----
    var drop = document.getElementById('drop');
    if (!drop) return;
    var fileInput = document.getElementById('file');
    var language = document.getElementById('language');
    var status = document.getElementById('status');
    var progress = document.getElementById('progress');
    var bar = progress.firstElementChild;
    var body = document.getElementById('rows');
    var tableWrap = document.getElementById('results');
    var download = document.getElementById('download');
    var clear = document.getElementById('clear');
    var rows = [];
    var busy = false;

    function setStatus(message, isError) {
        status.textContent = message;
        status.className = 'status' + (isError ? ' error' : '');
    }

    function render() {
        body.innerHTML = '';
        rows.forEach(function (row, index) {
            var tr = document.createElement('tr');
            ['vendor', 'date', 'total', 'tax', 'currency'].forEach(function (field) {
                var td = document.createElement('td');
                var input = document.createElement('input');
                input.type = 'text';
                input.value = row[field];
                input.setAttribute('aria-label', field + ' for receipt ' + (index + 1));
                input.addEventListener('input', function () { row[field] = input.value; });
                td.appendChild(input);
                tr.appendChild(td);
            });
            var remove = document.createElement('td');
            var button = document.createElement('button');
            button.className = 'icon-button';
            button.type = 'button';
            button.textContent = '×';
            button.setAttribute('aria-label', 'Remove receipt ' + (index + 1));
            button.addEventListener('click', function () { rows.splice(index, 1); render(); });
            remove.appendChild(button);
            tr.appendChild(remove);
            body.appendChild(tr);
        });
        tableWrap.style.display = rows.length ? 'block' : 'none';
        download.disabled = clear.disabled = rows.length === 0;
    }

    async function recognise(source, name) {
        if (typeof Tesseract === 'undefined') {
            setStatus('The text reader could not be loaded. Check your connection and reload the page.', true);
            return;
        }
        progress.style.display = 'block';
        bar.style.width = '0%';
        setStatus('Reading ' + name + '…');
        var worker;
        try {
            worker = await Tesseract.createWorker(language.value, 1, {
                logger: function (m) {
                    if (m.status === 'recognizing text') bar.style.width = Math.round(m.progress * 100) + '%';
                }
            });
            var output = await worker.recognize(source);
            var parsed = parseReceipt(output.data.text || '');
            parsed.file = name;
            rows.push(parsed);
            render();
            if (!parsed.total) setStatus('Read ' + name + ', but no total was found. Type it in the table.', true);
            else setStatus('Read ' + name + '. Check the values, then download.');
        } catch (error) {
            setStatus('Could not read ' + name + '. Try a sharper, well-lit photo.', true);
        } finally {
            if (worker) await worker.terminate();
            progress.style.display = 'none';
        }
    }

    async function handleFiles(list) {
        if (busy) return;
        busy = true;
        var files = Array.prototype.slice.call(list).filter(function (f) { return /^image\//.test(f.type); });
        if (!files.length) setStatus('Please choose a photo of a receipt (JPG or PNG).', true);
        for (var i = 0; i < files.length; i++) await recognise(files[i], files[i].name);
        busy = false;
        fileInput.value = '';
    }

    drop.addEventListener('click', function () { fileInput.click(); });
    drop.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); } });
    fileInput.addEventListener('change', function () { handleFiles(fileInput.files); });
    ['dragenter', 'dragover'].forEach(function (name) {
        drop.addEventListener(name, function (e) { e.preventDefault(); drop.classList.add('over'); });
    });
    ['dragleave', 'drop'].forEach(function (name) {
        drop.addEventListener(name, function (e) { e.preventDefault(); drop.classList.remove('over'); });
    });
    drop.addEventListener('drop', function (e) { handleFiles(e.dataTransfer.files); });

    document.getElementById('sample').addEventListener('click', async function () {
        if (busy) return;
        busy = true;
        await recognise('assets/sample-receipt.png', 'sample-receipt.png');
        busy = false;
    });

    download.addEventListener('click', function () {
        var blob = new Blob([toCSV(rows)], { type: 'text/csv;charset=utf-8' });
        var link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'receipts-' + new Date().toISOString().slice(0, 10) + '.csv';
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(function () { URL.revokeObjectURL(link.href); }, 1000);
    });

    clear.addEventListener('click', function () { rows = []; render(); setStatus(''); });
    render();
})();
