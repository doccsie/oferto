/* Оферто ЛК: xlsx → docx → zip. Всё на клиенте. */
(function () {
  'use strict';

  var TEMPLATE_URL = 'templates/contract_template.docx';

  // поля договора: ключ плейсхолдера → человекочитаемое имя + правила распознавания колонки
  var FIELDS = [
    { key: 'FIO',           label: 'ФИО',            required: true },
    { key: 'SerialNumbPas', label: 'Паспорт' },
    { key: 'Issued',        label: 'Кем выдан' },
    { key: 'IssuedDate',    label: 'Дата выдачи' },
    { key: 'Registration',  label: 'Регистрация' },
    { key: 'INN',           label: 'ИНН' },
    { key: 'SNILS',         label: 'СНИЛС' },
    { key: 'Status',        label: 'Статус' },
    { key: 'Specialty',     label: 'Специальность' },
    { key: 'EMAIL',         label: 'Email' },
    { key: 'AccNumb',       label: 'Счёт' },
    { key: 'BANK',          label: 'Банк' },
    { key: 'BIC',           label: 'БИК' },
    { key: 'CorrAcc',       label: 'Корр. счёт' }
  ];

  function norm(s) {
    return String(s || '').toLowerCase().replace(/ё/g, 'е').replace(/[^a-zа-я0-9]/g, '');
  }

  // порядок проверок важен: более специфичные — раньше
  function matchHeader(h) {
    var n = norm(h);
    if (!n) return null;
    if (n.indexOf('датавыдачи') !== -1) return 'IssuedDate';
    if (n.indexOf('кемвыдан') !== -1 || n.indexOf('выдан') !== -1) return 'Issued';
    if (n.indexOf('фио') !== -1) return 'FIO';
    if (n.indexOf('паспорт') !== -1) return 'SerialNumbPas';
    if (n.indexOf('регистр') !== -1) return 'Registration';
    if (/сни+лс/.test(n) || n.indexOf('snils') !== -1) return 'SNILS';
    if (n.indexOf('статус') !== -1) return 'Status';
    if (n.indexOf('специальн') !== -1 || n.indexOf('квалифи') !== -1) return 'Specialty';
    if (n.indexOf('email') !== -1 || n.indexOf('почта') !== -1 || n.indexOf('мейл') !== -1) return 'EMAIL';
    if (n.indexOf('бик') !== -1) return 'BIC';
    if (n.indexOf('корр') !== -1) return 'CorrAcc';
    if (n.indexOf('счет') !== -1) return 'AccNumb';
    if (n.indexOf('банк') !== -1) return 'BANK';
    if (n === 'инн') return 'INN';
    if (n.indexOf('инн') !== -1 && n.indexOf('банк') === -1) return 'INN';
    return null;
  }

  function pad2(x) { return x < 10 ? '0' + x : '' + x; }

  function cellToString(cell) {
    if (!cell) return '';
    if (cell.t === 'd' && cell.v instanceof Date) {
      var d = cell.v;
      return pad2(d.getUTCDate()) + '.' + pad2(d.getUTCMonth() + 1) + '.' + d.getUTCFullYear();
    }
    if (cell.t === 'n') {
      // большие номера счетов/ИНН: без экспоненты
      if (cell.w && /[eE]/.test(cell.w) === false && cell.w !== 'General') return cell.w;
      return cell.v.toLocaleString('fullwide', { useGrouping: false });
    }
    return String(cell.w != null ? cell.w : (cell.v != null ? cell.v : '')).trim();
  }

  function parseWorkbook(data) {
    var wb = XLSX.read(data, { type: 'array', cellDates: true });
    var ws = wb.Sheets[wb.SheetNames[0]];
    var range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    // ищем строку заголовков в первых 5 строках: та, где распозналось больше всего колонок
    var headerRow = range.s.r, best = 0, colMap = {};
    for (var r = range.s.r; r <= Math.min(range.s.r + 4, range.e.r); r++) {
      var map = {}, hits = 0;
      for (var c = range.s.c; c <= range.e.c; c++) {
        var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
        var key = matchHeader(cell && (cell.w || cell.v));
        if (key && !(key in map)) { map[key] = c; hits++; }
      }
      if (hits > best) { best = hits; headerRow = r; colMap = map; }
    }
    if (best < 3) throw new Error('Не удалось распознать заголовки. Скачайте пример реестра и сверьте названия колонок.');

    var rows = [];
    for (var rr = headerRow + 1; rr <= range.e.r; rr++) {
      var rec = {}, empty = true;
      FIELDS.forEach(function (f) {
        var col = colMap[f.key];
        var v = col == null ? '' : cellToString(ws[XLSX.utils.encode_cell({ r: rr, c: col })]);
        rec[f.key] = v;
        if (v) empty = false;
      });
      if (!empty) rows.push(rec);
    }
    return { rows: rows, colMap: colMap };
  }

  // ── UI ──
  var dropzone = document.getElementById('dropzone');
  var fileInput = document.getElementById('fileInput');
  var fileChip = document.getElementById('fileChip');
  var previewPanel = document.getElementById('previewPanel');
  var previewHint = document.getElementById('previewHint');
  var previewTable = document.getElementById('previewTable');
  var generateBtn = document.getElementById('generateBtn');
  var statusLine = document.getElementById('statusLine');

  var parsed = null;

  dropzone.addEventListener('click', function () { fileInput.click(); });
  dropzone.addEventListener('keydown', function (e) {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); }
  });
  ['dragover', 'dragenter'].forEach(function (ev) {
    dropzone.addEventListener(ev, function (e) { e.preventDefault(); dropzone.classList.add('drag'); });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    dropzone.addEventListener(ev, function (e) { e.preventDefault(); dropzone.classList.remove('drag'); });
  });
  dropzone.addEventListener('drop', function (e) {
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener('change', function () {
    if (fileInput.files.length) handleFile(fileInput.files[0]);
  });

  function setStatus(text, cls) {
    statusLine.textContent = text || '';
    statusLine.className = 'status-line' + (cls ? ' ' + cls : '');
  }

  function handleFile(file) {
    if (!/\.xlsx?$/i.test(file.name)) {
      setStatus('Нужен файл Excel (.xlsx)', 'err');
      return;
    }
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        parsed = parseWorkbook(new Uint8Array(e.target.result));
      } catch (err) {
        parsed = null;
        previewPanel.hidden = true;
        fileChip.innerHTML = '';
        setStatus('', '');
        alertError(err.message);
        return;
      }
      fileChip.innerHTML = '';
      var chip = document.createElement('span');
      chip.className = 'file-chip';
      chip.appendChild(document.createTextNode(file.name + ' · строк: ' + parsed.rows.length + ' '));
      var x = document.createElement('button');
      x.type = 'button'; x.textContent = '✕'; x.title = 'Убрать файл';
      x.addEventListener('click', function () {
        parsed = null; fileChip.innerHTML = ''; previewPanel.hidden = true; fileInput.value = '';
      });
      chip.appendChild(x);
      fileChip.appendChild(chip);
      renderPreview();
    };
    reader.readAsArrayBuffer(file);
  }

  function alertError(msg) {
    fileChip.innerHTML = '<span class="file-chip" style="color:#B3402A">' + msg + '</span>';
  }

  function renderPreview() {
    var rows = parsed.rows;
    var recognized = FIELDS.filter(function (f) { return f.key in parsed.colMap; });
    var missingCols = FIELDS.filter(function (f) { return !(f.key in parsed.colMap); });
    previewHint.textContent = 'Исполнителей: ' + rows.length +
      '. Распознано колонок: ' + recognized.length + ' из ' + FIELDS.length +
      (missingCols.length ? '. Не найдено: ' + missingCols.map(function (f) { return f.label; }).join(', ') + ' — в договоре будет прочерк.' : '.');

    var html = '<table class="preview-table"><thead><tr><th>№</th>';
    FIELDS.forEach(function (f) { html += '<th>' + f.label + '</th>'; });
    html += '<th>Готовность</th></tr></thead><tbody>';
    rows.forEach(function (rec, i) {
      var missing = FIELDS.filter(function (f) { return !rec[f.key]; });
      html += '<tr><td>' + (i + 1) + '</td>';
      FIELDS.forEach(function (f) {
        var v = rec[f.key];
        html += v ? '<td>' + escapeHtml(v) + '</td>' : '<td class="missing">—</td>';
      });
      html += '<td>' + (missing.length === 0
        ? '<span class="badge-ok">полные данные</span>'
        : '<span class="badge-warn">пусто: ' + missing.length + '</span>') + '</td></tr>';
    });
    html += '</tbody></table>';
    previewTable.innerHTML = html;
    previewPanel.hidden = false;
    setStatus('');
    previewPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function escapeHtml(s) {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function safeName(fio, idx, used) {
    var base = fio ? fio.replace(/[\\/:*?"<>|]/g, '').trim().split(/\s+/).slice(0, 3).join('_') : 'Исполнитель_' + (idx + 1);
    var name = 'Договор_' + base, n = name, k = 2;
    while (used[n]) { n = name + '_' + k++; }
    used[n] = true;
    return n + '.docx';
  }

  generateBtn.addEventListener('click', function () {
    if (!parsed || !parsed.rows.length) return;
    generateBtn.disabled = true;
    setStatus('Готовим договоры…');

    fetch(TEMPLATE_URL)
      .then(function (res) {
        if (!res.ok) throw new Error('Не удалось загрузить шаблон договора');
        return res.arrayBuffer();
      })
      .then(function (tplBuf) {
        var outZip = new JSZip();
        var used = {};
        parsed.rows.forEach(function (rec, i) {
          var zip = new PizZip(tplBuf);
          var doc = new window.docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            nullGetter: function () { return '—'; }
          });
          var data = {};
          FIELDS.forEach(function (f) { data[f.key] = rec[f.key] || '—'; });
          doc.render(data);
          outZip.file(safeName(rec.FIO, i, used), doc.getZip().generate({ type: 'uint8array' }));
          setStatus('Готовим договоры… ' + (i + 1) + ' / ' + parsed.rows.length);
        });
        return outZip.generateAsync({ type: 'blob' });
      })
      .then(function (blob) {
        var d = new Date();
        saveAs(blob, 'Оферто_договоры_' + d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate()) + '.zip');
        setStatus('Готово: ' + parsed.rows.length + ' договоров в архиве', 'ok');
      })
      .catch(function (err) {
        setStatus('Ошибка: ' + (err.message || err), 'err');
      })
      .then(function () { generateBtn.disabled = false; });
  });
})();
