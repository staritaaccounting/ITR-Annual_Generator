/* ITR Generator – Annual V2.0 · Excel outputs (ExcelJS, loaded on demand)
 *   computation(M)            → ITR Computation – Annual (.xlsx)
 *   statements(M)             → Financial Statements (.xlsx): Income Statement, Balance Sheet, Cash Flow, Changes in Equity, Income Tax
 *   correctedTemplate(M, buf) → the office ITR Template with corrected cells + "Corrections" and "Annual ITR Summary" sheets
 */
(function () {
  'use strict';
  const NAVY = 'FF0B3A7E', LIGHT = 'FFF2F6FC', SEC = 'FFE8EEF7', WHITE = 'FFFFFFFF', GOLD = 'FFF5B301', YELLOW = 'FFFFF59D', GREY = 'FF64748B';
  const NUM = '#,##0.00;(#,##0.00);"-"';
  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
  const fill = (c, argb) => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } }; };
  const topLine = (c, style) => { c.border = Object.assign({}, c.border, { top: { style: style || 'thin', color: { argb: 'FF0F172A' } } }); };
  const dbl = (c) => { c.border = Object.assign({}, c.border, { bottom: { style: 'double', color: { argb: 'FF0F172A' } } }); };
  const blob = async (wb) => new Blob([await wb.xlsx.writeBuffer()], { type: XLSX_MIME });
  const yearsOf = (M) => [M.year, M.year - 1];
  const entityLine = (M) => M.isProp ? '(A PROPRIETORSHIP)' : '';

  function newBook() {
    const wb = new window.ExcelJS.Workbook();
    wb.creator = 'ITR Generator – Annual V2.0 · Sta. Rita Accounting Office';
    wb.created = new Date();
    wb.calcProperties.fullCalcOnLoad = true;
    return wb;
  }
  function header(ws, M, title, sub, span) {
    const rows = [M.client.name, entityLine(M), title, sub].filter((x) => x !== '');
    rows.forEach((t, i) => {
      const r = ws.addRow([t]);
      ws.mergeCells(r.number, 1, r.number, span);
      r.getCell(1).alignment = { horizontal: 'center' };
      r.getCell(1).font = { bold: i === 0 || t === title, size: i === 0 ? 13 : 11 };
    });
    ws.addRow([]);
  }

  /* ===================================================== ITR COMPUTATION */
  async function computation(M) {
    const wb = newBook();
    const T = M.tax, E = M.expenses;
    const ws = wb.addWorksheet('Annual ITR', { views: [{ showGridLines: false }], pageSetup: { paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    ws.columns = [{ width: 64 }, { width: 22 }, { width: 40 }];
    const t = ws.addRow([T.formLabel + ' — Annual Income Tax Return Computation · ' + M.year]);
    ws.mergeCells(t.number, 1, t.number, 3); t.height = 26;
    fill(t.getCell(1), NAVY); t.getCell(1).font = { bold: true, size: 14, color: { argb: WHITE } }; t.getCell(1).alignment = { vertical: 'middle', indent: 1 };
    const s = ws.addRow(['Sta. Rita Accounting Office · ITR Generator – Annual V2.0 · figures read by Muse Spark 1.3 and computed by formula']);
    ws.mergeCells(s.number, 1, s.number, 3); s.getCell(1).font = { italic: true, size: 9, color: { argb: GREY } };
    ws.addRow([]);
    [['Taxpayer / Registered name', M.client.name], ['TIN', M.client.tin], ['RDO', M.client.rdo || ''], ['Taxable year', 'January 1 – December 31, ' + M.year],
      ['Return', T.formLabel], ['Tax method', M.methodLabel], ['VAT status', M.vat], ['Expense basis', { documents: 'Uploaded expense documents', template: 'ITR Template', prior: 'Prior-year (' + (M.year - 1) + ') ITR/AFS rates' }[E.basis] || '—'],
      ['Prepared by', (M.user && M.user.name) || ''], ['Generated on', new Date().toLocaleString('en-PH')]]
      .forEach(([k, v]) => { const r = ws.addRow([k, v]); ws.mergeCells(r.number, 2, r.number, 3); r.getCell(1).font = { bold: true, color: { argb: 'FF334155' } }; r.getCell(2).alignment = { horizontal: 'left' }; });
    ws.addRow([]);
    const section = (label) => { const r = ws.addRow([label, '', '']); [1, 2, 3].forEach((i) => { fill(r.getCell(i), NAVY); r.getCell(i).font = { bold: true, color: { argb: WHITE } }; }); };
    const line = (label, amt, kind, note) => {
      const r = ws.addRow([label, amt, note || '']);
      if (typeof amt === 'number') r.getCell(2).numFmt = NUM;
      r.getCell(3).font = { size: 9, color: { argb: GREY } };
      if (kind === 'ind') r.getCell(1).alignment = { indent: 2 };
      if (kind === 'sub') r.getCell(1).font = r.getCell(2).font = { bold: true };
      if (kind === 'tot') { [1, 2].forEach((i) => { fill(r.getCell(i), LIGHT); topLine(r.getCell(i)); r.getCell(i).font = { bold: true }; }); }
      if (kind === 'grand') { [1, 2, 3].forEach((i) => { fill(r.getCell(i), NAVY); r.getCell(i).font = { bold: true, size: 12, color: { argb: i === 2 ? GOLD : WHITE } }; }); }
      return r;
    };
    section('REVENUE — QUARTERLY RETURNS');
    M.revenue.byQ.forEach((q) => line('Q' + q.quarter + (M.vat === 'VAT' ? ' — 2550Q net sales' : ' — 2551Q gross receipts'), q.total, 'ind', q.found ? (M.vat === 'VAT' ? 'Output tax ₱' + r2(q.outputTax).toLocaleString('en-PH') : 'Percentage tax ₱' + r2(q.ptDue).toLocaleString('en-PH')) : 'Not uploaded'));
    line('Total sales / receipts for the year', M.revenueTotal, 'tot');
    ws.addRow([]);
    section('COMPUTATION OF TAX — ' + T.formLabel.toUpperCase());
    T.lines.forEach((l) => {
      if (l.type === 'rate') { const r = line(l.label, l.amount / 100, 'ind'); r.getCell(2).numFmt = '0%'; return; }
      line(l.label, l.amount, l.type === 'total' ? 'tot' : l.type === 'sub' ? 'sub' : 'ind', l.note);
    });
    section('LESS: TAX CREDITS / PAYMENTS');
    T.credits.forEach((c) => line(c.label, c.amount, 'ind', c.note));
    line('Total tax credits / payments', T.totalCredits, 'tot');
    line(T.overpayment ? 'OVERPAYMENT' : 'TAX STILL DUE', T.overpayment || T.payable, 'grand');
    ws.addRow([]);
    const n = ws.addRow(['Graduated income tax table applied: ' + (M.year >= 2023 ? 'TRAIN Law, 2023 onwards' : 'TRAIN Law, 2018–2022') + '. MCIT/RCIT per CREATE Act. Check every figure against the source documents before filing in eBIRForms / eFPS.']);
    ws.mergeCells(n.number, 1, n.number, 3); n.getCell(1).font = { italic: true, size: 9, color: { argb: GREY } }; n.getCell(1).alignment = { wrapText: true }; n.height = 28;

    // schedules
    table(wb, 'Deductions Schedule', ['Category', 'Account', 'Amount ' + M.year, 'Basis'], [].concat(
      E.cos.map((l) => ['Cost of sales', l.label, l.amount, l.basis || '']),
      E.opex.map((l) => ['Operating expense', l.label, l.amount, l.basis || ''])), [18, 44, 18, 60], [3],
      ['TOTAL', '', r2(E.cosTotal + E.opexTotal), '']);
    table(wb, 'Quarterly Data', ['Quarter', M.vat === 'VAT' ? 'Net sales (2550Q)' : 'Gross receipts (2551Q)', M.vat === 'VAT' ? 'Output tax' : 'Percentage tax', 'Income tax paid (quarterly ITR)', 'Source'],
      M.revenue.byQ.map((q) => { const p = (T.qPaid || []).find((x) => x.q === q.quarter); return ['Q' + q.quarter, q.total, M.vat === 'VAT' ? q.outputTax : q.ptDue, q.quarter < 4 && p ? p.amount : null, p ? p.src : '']; }),
      [10, 22, 18, 26, 22], [2, 3, 4], ['TOTAL', M.revenueTotal, M.vat === 'VAT' ? M.revenue.byQ.reduce((s, q) => s + (q.outputTax || 0), 0) : M.revenue.ptTotal, (T.qPaid || []).reduce((s, x) => s + x.amount, 0), '']);
    if (M.docs) {
      const certs = (M.docs.filter((d) => d.type === 'form_2307' && d.data).reduce((a, d) => a.concat(d.data.certificates.map((c) => Object.assign({ file: d.file }, c))), []));
      if (certs.length) table(wb, 'BIR 2307', ['Payor', 'Payor TIN', 'Quarter', 'Period', 'ATC', 'Income payment', 'Tax withheld', 'File'],
        certs.map((c) => [c.payor_name, c.payor_tin, c.quarter ? 'Q' + c.quarter : '', [c.period_from, c.period_to].filter(Boolean).join(' – '), (c.lines || []).map((l) => l.atc).filter(Boolean).join(', '), c.total_income, c.total_tax_withheld, c.file]),
        [34, 18, 9, 24, 12, 16, 14, 24], [6, 7], ['TOTAL', '', '', '', '', certs.reduce((s, c) => s + c.total_income, 0), certs.reduce((s, c) => s + c.total_tax_withheld, 0), '']);
    }
    if (M.warnings.length) table(wb, 'Review Notes', ['Items to review before filing'], M.warnings.map((w) => [w]), [130], []);
    return blob(wb);
  }

  function table(wb, name, headers, rows, widths, numCols, totals) {
    const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = headers.map((h, i) => ({ header: h, width: widths[i] || 16 }));
    ws.getRow(1).eachCell((c) => { fill(c, NAVY); c.font = { bold: true, color: { argb: WHITE } }; });
    rows.forEach((r) => ws.addRow(r));
    numCols.forEach((ci) => { ws.getColumn(ci).numFmt = NUM; });
    if (totals && rows.length) { const t = ws.addRow(totals); t.font = { bold: true }; t.eachCell((c) => { fill(c, LIGHT); topLine(c); }); }
    return ws;
  }

  /* ===================================================== FINANCIAL STATEMENTS */
  // Layout follows the office samples: label | notes | current | prior | increase (decrease) | basis / notes
  function fsSheet(wb, name) {
    const ws = wb.addWorksheet(name, { views: [{ showGridLines: false }], pageSetup: { paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    ws.columns = [{ width: 52 }, { width: 8 }, { width: 18 }, { width: 18 }, { width: 16 }, { width: 58 }];
    return ws;
  }
  function colHeads(ws, M, kind) {
    const [y, py] = yearsOf(M);
    const lab = kind === 'bs' ? ['As of\nDec 31, ' + y, 'As of\nDec 31, ' + py, 'Change'] : ['Year Ended\nDec 31, ' + y, 'Year Ended\nDec 31, ' + py, 'Increase\n(Decrease)'];
    const r = ws.addRow(['', 'Notes', lab[0], lab[1], lab[2], 'Basis / Notes']);
    r.height = 32;
    r.eachCell((c, i) => { c.font = { bold: true, size: 9.5 }; c.alignment = { horizontal: i === 1 || i === 6 ? 'left' : 'center', vertical: 'middle', wrapText: true }; c.border = { bottom: { style: 'thin' } }; });
  }
  function fsLine(ws, label, cur, prior, opts) {
    opts = opts || {};
    const r = ws.addRow([label, opts.note || '', null, null, null, opts.basis || '']);
    const n = r.number;
    if (opts.formula) r.getCell(3).value = { formula: opts.formula.replace(/#/g, 'C'), result: r2(cur) };
    else r.getCell(3).value = cur == null ? null : r2(cur);
    if (opts.priorFormula && prior != null) r.getCell(4).value = { formula: opts.priorFormula.replace(/#/g, 'D'), result: r2(prior) };
    else r.getCell(4).value = prior == null ? null : r2(prior);
    if (cur != null || prior != null) r.getCell(5).value = { formula: 'C' + n + '-D' + n, result: r2((cur || 0) - (prior || 0)) };
    [3, 4, 5].forEach((i) => { r.getCell(i).numFmt = NUM; });
    r.getCell(6).font = { size: 8.5, italic: true, color: { argb: GREY } };
    r.getCell(6).alignment = { wrapText: true, vertical: 'top' };
    if (opts.indent) r.getCell(1).alignment = { indent: 1 };
    if (opts.bold) [1, 3, 4, 5].forEach((i) => { r.getCell(i).font = { bold: true }; });
    if (opts.total) [3, 4, 5].forEach((i) => topLine(r.getCell(i)));
    if (opts.grand) [3, 4, 5].forEach((i) => { topLine(r.getCell(i)); dbl(r.getCell(i)); r.getCell(i).font = { bold: true }; });
    if (opts.grand) r.getCell(1).font = { bold: true };
    return n;
  }
  function secRow(ws, label) { const r = ws.addRow([label]); r.getCell(1).font = { bold: true }; return r.number; }

  async function statements(M) {
    const wb = newBook();
    const F = M.fs, s = F.sci, B = F.sfp, C = F.scf, [y, py] = yearsOf(M);

    // ---------- Income Statement
    let ws = fsSheet(wb, 'Income Statement');
    header(ws, M, s.title, 'For the Years Ended December 31, ' + y + ' and ' + py, 6);
    colHeads(ws, M, 'is');
    const revRow = fsLine(ws, M.isProp ? 'REVENUES' : 'SALES', s.revenue, s.priorRevenue, { bold: true, basis: 'Sum of Q1–Q4 ' + (M.vat === 'VAT' ? '2550Q (net of VAT, incl. exempt)' : '2551Q gross receipts') });
    let gpRow = null, cosTotRow = null;
    if (s.cos.length) {
      ws.addRow([]); secRow(ws, 'COST OF SALES');
      const a = ws.rowCount + 1;
      s.cos.forEach((l) => fsLine(ws, l.label, l.amount, l.prior, { indent: true, basis: l.basis }));
      const b = ws.rowCount;
      cosTotRow = fsLine(ws, 'Total cost of sales', s.cosTotal, s.priorCos, { total: true, formula: 'SUM(#' + a + ':#' + b + ')', priorFormula: s.cos.some((l) => l.prior != null) ? 'SUM(#' + a + ':#' + b + ')' : null });
      gpRow = fsLine(ws, 'GROSS PROFIT', s.grossProfit, s.priorGross, { bold: true, formula: '#' + revRow + '-#' + cosTotRow, priorFormula: '#' + revRow + '-#' + cosTotRow });
    }
    ws.addRow([]); secRow(ws, 'OPERATING EXPENSES');
    const oa = ws.rowCount + 1;
    s.opex.forEach((l) => fsLine(ws, l.label, l.amount, l.prior, { indent: true, basis: l.basis }));
    const ob = ws.rowCount;
    const opTot = fsLine(ws, 'TOTAL OPERATING EXPENSES', s.opexTotal, s.priorOpex, { total: true, bold: true, formula: 'SUM(#' + oa + ':#' + ob + ')', priorFormula: 'SUM(#' + oa + ':#' + ob + ')' });
    ws.addRow([]);
    const base = gpRow || revRow;
    let ibtRow;
    if (s.otherIncome || s.priorOtherIncome || !M.isProp) {
      const oi = fsLine(ws, 'OPERATING INCOME', s.operatingIncome, s.priorGross != null ? r2((s.priorGross || s.priorRevenue) - s.priorOpex) : null, { bold: true, formula: '#' + base + '-#' + opTot, priorFormula: '#' + base + '-#' + opTot });
      const ot = fsLine(ws, 'Other income', s.otherIncome, s.priorOtherIncome, { indent: true });
      ibtRow = fsLine(ws, 'INCOME BEFORE INCOME TAX', s.ibt, s.priorIbt, { bold: true, total: true, formula: '#' + oi + '+#' + ot, priorFormula: '#' + oi + '+#' + ot });
    } else ibtRow = fsLine(ws, 'INCOME BEFORE INCOME TAX', s.ibt, s.priorIbt, { bold: true, total: true, formula: '#' + base + '-#' + opTot, priorFormula: '#' + base + '-#' + opTot });
    const prov = fsLine(ws, 'Provision for income tax', s.provision, s.priorProvision, { indent: true, basis: M.tax.formLabel + ' — ' + M.tax.schedule });
    ws.addRow([]);
    fsLine(ws, 'NET INCOME', s.netIncome, s.priorNetIncome, { grand: true, formula: '#' + ibtRow + '-#' + prov, priorFormula: '#' + ibtRow + '-#' + prov });
    ws.addRow([]); ws.addRow(['See accompanying Notes to Financial Statements.']).getCell(1).font = { italic: true, size: 9 };

    // ---------- Balance Sheet
    ws = fsSheet(wb, 'Balance Sheet');
    header(ws, M, B.title, 'As of December 31, ' + y + ' and ' + py, 6);
    colHeads(ws, M, 'bs');
    const c = B.cur, p = B.prior, L = B.labels || {};
    const lab = (k, d) => L[k] || d;
    secRow(ws, 'A S S E T S'); secRow(ws, 'Current Assets');
    const ca1 = ws.rowCount + 1;
    fsLine(ws, lab('cash', 'Cash and cash equivalents'), c.cash, p.cash, { indent: true, basis: 'Balancing figure: total liabilities and equity less other assets' });
    [['receivables', 'Accounts receivable'], ['inventory', 'Inventory'], ['prepaid', 'Prepaid expenses'], ['other_ca', 'Other current assets']].forEach(([k, d]) => { if (c[k] || p[k]) fsLine(ws, lab(k, d), c[k], p[k], { indent: true }); });
    const caT = fsLine(ws, 'Total Current Assets', B.totalCA, B.priorCA, { total: true, bold: true, formula: 'SUM(#' + ca1 + ':#' + ws.rowCount + ')', priorFormula: 'SUM(#' + ca1 + ':#' + ws.rowCount + ')' });
    ws.addRow([]); secRow(ws, 'Non-current Assets');
    let ppeRow;
    if (B.hasGrossPPE) {
      const pc = fsLine(ws, lab('ppe_cost', 'Property and equipment, at cost'), c.ppe_cost, p.ppe_cost, { indent: true, basis: 'Prior cost + additions' });
      const ad = fsLine(ws, 'Less: Accumulated depreciation', c.accumulated_depreciation, p.accumulated_depreciation, { indent: true, basis: 'Prior + depreciation for the year ₱' + r2(F.dep).toLocaleString('en-PH') });
      ppeRow = fsLine(ws, 'Property and equipment, net', c.ppe_net, p.ppe_net, { indent: true, total: true, formula: '#' + pc + '-#' + ad, priorFormula: '#' + pc + '-#' + ad });
    } else ppeRow = fsLine(ws, lab('ppe_net', 'Property and equipment, net'), c.ppe_net, p.ppe_net, { indent: true, basis: 'Prior net + additions − depreciation' });
    let ncaLast = ppeRow;
    if (c.other_nca || p.other_nca) ncaLast = fsLine(ws, lab('other_nca', 'Other non-current assets'), c.other_nca, p.other_nca, { indent: true });
    const ncaT = fsLine(ws, 'Total Non-current Assets', B.totalNCA, B.priorNCA, { total: true, bold: true, formula: '#' + ppeRow + (ncaLast !== ppeRow ? '+#' + ncaLast : ''), priorFormula: '#' + ppeRow + (ncaLast !== ppeRow ? '+#' + ncaLast : '') });
    ws.addRow([]);
    fsLine(ws, 'TOTAL ASSETS', B.totalAssets, B.priorAssets, { grand: true, formula: '#' + caT + '+#' + ncaT, priorFormula: '#' + caT + '+#' + ncaT });
    ws.addRow([]);
    secRow(ws, M.isProp ? "LIABILITIES AND OWNER'S EQUITY" : "LIABILITIES AND STOCKHOLDERS' EQUITY"); secRow(ws, 'Liabilities');
    const l1 = ws.rowCount + 1;
    [['payables', 'Accounts payable'], ['accrued', 'Accrued expenses payable'], ['vat_payable', 'VAT payable'], ['income_tax_payable', 'Income tax payable'], ['other_cl', 'Other current liabilities'], ['loans', 'Loans payable'], ['other_ncl', 'Other non-current liabilities']]
      .forEach(([k, d]) => { if (c[k] || p[k]) fsLine(ws, lab(k, d), c[k], p[k], { indent: true, basis: k === 'income_tax_payable' ? 'Tax still due per ' + M.tax.formLabel : k === 'payables' || k === 'accrued' ? 'Prior year adjusted for revenue growth (editable in the app)' : k === 'loans' ? 'Outstanding balance per ITR Template / prior year' : '' }); });
    const lT = fsLine(ws, 'Total Liabilities', B.totalLiab, B.priorLiab, { total: true, bold: true, formula: 'SUM(#' + l1 + ':#' + ws.rowCount + ')', priorFormula: 'SUM(#' + l1 + ':#' + ws.rowCount + ')' });
    ws.addRow([]);
    const eq = B.equity;
    let eT;
    if (M.isProp) {
      secRow(ws, "Proprietor's Equity");
      const e1 = fsLine(ws, 'Balance, beginning of year', eq.begin, M.prior.equity.beginning || null, { indent: true });
      if (eq.capitalAdd) fsLine(ws, 'Additional capital', eq.capitalAdd, null, { indent: true });
      fsLine(ws, 'Drawings', -eq.drawings, M.prior.equity.drawings ? -M.prior.equity.drawings : null, { indent: true });
      const e3 = fsLine(ws, 'Net income', eq.netIncome, M.prior.sci.net_income || null, { indent: true });
      eT = fsLine(ws, 'Balance, end of year', eq.end, B.priorEquity, { total: true, bold: true, formula: 'SUM(#' + e1 + ':#' + e3 + ')' });
    } else {
      secRow(ws, "Stockholders' Equity");
      const e1 = ws.rowCount + 1;
      const lines = (eq.lines || []).length ? eq.lines : [{ label: 'Retained earnings', amount: eq.begin, kind: 'retained_earnings' }];
      let reDone = false;
      lines.forEach((l) => {
        let cur = Number(l.amount) || 0;
        let basis = '';
        if (!reDone && l.kind === 'retained_earnings') { cur = cur + eq.netIncome - eq.dividends; reDone = true; basis = 'Prior + net income' + (eq.dividends ? ' − dividends' : ''); }
        if (l.kind === 'capital' && eq.capitalAdd) { cur += eq.capitalAdd; basis = 'Prior + additional capital'; }
        fsLine(ws, l.label, cur, l.amount, { indent: true, basis });
      });
      if (!reDone) fsLine(ws, 'Net income for the year', eq.netIncome, null, { indent: true });
      eT = fsLine(ws, "Total Stockholders' Equity", eq.end, B.priorEquity, { total: true, bold: true, formula: 'SUM(#' + e1 + ':#' + ws.rowCount + ')', priorFormula: 'SUM(#' + e1 + ':#' + ws.rowCount + ')' });
    }
    ws.addRow([]);
    fsLine(ws, M.isProp ? "TOTAL LIABILITIES AND OWNER'S EQUITY" : "TOTAL LIABILITIES AND STOCKHOLDERS' EQUITY", B.totalLE, B.priorLE, { grand: true, formula: '#' + lT + '+#' + eT, priorFormula: '#' + lT + '+#' + eT });
    ws.addRow([]); ws.addRow(['See accompanying Notes to Financial Statements.']).getCell(1).font = { italic: true, size: 9 };

    // ---------- Cash Flow
    ws = fsSheet(wb, 'Cash Flow Statement');
    header(ws, M, 'STATEMENTS OF CASH FLOWS', 'For the Year Ended December 31, ' + y + ' — Indirect Method', 6);
    colHeads(ws, M, 'cf');
    const block = (title, list, total, label, priorTotal) => {
      secRow(ws, title);
      const a = ws.rowCount + 1;
      (list.length ? list : [{ label: 'None', amount: 0 }]).forEach((l) => fsLine(ws, l.label, l.amount, null, { indent: true }));
      return fsLine(ws, label, total, priorTotal != null ? priorTotal : null, { total: true, bold: true, formula: 'SUM(#' + a + ':#' + ws.rowCount + ')' });
    };
    const r1 = block('CASH FLOWS FROM OPERATING ACTIVITIES', C.op, C.cfo, 'Net cash provided by (used in) operating activities', C.prior.cfo || null);
    ws.addRow([]);
    const r2_ = block('CASH FLOWS FROM INVESTING ACTIVITIES', C.inv, C.cfi, 'Net cash provided by (used in) investing activities', C.prior.cfi || null);
    ws.addRow([]);
    const r3 = block('CASH FLOWS FROM FINANCING ACTIVITIES', C.fin, C.cff, 'Net cash provided by (used in) financing activities', C.prior.cff || null);
    ws.addRow([]);
    const nc = fsLine(ws, 'NET INCREASE (DECREASE) IN CASH', C.netChange, C.prior.net_change || null, { bold: true, formula: '#' + r1 + '+#' + r2_ + '+#' + r3 });
    const cb = fsLine(ws, 'CASH AT BEGINNING OF YEAR', C.cashBegin, C.prior.cash_beginning || null, {});
    fsLine(ws, 'CASH AT END OF YEAR', C.cashEnd, C.prior.cash_end || null, { grand: true, formula: '#' + nc + '+#' + cb, basis: C.reconciles ? 'Agrees with the balance sheet' : 'Check: does not agree with the balance sheet' });
    ws.addRow([]); ws.addRow(['See accompanying Notes to Financial Statements.']).getCell(1).font = { italic: true, size: 9 };

    // ---------- Changes in Equity
    ws = fsSheet(wb, 'Changes in Equity');
    header(ws, M, M.isProp ? "STATEMENTS OF CHANGES IN OWNER'S EQUITY" : "STATEMENTS OF CHANGES IN EQUITY", 'For the Years Ended December 31, ' + y + ' and ' + py, 6);
    colHeads(ws, M, 'is');
    const q = F.sce;
    fsLine(ws, 'Balance, beginning of year', q.begin, q.priorBegin || null, {});
    if (q.capitalAdd) fsLine(ws, 'Additional capital', q.capitalAdd, null, { indent: true });
    fsLine(ws, 'Net income for the year', q.netIncome, q.priorNetIncome || null, { indent: true });
    if (M.isProp) fsLine(ws, 'Drawings', -q.drawings, q.priorDrawings ? -q.priorDrawings : null, { indent: true });
    else if (q.dividends || q.priorDividends) fsLine(ws, 'Dividends', -q.dividends, q.priorDividends ? -q.priorDividends : null, { indent: true });
    fsLine(ws, 'Balance, end of year', q.end, q.priorEnd || null, { grand: true });

    // ---------- Income tax reconciliation
    ws = fsSheet(wb, 'Income Tax');
    header(ws, M, 'INCOME TAX PAYABLE RECONCILIATION', 'Year Ended December 31, ' + y, 6);
    const T = M.tax;
    fsLine(ws, 'Income before income tax (per books)', F.sci.ibt, null, {});
    fsLine(ws, 'Taxable income per ' + T.formLabel, T.taxable, null, { basis: T.schedule });
    fsLine(ws, 'Income tax due', T.taxDue, null, { bold: true });
    T.credits.filter((x) => x.amount).forEach((x) => fsLine(ws, 'Less: ' + x.label, -x.amount, null, { indent: true, basis: x.note }));
    fsLine(ws, T.overpayment ? 'OVERPAYMENT' : 'INCOME TAX PAYABLE', T.overpayment ? -T.overpayment : T.payable, null, { grand: true });

    // ---------- PPE schedule (if the prior AFS listed classes)
    const ppe = M.prior.ppe || [];
    if (ppe.length) {
      ws = wb.addWorksheet('PPE Schedule', { views: [{ showGridLines: false }] });
      ws.columns = [{ width: 34 }, { width: 18 }, { width: 20 }, { width: 18 }];
      header(ws, M, 'SCHEDULE OF PROPERTY AND EQUIPMENT', 'As of December 31, ' + py + ' (per prior-year AFS)', 4);
      const h = ws.addRow(['Class', 'Cost', 'Accumulated depreciation', 'Net book value']); h.eachCell((cc) => { cc.font = { bold: true }; cc.border = { bottom: { style: 'thin' } }; });
      ppe.forEach((x) => { const r = ws.addRow([x.class, x.cost, x.accumulated_depreciation, x.net]); [2, 3, 4].forEach((i) => { r.getCell(i).numFmt = NUM; }); });
      const add = ws.addRow(['Additions in ' + y, M.fs.sfp.cur.ppe_cost != null ? r2(M.fs.sfp.cur.ppe_cost - M.fs.sfp.prior.ppe_cost) : null, null, null]); add.getCell(2).numFmt = NUM;
      const dp = ws.addRow(['Depreciation for ' + y, null, r2(F.dep), null]); dp.getCell(3).numFmt = NUM;
    }
    return blob(wb);
  }

  /* ===================================================== CORRECTED TEMPLATE */
  function parseRef(ref, defaultSheet) {
    const m = String(ref || '').replace(/\$/g, '').match(/^(?:'?([^'!]+)'?!)?([A-Z]{1,3}\d{1,6})$/i);
    if (!m) return null;
    return { sheet: m[1] || defaultSheet, addr: m[2].toUpperCase() };
  }
  async function correctedTemplate(M, buf) {
    const wb = newBook();
    let loaded = false, note = '';
    if (buf) {
      try { await wb.xlsx.load(buf); loaded = true; }
      catch (e) { note = 'The original template could not be opened (' + e.message + '). Use an .xlsx file to get the corrections written into it.'; }
    }
    const first = loaded ? wb.worksheets[0] : null;
    const results = [];
    M.corrections.forEach((c) => {
      let status = c.ok ? 'Matches' : 'Listed only';
      let where = c.cell;
      if (!c.ok && loaded) {
        let ref, value;
        if (!c.isFormula) { ref = parseRef(c.cell, first.name); value = c.verified; }
        else if (c.inputCell && c.writeValue != null) { ref = parseRef(c.inputCell, first.name); value = c.writeValue; }
        if (ref) {
          const ws = wb.getWorksheet(ref.sheet) || first;
          const cell = ws.getCell(ref.addr);
          const was = cell.value && typeof cell.value === 'object' && 'formula' in cell.value ? null : cell.value;
          if (was === null && cell.value && cell.value.formula) { status = 'Listed only (cell holds a formula)'; }
          else {
            cell.value = r2(value);
            fill(cell, YELLOW);
            cell.note = 'Corrected by ITR Generator (' + new Date().toLocaleDateString('en-PH') + '): was ' + (was == null ? 'blank' : was) + '; per ' + c.source + '.';
            status = 'Corrected'; where = ws.name + '!' + ref.addr;
          }
        }
      }
      results.push([c.label, c.template, c.verified, c.ok ? 0 : c.diff, c.source, where, status]);
    });

    // Corrections sheet
    const cs = wb.addWorksheet('Corrections', { views: [{ state: 'frozen', ySplit: 3 }] });
    cs.columns = [{ width: 34 }, { width: 18 }, { width: 18 }, { width: 16 }, { width: 22 }, { width: 22 }, { width: 30 }];
    const t = cs.addRow(['ITR Template check — ' + M.client.name + ' · ' + M.year]); t.font = { bold: true, size: 13 };
    cs.addRow([note || 'Figures in the template compared with the quarterly returns and quarterly ITRs read by Muse Spark. Corrected cells are highlighted yellow with a note.']).font = { italic: true, size: 9, color: { argb: GREY } };
    const h = cs.addRow(['Item', 'Template', 'Per returns', 'Difference', 'Source', 'Cell', 'Status']);
    h.eachCell((cc) => { fill(cc, NAVY); cc.font = { bold: true, color: { argb: WHITE } }; });
    results.forEach((r) => { const row = cs.addRow(r); [2, 3, 4].forEach((i) => { row.getCell(i).numFmt = NUM; }); if (r[6] === 'Corrected') fill(row.getCell(7), YELLOW); });
    if (!results.length) cs.addRow(['No checkpoints were found in the template.']);

    // Annual ITR Summary sheet (final figures used)
    const ss = wb.addWorksheet('Annual ITR Summary', { views: [{ showGridLines: false }] });
    ss.columns = [{ width: 52 }, { width: 20 }, { width: 50 }];
    ss.addRow([M.client.name]).font = { bold: true, size: 13 };
    ss.addRow([M.client.tin + ' · ' + M.tax.formLabel + ' · FOR THE YEAR ' + M.year]).font = { size: 10 };
    ss.addRow([]);
    const put = (label, v, bold, basis) => { const r = ss.addRow([label, v, basis || '']); if (typeof v === 'number') r.getCell(2).numFmt = NUM; if (bold) r.font = { bold: true }; r.getCell(3).font = { size: 9, italic: true, color: { argb: GREY } }; };
    put('SALES / REVENUE', null, true);
    M.revenue.byQ.forEach((q) => put('  Q' + q.quarter, q.total, false, q.found ? (M.vat === 'VAT' ? '2550Q' : '2551Q') : 'not uploaded'));
    put('Total', M.revenueTotal, true);
    ss.addRow([]);
    if (M.expenses.cos.length) { put('COST OF SALES', null, true); M.expenses.cos.forEach((l) => put('  ' + l.label, l.amount, false, l.basis)); put('Total cost of sales', M.expenses.cosTotal, true); ss.addRow([]); }
    put('OPERATING EXPENSES', null, true); M.expenses.opex.forEach((l) => put('  ' + l.label, l.amount, false, l.basis)); put('Total operating expenses', M.expenses.opexTotal, true);
    ss.addRow([]);
    put('QUARTERLY ITRs — tax paid', null, true); (M.tax.qPaid || []).forEach((x) => put('  Q' + x.q, x.amount, false, x.src));
    ss.addRow([]);
    put('ANNUAL INCOME TAX', null, true);
    put('  Taxable income', M.tax.taxable); put('  Income tax due', M.tax.taxDue); put('  Less: total credits', M.tax.totalCredits); put(M.tax.overpayment ? '  Overpayment' : '  Tax still due', M.tax.overpayment || M.tax.payable, true);
    put('  Net income per books', M.fs.sci.netIncome);
    return blob(wb);
  }

  window.ITRExcel = { computation, statements, correctedTemplate };
})();
