/* ITR Generator – Annual V2.0 · Audited Financial Statements draft (Word + PDF)
 * One content model → two renderers: docx (docx.js) and PDF (pdfmake).
 * The independent auditor's report is NOT generated: a placeholder page is left for the CPA.
 */
(function () {
  'use strict';
  const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
  const f = (n) => { if (n == null || n === '') return ''; const v = r2(n); if (!v) return '-'; const s = Math.abs(v).toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); return v < 0 ? '(' + s + ')' : s; };
  const peso = (n) => '₱' + Math.abs(r2(n)).toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  /* ------------------------------------------------------------ content model */
  function build(M) {
    const F = M.fs, s = F.sci, B = F.sfp, C = F.scf, T = M.tax, y = M.year, py = M.year - 1;
    const E = M.entity || {};
    const name = M.client.name;
    const prop = M.isProp;
    const framework = E.framework || 'Philippine Financial Reporting Standards for Small and Medium-sized Entities (PFRS for SMEs)';
    const officers = (E.officers || []).filter((o) => o && o.name);
    const signers = prop ? (officers.length ? officers.slice(0, 1) : [{ name: '______________________', title: 'Proprietor' }])
      : (officers.length ? officers.slice(0, 3) : [{ name: '______________________', title: 'Chairman of the Board' }, { name: '______________________', title: 'Chief Executive Officer' }, { name: '______________________', title: 'Chief Financial Officer' }]);
    const blocks = [];
    const H = (text, sub) => { blocks.push({ t: 'title', lines: [name, prop ? '(A PROPRIETORSHIP)' : null, text, sub].filter(Boolean) }); };
    const P = (text, o) => blocks.push(Object.assign({ t: 'p', text }, o || {}));
    const PB = () => blocks.push({ t: 'pb' });
    const tbl = (head, rows, widths) => blocks.push({ t: 'table', head, rows, widths: widths || [52, 8, 20, 20] });
    const see = () => P('See accompanying Notes to Financial Statements.', { italic: true, size: 8.5 });
    const row = (label, cur, prior, style, note) => ({ cells: [label, note || '', f(cur), prior == null ? '' : f(prior)], style });
    const head = (a, b) => ['', 'Note', a, b];

    // ---- cover
    blocks.push({ t: 'cover', lines: [name, prop ? '(A Proprietorship)' : '', 'FINANCIAL STATEMENTS', 'December 31, ' + y + ' and ' + py, '', M.client.address || E.address || '', 'TIN ' + (M.client.tin || '')].filter((x) => x !== '') });
    PB();

    // ---- statement of management's responsibility
    blocks.push({ t: 'h', text: "STATEMENT OF MANAGEMENT'S RESPONSIBILITY FOR FINANCIAL STATEMENTS" });
    P('The management of ' + name + (prop ? ' (A Proprietorship)' : '') + ' is responsible for the preparation and fair presentation of the financial statements, including the schedules attached therein, for the years ended December 31, ' + y + ' and ' + py + ', in accordance with ' + framework + ', and for such internal control as management determines is necessary to enable the preparation of financial statements that are free from material misstatement, whether due to fraud or error.');
    P('In preparing the financial statements, management is responsible for assessing the ' + (prop ? 'Proprietorship' : 'Company') + "'s ability to continue as a going concern, disclosing, as applicable, matters related to going concern and using the going concern basis of accounting unless management either intends to liquidate the " + (prop ? 'Proprietorship' : 'Company') + ' or to cease operations, or has no realistic alternative but to do so.');
    P((prop ? 'The Proprietor is' : 'The Board of Directors is') + ' responsible for overseeing the ' + (prop ? 'Proprietorship' : 'Company') + "'s financial reporting process" + (prop ? '.' : ', and reviews and approves the financial statements and submits the same to the stockholders or members.'));
    P('The independent auditor appointed ' + (prop ? 'by the Proprietor' : 'by the stockholders') + ' has audited the financial statements in accordance with Philippine Standards on Auditing and, in its report to ' + (prop ? 'the Proprietor' : 'the stockholders or members') + ', has expressed its opinion on the fairness of presentation upon completion of such audit.');
    blocks.push({ t: 'sig', people: signers, date: 'Signed this ___ day of __________ ' + (y + 1) });
    PB();

    // ---- statement of responsibility for the annual ITR
    blocks.push({ t: 'h', text: 'STATEMENT OF RESPONSIBILITY FOR ANNUAL INCOME TAX RETURN' });
    P('The management of ' + name + ' is responsible for all information and representations contained in the Annual Income Tax Return for the year ended December 31, ' + y + '. Management is likewise responsible for all information and representations contained in the financial statements accompanying the Annual Income Tax Return covering the same reporting period. Furthermore, the management is responsible for all information and representations contained in all the other tax returns filed for the reporting period, including, but not limited to, the value added tax and/or percentage tax returns, withholding tax returns, documentary stamp tax returns, and any and all other tax returns.');
    P('In this regard, the management affirms that the attached audited financial statements for the year ended December 31, ' + y + ' and the accompanying Annual Income Tax Return are in accordance with the books and records of ' + name + ', complete and correct in all material respects. Management likewise affirms that:');
    P('(a) the Annual Income Tax Return has been prepared in accordance with the provisions of the National Internal Revenue Code, as amended, and pertinent tax regulations and other issuances of the Department of Finance and the Bureau of Internal Revenue;', { indent: true });
    P('(b) any disparity of figures in the submitted reports arising from the preparation of financial statements pursuant to financial accounting standards and the preparation of the income tax return pursuant to tax accounting rules has been reported as reconciling items and maintained in the ' + (prop ? "Proprietorship's" : "Company's") + ' books and records in accordance with the requirements of Revenue Regulations No. 8-2007 and other relevant issuances;', { indent: true });
    P('(c) ' + name + ' has filed all applicable tax returns, reports and statements required to be filed under Philippine tax laws for the reporting period, and all taxes and other impositions shown thereon to be due and payable have been paid for the reporting period, except those contested in good faith.', { indent: true });
    blocks.push({ t: 'sig', people: signers, date: 'Signed this ___ day of __________ ' + (y + 1) });
    PB();

    // ---- auditor's report placeholder
    blocks.push({ t: 'h', text: 'REPORT OF INDEPENDENT AUDITOR' });
    blocks.push({ t: 'box', text: "[To be issued and signed by the independent Certified Public Accountant after completing the audit in accordance with Philippine Standards on Auditing. This page is a placeholder — ITR Generator does not issue an auditor's opinion.]" + (E.auditor && E.auditor.name ? '\n\nPrior-year auditor: ' + E.auditor.name + (E.auditor.firm ? ', ' + E.auditor.firm : '') + '.' : '') });
    PB();

    // ---- balance sheet
    H(B.title, 'As of December 31, ' + y + ' and ' + py);
    const c = B.cur, p = B.prior, L = B.labels || {};
    const lab = (k, d) => L[k] || d;
    const bs = [];
    bs.push({ cells: ['A S S E T S', '', '', ''], style: 'sec' });
    bs.push({ cells: ['Current Assets', '', '', ''], style: 'sub' });
    bs.push(row(lab('cash', 'Cash and cash equivalents'), c.cash, p.cash, 'ind', '3'));
    [['receivables', 'Accounts receivable', '4'], ['inventory', 'Inventory', '5'], ['prepaid', 'Prepaid expenses', ''], ['other_ca', 'Other current assets', '']].forEach(([k, d, n]) => { if (c[k] || p[k]) bs.push(row(lab(k, d), c[k], p[k], 'ind', n)); });
    bs.push(row('Total Current Assets', B.totalCA, B.priorCA, 'tot'));
    bs.push({ cells: ['Non-current Assets', '', '', ''], style: 'sub' });
    if (B.hasGrossPPE) { bs.push(row('Property and equipment, at cost', c.ppe_cost, p.ppe_cost, 'ind', '6')); bs.push(row('Less: Accumulated depreciation', -c.accumulated_depreciation, -p.accumulated_depreciation, 'ind')); }
    bs.push(row('Property and equipment, net', c.ppe_net, p.ppe_net, 'ind', B.hasGrossPPE ? '' : '6'));
    if (c.other_nca || p.other_nca) bs.push(row(lab('other_nca', 'Other non-current assets'), c.other_nca, p.other_nca, 'ind'));
    bs.push(row('Total Non-current Assets', B.totalNCA, B.priorNCA, 'tot'));
    bs.push(row('TOTAL ASSETS', B.totalAssets, B.priorAssets, 'grand'));
    bs.push({ cells: [prop ? "LIABILITIES AND OWNER'S EQUITY" : "LIABILITIES AND STOCKHOLDERS' EQUITY", '', '', ''], style: 'sec' });
    bs.push({ cells: ['Liabilities', '', '', ''], style: 'sub' });
    [['payables', 'Accounts payable', '7'], ['accrued', 'Accrued expenses payable', '7'], ['vat_payable', 'VAT payable', ''], ['income_tax_payable', 'Income tax payable', '9'], ['other_cl', 'Other current liabilities', ''], ['loans', 'Loans payable', '8'], ['other_ncl', 'Other non-current liabilities', '']]
      .forEach(([k, d, n]) => { if (c[k] || p[k]) bs.push(row(lab(k, d), c[k], p[k], 'ind', n)); });
    bs.push(row('Total Liabilities', B.totalLiab, B.priorLiab, 'tot'));
    const eq = B.equity;
    bs.push({ cells: [prop ? "Owner's Equity" : "Stockholders' Equity", '10', '', ''], style: 'sub' });
    if (prop) {
      bs.push(row('Balance, beginning of year', eq.begin, M.prior.equity.beginning || null, 'ind'));
      if (eq.capitalAdd) bs.push(row('Additional capital', eq.capitalAdd, null, 'ind'));
      bs.push(row('Drawings', -eq.drawings, M.prior.equity.drawings ? -M.prior.equity.drawings : null, 'ind'));
      bs.push(row('Net income', eq.netIncome, M.prior.sci.net_income || null, 'ind'));
      bs.push(row('Balance, end of year', eq.end, B.priorEquity, 'tot'));
    } else {
      const lines = (eq.lines || []).length ? eq.lines : [{ label: 'Retained earnings', amount: eq.begin, kind: 'retained_earnings' }];
      let reDone = false;
      lines.forEach((l) => { let cur = Number(l.amount) || 0; if (!reDone && l.kind === 'retained_earnings') { cur = cur + eq.netIncome - eq.dividends; reDone = true; } if (l.kind === 'capital' && eq.capitalAdd) cur += eq.capitalAdd; bs.push(row(l.label, cur, l.amount, 'ind')); });
      if (!reDone) bs.push(row('Net income for the year', eq.netIncome, null, 'ind'));
      bs.push(row("Total Stockholders' Equity", eq.end, B.priorEquity, 'tot'));
    }
    bs.push(row(prop ? "TOTAL LIABILITIES AND OWNER'S EQUITY" : "TOTAL LIABILITIES AND STOCKHOLDERS' EQUITY", B.totalLE, B.priorLE, 'grand'));
    tbl(head('December 31, ' + y, 'December 31, ' + py), bs);
    see(); PB();

    // ---- income statement
    H(s.title, 'For the Years Ended December 31, ' + y + ' and ' + py);
    const is = [];
    is.push(row(prop ? 'REVENUES' : 'SALES', s.revenue, s.priorRevenue, 'sub', '11'));
    if (s.cos.length) {
      is.push({ cells: ['COST OF SALES', '12', '', ''], style: 'sec' });
      s.cos.forEach((l) => is.push(row(l.label, l.amount, l.prior, 'ind')));
      is.push(row('Total cost of sales', s.cosTotal, s.priorCos, 'tot'));
      is.push(row('GROSS PROFIT', s.grossProfit, s.priorGross, 'sub'));
    }
    is.push({ cells: ['OPERATING EXPENSES', '13', '', ''], style: 'sec' });
    s.opex.forEach((l) => is.push(row(l.label, l.amount, l.prior, 'ind')));
    is.push(row('Total operating expenses', s.opexTotal, s.priorOpex, 'tot'));
    if (s.otherIncome || s.priorOtherIncome) { is.push(row('OPERATING INCOME', s.operatingIncome, null, 'sub')); is.push(row('Other income', s.otherIncome, s.priorOtherIncome, 'ind')); }
    is.push(row('INCOME BEFORE INCOME TAX', s.ibt, s.priorIbt, 'sub'));
    is.push(row('Provision for income tax', s.provision, s.priorProvision, 'ind', '9'));
    is.push(row('NET INCOME', s.netIncome, s.priorNetIncome, 'grand'));
    tbl(head(String(y), String(py)), is);
    see(); PB();

    // ---- changes in equity
    H(prop ? "STATEMENTS OF CHANGES IN OWNER'S EQUITY" : 'STATEMENTS OF CHANGES IN EQUITY', 'For the Years Ended December 31, ' + y + ' and ' + py);
    const q = F.sce;
    const ce = [row('Balance, beginning of year', q.begin, q.priorBegin || null, '')];
    if (q.capitalAdd) ce.push(row('Additional capital', q.capitalAdd, null, 'ind'));
    ce.push(row('Net income for the year', q.netIncome, q.priorNetIncome || null, 'ind'));
    if (prop) ce.push(row('Drawings', -q.drawings, q.priorDrawings ? -q.priorDrawings : null, 'ind'));
    else if (q.dividends || q.priorDividends) ce.push(row('Dividends', -q.dividends, q.priorDividends ? -q.priorDividends : null, 'ind'));
    ce.push(row('Balance, end of year', q.end, q.priorEnd || null, 'grand'));
    tbl(head(String(y), String(py)), ce);
    see(); PB();

    // ---- cash flows
    H('STATEMENTS OF CASH FLOWS', 'For the Years Ended December 31, ' + y + ' and ' + py);
    const cf = [];
    cf.push({ cells: ['CASH FLOWS FROM OPERATING ACTIVITIES', '', '', ''], style: 'sec' });
    C.op.forEach((l) => cf.push(row(l.label, l.amount, null, 'ind')));
    cf.push(row('Net cash provided by (used in) operating activities', C.cfo, C.prior.cfo || null, 'tot'));
    cf.push({ cells: ['CASH FLOWS FROM INVESTING ACTIVITIES', '', '', ''], style: 'sec' });
    (C.inv.length ? C.inv : [{ label: 'None', amount: 0 }]).forEach((l) => cf.push(row(l.label, l.amount, null, 'ind')));
    cf.push(row('Net cash provided by (used in) investing activities', C.cfi, C.prior.cfi || null, 'tot'));
    cf.push({ cells: ['CASH FLOWS FROM FINANCING ACTIVITIES', '', '', ''], style: 'sec' });
    (C.fin.length ? C.fin : [{ label: 'None', amount: 0 }]).forEach((l) => cf.push(row(l.label, l.amount, null, 'ind')));
    cf.push(row('Net cash provided by (used in) financing activities', C.cff, C.prior.cff || null, 'tot'));
    cf.push(row('NET INCREASE (DECREASE) IN CASH', C.netChange, C.prior.net_change || null, 'sub'));
    cf.push(row('CASH AT BEGINNING OF YEAR', C.cashBegin, C.prior.cash_beginning || null, ''));
    cf.push(row('CASH AT END OF YEAR', C.cashEnd, C.prior.cash_end || null, 'grand', '3'));
    tbl(head(String(y), String(py)), cf);
    see(); PB();

    // ---- notes
    H('NOTES TO FINANCIAL STATEMENTS', 'For the Years Ended December 31, ' + y + ' and ' + py);
    let n = 0;
    const N = (title) => blocks.push({ t: 'nh', text: (++n) + '. ' + title });
    N(prop ? 'General Information' : 'Corporate Information');
    P(name + (prop ? ' is a sole proprietorship' : ' is a corporation') + ' registered in the Philippines' + (E.nature_of_business || M.client.lob ? ' engaged in ' + (E.nature_of_business || M.client.lob).replace(/\.$/, '') : '') + '. ' + (M.client.address || E.address ? 'Its registered address is ' + (M.client.address || E.address) + '. ' : '') + 'Its Taxpayer Identification Number is ' + M.client.tin + '.');
    N('Summary of Significant Accounting Policies');
    P('Basis of preparation — The financial statements have been prepared in accordance with ' + framework + ', using the historical cost basis. The financial statements are presented in Philippine Peso, the ' + (prop ? "Proprietorship's" : "Company's") + ' functional currency.');
    P('Cash and cash equivalents include cash on hand and in banks. Receivables are carried at the amounts due less any allowance for uncollectible accounts. Inventories are stated at the lower of cost and estimated selling price less costs to complete and sell.');
    P('Property and equipment are carried at cost less accumulated depreciation. Depreciation is computed on the straight-line method over the estimated useful lives of the assets.');
    P('Revenue is recognized when goods are delivered or services are rendered, measured at the fair value of the consideration received or receivable, net of value-added tax, discounts and returns. Expenses are recognized when incurred.');
    P('Income tax expense is the amount of income tax currently payable computed in accordance with the National Internal Revenue Code, as amended, including the CREATE Act' + (prop ? ' and the TRAIN Law graduated rates' : '') + '.');
    N('Cash and Cash Equivalents');
    P('This account consists of cash on hand and in banks amounting to ' + peso(c.cash) + ' and ' + peso(p.cash) + ' as of December 31, ' + y + ' and ' + py + ', respectively.');
    if (c.receivables || p.receivables) { N('Receivables'); P('Accounts receivable amount to ' + peso(c.receivables) + ' (' + py + ': ' + peso(p.receivables) + ') and are expected to be collected within one year.'); }
    if (c.inventory || p.inventory) { N('Inventories'); P('Inventories amount to ' + peso(c.inventory) + ' (' + py + ': ' + peso(p.inventory) + ') and are stated at cost.'); }
    N('Property and Equipment');
    const ppe = M.prior.ppe || [];
    if (B.hasGrossPPE) tbl(['', '', String(y), String(py)], [row('Cost', c.ppe_cost, p.ppe_cost, ''), row('Accumulated depreciation', -c.accumulated_depreciation, -p.accumulated_depreciation, ''), row('Net book value', c.ppe_net, p.ppe_net, 'tot')]);
    else tbl(['', '', String(y), String(py)], [row('Net book value, beginning', p.ppe_net, null, ''), row('Additions', r2(c.ppe_net - p.ppe_net + F.dep), null, ''), row('Depreciation', -F.dep, null, ''), row('Net book value, end', c.ppe_net, p.ppe_net, 'tot')]);
    if (ppe.length) P('Property and equipment by class as of December 31, ' + py + ': ' + ppe.map((x) => x.class + ' ' + peso(x.net)).join('; ') + '.', { size: 8.5 });
    P('Depreciation for the year amounted to ' + peso(F.dep) + '.');
    if (c.payables || c.accrued || p.payables || p.accrued) { N('Accounts Payable and Accrued Expenses'); P('This account consists of amounts due to suppliers and accrued expenses for goods and services received, amounting to ' + peso(r2(c.payables + c.accrued)) + ' (' + py + ': ' + peso(r2(p.payables + p.accrued)) + ').'); }
    if (c.loans || p.loans) {
      N('Loans Payable');
      const loans = (M.templateLoans || []);
      P('Loans payable amount to ' + peso(c.loans) + ' (' + py + ': ' + peso(p.loans) + ').' + (loans.length ? ' ' + loans.map((l) => l.bank + ': outstanding ' + peso(l.outstanding) + ', interest paid during the year ' + peso(l.interest_paid)).join('; ') + '.' : ''));
    }
    N('Income Taxes');
    const itx = [row('Income before income tax', s.ibt, null, ''), row('Taxable income per ' + T.formLabel, T.taxable, null, ''), row('Income tax due', T.taxDue, null, 'sub')];
    T.credits.filter((x) => x.amount).forEach((x) => itx.push(row('Less: ' + x.label, -x.amount, null, 'ind')));
    itx.push(row(T.overpayment ? 'Overpayment' : 'Income tax payable', T.overpayment ? -T.overpayment : T.payable, null, 'grand'));
    tbl(['', '', String(y), ''], itx);
    P('Income tax was computed under ' + T.schedule + (M.form === '1701' ? ', using the graduated income tax rates under the TRAIN Law (' + (y >= 2023 ? '2023 onwards' : '2018–2022') + ' table)' : '') + '.', { size: 8.5 });
    N(prop ? "Owner's Equity" : 'Equity');
    P(prop ? "Owner's equity increased by the net income of " + peso(eq.netIncome) + ' and decreased by drawings of ' + peso(eq.drawings) + ' during the year.' : 'Retained earnings increased by the net income of ' + peso(eq.netIncome) + (eq.dividends ? ' and decreased by dividends of ' + peso(eq.dividends) : '') + ' during the year.');
    N('Revenue');
    P('Revenue for the year amounted to ' + peso(s.revenue) + ' (' + py + ': ' + peso(s.priorRevenue) + '), as declared in the quarterly ' + (M.vat === 'VAT' ? 'value-added tax returns (BIR Form 2550Q)' : 'percentage tax returns (BIR Form 2551Q)') + ': ' + M.revenue.byQ.filter((x) => x.found).map((x) => 'Q' + x.quarter + ' ' + peso(x.total)).join('; ') + '.');
    N('Events after the Reporting Period');
    P('There were no significant events after the reporting period that require adjustment to or disclosure in the financial statements.');
    N('Supplementary Information Required under Revenue Regulations No. 15-2010');
    if (M.vat === 'VAT') {
      const rows = M.revenue.byQ.filter((x) => x.found).map((x) => ({ cells: ['Q' + x.quarter, '', f(x.total), f(x.outputTax)] }));
      tbl(['Value-added tax', '', 'Net sales', 'Output VAT'], rows.concat([{ cells: ['Total', '', f(M.revenueTotal), f(M.revenue.byQ.reduce((a, x) => a + (x.outputTax || 0), 0))], style: 'tot' }]));
    } else {
      const rows = M.revenue.byQ.filter((x) => x.found).map((x) => ({ cells: ['Q' + x.quarter, '', f(x.total), f(x.ptDue)] }));
      tbl(['Percentage tax', '', 'Gross receipts', 'Percentage tax'], rows.concat([{ cells: ['Total', '', f(M.revenueTotal), f(M.revenue.ptTotal)], style: 'tot' }]));
    }
    const tl = M.templateTaxesLicenses || [];
    if (tl.length) { P('Taxes and licenses paid during the year:'); tbl(['Description', '', 'Amount', ''], tl.map((x) => ({ cells: [x.description, '', f(x.amount), ''] })).concat([{ cells: ['Total', '', f(tl.reduce((a, x) => a + x.amount, 0)), ''], style: 'tot' }])); }
    const cwt = T.credits.filter((x) => /2307/.test(x.label)).reduce((a, x) => a + x.amount, 0);
    if (cwt) P('Creditable withholding taxes claimed (BIR Form 2307) amounted to ' + peso(cwt) + '.');
    return blocks;
  }

  /* ------------------------------------------------------------ docx renderer */
  async function docx(M) {
    const D = window.docx;
    const blocks = build(M);
    const FONT = 'Arial';
    const run = (text, o) => new D.TextRun(Object.assign({ text: String(text), font: FONT, size: (o && o.size ? o.size : 10) * 2 }, o || {}, { size: ((o && o.size) || 10) * 2 }));
    const children = [];
    const none = { style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    const noBorders = { top: none, bottom: none, left: none, right: none };
    blocks.forEach((b) => {
      if (b.t === 'pb') { children.push(new D.Paragraph({ children: [new D.PageBreak()] })); return; }
      if (b.t === 'cover') {
        children.push(new D.Paragraph({ spacing: { before: 3200 } }));
        b.lines.forEach((l, i) => children.push(new D.Paragraph({ alignment: D.AlignmentType.CENTER, spacing: { after: i === 1 ? 600 : 160 }, children: [run(l, { bold: i === 0 || l === 'FINANCIAL STATEMENTS', size: i === 0 ? 18 : l === 'FINANCIAL STATEMENTS' ? 16 : 11 })] })));
        return;
      }
      if (b.t === 'title') { b.lines.forEach((l, i) => children.push(new D.Paragraph({ alignment: D.AlignmentType.LEFT, spacing: { after: i === b.lines.length - 1 ? 240 : 0 }, children: [run(l, { bold: i === 0 || i === b.lines.length - 2, size: i === 0 ? 11 : 10 })] }))); return; }
      if (b.t === 'h') { children.push(new D.Paragraph({ alignment: D.AlignmentType.CENTER, spacing: { after: 300 }, children: [run(b.text, { bold: true, size: 11 })] })); return; }
      if (b.t === 'nh') { children.push(new D.Paragraph({ spacing: { before: 200, after: 80 }, children: [run(b.text, { bold: true })] })); return; }
      if (b.t === 'p') { children.push(new D.Paragraph({ alignment: D.AlignmentType.JUSTIFIED, indent: b.indent ? { left: 500 } : undefined, spacing: { after: 140 }, children: [run(b.text, { italics: !!b.italic, size: b.size || 10 })] })); return; }
      if (b.t === 'box') {
        children.push(new D.Table({ width: { size: 9300, type: D.WidthType.DXA }, columnWidths: [9300], rows: [new D.TableRow({ children: [new D.TableCell({ width: { size: 9300, type: D.WidthType.DXA }, margins: { top: 300, bottom: 300, left: 300, right: 300 }, children: b.text.split('\n').map((t) => new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [run(t, { italics: true })] })) })] })] }));
        return;
      }
      if (b.t === 'sig') {
        children.push(new D.Paragraph({ spacing: { before: 300 } }));
        b.people.forEach((p) => {
          children.push(new D.Paragraph({ spacing: { before: 400 }, children: [run('_______________________________')] }));
          children.push(new D.Paragraph({ children: [run(p.name, { bold: true })] }));
          children.push(new D.Paragraph({ spacing: { after: 120 }, children: [run(p.title, { size: 9 })] }));
        });
        children.push(new D.Paragraph({ spacing: { before: 240 }, children: [run(b.date, { size: 9 })] }));
        return;
      }
      if (b.t === 'table') {
        const W = b.widths, tot = W.reduce((a, x) => a + x, 0);
        const TEXT_W = 9300; // A4 text width in twips (page 11906 − margins)
        const cw = W.map((w) => Math.round(w / tot * TEXT_W));
        const mk = (cells, style, isHead) => new D.TableRow({
          children: cells.map((txt, i) => {
            const right = i >= 2;
            const bold = isHead || style === 'sub' || style === 'tot' || style === 'grand' || style === 'sec';
            const borders = Object.assign({}, noBorders);
            if (isHead) borders.bottom = { style: D.BorderStyle.SINGLE, size: 6, color: '000000' };
            if (right && (style === 'tot' || style === 'grand')) borders.top = { style: D.BorderStyle.SINGLE, size: 4, color: '000000' };
            if (right && style === 'grand') borders.bottom = { style: D.BorderStyle.DOUBLE, size: 6, color: '000000' };
            return new D.TableCell({
              width: { size: cw[i], type: D.WidthType.DXA }, borders,
              margins: { top: 30, bottom: 30, left: 60, right: 60 },
              children: [new D.Paragraph({ alignment: right ? D.AlignmentType.RIGHT : i === 1 ? D.AlignmentType.CENTER : D.AlignmentType.LEFT, indent: i === 0 && style === 'ind' ? { left: 280 } : undefined, children: [run(txt || '', { bold, size: 9.5 })] })]
            });
          })
        });
        children.push(new D.Table({ width: { size: TEXT_W, type: D.WidthType.DXA }, columnWidths: cw, layout: D.TableLayoutType.FIXED, rows: [mk(b.head, '', true)].concat(b.rows.map((r) => mk(r.cells, r.style, false))) }));
        children.push(new D.Paragraph({ spacing: { after: 160 } }));
      }
    });
    const doc = new D.Document({
      creator: 'ITR Generator – Annual V2.0', title: 'Financial Statements — ' + M.client.name + ' ' + M.year,
      styles: { default: { document: { run: { font: FONT, size: 20 } } } },
      sections: [{ properties: { page: { margin: { top: 1200, bottom: 1200, left: 1300, right: 1300 } } }, children }]
    });
    return D.Packer.toBlob(doc);
  }

  /* ------------------------------------------------------------ pdf renderer */
  function pdf(M) {
    const blocks = build(M);
    const content = [];
    blocks.forEach((b) => {
      if (b.t === 'pb') { content.push({ text: '', pageBreak: 'after' }); return; }
      if (b.t === 'cover') { content.push({ text: '', margin: [0, 180, 0, 0] }); b.lines.forEach((l, i) => content.push({ text: l, alignment: 'center', bold: i === 0 || l === 'FINANCIAL STATEMENTS', fontSize: i === 0 ? 17 : l === 'FINANCIAL STATEMENTS' ? 15 : 10.5, margin: [0, 0, 0, i === 1 ? 30 : 8] })); return; }
      if (b.t === 'title') { b.lines.forEach((l, i) => content.push({ text: l, bold: i === 0 || i === b.lines.length - 2, fontSize: i === 0 ? 11 : 10, margin: [0, 0, 0, i === b.lines.length - 1 ? 12 : 0] })); return; }
      if (b.t === 'h') { content.push({ text: b.text, alignment: 'center', bold: true, fontSize: 11, margin: [0, 0, 0, 16] }); return; }
      if (b.t === 'nh') { content.push({ text: b.text, bold: true, margin: [0, 10, 0, 4] }); return; }
      if (b.t === 'p') { content.push({ text: b.text, alignment: 'justify', italics: !!b.italic, fontSize: b.size || 9.5, margin: [b.indent ? 22 : 0, 0, 0, 7] }); return; }
      if (b.t === 'box') { content.push({ table: { widths: ['*'], body: [[{ text: b.text, alignment: 'center', italics: true, margin: [16, 24, 16, 24] }]] } }); return; }
      if (b.t === 'sig') {
        content.push({ text: '', margin: [0, 10, 0, 0] });
        b.people.forEach((p) => content.push({ stack: [{ text: '_______________________________', margin: [0, 22, 0, 0] }, { text: p.name, bold: true }, { text: p.title, fontSize: 8.5 }] }));
        content.push({ text: b.date, fontSize: 8.5, margin: [0, 14, 0, 0] });
        return;
      }
      if (b.t === 'table') {
        const tot = b.widths.reduce((a, x) => a + x, 0);
        const widths = b.widths.map((w, i) => (i === 0 ? '*' : Math.round(w / tot * 515)));
        const body = [b.head.map((h, i) => ({ text: h, bold: true, fontSize: 8.5, alignment: i >= 2 ? 'right' : i === 1 ? 'center' : 'left', border: [false, false, false, true] }))];
        b.rows.forEach((r) => body.push(r.cells.map((t, i) => {
          const bold = ['sub', 'tot', 'grand', 'sec'].indexOf(r.style) >= 0;
          const cell = { text: t || '', bold, fontSize: 9, alignment: i >= 2 ? 'right' : i === 1 ? 'center' : 'left', margin: [i === 0 && r.style === 'ind' ? 12 : 0, 1, 0, 1], border: [false, false, false, false] };
          if (i >= 2 && (r.style === 'tot' || r.style === 'grand')) cell.border = [false, true, false, r.style === 'grand'];
          return cell;
        })));
        content.push({ table: { headerRows: 1, widths, body }, layout: { hLineWidth: (i, node) => 0.6, vLineWidth: () => 0, hLineColor: () => '#111', paddingLeft: () => 3, paddingRight: () => 3 }, margin: [0, 0, 0, 10] });
      }
    });
    const dd = {
      pageSize: 'A4', pageMargins: [56, 56, 56, 56],
      info: { title: 'Financial Statements — ' + M.client.name + ' ' + M.year, creator: 'ITR Generator – Annual V2.0' },
      defaultStyle: { font: 'Roboto', fontSize: 9.5, lineHeight: 1.15 },
      footer: (cur, total) => (cur > 1 ? { text: M.client.name + ' · Financial Statements ' + M.year + ' · page ' + cur + ' of ' + total, alignment: 'center', fontSize: 7.5, color: '#888', margin: [0, 20, 0, 0] } : null),
      content
    };
    return new Promise((resolve, reject) => {
      try { window.pdfMake.createPdf(dd).getBlob((b) => resolve(b)); } catch (e) { reject(e); }
    });
  }

  window.AFSDoc = { build, docx, pdf };
})();
