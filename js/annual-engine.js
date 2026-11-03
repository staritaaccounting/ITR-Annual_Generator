/*!
 * ITR Generator – Annual V2.0 · computation engine (pure functions, no DOM)
 *
 * Builds, from figures extracted by Muse Spark:
 *   - Revenue (sum of the four quarterly 2550Q / 2551Q returns)
 *   - Cost of sales & operating expenses from one of three bases:
 *       documents  : uploaded expense documents, consolidated per account
 *       template   : operating expenses listed in the office ITR Template
 *       prior      : prior-year ITR/AFS line items re-scaled to this year's rates
 *   - Annual income tax: BIR Form 1702-RT (RCIT vs MCIT) or BIR Form 1701 (graduated / OSD / 8%, mixed income)
 *   - Financial statements: Income Statement, Balance Sheet (rolled forward, cash as balancing figure),
 *     Cash Flow (indirect, reconciles to cash), Changes in Equity
 *   - ITR Template check: template figures vs the returns, with corrections
 */
(function (root) {
  'use strict';

  const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
  const pos = (n) => Math.max(0, Number(n) || 0);
  const num = (n) => (isFinite(Number(n)) ? Number(n) : 0);
  const sum = (arr, f) => r2(arr.reduce((s, x) => s + num(f ? f(x) : x), 0));
  const norm = (s) => String(s || '').toLowerCase().replace(/&/g, 'and').replace(/\(.*?\)/g, '').replace(/[^a-z0-9]+/g, ' ').trim();

  /* ------------------------------------------------------------ tax tables */
  const BRACKETS_2018 = [[250000, 0, 0], [400000, 0, 0.20], [800000, 30000, 0.25], [2000000, 130000, 0.30], [8000000, 490000, 0.32], [Infinity, 2410000, 0.35]];
  const BRACKETS_2023 = [[250000, 0, 0], [400000, 0, 0.15], [800000, 22500, 0.20], [2000000, 102500, 0.25], [8000000, 402500, 0.30], [Infinity, 2202500, 0.35]];
  function graduatedTax(income, year) {
    const t = Number(year) >= 2023 ? BRACKETS_2023 : BRACKETS_2018;
    const x = num(income);
    if (x <= 0) return 0;
    let lower = 0;
    for (const [upper, base, rate] of t) { if (x <= upper) return r2(base + (x - lower) * rate); lower = upper; }
    return 0;
  }
  function bracketText(income, year) {
    const t = Number(year) >= 2023 ? BRACKETS_2023 : BRACKETS_2018;
    const x = num(income);
    if (x <= 250000) return 'Not over ₱250,000 — exempt';
    let lower = 0;
    for (const [upper, base, rate] of t) {
      if (x <= upper) return '₱' + fmt(base) + ' + ' + (rate * 100) + '% of excess over ₱' + fmt(lower) + ' (' + (Number(year) >= 2023 ? '2023 onwards' : '2018–2022') + ' table)';
      lower = upper;
    }
    return '';
  }
  const fmt = (n) => r2(n).toLocaleString('en-PH', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  const TAX_TYPES = {
    rcit25: { label: 'RCIT 25%', rate: 0.25, mcit: true },
    rcit20: { label: 'RCIT 20%', rate: 0.20, mcit: true },
    iaet10: { label: 'Special rate 10%', rate: 0.10, mcit: false },
    nonres25: { label: 'RCIT — Non-Resident 25%', rate: 0.25, mcit: false }
  };

  /* ------------------------------------------------------------ revenue */
  function buildRevenue(sales, vat) {
    const byQ = [1, 2, 3, 4].map((q) => {
      const docs = sales.filter((s) => Number(s.quarter) === q);
      if (vat === 'VAT') {
        const v = { quarter: q, vatable: sum(docs, (d) => d.vatable_sales), gov: sum(docs, (d) => d.sales_to_government), zero: sum(docs, (d) => d.zero_rated_sales), exempt: sum(docs, (d) => d.exempt_sales), outputTax: sum(docs, (d) => d.output_tax), inputTax: sum(docs, (d) => d.input_tax), vatPayable: sum(docs, (d) => d.vat_payable) };
        v.total = sum(docs, (d) => d.total_sales || (num(d.vatable_sales) + num(d.sales_to_government) + num(d.zero_rated_sales) + num(d.exempt_sales)));
        v.found = docs.length > 0;
        return v;
      }
      return { quarter: q, total: sum(docs, (d) => d.gross_receipts || d.revenue), ptDue: sum(docs, (d) => d.percentage_tax_due), found: docs.length > 0 };
    });
    const total = sum(byQ, (q) => q.total);
    const lines = vat === 'VAT'
      ? [['VATable sales', sum(byQ, (q) => q.vatable)], ['Sales to government', sum(byQ, (q) => q.gov)], ['Zero-rated sales', sum(byQ, (q) => q.zero)], ['Exempt sales', sum(byQ, (q) => q.exempt)]].filter((l) => l[1])
      : byQ.filter((q) => q.found).map((q) => ['Q' + q.quarter + ' gross receipts (2551Q)', q.total]);
    return { byQ, total, lines, ptTotal: sum(byQ, (q) => q.ptDue || 0), missing: byQ.filter((q) => !q.found).map((q) => q.quarter) };
  }

  /* ------------------------------------------------------------ expenses */
  const KIND = {
    dep: /depreci|amortiz/i,
    interest: /interest/i,
    tl: /tax(es)?\s*(and|&)?\s*licen|licen[sc]es|permits?/i,
    rep: /represent|entertain/i,
    begInv: /beginning\s*inv|inventory,?\s*beg|inv(entory)?\s*beg/i,
    endInv: /inventory,?\s*end|ending\s*inv|inv(entory)?\s*end/i
  };
  const kindOf = (label, given) => given && given !== 'regular' ? given
    : KIND.begInv.test(label) ? 'beginning_inventory' : KIND.endInv.test(label) ? 'ending_inventory'
      : KIND.dep.test(label) ? 'depreciation' : KIND.interest.test(label) ? 'interest' : KIND.tl.test(label) ? 'taxes_licenses' : 'regular';

  /** Prior-year basis: re-scale the prior year's lines to this year's cost-of-sales % and operating-expense % */
  function priorBasis(ctx) {
    const { prior, revenue, template, ptTotal, vat, adj } = ctx;
    const priorRev = num(prior.sci.revenue) || num(prior.itr.gross_sales);
    const cosLines = (prior.sci.cost_of_sales_lines || []).map((l) => ({ label: l.label, prior: num(l.amount), kind: kindOf(l.label, l.kind) }));
    const opLines = (prior.sci.operating_expenses || []).map((l) => ({ label: l.label, prior: num(l.amount), kind: kindOf(l.label, l.kind) }));
    if (!opLines.length && (prior.itr.deductions || []).length) prior.itr.deductions.forEach((d) => opLines.push({ label: d.account, prior: num(d.amount), kind: kindOf(d.account) }));
    if (!cosLines.length && num(prior.itr.cost_of_sales)) cosLines.push({ label: 'Cost of sales', prior: num(prior.itr.cost_of_sales), kind: 'regular' });
    if (!opLines.length && num(prior.itr.itemized_deductions_total)) opLines.push({ label: 'Operating expenses', prior: num(prior.itr.itemized_deductions_total), kind: 'regular' });

    const priorCos = sum(cosLines, (l) => l.prior);
    const priorOpex = sum(opLines, (l) => l.prior);
    const priorCosPct = priorRev ? r2(priorCos / priorRev * 100) : 0;
    const priorOpexPct = priorRev ? r2(priorOpex / priorRev * 100) : 0;
    const cosPct = adj.cosPct != null ? num(adj.cosPct) : priorCosPct;
    const opexPct = adj.opexPct != null ? num(adj.opexPct) : priorOpexPct;

    const tplTL = template && template.taxes_licenses_total ? num(template.taxes_licenses_total) : 0;
    const tplInterest = template && template.loans ? sum(template.loans, (l) => l.interest_paid) : 0;
    const invEndPrior = num(prior.balances.inventory);

    // fixed (non-scaled) amounts, with the reason shown in the statements
    const fixedOf = (l) => {
      if (l.kind === 'beginning_inventory') return { amt: invEndPrior || l.prior, basis: 'Prior year ending inventory' };
      if (l.kind === 'ending_inventory') { const e = adj.inventoryEnd != null ? num(adj.inventoryEnd) : (invEndPrior || Math.abs(l.prior)); return { amt: -Math.abs(e), basis: 'Ending inventory (count / carried over)' }; }
      if (l.kind === 'depreciation') return { amt: adj.depreciation != null ? num(adj.depreciation) : l.prior, basis: 'Depreciation — same as prior year (PPE)' };
      if (l.kind === 'interest' && tplInterest) return { amt: tplInterest, basis: 'Actual interest per ITR Template (bank loans)' };
      if (l.kind === 'taxes_licenses' && (tplTL || (vat !== 'VAT' && ptTotal))) {
        const amt = r2(tplTL + (vat !== 'VAT' ? ptTotal : 0));
        return { amt, basis: (tplTL ? 'Permits/licenses ₱' + fmt(tplTL) + ' per ITR Template' : '') + (vat !== 'VAT' && ptTotal ? (tplTL ? ' + ' : '') + 'percentage tax ₱' + fmt(ptTotal) + ' per 2551Q' : '') };
      }
      return null;
    };
    const scale = (lines, targetPct) => {
      const target = r2(revenue * targetPct / 100);
      const out = lines.map((l) => { const f = fixedOf(l); return Object.assign({}, l, f ? { amount: r2(f.amt), basis: f.basis, fixed: true } : { fixed: false }); });
      const fixed = sum(out.filter((l) => l.fixed), (l) => l.amount);
      const varPrior = sum(out.filter((l) => !l.fixed), (l) => l.prior);
      const room = r2(target - fixed);
      const factor = varPrior > 0 ? Math.max(0, room) / varPrior : 0;
      out.filter((l) => !l.fixed).forEach((l) => {
        l.amount = r2(l.prior * factor);
        l.basis = 'Prior year ₱' + fmt(l.prior) + ' × ' + (factor ? r2((factor - 1) * 100) >= 0 ? '+' + r2((factor - 1) * 100) : r2((factor - 1) * 100) : '0') + '%';
      });
      const vars = out.filter((l) => !l.fixed);
      const diff = r2(Math.max(target, fixed) - sum(out, (l) => l.amount));
      if (diff && vars.length) { const big = vars.reduce((a, b) => (b.amount > a.amount ? b : a)); big.amount = r2(big.amount + diff); }
      return { lines: out, target, fixed, factor: r2(factor), short: room < 0 };
    };
    const cos = scale(cosLines, cosPct);
    const op = scale(opLines, opexPct);
    return {
      basis: 'prior', cos: cos.lines, opex: op.lines, cosPct, opexPct, priorCosPct, priorOpexPct, priorRevenue: priorRev,
      factors: { cos: cos.factor, opex: op.factor },
      warnings: [].concat(cos.short ? ['Fixed cost-of-sales items exceed ' + cosPct + '% of revenue.'] : [], op.short ? ['Fixed operating expenses (depreciation, interest, taxes & licenses) exceed ' + opexPct + '% of revenue; other expenses were set to zero.'] : [])
    };
  }

  function documentsBasis(ctx) {
    const map = new Map();
    (ctx.expenseItems || []).forEach((it) => {
      if (it.category === 'non_deductible') return;
      const k = it.category + '|' + norm(it.account);
      if (!map.has(k)) map.set(k, { label: it.account, category: it.category, amount: 0, items: [] });
      const a = map.get(k); a.amount = r2(a.amount + num(it.amount)); a.items.push(it);
    });
    const all = Array.from(map.values()).map((a) => Object.assign(a, { basis: a.items.length + ' item' + (a.items.length === 1 ? '' : 's') + ' from uploaded documents', kind: kindOf(a.label) }));
    const opex = all.filter((a) => a.category === 'operating_expense');
    const cos = all.filter((a) => a.category === 'cost_of_sales');
    // depreciation is non-cash: never in receipts — add it from the prior-year PPE if missing
    const dep = ctx.adj.depreciation != null ? num(ctx.adj.depreciation) : num(ctx.prior && ctx.prior.depreciation);
    if (dep && !opex.concat(cos).some((l) => l.kind === 'depreciation')) opex.push({ label: 'Depreciation', amount: r2(dep), basis: 'Non-cash — per PPE / prior year', kind: 'depreciation', fixed: true });
    const nonded = (ctx.expenseItems || []).filter((i) => i.category === 'non_deductible');
    return { basis: 'documents', cos, opex, nondeductible: sum(nonded, (i) => i.amount) };
  }

  function templateBasis(ctx) {
    const t = ctx.template;
    const lines = (t.operating_expenses || []).map((e) => ({ label: e.account, amount: r2(e.amount), basis: 'Per ITR Template' + (e.cell ? ' (' + e.cell + ')' : ''), kind: kindOf(e.account) }));
    const cos = lines.filter((l) => /^cost of (sales|goods|services)|purchases|direct materials/i.test(l.label) && !/direct expenses/i.test(l.label));
    const opex = lines.filter((l) => cos.indexOf(l) < 0);
    return { basis: 'template', cos, opex };
  }

  function buildExpenses(ctx) {
    const avail = {
      documents: (ctx.expenseItems || []).length > 0,
      template: !!(ctx.template && (ctx.template.operating_expenses || []).length),
      prior: !!(ctx.prior && ((ctx.prior.sci.operating_expenses || []).length || (ctx.prior.itr.deductions || []).length || num(ctx.prior.itr.itemized_deductions_total) || num(ctx.prior.sci.cost_of_sales_total)))
    };
    let basis = ctx.adj.expenseBasis;
    if (!basis || !avail[basis]) basis = avail.documents ? 'documents' : avail.template ? 'template' : avail.prior ? 'prior' : 'none';
    let e = basis === 'documents' ? documentsBasis(ctx) : basis === 'template' ? templateBasis(ctx) : basis === 'prior' ? priorBasis(ctx) : { basis: 'none', cos: [], opex: [] };
    // per-line manual overrides
    const ov = ctx.adj.lineOverrides || {};
    ['cos', 'opex'].forEach((cat) => e[cat].forEach((l) => {
      const k = cat + '|' + norm(l.label);
      l.key = k;
      if (ov[k] != null) { l.extracted = l.amount; l.amount = r2(ov[k]); l.edited = true; l.basis = 'Edited (was ₱' + fmt(l.extracted) + ')'; }
    }));
    e.cosTotal = sum(e.cos, (l) => l.amount);
    e.opexTotal = sum(e.opex, (l) => l.amount);
    e.available = avail;
    e.warnings = e.warnings || [];
    if (basis === 'none') e.warnings.push('No expenses: no expense documents, no expense list in the ITR Template and no prior-year ITR/AFS to base them on.');
    return e;
  }

  /* ------------------------------------------------------------ annual tax */
  function credits(ctx, form) {
    const a = ctx.adj, q = ctx.quarterlyItrs || [], t = ctx.template;
    const paidQ = (n) => { const d = q.find((x) => Number(x.quarter) === n); return d ? pos(d.tax_payable) : null; };
    const qPaid = [1, 2, 3].map((n) => { const v = paidQ(n); if (v != null) return { q: n, amount: v, src: 'BIR Form ' + (form === '1701' ? '1701Q' : '1702Q') }; const tq = t && (t.quarterly_itr_paid || []).find((x) => Number(x.quarter) === n); return tq ? { q: n, amount: num(tq.amount), src: 'ITR Template' } : { q: n, amount: 0, src: '—' }; });
    const quarterlyPaid = a.quarterlyPaid != null ? num(a.quarterlyPaid) : sum(qPaid, (x) => x.amount);
    // CWT: 2307 certificates for the year, split previous quarters vs 4th quarter
    const certs = ctx.certs2307 || [];
    let cwtPrev = sum(certs.filter((c) => Number(c.quarter) && Number(c.quarter) < 4), (c) => c.total_tax_withheld);
    let cwtQ4 = sum(certs.filter((c) => !Number(c.quarter) || Number(c.quarter) === 4), (c) => c.total_tax_withheld);
    let cwtSrc = certs.length ? certs.length + ' BIR 2307 certificate' + (certs.length === 1 ? '' : 's') : '';
    if (!certs.length && q.length) { cwtPrev = sum(q, (x) => x.cwt_this_quarter); cwtSrc = cwtPrev ? 'CWT claimed on the quarterly ITRs' : ''; }
    if (a.cwtPrev != null) cwtPrev = num(a.cwtPrev);
    if (a.cwtQ4 != null) cwtQ4 = num(a.cwtQ4);
    const priorExcess = a.priorYearExcess != null ? num(a.priorYearExcess) : num(ctx.prior && ctx.prior.itr.excess_credits_carried_over);
    const comp = ctx.certs2316 || [];
    const compWithheld = form === '1701' ? (a.compWithheld != null ? num(a.compWithheld) : sum(comp, (c) => c.tax_withheld)) : 0;
    const items = [
      { key: 'priorYearExcess', label: "Prior Year's Excess Credits", amount: r2(priorExcess), note: ctx.prior && ctx.prior.itr.excess_credits_carried_over ? 'Per prior-year annual ITR' : '' },
      { key: 'quarterlyPaid', label: 'Tax Payments for the First Three Quarters', amount: r2(quarterlyPaid), note: qPaid.map((x) => 'Q' + x.q + ' ₱' + fmt(x.amount)).join(' · ') },
      { key: 'cwtPrev', label: 'Creditable Tax Withheld — First Three Quarters (BIR 2307)', amount: r2(cwtPrev), note: cwtSrc },
      { key: 'cwtQ4', label: 'Creditable Tax Withheld — 4th Quarter (BIR 2307)', amount: r2(cwtQ4), note: '' }
    ];
    if (form === '1701') items.push({ key: 'compWithheld', label: 'Tax Withheld on Compensation (BIR 2316)', amount: r2(compWithheld), note: comp.length ? comp.length + ' BIR 2316' : '' });
    if (form === '1702RT') items.push({ key: 'excessMcit', label: 'Excess MCIT Applied', amount: r2(a.excessMcit), note: '' });
    items.push({ key: 'otherCredits', label: 'Other Tax Credits / Payments', amount: r2(a.otherCredits), note: '' });
    return { items, total: sum(items, (i) => i.amount), qPaid };
  }

  function tax1702(ctx, exp) {
    const a = ctx.adj, o = ctx.opts, year = ctx.year;
    const tt = TAX_TYPES[o.taxType] || TAX_TYPES.rcit25;
    const sales = ctx.revenue;
    const cos = exp.cosTotal;
    const grossIncome = r2(sales - cos);
    const other = r2(a.otherIncome);
    const totalGross = r2(grossIncome + other);
    const ent = entertainment(exp, sales, ctx.client.sellerOf);
    const itemized = o.method !== 'osd';
    const ded = itemized ? r2(exp.opexTotal - ent.disallowed) : r2(pos(totalGross) * 0.40);
    const nolco = itemized ? r2(a.nolco) : 0;
    const taxable = r2(totalGross - ded - nolco);
    const regular = r2(pos(taxable) * tt.rate);
    const mcitRate = num(o.mcitRate) || 0.02;
    const mcitOn = tt.mcit && !o.mcitExempt;
    const mcit = mcitOn ? r2(pos(totalGross) * mcitRate) : 0;
    const taxDue = r2(Math.max(regular, mcit));
    const c = credits(ctx, '1702RT');
    const net = r2(taxDue - c.total);
    const lines = [
      { label: 'Sales/Receipts/Revenues/Fees', amount: sales },
      { label: 'Less: Cost of Sales/Services', amount: cos },
      { label: 'Gross Income from Operation', amount: grossIncome, type: 'sub' },
      { label: 'Add: Other Taxable Income Not Subjected to Final Tax', amount: other, edit: 'otherIncome' },
      { label: 'Total Gross Income', amount: totalGross, type: 'sub' },
      itemized ? { label: 'Less: Total Allowable Itemized Deductions', amount: ded, note: ent.disallowed ? 'Excludes ₱' + fmt(ent.disallowed) + ' entertainment over limit' : '' } : { label: 'Less: Optional Standard Deduction (40% of Gross Income)', amount: ded },
      itemized ? { label: 'Less: NOLCO (only for itemized deductions)', amount: nolco, edit: 'nolco' } : null,
      { label: 'Net Taxable Income', amount: taxable, type: 'sub' },
      { label: 'Applicable Income Tax Rate', amount: tt.rate * 100, type: 'rate' },
      { label: 'Income Tax Due other than MCIT', amount: regular },
      tt.mcit ? { label: 'Minimum Corporate Income Tax (' + (mcitRate * 100) + '% of Gross Income)' + (mcitOn ? '' : ' — not yet applicable'), amount: mcit } : null,
      { label: 'Total Income Tax Due (' + (mcit > regular ? 'MCIT' : tt.label) + ')', amount: taxDue, type: 'total' }
    ].filter(Boolean);
    const warn = [];
    if (mcit > regular) warn.push('MCIT (₱' + fmt(mcit) + ') is higher than the regular income tax, so MCIT is the tax due.');
    if (o.taxType === 'rcit20') warn.push('The 20% rate applies only if net taxable income is ₱5,000,000 or less and total assets (excluding land) are ₱100,000,000 or less.');
    if (o.taxType === 'rcit20' && taxable > 5000000) warn.push('Net taxable income exceeds ₱5,000,000 — the 20% rate does not apply; use 25%.');
    return { form: '1702RT', formLabel: 'BIR Form 1702-RT', schedule: 'Part IV — Computation of Tax · ' + (itemized ? 'Itemized Deductions' : 'OSD') + ' · ' + tt.label, lines, credits: c.items, qPaid: c.qPaid, totalCredits: c.total, taxDue, provision: taxDue, taxable, payable: r2(Math.max(0, net)), overpayment: r2(Math.max(0, -net)), regular, mcit, ent, warnings: warn };
  }

  function tax1701(ctx, exp) {
    const a = ctx.adj, o = ctx.opts, year = ctx.year;
    const sales = ctx.revenue;
    const other = r2(a.otherIncome);
    const comp = ctx.certs2316 || [];
    const compIncome = a.compIncome != null ? num(a.compIncome) : sum(comp, (c) => c.taxable_compensation);
    const mixed = compIncome > 0;
    const ent = entertainment(exp, sales, ctx.client.sellerOf);
    const m = o.method || 'itemized';
    const lines = []; const warn = [];
    let taxDue, provision, taxable, schedule;
    if (m === 'eight') {
      schedule = '8% Income Tax Rate on Gross Sales/Receipts' + (mixed ? ' (mixed income)' : '');
      const base = r2(sales + other);
      const reduction = mixed ? 0 : 250000;
      const taxable8 = r2(pos(base - reduction));
      const bizTax = r2(taxable8 * 0.08);
      const compTax = graduatedTax(compIncome, year);
      taxDue = r2(bizTax + compTax); provision = bizTax; taxable = taxable8;
      if (ctx.vat === 'VAT') warn.push('The 8% option is only for non-VAT taxpayers.');
      if (base > 3000000) warn.push('Gross sales/receipts exceed ₱3,000,000 — the 8% option does not apply.');
      if (mixed) lines.push({ label: 'Taxable Compensation Income (BIR 2316)', amount: compIncome }, { label: 'Tax Due on Compensation (graduated)', amount: compTax, type: 'sub' });
      lines.push(
        { label: 'Gross Sales/Receipts', amount: sales },
        { label: 'Add: Non-Operating Income', amount: other, edit: 'otherIncome' },
        { label: 'Total', amount: base, type: 'sub' },
        { label: mixed ? 'Less: ₱250,000 reduction (not applicable — mixed income)' : 'Less: Allowable reduction', amount: reduction },
        { label: 'Taxable Income — 8%', amount: taxable8, type: 'sub' },
        { label: 'Tax Due — 8%', amount: bizTax },
        { label: 'Total Income Tax Due', amount: taxDue, type: 'total' });
    } else {
      const osd = m === 'osd' || m === 'osd_gi';
      const cos = m === 'osd' ? 0 : exp.cosTotal;
      const gross = r2(sales - cos);
      const ded = m === 'itemized' ? r2(exp.opexTotal - ent.disallowed) : m === 'osd' ? r2(sales * 0.40) : r2(pos(gross) * 0.40);
      const nolco = m === 'itemized' ? r2(a.nolco) : 0;
      const bizTaxable = r2(gross - ded - nolco + other);
      taxable = r2(pos(bizTaxable) + compIncome);
      taxDue = graduatedTax(taxable, year);
      provision = r2(taxDue - graduatedTax(compIncome, year));
      schedule = 'Graduated Income Tax Rates · ' + ({ itemized: 'Itemized Deductions', osd: 'OSD 40% of gross sales', osd_gi: 'OSD 40% of gross income' }[m]) + (mixed ? ' · mixed income' : '');
      if (mixed) lines.push({ label: 'Taxable Compensation Income (BIR 2316)', amount: compIncome });
      lines.push(
        { label: 'Sales/Revenues/Receipts/Fees', amount: sales },
        { label: 'Less: Cost of Sales/Services', amount: cos, note: m === 'osd' ? 'Not deducted under OSD on gross sales' : '' },
        { label: 'Gross Income from Business/Profession', amount: gross, type: 'sub' },
        m === 'itemized' ? { label: 'Less: Total Allowable Itemized Deductions', amount: ded, note: ent.disallowed ? 'Excludes ₱' + fmt(ent.disallowed) + ' entertainment over limit' : '' } : { label: 'Less: Optional Standard Deduction (40%)', amount: ded },
        m === 'itemized' ? { label: 'Less: NOLCO', amount: nolco, edit: 'nolco' } : null,
        { label: 'Add: Other Non-Operating Income', amount: other, edit: 'otherIncome' },
        { label: 'Taxable Income from Business/Profession', amount: bizTaxable, type: 'sub' },
        mixed ? { label: 'Total Taxable Income (compensation + business)', amount: taxable, type: 'sub' } : null,
        { label: 'Tax Due — ' + bracketText(taxable, year), amount: taxDue, type: 'total' });
    }
    const c = credits(ctx, '1701');
    const net = r2(taxDue - c.total);
    return { form: '1701', formLabel: 'BIR Form 1701', schedule, lines: lines.filter(Boolean), credits: c.items, qPaid: c.qPaid, totalCredits: c.total, taxDue, provision: r2(provision), taxable, payable: r2(Math.max(0, net)), overpayment: r2(Math.max(0, -net)), compIncome, ent, warnings: warn };
  }

  function entertainment(exp, revenue, sellerOf) {
    const rate = sellerOf === 'Services' ? 0.01 : 0.005;
    const cap = r2(pos(revenue) * rate);
    const re = sum(exp.opex.filter((l) => KIND.rep.test(l.label)), (l) => l.amount);
    return { cap, rate, reTotal: re, disallowed: r2(Math.max(0, re - cap)) };
  }

  /* ------------------------------------------------------------ financial statements */
  const ASSET_ORDER = ['cash', 'receivables', 'inventory', 'prepaid', 'other_ca', 'ppe_cost', 'accumulated_depreciation', 'ppe_net', 'other_nca'];
  const LIAB_ORDER = ['payables', 'accrued', 'vat_payable', 'income_tax_payable', 'other_cl', 'loans', 'other_ncl'];

  function buildFS(ctx, exp, tax) {
    const P = ctx.prior, a = ctx.adj, year = ctx.year;
    const isProp = ctx.client.type !== 'Corporation';
    const priorRev = num(P.sci.revenue) || num(P.itr.gross_sales);
    const growth = priorRev > 0 ? ctx.revenue / priorRev : 1;

    // ---------- income statement
    const otherIncome = r2(a.otherIncome);
    const gross = r2(ctx.revenue - exp.cosTotal);
    const opIncome = r2(gross - exp.opexTotal);
    const ibt = r2(opIncome + otherIncome);
    const provision = tax.provision;
    const netIncome = r2(ibt - provision);
    const priorLine = (label, list) => { const f = (list || []).find((l) => norm(l.label) === norm(label)); return f ? num(f.amount) : null; };
    const sci = {
      title: isProp ? 'STATEMENTS OF REVENUES AND EXPENSES' : 'STATEMENTS OF COMPREHENSIVE INCOME',
      revenue: ctx.revenue, priorRevenue: priorRev,
      cos: exp.cos.map((l) => ({ label: l.label, amount: l.amount, prior: priorLine(l.label, P.sci.cost_of_sales_lines), basis: l.basis })),
      cosTotal: exp.cosTotal, priorCos: num(P.sci.cost_of_sales_total) || sum(P.sci.cost_of_sales_lines || [], (l) => l.amount),
      grossProfit: gross, priorGross: num(P.sci.gross_profit) || r2(priorRev - (num(P.sci.cost_of_sales_total) || 0)),
      opex: exp.opex.map((l) => ({ label: l.label, amount: l.amount, prior: priorLine(l.label, P.sci.operating_expenses), basis: l.basis })),
      opexTotal: exp.opexTotal, priorOpex: num(P.sci.opex_total) || sum(P.sci.operating_expenses || [], (l) => l.amount),
      operatingIncome: opIncome, otherIncome, priorOtherIncome: num(P.sci.other_income),
      ibt, priorIbt: num(P.sci.income_before_tax), provision, priorProvision: num(P.sci.income_tax), netIncome, priorNetIncome: num(P.sci.net_income)
    };
    // prior lines not used this year (shown with zero)
    (P.sci.operating_expenses || []).forEach((l) => { if (!sci.opex.some((x) => norm(x.label) === norm(l.label))) sci.opex.push({ label: l.label, amount: 0, prior: num(l.amount), basis: 'None this year' }); });
    (P.sci.cost_of_sales_lines || []).forEach((l) => { if (!sci.cos.some((x) => norm(x.label) === norm(l.label))) sci.cos.push({ label: l.label, amount: 0, prior: num(l.amount), basis: 'None this year' }); });

    // ---------- balance sheet (roll forward; cash is the balancing figure)
    const B = P.balances; // prior year-end balances by kind
    const dep = sum(exp.cos.concat(exp.opex).filter((l) => l.kind === 'depreciation'), (l) => l.amount);
    const additions = r2(a.ppeAdditions);
    const loansTpl = ctx.template && (ctx.template.loans || []).length ? sum(ctx.template.loans, (l) => l.outstanding) : null;
    const cur = {
      receivables: a.receivables != null ? num(a.receivables) : num(B.receivables),
      inventory: a.inventoryEnd != null ? num(a.inventoryEnd) : num(B.inventory),
      prepaid: a.prepaid != null ? num(a.prepaid) : num(B.prepaid),
      other_ca: a.otherCA != null ? num(a.otherCA) : num(B.other_ca),
      other_nca: a.otherNCA != null ? num(a.otherNCA) : num(B.other_nca),
      payables: a.payables != null ? num(a.payables) : r2(num(B.payables) * growth),
      accrued: a.accrued != null ? num(a.accrued) : r2(num(B.accrued) * growth),
      vat_payable: a.vatPayable != null ? num(a.vatPayable) : (ctx.vat === 'VAT' && ctx.q4VatPayable != null ? pos(ctx.q4VatPayable) : num(B.vat_payable)),
      income_tax_payable: tax.payable,
      other_cl: num(B.other_cl),
      loans: a.loans != null ? num(a.loans) : (loansTpl != null ? loansTpl : num(B.loans)),
      other_ncl: num(B.other_ncl)
    };
    const hasGrossPPE = num(B.ppe_cost) > 0;
    if (hasGrossPPE) { cur.ppe_cost = r2(num(B.ppe_cost) + additions); cur.accumulated_depreciation = r2(num(B.accumulated_depreciation) + dep); cur.ppe_net = r2(cur.ppe_cost - cur.accumulated_depreciation); }
    else cur.ppe_net = r2(num(B.ppe_net) + additions - dep);
    // equity
    const drawings = isProp ? (a.drawings != null ? num(a.drawings) : num(P.equity.drawings)) : 0;
    const dividends = isProp ? 0 : r2(a.dividends);
    const capitalAdd = r2(a.capitalAdditions);
    const eqBegin = num(B.total_equity);
    const eqEnd = r2(eqBegin + capitalAdd + netIncome - drawings - dividends);
    const liabTotal = sum(LIAB_ORDER, (k) => cur[k]);
    const nonCash = r2(cur.receivables + cur.inventory + cur.prepaid + cur.other_ca + cur.ppe_net + cur.other_nca);
    cur.cash = r2(liabTotal + eqEnd - nonCash);
    const totalAssets = r2(cur.cash + nonCash);
    const pr = {}; ASSET_ORDER.concat(LIAB_ORDER).forEach((k) => { pr[k] = num(B[k]); });
    if (hasGrossPPE) pr.ppe_net = r2(pr.ppe_cost - pr.accumulated_depreciation);
    const sfp = {
      title: isProp ? 'BALANCE SHEETS' : 'STATEMENTS OF FINANCIAL POSITION', cur, prior: pr, hasGrossPPE,
      totalCA: r2(cur.cash + cur.receivables + cur.inventory + cur.prepaid + cur.other_ca),
      priorCA: r2(pr.cash + pr.receivables + pr.inventory + pr.prepaid + pr.other_ca),
      totalNCA: r2(cur.ppe_net + cur.other_nca), priorNCA: r2((pr.ppe_net || r2(pr.ppe_cost - pr.accumulated_depreciation)) + pr.other_nca),
      totalAssets, priorAssets: num(B.total_assets) || r2(pr.cash + pr.receivables + pr.inventory + pr.prepaid + pr.other_ca + (pr.ppe_net || 0) + pr.other_nca),
      totalLiab: liabTotal, priorLiab: num(B.total_liabilities) || sum(LIAB_ORDER, (k) => pr[k]),
      equity: { begin: eqBegin, capitalAdd, netIncome, drawings, dividends, end: eqEnd, lines: P.equity.lines || [] },
      priorEquity: eqBegin, labels: B.labels || {}
    };
    sfp.totalLE = r2(liabTotal + eqEnd);
    sfp.priorLE = r2(sfp.priorLiab + eqBegin);

    // ---------- cash flow (indirect)
    const d = (k) => r2(num(cur[k]) - num(pr[k]));
    const op = [
      { label: isProp ? 'Excess of revenues over expenses for the year' : 'Net income for the year', amount: netIncome },
      { label: 'Depreciation and amortization', amount: dep },
      { label: '(Increase) decrease in receivables', amount: -d('receivables') },
      { label: '(Increase) decrease in inventory', amount: -d('inventory') },
      { label: '(Increase) decrease in prepaid expenses and other current assets', amount: -r2(d('prepaid') + d('other_ca')) },
      { label: 'Increase (decrease) in accounts payable and accrued expenses', amount: r2(d('payables') + d('accrued') + d('other_cl')) },
      { label: 'Increase (decrease) in VAT payable', amount: d('vat_payable') },
      { label: 'Increase (decrease) in income tax payable', amount: d('income_tax_payable') }
    ].filter((l) => l.amount || /net income|excess/i.test(l.label));
    const inv = [
      { label: 'Net additions to property and equipment', amount: -additions },
      { label: '(Increase) decrease in other non-current assets', amount: -d('other_nca') }
    ];
    if (isProp) inv.push({ label: 'Additional capital (Drawings)', amount: r2(capitalAdd - drawings) });
    const fin = [{ label: 'Proceeds from (payment of) loans', amount: r2(d('loans') + d('other_ncl')) }];
    if (!isProp) { if (dividends) fin.push({ label: 'Dividends paid', amount: -dividends }); if (capitalAdd) fin.push({ label: 'Additional paid-up capital', amount: capitalAdd }); }
    // prior-year AFS figures are often in whole pesos: absorb small rounding gaps so the statement ties exactly
    const preNet = r2(sum(op, (l) => l.amount) + sum(inv, (l) => l.amount) + sum(fin, (l) => l.amount));
    const roundGap = r2(cur.cash - (num(pr.cash) + preNet));
    if (roundGap && Math.abs(roundGap) < 5) op.push({ label: 'Rounding difference (prior-year figures in whole pesos)', amount: roundGap });
    const cfo = sum(op, (l) => l.amount), cfi = sum(inv, (l) => l.amount), cff = sum(fin, (l) => l.amount);
    const netChange = r2(cfo + cfi + cff);
    const scf = {
      op: op, inv: inv.filter((l) => l.amount), fin: fin.filter((l) => l.amount), cfo, cfi, cff, netChange,
      cashBegin: pr.cash, cashEnd: r2(pr.cash + netChange),
      prior: P.cash_flow || {}, reconciles: Math.abs(r2(pr.cash + netChange) - cur.cash) < 0.05
    };
    const sce = { begin: eqBegin, capitalAdd, netIncome, drawings, dividends, end: eqEnd, priorBegin: num(P.equity.beginning), priorNetIncome: num(P.sci.net_income), priorDrawings: num(P.equity.drawings), priorDividends: num(P.equity.dividends), priorEnd: eqBegin };
    const warnings = [];
    if (cur.cash < 0) warnings.push('Cash came out negative (₱' + fmt(cur.cash) + ') as the balancing figure. Review drawings/dividends, loans, payables or receivables in the Balance Sheet inputs.');
    // prior-year consistency: do the prior line items add up to its own totals?
    const priorAssetsSum = r2(pr.cash + pr.receivables + pr.inventory + pr.prepaid + pr.other_ca + pr.ppe_net + pr.other_nca);
    const priorLiabSum = sum(LIAB_ORDER, (k) => pr[k]);
    const priorGap = r2(priorAssetsSum - (priorLiabSum + eqBegin));
    sfp.priorGap = priorGap;
    if (Math.abs(priorGap) >= 5) {
      warnings.push('The prior-year balance sheet does not add up: its asset lines total ₱' + fmt(priorAssetsSum) + ' but liabilities + equity lines total ₱' + fmt(r2(priorLiabSum + eqBegin)) +
        ' (difference ₱' + fmt(priorGap) + '). The cash flow will be off by the same amount until the prior-year figures are corrected in the Balance Sheet tab.');
    } else if (!scf.reconciles) warnings.push('Cash flow does not reconcile to the balance sheet cash. Check the prior-year balance sheet figures.');
    if (num(B.total_assets) && Math.abs(num(B.total_assets) - priorAssetsSum) >= 5) warnings.push('Prior-year total assets per the AFS (₱' + fmt(B.total_assets) + ') differ from the sum of its asset lines (₱' + fmt(priorAssetsSum) + ').');
    return { sci, sfp, scf, sce, warnings, dep, isProp };
  }

  /* ------------------------------------------------------------ ITR Template check */
  function templateCheck(ctx, rev, tax) {
    const t = ctx.template; if (!t) return [];
    const out = [];
    const cp = (t.checkpoints || []);
    const get = (key) => cp.find((c) => c.key === key);
    const add = (key, label, verified, source) => {
      const c = get(key);
      if (!c || verified == null) return;
      const diff = r2(num(verified) - num(c.value));
      out.push({ key, label, template: r2(c.value), verified: r2(verified), diff, ok: Math.abs(diff) < 1, cell: c.cell || '', isFormula: !!c.is_formula, inputCell: c.input_cell || '', source,
        writeValue: c.is_formula && c.input_cell && c.input_factor ? r2(num(verified) * num(c.input_factor)) : null });
    };
    rev.byQ.forEach((q) => {
      if (!q.found) return;
      const src = (ctx.vat === 'VAT' ? '2550Q' : '2551Q') + ' Q' + q.quarter;
      if (ctx.vat === 'VAT') {
        add('q' + q.quarter + '_vatable_sales', 'Q' + q.quarter + ' VATable sales', r2(q.vatable + q.gov), src);
        add('q' + q.quarter + '_exempt_sales', 'Q' + q.quarter + ' exempt sales', q.exempt, src);
        add('q' + q.quarter + '_output_tax', 'Q' + q.quarter + ' output tax', q.outputTax, src);
        add('q' + q.quarter + '_total_sales', 'Q' + q.quarter + ' total sales', q.total, src);
      } else {
        add('q' + q.quarter + '_gross_sales', 'Q' + q.quarter + ' gross sales', q.total, src);
        add('q' + q.quarter + '_percentage_tax', 'Q' + q.quarter + ' percentage tax', q.ptDue, src);
      }
    });
    (tax.qPaid || []).forEach((x) => { if (x.src && x.src.indexOf('BIR') === 0) add('q' + x.q + '_itr_paid', 'Q' + x.q + ' income tax paid', x.amount, x.src); });
    add('annual_sales', 'Total sales for the year', rev.total, 'Sum of 4 quarterly returns');
    return out;
  }

  /* ------------------------------------------------------------ main */
  function emptyPrior() {
    return { itr: {}, sci: {}, balances: {}, equity: {}, cash_flow: {}, entity: {}, depreciation: 0 };
  }

  function build(input) {
    const ctx = Object.assign({ adj: {}, opts: {} }, input);
    ctx.prior = Object.assign(emptyPrior(), input.prior || {});
    ['itr', 'sci', 'balances', 'equity', 'cash_flow', 'entity'].forEach((k) => { ctx.prior[k] = ctx.prior[k] || {}; });
    const rev = buildRevenue(ctx.sales || [], ctx.vat);
    ctx.revenue = ctx.adj.revenue != null ? num(ctx.adj.revenue) : rev.total;
    ctx.ptTotal = rev.ptTotal;
    const q4 = rev.byQ[3]; ctx.q4VatPayable = q4 && q4.found ? q4.vatPayable : null;
    const exp = buildExpenses(ctx);
    const tax = ctx.form === '1702RT' ? tax1702(ctx, exp) : tax1701(ctx, exp);
    const fs = buildFS(ctx, exp, tax);
    const corrections = templateCheck(ctx, rev, tax);
    const warnings = [].concat(exp.warnings || [], tax.warnings || [], fs.warnings || []);
    if (rev.missing.length) warnings.push('Missing quarterly ' + (ctx.vat === 'VAT' ? '2550Q' : '2551Q') + ' for ' + rev.missing.map((q) => 'Q' + q).join(', ') + '. Revenue includes only the quarters uploaded.');
    if (tax.ent && tax.ent.disallowed) warnings.push('Representation & entertainment of ₱' + fmt(tax.ent.reTotal) + ' exceeds the ' + (tax.ent.rate * 100) + '% limit (₱' + fmt(tax.ent.cap) + '); ₱' + fmt(tax.ent.disallowed) + ' is not deductible.');
    if (tax.overpayment) warnings.push('Credits exceed the tax due by ₱' + fmt(tax.overpayment) + ' (overpayment — to be carried over, refunded or issued a TCC).');
    if (corrections.some((c) => !c.ok)) warnings.push(corrections.filter((c) => !c.ok).length + ' figure(s) in the ITR Template differ from the returns — see Template Check.');
    return { revenue: rev, revenueTotal: ctx.revenue, expenses: exp, tax, fs, corrections, warnings, year: ctx.year, form: ctx.form, vat: ctx.vat, client: ctx.client, prior: ctx.prior, isProp: fs.isProp };
  }

  const api = { build, graduatedTax, bracketText, TAX_TYPES, r2, norm, kindOf };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.AnnualEngine = api;
})(typeof window !== 'undefined' ? window : globalThis);
