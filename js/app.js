/* =====================================================================
   Charter Accounts — compta charter + fuel booklet
   Application statique (GitHub Pages) + Supabase (données, factures)
   ===================================================================== */
(() => {
  'use strict';

  // ------------------------------------------------------------------
  //  Constantes
  // ------------------------------------------------------------------
  const CFG = window.APP_CONFIG || {};
  const DEMO = !CFG.supabaseUrl || /VOTRE/i.test(CFG.supabaseUrl) || !CFG.supabaseAnonKey || /VOTRE/i.test(CFG.supabaseAnonKey);
  const SUPABASE_JS = 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.45.4/dist/umd/supabase.min.js';
  const XLSX_JS = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
  const PDFJS_JS = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js';
  const PDFJS_WORKER = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
  // Pays de début du charter → taux de TVA appliqué aux montants payés HT (modifiable par charter)
  const VAT_COUNTRIES = [['France', 20], ['Italie', 22], ['Espagne', 21], ['Croatie', 13], ['Grèce', 0], ['Monaco', 20], ['Autre', null]];
  const TESS_JS = 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js';
  const TESS_OPTS = { workerPath: 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/worker.min.js',
    corePath: 'https://cdn.jsdelivr.net/npm/tesseract.js-core@5.1.1' };
  const JSPDF_JS = 'https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js';
  const JSZIP_JS = 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js';
  const YACHT = CFG.yachtName || 'Motor Yacht';

  const METHODS = ['Credit Card APA', 'Bank Transfer from APA', 'Cash APA', 'ATM Fee', 'Bank Fee',
    'Bank Transfer from VIP', 'Credit Card VIP', 'Bank Transfer from TDB', 'Credit Card TDB', 'To be paid'];
  const CATEGORIES = ['Food', 'Drinks', 'Cigarette', 'Laundry', 'Berth', 'Hygiene', 'Medical', 'Fuel',
    'Agent Fee', 'Transport', 'Miscellaneous', 'Baby & Kid', 'Garbage', 'Administrative', 'Local Tax',
    'Bank Fee', 'TIPS', 'ATM fee'];

  // ------------------------------------------------------------------
  //  Utilitaires
  // ------------------------------------------------------------------
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => Array.from(el.querySelectorAll(s));
  const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const toNum = v => (v === null || v === undefined || v === '' || isNaN(+v)) ? null : +v;
  const r2 = v => Math.round((+v + Number.EPSILON) * 100) / 100;
  const r3 = v => Math.round((+v + Number.EPSILON) * 1000) / 1000;
  const sum = (arr, f = x => x) => arr.reduce((s, x) => s + (toNum(f(x)) || 0), 0);
  const uuid = () => (crypto.randomUUID ? crypto.randomUUID()
    : 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => { const r = Math.random() * 16 | 0; return (c === 'x' ? r : (r & 3 | 8)).toString(16); }));
  const today = () => new Date().toISOString().slice(0, 10);
  const store = {
    get(k) { try { return localStorage.getItem(k); } catch { return null; } },
    set(k, v) { try { localStorage.setItem(k, v); } catch { /* stockage indisponible */ } },
  };

  let LANG = 'fr';
  const locale = () => (LANG === 'fr' ? 'fr-FR' : 'en-GB');
  const money = (v, cur = 'EUR') => (v === null || v === undefined || v === '') ? ''
    : new Intl.NumberFormat(locale(), { style: 'currency', currency: cur || 'EUR' }).format(+v);
  const fmt = (v, d = 0) => (v === null || v === undefined || v === '' || isNaN(+v)) ? ''
    : new Intl.NumberFormat(locale(), { minimumFractionDigits: d, maximumFractionDigits: d }).format(+v);
  const fdate = s => s ? String(s).slice(0, 10).split('-').reverse().join('/') : '';
  const fdatetime = s => s ? new Date(s).toLocaleString(locale(), { dateStyle: 'medium', timeStyle: 'short' }) : '';

  function loadScript(src) {
    return new Promise((res, rej) => {
      if ($(`script[src="${src}"]`)) return res();
      const s = document.createElement('script');
      s.src = src; s.onload = res; s.onerror = () => rej(new Error('Chargement impossible : ' + src));
      document.head.appendChild(s);
    });
  }

  // ------------------------------------------------------------------
  //  Traductions (vue client FR / EN)
  // ------------------------------------------------------------------
  const I18N = {
    fr: {
      apaReceived: 'APA reçu', totalExpenses: 'Total des dépenses', apaBalance: 'Solde APA',
      toRefund: 'à rembourser au client', toPay: 'à régler par le client', fuelBilled: 'Fuel facturé',
      expenses: 'Dépenses', num: 'N°', date: 'Date', supplier: 'Fournisseur', ref: 'Réf.', category: 'Catégorie',
      description: 'Description', exclVat: 'HT', vat: 'TVA', total: 'Total', payment: 'Paiement', invoice: 'Facture',
      fuelConsumed: 'Fuel consommé', fuelDetail: 'Détail du fuel consommé', robBefore: 'ROB au départ',
      robAfter: 'ROB à l’arrivée', consumption: 'Consommation', lot: 'Soutage', port: 'Port', liters: 'Litres',
      pricePerL: 'Prix / L', amount: 'Montant', view: 'Voir', byCategory: 'Répartition par catégorie',
      cash: 'Espèces', cashWithdrawn: 'Retraits', cashSpent: 'Dépenses en espèces', cashInHand: 'Espèces restantes',
      signatures: 'Signatures', charterer: 'Charterer', master: 'Capitaine', notSigned: 'Non signé',
      signedOn: 'Signé le', by: 'par', signCta: 'Valider et signer la comptabilité', live: 'En direct',
      updated: 'mis à jour à', print: 'Imprimer / PDF', excel: 'Excel', noExpenses: 'Aucune dépense pour le moment.',
      signedBanner: 'Comptabilité validée et signée', signTitle: 'Signature de la comptabilité', yourName: 'Votre nom',
      signHere: 'Signez dans le cadre', clear: 'Effacer', approve: 'J’approuve le décompte des dépenses de ce charter',
      sign: 'Signer', cancel: 'Annuler', signedThanks: 'Merci, la comptabilité est signée.', invalidLink: 'Lien invalide ou expiré.',
      invalidLinkHelp: 'Demandez un nouveau lien au capitaine.', tipsOfFee: 'du prix du charter', fuelLine: 'Fuel',
      pending: 'En attente', voyage: 'Voyage', sigNeeded: 'Merci de signer dans le cadre.', nameNeeded: 'Merci d’indiquer votre nom.',
      approveNeeded: 'Merci de cocher la case d’approbation.', invoiceFuel: 'Facture soutage', from: 'du', to: 'au',
      noInvoice: '—', statement: 'Relevé des dépenses',
      openBanner: 'Comptabilité en cours — vous pourrez la signer dès que le capitaine l’aura clôturée.',
      closedBanner: 'La comptabilité est clôturée : vous pouvez la vérifier et la signer.',
      closedPopupTitle: 'Comptabilité clôturée', closedPopupText: 'Le capitaine a clôturé la comptabilité de votre charter. Vous pouvez maintenant la vérifier puis la signer.',
      signNow: 'Signer maintenant', later: 'Plus tard', tipsTitle: 'Tips pour l’équipage',
      tipsIntro: 'Souhaitez-vous laisser un pourboire à l’équipage ? Il est calculé sur le prix du charter',
      tipsIntroNoFee: 'Souhaitez-vous laisser un pourboire à l’équipage ?', tipsCustom: 'Personnalisé',
      tipsCustomHint: 'Saisissez un pourcentage ou un montant', percent: '%', amountUnit: 'Montant',
      continue: 'Continuer', back: 'Retour', tipsLine: 'Tips équipage', ofCharterFee: 'du charter fee',
      tipsChosen: 'Tips', totalToSign: 'Total à approuver', charterFee: 'Charter fee', tipsInvalid: 'Montant de tips invalide.',
      step: 'Étape', reportTitle: 'Rapport de charter', zipAll: 'Toutes les factures (.zip)',
      paidAs: 'Payé', inclVatChip: 'TTC', pricePerLPaid: 'Prix / L',
      vatEstimated: 'TVA estimée', apaLeftInclVat: 'Solde APA après TVA estimée', vatEstNote: 'Montant payé HT : la TVA sera facturée par le représentant fiscal après envoi du rapport',
      htChip: 'HT', apaLeft: 'Solde APA',
      boatInvoice: 'Facture', boatInvoicePending: 'Facture émise à la clôture',
      reportUnavailable: 'Ce rapport n’est pas disponible (compta non signée ou lien invalide).', preparingZip: 'Préparation du zip…',
    },
    en: {
      apaReceived: 'APA received', totalExpenses: 'Total expenses', apaBalance: 'APA balance',
      toRefund: 'to be refunded to the client', toPay: 'to be paid by the client', fuelBilled: 'Fuel charged',
      expenses: 'Expenses', num: 'No.', date: 'Date', supplier: 'Company', ref: 'Invoice', category: 'Category',
      description: 'Description', exclVat: 'Excl. VAT', vat: 'VAT', total: 'Total', payment: 'Method', invoice: 'Receipt',
      fuelConsumed: 'Fuel consumed', fuelDetail: 'Fuel consumption detail', robBefore: 'ROB before departure',
      robAfter: 'ROB after arrival', consumption: 'Consumption', lot: 'Bunkering', port: 'Port', liters: 'Liters',
      pricePerL: 'Price / L', amount: 'Amount', view: 'View', byCategory: 'Breakdown by category',
      cash: 'Cash', cashWithdrawn: 'Withdrawals', cashSpent: 'Cash expenses', cashInHand: 'Cash left',
      signatures: 'Signatures', charterer: 'Charterer', master: 'Master', notSigned: 'Not signed',
      signedOn: 'Signed on', by: 'by', signCta: 'Approve & sign the accounts', live: 'Live',
      updated: 'updated at', print: 'Print / PDF', excel: 'Excel', noExpenses: 'No expenses yet.',
      signedBanner: 'Accounts approved and signed', signTitle: 'Sign the charter accounts', yourName: 'Your full name',
      signHere: 'Sign inside the box', clear: 'Clear', approve: 'I approve the statement of expenses for this charter',
      sign: 'Sign', cancel: 'Cancel', signedThanks: 'Thank you, the accounts are signed.', invalidLink: 'Invalid or expired link.',
      invalidLinkHelp: 'Please ask the captain for a new link.', tipsOfFee: 'of the charter fee', fuelLine: 'Fuel',
      pending: 'Pending', voyage: 'Voyage', sigNeeded: 'Please sign inside the box.', nameNeeded: 'Please enter your name.',
      approveNeeded: 'Please tick the approval box.', invoiceFuel: 'Bunkering invoice', from: 'from', to: 'to',
      noInvoice: '—', statement: 'Statement of expenses',
      openBanner: 'Accounts in progress — you will be able to sign once the captain has closed them.',
      closedBanner: 'The accounts are closed: you can now review and sign them.',
      closedPopupTitle: 'Accounts closed', closedPopupText: 'The captain has closed the accounts of your charter. You can now review and sign them.',
      signNow: 'Sign now', later: 'Later', tipsTitle: 'Crew gratuity',
      tipsIntro: 'Would you like to leave a gratuity for the crew? It is calculated on the charter fee',
      tipsIntroNoFee: 'Would you like to leave a gratuity for the crew?', tipsCustom: 'Custom',
      tipsCustomHint: 'Enter a percentage or an amount', percent: '%', amountUnit: 'Amount',
      continue: 'Continue', back: 'Back', tipsLine: 'Crew gratuity', ofCharterFee: 'of the charter fee',
      tipsChosen: 'Gratuity', totalToSign: 'Total to approve', charterFee: 'Charter fee', tipsInvalid: 'Invalid gratuity amount.',
      step: 'Step', reportTitle: 'Charter report', zipAll: 'All invoices (.zip)',
      paidAs: 'Paid', inclVatChip: 'incl. VAT', pricePerLPaid: 'Price / L',
      vatEstimated: 'VAT estimated', apaLeftInclVat: 'APA left including VAT estimated', vatEstNote: 'Paid excl. VAT: this VAT will be invoiced by the fiscal representative once the report is sent',
      htChip: 'excl. VAT', apaLeft: 'APA left',
      boatInvoice: 'Invoice', boatInvoicePending: 'Invoice issued at closing',
      reportUnavailable: 'This report is not available (accounts not signed or invalid link).', preparingZip: 'Preparing zip…',
    },
  };
  const t = k => (I18N[LANG] && I18N[LANG][k]) || I18N.fr[k] || k;

  // ------------------------------------------------------------------
  //  Calculs métier
  // ------------------------------------------------------------------
  const tipsEnabled = c => c.tips_prompt !== false;
  const ttcL = b => (toNum(b.price_ht_l) || 0) + (toNum(b.vat_l) || 0);
  const consumption = v => (toNum(v.rob_before) !== null && toNum(v.rob_after) !== null) ? v.rob_before - v.rob_after : null;
  const byPos = (a, b) => (a.position ?? 0) - (b.position ?? 0) || String(a.created_at || '').localeCompare(String(b.created_at || ''));
  const byLot = (a, b) => String(a.date || '9999').localeCompare(String(b.date || '9999')) || (a.num ?? 0) - (b.num ?? 0);
  const sortExpenses = list => list.slice().sort((a, b) =>
    (a.num ?? 1e9) - (b.num ?? 1e9) || String(a.date || '9999').localeCompare(String(b.date || '9999')) ||
    String(a.created_at || '').localeCompare(String(b.created_at || '')));

  // Même structure que la fonction SQL client_view()
  function buildBundle(S, c) {
    const vs = S.voyages.filter(v => v.charter_id === c.id).sort(byPos);
    const fuel = [];
    vs.forEach(v => S.allocations.filter(a => a.voyage_id === v.id).forEach(a => {
      const b = S.bunkers.find(x => x.id === a.bunker_id);
      if (b) fuel.push({ voyage_id: v.id, liters: a.liters, bunker_num: b.num, bunker_date: b.date, town: b.town,
        country: b.country, price_ht_l: b.price_ht_l, vat_l: b.vat_l, invoice_ref: b.invoice_ref, fuel_type: b.fuel_type,
        file_path: b.file_path, file_name: b.file_name });
    }));
    return {
      charter: c,
      expenses: sortExpenses(S.expenses.filter(e => e.charter_id === c.id)),
      cash: S.cash_withdrawals.filter(w => w.charter_id === c.id).sort((a, b) => String(a.date).localeCompare(String(b.date))),
      voyages: vs.map(v => ({ id: v.id, description: v.description, start_date: v.start_date, end_date: v.end_date, rob_before: v.rob_before, rob_after: v.rob_after })),
      fuel,
      invoices: (S.invoices || []).filter(i => i.charter_id === c.id && i.status === 'issued'),
    };
  }

  // TVA estimée d'une dépense payée (en partie) HT : saisie manuelle, sinon base HT × taux du charter
  const charterVatRate = c => toNum(c.vat_rate) ?? 0;
  const vatEstOf = (e, c) => {
    if (toNum(e.vat_est) != null) return r2(e.vat_est);
    const base = toNum(e.vat_est_base);
    return base ? r2(base * charterVatRate(c) / 100) : 0;
  };

  function computeStatement(B, opts = {}) {
    const c = B.charter;
    const tipsAuto = r2(toNum(opts.tipsAmount !== undefined ? opts.tipsAmount : c.tips_amount) || 0);
    const tipsAutoPct = opts.tipsAmount !== undefined ? opts.tipsPct : c.tips_pct;
    // Prix facturé au client = prix TTC du lot arrondi au millième (comme dans l'Excel : 2,228 €/L)
    // Fuel facturé au prix payé : HT si le lot a été payé HT (TVA = 0), sinon TTC.
    // Lot payé HT → le client doit le TTC : la TVA passe en « VAT estimated » au taux du charter.
    const fuelLines = B.fuel.map(f => {
      const price = r3((toNum(f.price_ht_l) || 0) + (toNum(f.vat_l) || 0));
      const amount = r2((toNum(f.liters) || 0) * price), paidHt = !(toNum(f.vat_l) > 0);
      return { ...f, price, amount, paidHt, vatEst: paidHt ? r2(amount * charterVatRate(c) / 100) : 0 };
    });
    const fuelTotal = r2(sum(fuelLines, f => f.amount));
    const fuelLiters = sum(fuelLines, f => f.liters);
    const expTotal = r2(sum(B.expenses, e => e.total));
    const total = r2(expTotal + fuelTotal + tipsAuto);
    const apa = toNum(c.apa_paid) || 0;
    const balance = r2(apa - total);
    const cashWithdrawn = r2(sum(B.cash, w => w.amount));
    const cashSpent = r2(sum(B.expenses.filter(e => e.method === 'Cash APA'), e => e.total));
    const byCat = {};
    if (fuelLines.length) byCat.Fuel = fuelTotal;
    B.expenses.forEach(e => { const k = (e.category || 'Other').trim(); byCat[k] = r2((byCat[k] || 0) + (toNum(e.total) || 0)); });
    if (tipsAuto) byCat.TIPS = r2((byCat.TIPS || 0) + tipsAuto);
    const tips = sum(B.expenses.filter(e => (e.category || '').toUpperCase() === 'TIPS'), e => e.total) + tipsAuto;
    const fee = toNum(c.charter_fee);
    return { tipsAuto, tipsAutoPct, fuelLines, fuelTotal, fuelLiters, expTotal, total, apa, balance, cashWithdrawn, cashSpent,
      cashInHand: r2(cashWithdrawn - cashSpent), byCat, tips, tipsPct: fee ? tips / fee : null,
      fuelVatEst: r2(sum(fuelLines, f => f.vatEst)),
      vatEst: r2(sum(B.expenses, e => vatEstOf(e, c)) + sum(fuelLines, f => f.vatEst)), get balanceInclVat() { return r2(this.balance - this.vatEst); } };
  }

  function lotUsage(S) {
    const used = {};
    S.allocations.forEach(a => { used[a.bunker_id] = (used[a.bunker_id] || 0) + (toNum(a.liters) || 0); });
    return used;
  }

  // Répartition automatique FIFO (lots les plus anciens d'abord)
  function autoAllocate(S, voyageId, endDate, liters) {
    const used = {};
    S.allocations.filter(a => a.voyage_id !== voyageId)
      .forEach(a => { used[a.bunker_id] = (used[a.bunker_id] || 0) + (toNum(a.liters) || 0); });
    let left = liters;
    const out = [];
    for (const b of S.bunkers.slice().sort(byLot)) {
      if (left <= 0.0001) break;
      if (endDate && b.date && b.date > endDate) continue;
      const rem = (toNum(b.quantity) || 0) - (used[b.id] || 0);
      if (rem <= 0.0001) continue;
      const take = Math.min(left, rem);
      out.push({ bunker_id: b.id, liters: r2(take) });
      left -= take;
    }
    return { out, left: r2(left) };
  }

  // ------------------------------------------------------------------
  //  Back-ends : Supabase (production) ou Démo (navigateur)
  // ------------------------------------------------------------------
  const TABLES = ['charters', 'expenses', 'cash_withdrawals', 'bunkers', 'voyages', 'allocations', 'boat_profile', 'invoices'];

  class SupabaseBackend {
    async init() {
      await loadScript(SUPABASE_JS);
      this.sb = window.supabase.createClient(CFG.supabaseUrl, CFG.supabaseAnonKey);
      this.channels = {};
    }
    async session() { const { data } = await this.sb.auth.getSession(); return data.session; }
    async signIn(email, password) {
      const { error } = await this.sb.auth.signInWithPassword({ email, password });
      if (error) throw new Error(error.message === 'Invalid login credentials' ? 'E-mail ou mot de passe incorrect.' : error.message);
    }
    async signOut() { await this.sb.auth.signOut(); }
    async isAdmin() { const { data, error } = await this.sb.rpc('is_admin'); if (error) throw error; return !!data; }
    async userEmail() { const s = await this.session(); return s?.user?.email || ''; }
    async loadAll() {
      const res = await Promise.all(TABLES.map(tb => this.sb.from(tb).select('*').limit(10000)));
      const S = {};
      res.forEach((r, i) => { if (r.error) throw r.error; S[TABLES[i]] = r.data; });
      return S;
    }
    async insert(table, row) {
      const { data, error } = await this.sb.from(table).insert(row).select().single();
      if (error) throw error; return data;
    }
    async insertMany(table, rows) {
      if (!rows.length) return [];
      const { data, error } = await this.sb.from(table).insert(rows).select();
      if (error) throw error; return data;
    }
    async update(table, id, patch) {
      const { data, error } = await this.sb.from(table).update(patch).eq('id', id).select().single();
      if (error) throw error; return data;
    }
    async remove(table, id) { const { error } = await this.sb.from(table).delete().eq('id', id); if (error) throw error; }
    async removeWhere(table, col, val) { const { error } = await this.sb.from(table).delete().eq(col, val); if (error) throw error; }
    async saveProfile(data) { const { error } = await this.sb.from('boat_profile').upsert({ id: 1, data }); if (error) throw error; }
    async issueInvoice(a) {
      const { data, error } = await this.sb.rpc('issue_invoice', { p_charter: a.charter_id, p_prefix: a.prefix, p_seller: a.seller,
        p_buyer: a.buyer, p_lines: a.lines, p_total_ht: a.total_ht, p_total_vat: a.total_vat, p_total_ttc: a.total_ttc,
        p_currency: a.currency, p_fingerprint: a.fingerprint });
      if (error) throw error; return data;
    }
    async cancelInvoice(id) { const { error } = await this.sb.rpc('cancel_invoice', { p_id: id }); if (error) throw error; }
    async upload(file, prefix) {
      const safe = file.name.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\w.\-]+/g, '_').slice(-80);
      const path = `${prefix}/${uuid()}/${safe}`;
      const { error } = await this.sb.storage.from('invoices').upload(path, file, { contentType: file.type || undefined, upsert: false });
      if (error) throw error;
      return { file_path: path, file_name: file.name };
    }
    async removeFile(path) { if (path) await this.sb.storage.from('invoices').remove([path]); }
    fileUrl(path) { return path ? this.sb.storage.from('invoices').getPublicUrl(path).data.publicUrl : null; }
    async clientView(token) {
      const { data, error } = await this.sb.rpc('client_view', { p_token: token });
      if (error) throw error; return data;
    }
    async reportView(token) {
      const { data, error } = await this.sb.rpc('report_view', { p_token: token });
      if (error) throw error; return data;
    }
    async clientSign(token, name, signature, tipsPct, tipsAmount) {
      const { data, error } = await this.sb.rpc('client_sign', { p_token: token, p_name: name, p_signature: signature,
        p_tips_pct: tipsPct ?? null, p_tips_amount: tipsAmount || 0 });
      if (error) throw error; return data;
    }
    channel(token) {
      if (!this.channels[token]) {
        const ch = this.sb.channel('charter-' + token);
        this.channels[token] = new Promise(res => ch.subscribe(st => { if (st === 'SUBSCRIBED') res(ch); }));
      }
      return this.channels[token];
    }
    async notify(token) {
      try { const ch = await Promise.race([this.channel(token), new Promise((_, rej) => setTimeout(rej, 4000))]);
        await ch.send({ type: 'broadcast', event: 'refresh', payload: { at: Date.now() } }); } catch { /* le client se rafraîchit aussi seul */ }
    }
    subscribe(token, cb) {
      this.sb.channel('charter-' + token).on('broadcast', { event: 'refresh' }, cb).subscribe();
    }
  }

  class DemoBackend {
    constructor() { this.KEY = 'charter-accounts-demo-v1'; this.files = {}; }
    async init() {
      const saved = store.get(this.KEY);
      this.S = saved ? JSON.parse(saved) : JSON.parse(JSON.stringify(window.DEMO_DATA || {}));
      TABLES.forEach(tb => { this.S[tb] = this.S[tb] || []; });
      this.S.charters.forEach(c => { if (!c.report_token) c.report_token = 'demo-report-' + c.id; });
      window.addEventListener('storage', e => { if (e.key === this.KEY && e.newValue) { this.S = JSON.parse(e.newValue); this._cbs?.forEach(cb => cb()); } });
    }
    save() { store.set(this.KEY, JSON.stringify(this.S)); }
    reset() { this.S = JSON.parse(JSON.stringify(window.DEMO_DATA || {})); TABLES.forEach(tb => { this.S[tb] = this.S[tb] || []; }); this.save(); }
    async session() { return store.get(this.KEY + '-session') ? { user: { email: 'demo' } } : null; }
    async signIn() { store.set(this.KEY + '-session', '1'); }
    async signOut() { try { localStorage.removeItem(this.KEY + '-session'); } catch { /* */ } }
    async isAdmin() { return true; }
    async userEmail() { return 'Mode démo'; }
    async loadAll() { return JSON.parse(JSON.stringify(this.S)); }
    _checkLock(table, row) {
      if (!['expenses', 'cash_withdrawals'].includes(table)) return;
      const c = this.S.charters.find(x => x.id === row.charter_id);
      if (c && (c.status === 'signed' || c.status === 'closed')) throw new Error('Compta clôturée : rouvrez-la avant de modifier.');
    }
    async insert(table, row) {
      const r = { id: uuid(), created_at: new Date().toISOString(), ...row };
      if (table === 'charters') { r.client_token = r.client_token || uuid(); r.report_token = r.report_token || uuid(); }
      if (table === 'charters') r.status = r.status || 'open';
      this._checkLock(table, r);
      this.S[table].push(r); this.save(); return r;
    }
    async insertMany(table, rows) { const out = []; for (const r of rows) out.push(await this.insert(table, r)); return out; }
    async update(table, id, patch) {
      const r = this.S[table].find(x => x.id === id); if (!r) throw new Error('Introuvable');
      this._checkLock(table, r); Object.assign(r, patch); this._checkLock(table, r); this.save(); return r;
    }
    async remove(table, id) {
      const r = this.S[table].find(x => x.id === id); if (r) this._checkLock(table, r);
      if (table === 'bunkers' && this.S.allocations.some(a => a.bunker_id === id)) throw new Error('Ce soutage est utilisé dans des voyages : retirez d’abord ses répartitions.');
      this.S[table] = this.S[table].filter(x => x.id !== id);
      if (table === 'charters') {
        ['expenses', 'cash_withdrawals'].forEach(tb => { this.S[tb] = this.S[tb].filter(x => x.charter_id !== id); });
        this.S.voyages.forEach(v => { if (v.charter_id === id) v.charter_id = null; });
      }
      if (table === 'voyages') this.S.allocations = this.S.allocations.filter(a => a.voyage_id !== id);
      this.save();
    }
    async removeWhere(table, col, val) { this.S[table] = this.S[table].filter(x => x[col] !== val); this.save(); }
    async saveProfile(data) { this.S.boat_profile = [{ id: 1, data }]; this.save(); }
    async issueInvoice(a) {
      const y = new Date().getFullYear(), pfx = String(a.prefix || 'BOAT').replace(/[^A-Za-z0-9]/g, '').toUpperCase() || 'BOAT';
      const seq = Math.max(0, ...this.S.invoices.filter(i => i.prefix === pfx && i.year === y).map(i => i.seq)) + 1;
      const inv = { id: uuid(), ...a, prefix: pfx, year: y, seq, number: `${pfx}-${y}-${String(seq).padStart(4, '0')}`,
        issue_date: today(), status: 'issued', created_at: new Date().toISOString() };
      this.S.invoices.push(inv); this.save(); return inv;
    }
    async cancelInvoice(id) { const i = this.S.invoices.find(x => x.id === id); if (i) { i.status = 'cancelled'; i.cancelled_at = new Date().toISOString(); } this.save(); }
    async upload(file, prefix) {
      const path = `${prefix}/${uuid()}/${file.name}`;
      this.files[path] = URL.createObjectURL(file);
      return { file_path: path, file_name: file.name };
    }
    async removeFile(path) { delete this.files[path]; }
    fileUrl(path) { return path ? (this.files[path] || null) : null; }
    async clientView(token) {
      const c = this.S.charters.find(x => x.client_token === token);
      if (!c) return null;
      const B = buildBundle(this.S, c);
      const { client_token, notes, signed_snapshot, report_token, report_to, report_cc, ...pub } = c; // eslint-disable-line no-unused-vars
      return { ...B, charter: pub, invoices: B.invoices.map(({ fingerprint, ...i }) => i) }; // eslint-disable-line no-unused-vars
    }
    async reportView(token) {
      const c = this.S.charters.find(x => x.report_token === token && x.status === 'signed');
      return c ? this.clientView(c.client_token) : null;
    }
    async clientSign(token, name, signature, tipsPct, tipsAmount) {
      const c = this.S.charters.find(x => x.client_token === token);
      if (!c) throw new Error('Lien invalide');
      if (c.status === 'signed') throw new Error('Ce charter est déjà signé');
      if (c.status !== 'closed') throw new Error('La compta n’est pas encore clôturée par le capitaine');
      if (!tipsEnabled(c)) { tipsPct = null; tipsAmount = 0; }
      if (!(tipsAmount >= 0)) throw new Error('Montant de tips invalide');
      Object.assign(c, { tips_pct: tipsPct ?? null, tips_amount: r2(tipsAmount || 0), client_signature: signature,
        client_signed_name: name, client_signed_at: new Date().toISOString(), status: 'signed' });
      c.signed_snapshot = await this.clientView(token);
      this.save(); return this.clientView(token);
    }
    async notify() { /* l'événement « storage » prévient les autres onglets */ }
    subscribe(token, cb) { (this._cbs = this._cbs || []).push(cb); }
  }

  const api = DEMO ? new DemoBackend() : new SupabaseBackend();
  const fileLink = (path, label, cls = '') => {
    if (!path) return '';
    const url = api.fileUrl(path);
    return url ? `<a class="file-link ${cls}" href="${esc(url)}" target="_blank" rel="noopener" title="${esc(label)}">${ICON.file}<span>${esc(label)}</span></a>`
      : `<span class="muted small" title="Fichier disponible uniquement pendant la session démo">${ICON.file} ${esc(label)}</span>`;
  };

  // ------------------------------------------------------------------
  //  Icônes (SVG inline)
  // ------------------------------------------------------------------
  const svg = p => `<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${p}</svg>`;
  const ICON = {
    file: svg('<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/>'),
    plus: svg('<path d="M12 5v14M5 12h14"/>'),
    edit: svg('<path d="M4 20h4L19 9l-4-4L4 16z"/>'),
    trash: svg('<path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/>'),
    share: svg('<circle cx="18" cy="5" r="3"/><circle cx="6" cy="12" r="3"/><circle cx="18" cy="19" r="3"/><path d="M8.6 13.5l6.8 4M15.4 6.5l-6.8 4"/>'),
    eye: svg('<path d="M2 12s3.6-7 10-7 10 7 10 7-3.6 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>'),
    print: svg('<path d="M6 9V3h12v6M6 18H4v-7h16v7h-2"/><path d="M6 14h12v7H6z"/>'),
    sheet: svg('<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18"/>'),
    back: svg('<path d="M15 18l-6-6 6-6"/>'),
    lock: svg('<rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/>'),
    unlock: svg('<rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 7.5-2"/>'),
    pen: svg('<path d="M3 21c3-1 5-3 7-6l7-7a2.1 2.1 0 0 0-3-3l-7 7c-3 2-5 4-6 7z"/>'),
    drop: svg('<path d="M12 3s6 6.5 6 11a6 6 0 0 1-12 0c0-4.5 6-11 6-11z"/>'),
    copy: svg('<rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15V5a2 2 0 0 1 2-2h10"/>'),
    link: svg('<path d="M10 14a5 5 0 0 0 7 0l3-3a5 5 0 0 0-7-7l-1 1"/><path d="M14 10a5 5 0 0 0-7 0l-3 3a5 5 0 0 0 7 7l1-1"/>'),
    check: svg('<path d="M5 12l5 5L20 7"/>'),
    logout: svg('<path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4M16 17l5-5-5-5M21 12H9"/>'),
    anchor: svg('<circle cx="12" cy="5" r="2"/><path d="M12 7v14M5 12H3a9 9 0 0 0 18 0h-2M8 10h8"/>'),
    mail: svg('<rect x="3" y="5" width="18" height="14" rx="2"/><path d="M3 7l9 6 9-6"/>'),
    wand: svg('<path d="M15 4V2M15 10V8M11 6h2M17 6h2M4 20L14 10"/>'),
  };

  // ------------------------------------------------------------------
  //  UI : modales, confirmations, toasts
  // ------------------------------------------------------------------
  function toast(msg, type = 'ok') {
    const el = document.createElement('div');
    el.className = `toast toast-${type}`;
    el.textContent = msg;
    $('#toast-root').appendChild(el);
    setTimeout(() => el.classList.add('show'), 10);
    setTimeout(() => { el.classList.remove('show'); setTimeout(() => el.remove(), 300); }, type === 'err' ? 6000 : 3000);
  }
  const errMsg = e => (e && (e.message || e.error_description || e.details)) || String(e);

  function openModal({ title, body, footer = '', wide = false, onMount }) {
    const root = $('#modal-root');
    root.innerHTML = `<div class="modal-backdrop"><div class="modal ${wide ? 'modal-wide' : ''}" role="dialog" aria-modal="true" aria-label="${esc(title)}">
      <div class="modal-head"><h3>${esc(title)}</h3><button class="icon-btn" data-close aria-label="Fermer">✕</button></div>
      <div class="modal-body">${body}</div>${footer ? `<div class="modal-foot">${footer}</div>` : ''}</div></div>`;
    const m = $('.modal', root);
    $$('[data-close]', root).forEach(b => b.addEventListener('click', closeModal));
    $('.modal-backdrop', root).addEventListener('mousedown', e => { if (e.target.classList.contains('modal-backdrop')) closeModal(); });
    document.body.classList.add('modal-open');
    onMount && onMount(m);
    const first = $('input:not([type=hidden]):not([type=checkbox]), select, textarea', m);
    if (first && window.matchMedia('(pointer:fine)').matches) first.focus();
    return m;
  }
  function closeModal() { $('#modal-root').innerHTML = ''; document.body.classList.remove('modal-open'); }
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && $('#modal-root').innerHTML) closeModal(); });

  function confirmDialog(message, okLabel = 'Confirmer', danger = true) {
    return new Promise(res => {
      const m = openModal({ title: 'Confirmation', body: `<p>${message}</p>`,
        footer: `<button class="btn btn-ghost" data-close>Annuler</button><button class="btn ${danger ? 'btn-danger' : 'btn-primary'}" data-ok>${esc(okLabel)}</button>` });
      let done = false;
      $('[data-ok]', m).addEventListener('click', () => { done = true; closeModal(); res(true); });
      const obs = new MutationObserver(() => { if (!$('#modal-root').innerHTML) { obs.disconnect(); if (!done) res(false); } });
      obs.observe($('#modal-root'), { childList: true });
    });
  }

  function busy(btn, on) {
    if (!btn) return;
    btn.disabled = on;
    btn.classList.toggle('is-busy', on);
  }

  // ------------------------------------------------------------------
  //  Pavé de signature
  // ------------------------------------------------------------------
  function signaturePad(canvas) {
    const ctx = canvas.getContext('2d');
    let drawing = false, empty = true, last = null;
    function resize() {
      const dpr = window.devicePixelRatio || 1;
      const w = canvas.clientWidth, h = canvas.clientHeight;
      canvas.width = w * dpr; canvas.height = h * dpr;
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      ctx.lineWidth = 2.2; ctx.lineCap = 'round'; ctx.lineJoin = 'round'; ctx.strokeStyle = '#0f2a44';
      empty = true;
    }
    const pos = e => { const r = canvas.getBoundingClientRect(); return { x: e.clientX - r.left, y: e.clientY - r.top }; };
    canvas.addEventListener('pointerdown', e => { e.preventDefault(); canvas.setPointerCapture(e.pointerId); drawing = true; last = pos(e);
      ctx.beginPath(); ctx.arc(last.x, last.y, 1.1, 0, Math.PI * 2); ctx.fillStyle = '#0f2a44'; ctx.fill(); empty = false; });
    canvas.addEventListener('pointermove', e => { if (!drawing) return; e.preventDefault(); const p = pos(e);
      ctx.beginPath(); ctx.moveTo(last.x, last.y); ctx.lineTo(p.x, p.y); ctx.stroke(); last = p; empty = false; });
    ['pointerup', 'pointercancel', 'pointerleave'].forEach(ev => canvas.addEventListener(ev, () => { drawing = false; }));
    resize();
    return {
      clear: resize,
      isEmpty: () => empty,
      toDataURL() { // export à l'échelle 1 pour limiter la taille
        const out = document.createElement('canvas');
        out.width = canvas.clientWidth; out.height = canvas.clientHeight;
        out.getContext('2d').drawImage(canvas, 0, 0, out.width, out.height);
        return out.toDataURL('image/png');
      },
    };
  }

  // ------------------------------------------------------------------
  //  Rendu commun : relevé de compte (capitaine ET client)
  // ------------------------------------------------------------------
  function renderKpis(B, st) {
    const cur = B.charter.currency;
    const pos = st.balance >= 0;
    return `<div class="kpis ${st.vatEst ? 'kpis-6' : ''}">
      <div class="kpi"><div class="kpi-label">${t('apaReceived')}</div><div class="kpi-value">${money(st.apa, cur)}</div></div>
      <div class="kpi"><div class="kpi-label">${t('totalExpenses')}</div><div class="kpi-value">${money(st.total, cur)}</div></div>
      <div class="kpi kpi-balance ${pos ? 'is-pos' : 'is-neg'}"><div class="kpi-label">${t('apaBalance')}</div>
        <div class="kpi-value">${money(st.balance, cur)}</div><div class="kpi-sub">${pos ? t('toRefund') : t('toPay')}</div></div>
      <div class="kpi"><div class="kpi-label">${t('fuelBilled')}</div><div class="kpi-value">${money(st.fuelTotal, cur)}</div>
        <div class="kpi-sub">${fmt(st.fuelLiters)} L</div></div>
      ${st.vatEst ? `<div class="kpi kpi-warn"><div class="kpi-label">${t('vatEstimated')}</div><div class="kpi-value">${money(st.vatEst, cur)}</div>
        <div class="kpi-sub">${fmt(charterVatRate(B.charter), charterVatRate(B.charter) % 1 ? 1 : 0)} %${B.charter.vat_country ? ' · ' + esc(B.charter.vat_country) : ''}</div></div>
      <div class="kpi kpi-balance ${st.balanceInclVat >= 0 ? 'is-pos' : 'is-neg'}"><div class="kpi-label">${t('apaLeftInclVat')}</div>
        <div class="kpi-value">${money(st.balanceInclVat, cur)}</div><div class="kpi-sub">${st.balanceInclVat >= 0 ? t('toRefund') : t('toPay')}</div></div>` : ''}
    </div>`;
  }

  const boatInvCell = (B, label) => (B.invoices || []).length ? B.invoices.map(invoiceLink).join('')
    : IS_CLIENT ? `<span class="chip chip-gold" title="${t('boatInvoicePending')}">${t('boatInvoice')} ${esc(label || '')}</span>`
    : `<button type="button" class="chip chip-gold link-plain" data-act="inv-draft" title="Aperçu de la facture (brouillon) — elle sera numérotée à la clôture">${ICON.eye} Facture ${esc(label || '')} (aperçu)</button>`;

  function renderExpensesTable(B, st, { admin = false, locked = false } = {}) {
    const c = B.charter, cur = c.currency;
    const fuelInvoices = [];
    const seen = new Set();
    st.fuelLines.forEach(f => { if (f.file_path && !seen.has(f.file_path)) { seen.add(f.file_path); fuelInvoices.push(f); } });
    const rows = [];
    if (st.fuelLines.length) {
      const prices = [...new Set(st.fuelLines.map(f => f.price))];
      rows.push(`<tr class="row-fuel">
        <td data-label="${t('num')}">1</td><td data-label="${t('date')}">${fdate(c.end_date)}</td>
        <td data-label="${t('supplier')}">${esc(c.fuel_company || '')}</td><td data-label="${t('ref')}"></td>
        <td data-label="${t('category')}"><span class="chip chip-fuel">${ICON.drop}Fuel</span></td>
        <td data-label="${t('description')}">${esc(c.fuel_description || '')} — ${fmt(st.fuelLiters)} L × ${prices.map(p => fmt(p, 3)).join(' / ')} €/L${st.fuelVatEst ? ` <span class="chip chip-warn" title="${t('vatEstNote')}">${t('htChip')} · ${t('vatEstimated')} ${money(st.fuelVatEst, cur)}</span>` : ''}</td>
        <td data-label="${t('exclVat')}" class="num"></td><td data-label="${t('vat')}" class="num"></td>
        <td data-label="${t('total')}" class="num strong">${money(st.fuelTotal, cur)}</td>
        <td data-label="${t('payment')}">${esc(c.fuel_method || '')}</td>
        <td data-label="${t('invoice')}" class="files">${c.fuel_by_boat !== false ? boatInvCell(B, c.fuel_company) : ''}${fuelInvoices.map(f => fileLink(f.file_path, `${t('lot')} ${f.bunker_num ?? ''}`)).join('') || ''}</td>
        ${admin ? `<td class="actions"><a class="icon-btn" href="#/fuel/voyages" title="Gérer dans le fuel booklet">${ICON.drop}</a></td>` : ''}
      </tr>`);
    }
    B.expenses.forEach(e => {
      rows.push(`<tr class="${admin ? 'row-edit' : ''}" ${admin ? `data-row-exp="${e.id}"` : ''}>
        <td data-label="${t('num')}">${esc(e.num ?? '')}</td><td data-label="${t('date')}">${fdate(e.date)}</td>
        <td data-label="${t('supplier')}">${esc(e.company)}</td><td data-label="${t('ref')}">${esc(e.invoice_ref)}</td>
        <td data-label="${t('category')}"><span class="chip">${esc(e.category)}</span></td>
        <td data-label="${t('description')}">${esc(e.description)}${vatEstOf(e, c) ? ` <span class="chip chip-warn" title="${t('vatEstNote')}">${t('htChip')} ${money(e.vat_est_base ?? '', cur)} · ${t('vatEstimated')} ${money(vatEstOf(e, c), cur)}</span>` : ''}</td>
        <td data-label="${t('exclVat')}" class="num">${e.amount_ht != null ? money(e.amount_ht, cur) : ''}</td>
        <td data-label="${t('vat')}" class="num">${e.vat != null ? money(e.vat, cur) : ''}</td>
        <td data-label="${t('total')}" class="num strong">${money(e.total, cur)}</td>
        <td data-label="${t('payment')}">${esc(e.method)}</td>
        <td data-label="${t('invoice')}" class="files">${e.billed_by_boat ? boatInvCell(B, e.company) : ''}${e.file_path ? fileLink(e.file_path, t('view')) : ''}</td>
        ${admin ? `<td class="actions">${locked ? `<button class="icon-btn" data-act="locked-edit" data-id="${e.id}" title="Compta clôturée : rouvrir pour modifier">${ICON.lock}</button>`
          : `<button class="icon-btn" data-act="edit-expense" data-id="${e.id}" title="Modifier">${ICON.edit}</button>
          <button class="icon-btn danger" data-act="del-expense" data-id="${e.id}" title="Supprimer">${ICON.trash}</button>`}</td>` : ''}
      </tr>`);
    });
    if (st.tipsAuto) {
      rows.push(`<tr class="row-tips">
        <td data-label="${t('num')}"></td><td data-label="${t('date')}">${fdate((c.client_signed_at || '').slice(0, 10) || today())}</td>
        <td data-label="${t('supplier')}">Crew</td><td data-label="${t('ref')}"></td>
        <td data-label="${t('category')}"><span class="chip chip-gold">TIPS</span></td>
        <td data-label="${t('description')}">${t('tipsLine')}${st.tipsAutoPct != null ? ` — ${fmt(st.tipsAutoPct, st.tipsAutoPct % 1 ? 1 : 0)} % ${t('ofCharterFee')}` : ''}</td>
        <td data-label="${t('exclVat')}" class="num"></td><td data-label="${t('vat')}" class="num"></td>
        <td data-label="${t('total')}" class="num strong">${money(st.tipsAuto, cur)}</td>
        <td data-label="${t('payment')}">APA</td><td data-label="${t('invoice')}"></td>${admin ? '<td class="actions"></td>' : ''}</tr>`);
    }
    const empty = !rows.length ? `<tr><td colspan="${admin ? 12 : 11}" class="empty">${t('noExpenses')}</td></tr>` : '';
    return `<div class="table-wrap"><table class="table table-cards">
      <thead><tr><th>${t('num')}</th><th>${t('date')}</th><th>${t('supplier')}</th><th>${t('ref')}</th><th>${t('category')}</th>
      <th>${t('description')}</th><th class="num">${t('exclVat')}</th><th class="num">${t('vat')}</th><th class="num">${t('total')}</th>
      <th>${t('payment')}</th><th>${t('invoice')}</th>${admin ? '<th class="actions"></th>' : ''}</tr></thead>
      <tbody>${rows.join('')}${empty}</tbody>
      <tfoot><tr><td colspan="8" class="tfoot-label">${t('total')}</td><td class="num strong">${money(st.total, cur)}</td><td colspan="${admin ? 3 : 2}"></td></tr>
      ${st.vatEst ? `<tr class="tfoot-vat"><td colspan="8" class="tfoot-label">${t('vatEstimated')} (${fmt(charterVatRate(c), charterVatRate(c) % 1 ? 1 : 0)} %)</td><td class="num strong">${money(st.vatEst, cur)}</td><td colspan="${admin ? 3 : 2}"></td></tr>` : ''}</tfoot>
    </table></div>`;
  }

  function renderFuelDetail(B, st) {
    if (!B.voyages.length) return '';
    const cur = B.charter.currency;
    return B.voyages.map(v => {
      const lines = st.fuelLines.filter(f => f.voyage_id === v.id);
      return `<div class="fuel-voyage">
        <div class="fuel-voyage-head"><strong>${esc(v.description || t('voyage'))}</strong>
          <span class="muted">${fdate(v.start_date)} → ${fdate(v.end_date)}</span></div>
        <div class="fuel-figs">
          <div><span>${t('robBefore')}</span><b>${fmt(v.rob_before)} L</b></div>
          <div><span>${t('robAfter')}</span><b>${fmt(v.rob_after)} L</b></div>
          <div><span>${t('consumption')}</span><b>${fmt(consumption(v))} L</b></div>
        </div>
        ${lines.length ? `<div class="table-wrap"><table class="table table-compact table-cards">
          <thead><tr><th>${t('lot')}</th><th>${t('date')}</th><th>${t('port')}</th><th class="num">${t('liters')}</th>
          <th class="num">${t('pricePerL')}</th><th>${t('paidAs')}</th><th class="num">${t('amount')}</th><th>${t('invoice')}</th></tr></thead>
          <tbody>${lines.map(f => `<tr><td data-label="${t('lot')}">#${esc(f.bunker_num ?? '')}</td><td data-label="${t('date')}">${fdate(f.bunker_date)}</td>
            <td data-label="${t('port')}" title="${esc(f.country || '')}">${esc(f.town || '')}</td><td data-label="${t('liters')}" class="num">${fmt(f.liters)}</td>
            <td data-label="${t('pricePerL')}" class="num">${fmt(f.price, 3)} €</td>
            <td data-label="${t('paidAs')}">${f.paidHt ? `<span class="chip chip-warn">${t('htChip')}</span>${f.vatEst ? `<br><span class="small warn">${t('vatEstimated')} ${money(f.vatEst, cur)}</span>` : ''}` : `<span class="chip">${t('inclVatChip')}</span>`}</td><td data-label="${t('amount')}" class="num strong">${money(f.amount, cur)}</td>
            <td data-label="${t('invoice')}" class="files">${f.file_path ? fileLink(f.file_path, f.invoice_ref || t('view')) : `<span class="muted">${esc(f.invoice_ref || '')}</span>`}</td></tr>`).join('')}</tbody>
        </table></div>` : `<p class="muted small">${t('pending')}</p>`}
      </div>`;
    }).join('');
  }

  function renderCategories(B, st) {
    const entries = Object.entries(st.byCat).sort((a, b) => b[1] - a[1]);
    if (!entries.length) return '';
    const max = Math.max(...entries.map(e => Math.abs(e[1]))) || 1;
    const cur = B.charter.currency;
    return `<div class="cats">${entries.map(([k, v]) => `<div class="cat-row">
      <div class="cat-name">${esc(k)}</div><div class="cat-bar"><span style="width:${Math.max(2, Math.abs(v) / max * 100)}%"></span></div>
      <div class="cat-val">${money(v, cur)} <span class="muted small">${st.total ? fmt(v / st.total * 100, 1) + ' %' : ''}</span></div></div>`).join('')}
      ${st.tipsPct != null && st.tips ? `<p class="muted small cat-note">TIPS : ${fmt(st.tipsPct * 100, 1)} % ${t('tipsOfFee')} (${money(B.charter.charter_fee, cur)})</p>` : ''}
    </div>`;
  }

  function renderSignatureBox(label, sig, name, at) {
    return `<div class="sig-box ${sig ? 'is-signed' : ''}">
      <div class="sig-label">${esc(label)}</div>
      ${sig ? `<img class="sig-img" src="${esc(sig)}" alt="Signature ${esc(label)}">
        <div class="sig-meta">${esc(name || '')}<br><span class="muted small">${t('signedOn')} ${fdatetime(at)}</span></div>`
        : `<div class="sig-empty">${t('notSigned')}</div>`}
    </div>`;
  }

  // ------------------------------------------------------------------
  //  Facturation par le bateau
  // ------------------------------------------------------------------
  const PROFILE_DEFAULTS = {
    short_name: 'Boat', legal_name: '', address: '', vat_number: '', registration: '', email: '', phone: '',
    iban: '', bic: '', bank: '', invoice_prefix: 'BOAT', vat_rate: 0,
    vat_mention: '', payment_terms: 'Montant prélevé sur l’APA (Advance Provisioning Allowance) du charter.', footer: '',
    suppliers: [],
  };
  let PROFILE = { ...PROFILE_DEFAULTS };
  const boatName = () => (PROFILE.short_name || 'Boat').trim();
  const isBoatName = s => !!s && String(s).trim().toLowerCase() === boatName().toLowerCase();
  const INV_CACHE = {};

  // Lignes facturées par le bateau : fuel (si coché sur le charter) + dépenses marquées « facturée par le bateau »
  function boatInvoiceDraft(B, prof = PROFILE) {
    const c = B.charter, st = computeStatement(B), rateDefault = toNum(prof.vat_rate) || 0;
    // ligne : HT + TVA = TTC (le TTC est toujours le montant de la compta)
    const mk = (designation, qty, unit, ht, ttc, unitPrice) => {
      ht = r2(ht); ttc = r2(ttc);
      const vat = r2(ttc - ht), rate = ht ? r2(vat / ht * 100) : 0;
      return { designation, qty, unit, rate, unit_price_ht: unitPrice ?? (qty ? r3(ht / qty) : ht), total_ht: ht, total_vat: vat, total_ttc: ttc };
    };
    const lines = [];
    // Fuel : une ligne par soutage consommé, au prix payé (HT si payé HT, sinon HT + TVA du fournisseur)
    if (c.fuel_by_boat !== false) st.fuelLines.forEach(f => {
      const ht = f.paidHt ? f.amount : (toNum(f.liters) || 0) * (toNum(f.price_ht_l) || 0);
      lines.push(mk(`${f.fuel_type || 'Gasoil'} — ${c.fuel_description || 'MEs & DGs'} — soutage n°${f.bunker_num ?? ''} du ${fdate(f.bunker_date)}${f.town ? ` (${f.town})` : ''}${f.paidHt ? ' — payé HT' : ''}`,
        toNum(f.liters), 'L', ht, f.amount, f.paidHt ? f.price : toNum(f.price_ht_l)));
    });
    B.expenses.filter(e => e.billed_by_boat).forEach(e => {
      const ttc = toNum(e.total) || 0;
      const ht = toNum(e.amount_ht) != null && toNum(e.vat) != null ? e.amount_ht : r2(ttc / (1 + rateDefault / 100));
      lines.push(mk([e.category, e.description].filter(Boolean).join(' — ') || 'Prestation', 1, '', ht, ttc));
    });
    const fuelHt = st.fuelLines.filter(f => f.paidHt && f.vatEst);
    const vatEstNote = c.fuel_by_boat !== false && fuelHt.length
      ? `Fuel payé HT : TVA estimée de ${money(sum(fuelHt, f => f.vatEst), c.currency)} (${charterVatRate(c)} %${c.vat_country ? ' — ' + c.vat_country : ''}) facturée séparément par le représentant fiscal.` : '';
    const seller = { short_name: prof.short_name, legal_name: prof.legal_name || prof.short_name, address: prof.address,
      vat_number: prof.vat_number, registration: prof.registration, email: prof.email, phone: prof.phone,
      iban: prof.iban, bic: prof.bic, bank: prof.bank, yacht: c.yacht || YACHT };
    const buyer = { name: c.client_name || '', address: c.client_address || '', vat_number: c.client_vat || '',
      charter: c.title, start_date: c.start_date, end_date: c.end_date };
    const totals = { total_ht: r2(sum(lines, l => l.total_ht)), total_vat: r2(sum(lines, l => l.total_vat)), total_ttc: r2(sum(lines, l => l.total_ttc)) };
    const notes = { vat_mention: [prof.vat_mention, vatEstNote].filter(Boolean).join('\n'), payment_terms: prof.payment_terms || '', footer: prof.footer || '' };
    const fingerprint = JSON.stringify({ lines, buyer, totals, seller: { ...seller }, notes });
    return { charter_id: c.id, prefix: prof.invoice_prefix || prof.short_name || 'BOAT', seller: { ...seller, notes }, buyer,
      lines, ...totals, currency: c.currency || 'EUR', fingerprint };
  }

  // Émet (ou ré-émet si le contenu a changé) la facture du charter — appelé à la clôture
  async function ensureBoatInvoice(c) {
    const draft = boatInvoiceDraft(buildBundle(S, c));
    const current = S.invoices.find(i => i.charter_id === c.id && i.status === 'issued');
    if (current && current.fingerprint === draft.fingerprint) return { inv: current, changed: false };
    if (current) await api.cancelInvoice(current.id);
    if (!draft.lines.length) return { inv: null, changed: !!current };
    const inv = await api.issueInvoice(draft);
    return { inv, changed: true, replaced: current?.number };
  }

  const pdfText = s => String(s ?? '').replace(/\u03bc/g, '\u00b5').replace(/[  ]/g, ' ').replace(/→/g, '->').replace(/[^\x00-\xff€—–’‘“”…•]/g, '');
  const pdfMoney = (v, cur) => pdfText(money(v, cur));

  async function invoicePdf(inv, draft = false) {
    await loadScript(JSPDF_JS);
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: 'mm', format: 'a4' });
    const W = 210, M = 16, cur = inv.currency || 'EUR', s = inv.seller || {}, b = inv.buyer || {}, n = s.notes || {};
    const navy = [15, 42, 68], gold = [184, 147, 90], grey = [110, 120, 135];
    if (draft) { doc.setFont('helvetica', 'bold'); doc.setFontSize(70); doc.setTextColor(244, 238, 226); doc.text('BROUILLON', W / 2, 200, { align: 'center', angle: 30 }); }
    let y = 18;
    doc.setFont('helvetica', 'bold'); doc.setFontSize(18); doc.setTextColor(...navy);
    doc.text(pdfText(s.legal_name || s.short_name || 'Boat'), M, y);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(...grey);
    const sellerLines = [s.yacht, ...(String(s.address || '').split('\n')), s.vat_number && `TVA : ${s.vat_number}`,
      s.registration, [s.email, s.phone].filter(Boolean).join(' · ')].filter(Boolean).map(pdfText);
    sellerLines.forEach((l, i) => doc.text(l, M, y + 6 + i * 4.3));
    // titre facture
    doc.setFont('helvetica', 'bold'); doc.setFontSize(22); doc.setTextColor(...gold);
    doc.text(draft ? 'BROUILLON' : 'FACTURE', W - M, y, { align: 'right' });
    doc.setFontSize(10); doc.setTextColor(...navy);
    doc.text(pdfText(draft ? 'Non numérotée — aperçu' : `N° ${inv.number}`), W - M, y + 7, { align: 'right' });
    doc.setFont('helvetica', 'normal'); doc.setTextColor(...grey);
    doc.text(pdfText(`Date : ${fdate(inv.issue_date || today())}`), W - M, y + 12, { align: 'right' });
    y = Math.max(y + 8 + sellerLines.length * 4.3, 48);
    // client
    doc.setDrawColor(...gold); doc.setLineWidth(0.4); doc.line(M, y, W - M, y); y += 7;
    doc.setFontSize(8); doc.setTextColor(...gold); doc.setFont('helvetica', 'bold');
    doc.text('FACTURÉ À', M, y); doc.text('CHARTER', W / 2 + 5, y); y += 5;
    doc.setFont('helvetica', 'normal'); doc.setFontSize(10); doc.setTextColor(20, 34, 51);
    const buyerLines = [b.name || '—', ...String(b.address || '').split('\n'), b.vat_number && `TVA : ${b.vat_number}`].filter(Boolean).map(pdfText);
    buyerLines.forEach((l, i) => doc.text(l, M, y + i * 4.6));
    const chLines = doc.splitTextToSize(pdfText(b.charter || ''), W / 2 - M - 5);
    chLines.forEach((l, i) => doc.text(l, W / 2 + 5, y + i * 4.6));
    doc.setTextColor(...grey); doc.setFontSize(9);
    doc.text(pdfText(`Du ${fdate(b.start_date)} au ${fdate(b.end_date)}`), W / 2 + 5, y + chLines.length * 4.6);
    y += Math.max(buyerLines.length, chLines.length + 1) * 4.6 + 8;
    // tableau
    const cols = [{ x: M, w: 58, h: 'Désignation' }, { x: M + 58, w: 20, h: 'Qté', r: 1 }, { x: M + 78, w: 20, h: 'PU HT', r: 1 },
      { x: M + 98, w: 26, h: 'Total HT', r: 1 }, { x: M + 124, w: 26, h: 'TVA', r: 1 }, { x: M + 150, w: 28, h: 'Total TTC', r: 1 }];
    doc.setFillColor(...navy); doc.rect(M, y - 5, W - 2 * M, 8, 'F');
    doc.setFont('helvetica', 'bold'); doc.setFontSize(8.5); doc.setTextColor(255, 255, 255);
    cols.forEach(c => doc.text(c.h, c.r ? c.x + c.w - 2 : c.x + 2, y, { align: c.r ? 'right' : 'left' }));
    y += 8; doc.setFont('helvetica', 'normal'); doc.setTextColor(20, 34, 51);
    (inv.lines || []).forEach(l => {
      const d = doc.splitTextToSize(pdfText(l.designation), cols[0].w - 4);
      if (y + d.length * 4.4 > 250) { doc.addPage(); y = 20; }
      doc.text(d, cols[0].x + 2, y);
      doc.text(pdfText(`${fmt(l.qty, l.unit === 'L' ? 0 : (l.qty % 1 ? 2 : 0))}${l.unit ? ' ' + l.unit : ''}`), cols[1].x + cols[1].w - 2, y, { align: 'right' });
      doc.setFontSize(8.5);
      doc.text(pdfText(fmt(l.unit_price_ht, l.unit === 'L' ? 4 : 2)), cols[2].x + cols[2].w - 2, y, { align: 'right' });
      doc.text(pdfMoney(l.total_ht, cur), cols[3].x + cols[3].w - 2, y, { align: 'right' });
      doc.text(pdfMoney(l.total_vat, cur), cols[4].x + cols[4].w - 2, y, { align: 'right' });
      doc.setFontSize(7); doc.setTextColor(...grey);
      doc.text(pdfText(l.total_vat ? `${fmt(l.rate, l.rate % 1 ? 1 : 0)} %` : 'payé HT'), cols[4].x + cols[4].w - 2, y + 3.6, { align: 'right' });
      doc.setFontSize(8.5); doc.setTextColor(20, 34, 51);
      doc.text(pdfMoney(l.total_ttc, cur), cols[5].x + cols[5].w - 2, y, { align: 'right' });
      y += Math.max(d.length * 4.4, 7.4) + 2.5;
      doc.setDrawColor(228, 222, 211); doc.setLineWidth(0.2); doc.line(M, y - 3, W - M, y - 3);
    });
    // totaux
    y += 3;
    const tot = [['Total HT', inv.total_ht], ['TVA', inv.total_vat], ['Total TTC', inv.total_ttc]];
    tot.forEach(([k, v], i) => {
      const last = i === tot.length - 1;
      doc.setFont('helvetica', last ? 'bold' : 'normal'); doc.setFontSize(last ? 11.5 : 9.5);
      doc.setTextColor(...(last ? navy : [20, 34, 51]));
      doc.text(k, W - M - 50, y); doc.text(pdfMoney(v, cur), W - M, y, { align: 'right' }); y += last ? 7 : 5.5;
    });
    // mentions
    y += 6; doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.setTextColor(...grey);
    const mentions = [n.payment_terms, n.vat_mention,
      (s.iban || s.bic) && `Coordonnées bancaires : ${[s.bank, s.iban && `IBAN ${s.iban}`, s.bic && `BIC ${s.bic}`].filter(Boolean).join(' · ')}`,
      n.footer].filter(Boolean);
    mentions.flatMap(x => String(x).split('\n')).forEach(txt => { const l = doc.splitTextToSize(pdfText(txt), W - 2 * M); if (y + l.length * 4 > 285) { doc.addPage(); y = 20; } doc.text(l, M, y); y += l.length * 4 + 2; });
    return doc;
  }

  // Enregistrement d'un fichier : téléchargement direct sur ordinateur ; sur iPhone/iPad (surtout l'app
  // installée sur l'écran d'accueil, où iOS bloque les téléchargements), fenêtre « Ouvrir / Enregistrer ».
  const IS_IOS = /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  function saveBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    setTimeout(() => URL.revokeObjectURL(url), 10 * 60 * 1000);
    if (!IS_IOS) {
      const a = document.createElement('a'); a.href = url; a.download = name;
      document.body.appendChild(a); a.click(); a.remove(); return;
    }
    const file = new File([blob], name, { type: blob.type || 'application/octet-stream' });
    const canShare = !!(navigator.canShare && navigator.canShare({ files: [file] }));
    const m = openModal({
      title: LANG === 'fr' ? 'Fichier prêt' : 'File ready',
      body: `<p class="center"><b>${esc(name)}</b></p><p class="muted small center">${LANG === 'fr'
        ? 'Ouvrez-le pour le consulter, ou utilisez « Enregistrer / Partager » pour l’enregistrer dans Fichiers, l’envoyer par e-mail…'
        : 'Open it to view, or use “Save / Share” to save it to Files, send it by e-mail…'}</p>`,
      footer: `<a class="btn btn-ghost" href="${url}" target="_blank" rel="noopener" download="${esc(name)}">${ICON.eye}${LANG === 'fr' ? 'Ouvrir' : 'Open'}</a>
        ${canShare ? `<button class="btn btn-primary" data-share>${ICON.share}${LANG === 'fr' ? 'Enregistrer / Partager' : 'Save / Share'}</button>` : ''}`,
    });
    const sh = $('[data-share]', m);
    if (sh) sh.onclick = () => navigator.share({ files: [file], title: name }).then(closeModal).catch(() => {});
  }

  async function downloadInvoice(inv, draft = false) {
    try {
      const doc = await invoicePdf(inv, draft);
      saveBlob(doc.output('blob'), draft ? `Facture_brouillon_${safeName(inv.buyer?.charter || '')}.pdf` : `Facture_${inv.number}.pdf`);
    } catch (e) { toast(errMsg(e), 'err'); }
  }
  const invoiceLink = inv => { INV_CACHE[inv.id] = inv; return `<button type="button" class="file-link link-plain" data-inv="${esc(inv.id)}">${ICON.file}<span>${t('boatInvoice')} ${esc(inv.number)}</span></button>`; };
  document.addEventListener('click', e => {
    const b = e.target.closest('[data-inv]'); if (!b) return;
    e.preventDefault(); const inv = INV_CACHE[b.dataset.inv]; if (inv) downloadInvoice(inv);
  });

  // ------------------------------------------------------------------
  //  Analyse d'une facture fournisseur : mentions de facturation HT (PDF texte)
  //  et lecture ligne par ligne des factures électroniques italiennes (FatturaPA XML)
  // ------------------------------------------------------------------
  const HT_PATTERNS = [
    [/art(?:icolo|\.)?\s*7\s*[-‐ ]?\s*(?:bis|ter|quater|quinquies|sexies|septies)\b/i, m => m],
    [/\bart(?:icolo|\.)?\s*7\b(?![,.]?\d)/i, () => 'art. 7'],
    [/non\s+imponibil\w*/i, m => m], [/fuori\s+campo\s+(?:di\s+applicazione\s+)?(?:dell['’]\s*)?iva/i, m => m],
    [/inversione\s+contabile/i, m => m], [/reverse\s*[-‐ ]?\s*charge/i, m => m], [/autoliquidation/i, m => m],
    [/esent[ei]\s+(?:da\s+)?iva|esenzione\s+iva/i, m => m], [/natura\s*:?\s*N[2-6](?:\.\d)?/i, m => m],
    [/iva\s*0\s*%|vat\s*0\s*%|tva\s*0\s*%/i, m => m], [/exon[ée]ration\s+de\s+tva|exon[ée]r[ée]e?\s+de\s+tva/i, m => m],
    [/art(?:icle|\.)?\s*(?:262|283)\b/i, m => m], [/vat\s+exempt\w*|exempt\s+from\s+vat/i, m => m],
    [/operaci[oó]n\s+exenta|inversi[oó]n\s+del\s+sujeto\s+pasivo/i, m => m],
    [/(?:operazione\s+)?non\s+soggett\w*\s+(?:ad\s+)?iva/i, m => m], [/senza\s+(?:applicazione\s+(?:dell['’]\s*)?)?iva/i, m => m],
    [/\bOp\.?\s*non\s+imp\w*/i, m => m],
  ];
  function findHtMentions(text) {
    const out = new Set();
    HT_PATTERNS.forEach(([re, label]) => { const m = text.match(re); if (m) out.add(label(m[0].replace(/\s+/g, ' ').trim())); });
    return [...out];
  }
  async function pdfExtractText(file) {
    await loadScript(PDFJS_JS);
    const lib = window.pdfjsLib; lib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER;
    const pdf = await lib.getDocument({ data: await file.arrayBuffer() }).promise;
    const lines = [];
    for (let i = 1; i <= Math.min(pdf.numPages, 6); i++) {
      const page = await pdf.getPage(i); const c = await page.getTextContent();
      const rows = [];
      c.items.filter(it => it.str && it.str.trim()).forEach(it => {
        const y = it.transform[5], x = it.transform[4], h = Math.abs(it.transform[3]) || 8;
        let row = rows.find(r => Math.abs(r.y - y) < Math.max(2.5, h * 0.45));
        if (!row) { row = { y, items: [] }; rows.push(row); }
        row.items.push({ x, s: it.str });
      });
      rows.sort((a, b) => b.y - a.y).forEach(r => lines.push(r.items.sort((a, b) => a.x - b.x).map(i => i.s).join(' ').replace(/\s+/g, ' ').trim()));
    }
    return lines.join('\n');
  }

  // ---- Lecture des montants d'une facture (texte PDF ou photo) : total, HT, TVA, part non taxée
  const AMOUNT_RE = /(?:€|eur\.?)?\s*(-?\d{1,3}(?:[ .\u00a0\u202f]\d{3})+(?:,\d{2})|-?\d{1,3}(?:,\d{3})+\.\d{2}|-?\d+[,.]\d{2})(?!\d|\s?%)/gi;
  function parseAmount(s) {
    s = s.replace(/[\s\u00a0\u202f]/g, '');
    if (/,\d{2}$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
    else s = s.replace(/,/g, '');
    const v = parseFloat(s); return isNaN(v) ? null : v;
  }
  const amountsIn = line => [...line.matchAll(AMOUNT_RE)].map(m => parseAmount(m[1])).filter(v => v != null);
  function parseInvoiceAmounts(text) {
    const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
    const norm = s => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
    // montant sur la ligne qui porte le libellé (et éventuellement la suivante si la ligne ne contient QUE le libellé)
    const pick = (re, exclude, allowNext = true, firstAfter = false) => {
      for (let i = lines.length - 1; i >= 0; i--) {
        const l = norm(lines[i]); const mt = l.match(re); if (!mt || (exclude && exclude.test(l))) continue;
        const after = amountsIn(lines[i].slice(mt.index)); if (after.length) return firstAfter ? after[0] : after[after.length - 1];
        if (allowNext && l.replace(re, '').replace(/[^a-z0-9]/g, '').length < 4) {
          const b = lines[i + 1] ? amountsIn(lines[i + 1]) : []; if (b.length) return b[b.length - 1];
        }
      }
      return null;
    };
    const POINTS = /point|fidelit|bonus|consigne|colis|poids|volume|page/;
    const lTotal = pick(/total(e)?\s*(documento|fattura|da pagare|a pagare|a payer|ttc|eur|general|importo)|net a payer|montant ttc|importo totale|total (due|amount|to pay)|amount due|(^|[^a-z])total(e)?\b(?!.*(ht|h\.t|imponibile|iva|tva|t\.v\.a|vat|excl|net|hors))/, POINTS);
    const lVat = pick(/(totale\s+|total\s+|montant\s+)?(iva|tva|t\.v\.a\.?|vat|imposta)\b(?!.*(esclus|exclu|non\s+imp|art\.?\s*7|intra|n°|nº|no\.))/, /imponibile\s+\d|base|intra|taux|rate|aliquota|collectee|client/, false, true);
    const lHt = pick(/total(e)?\s*(ht|h\.t\.?|imponibile|hors taxe|excl|net)|imponibile|montant ht|montant hors|net amount|subtotal|base imponible/, /iva|tva|vat|imposta|taux/);
    // htBase : récapitulatif (« non imponibile art. … 1.700,00 ») en priorité, sinon somme des lignes codées N2…N6
    let sumMention = 0, sumN = 0;
    for (const l of lines) {
      const a = amountsIn(l); if (!a.length) continue;
      if (findHtMentions(l).length) sumMention = r2(sumMention + a[a.length - 1]);
      else if (/\bN[2-6](\.\d)?\b/.test(l)) sumN = r2(sumN + a[a.length - 1]);
    }
    const htBase = sumMention || sumN || null;
    // Cohérence : le total est le plus grand montant ; on cherche HT + TVA = TOTAL parmi les montants de la facture
    const all = [...new Set(lines.flatMap(amountsIn).filter(v => v > 0).map(v => r2(v)))];
    const max = all.length ? Math.max(...all) : null;
    const T = lTotal && lTotal > 0 ? lTotal : max;
    let best = null;
    if (T) {
      const has = new Set(all.map(v => Math.round(v * 100)));
      for (const v of all) {
        if (v <= 0 || v >= T / 2) continue;
        const h = r2(T - v); if (!has.has(Math.round(h * 100))) continue;
        const sameLine = lines.some(l => { const a = amountsIn(l).map(r2); return a.includes(h) && a.includes(v); });
        const withT = lines.some(l => { const a = amountsIn(l).map(r2); return a.includes(h) && a.includes(v) && a.includes(r2(T)); });
        const score = (withT ? 4 : 0) + (sameLine ? 2 : 0) + (lVat != null && r2(lVat) === v ? 3 : 0);
        if (!best || score > best.score) best = { h, v, score };
      }
    }
    let totalF = T ?? null, vat, ht;
    if (best && (best.score > 0 || lVat == null)) { vat = best.v; ht = best.h; }
    else { vat = lVat; ht = lHt ?? (totalF != null && vat != null ? r2(totalF - vat) : null); }
    if (ht != null && totalF != null && (ht <= 0 || ht > totalF)) ht = null;
    if (vat != null && totalF != null && (vat < 0 || vat >= totalF)) vat = null;
    return { total: totalF, vat, ht, htBase };
  }
  function parseFatturaPA(xmlText) {
    const doc = new DOMParser().parseFromString(xmlText, 'application/xml');
    const all = n => Array.from(doc.getElementsByTagNameNS('*', n));
    const one = (n, el = doc) => { const x = el.getElementsByTagNameNS('*', n)[0]; return x ? x.textContent.trim() : null; };
    if (!all('FatturaElettronicaBody').length) return null;
    const ced = doc.getElementsByTagNameNS('*', 'CedentePrestatore')[0];
    const company = ced ? (one('Denominazione', ced) || [one('Nome', ced), one('Cognome', ced)].filter(Boolean).join(' ')) : null;
    const riep = all('DatiRiepilogo').map(r => ({ rate: toNum(one('AliquotaIVA', r)) || 0, natura: one('Natura', r),
      base: toNum(one('ImponibileImporto', r)) || 0, tax: toNum(one('Imposta', r)) || 0 }));
    const lines = all('DettaglioLinee').map(l => ({ desc: one('Descrizione', l), amount: toNum(one('PrezzoTotale', l)) || 0,
      qty: toNum(one('Quantita', l)), unitPrice: toNum(one('PrezzoUnitario', l)),
      rate: toNum(one('AliquotaIVA', l)) || 0, natura: one('Natura', l) }));
    const htLines = lines.filter(l => !l.rate);
    const htBase = r2(riep.length ? sum(riep.filter(r => !r.rate), r => r.base) : sum(htLines, l => l.amount));
    const ht = r2(riep.length ? sum(riep, r => r.base) : sum(lines, l => l.amount));
    const vat = r2(sum(riep, r => r.tax));
    const total = toNum(one('ImportoTotaleDocumento')) ?? r2(ht + vat);
    const naturas = [...new Set([...riep, ...lines].map(x => x.natura).filter(Boolean))];
    return { exact: true, htBase, lines, mentions: naturas.map(n => `Natura ${n}`),
      fill: { company, invoice_ref: one('Numero'), date: one('Data'), amount_ht: ht, vat, total },
      summary: `Facture électronique italienne : ${lines.length} ligne(s), dont <b>${htLines.length} sans TVA</b> (${money(htBase)}) et ${lines.length - htLines.length} avec TVA. Total ${money(total)}.` };
  }
  // Lecture de texte sur photo / PDF scanné (OCR, dans le navigateur — quelques secondes)
  async function downscale(src, max = 2000) {
    const img = src instanceof HTMLCanvasElement ? src : await createImageBitmap(src);
    const k = Math.min(1, max / Math.max(img.width, img.height));
    if (k === 1 && src instanceof HTMLCanvasElement) return src;
    const cv = document.createElement('canvas'); cv.width = Math.round(img.width * k); cv.height = Math.round(img.height * k);
    cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height); return cv;
  }
  async function ocr(sources, onProgress) {
    await loadScript(TESS_JS);
    const worker = await window.Tesseract.createWorker(['ita', 'eng'], 1, { ...TESS_OPTS,
      logger: m => { if (onProgress && m.status === 'recognizing text') onProgress(m.progress); } });
    let txt = '';
    try { for (const s of sources) txt += (await worker.recognize(await downscale(s))).data.text + '\n'; }
    finally { await worker.terminate(); }
    return txt;
  }
  async function pdfPagesAsCanvas(file, maxPages = 2) {
    await loadScript(PDFJS_JS);
    const lib = window.pdfjsLib; lib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER;
    const pdf = await lib.getDocument({ data: await file.arrayBuffer() }).promise, out = [];
    for (let i = 1; i <= Math.min(pdf.numPages, maxPages); i++) {
      const page = await pdf.getPage(i), vp = page.getViewport({ scale: 2 });
      const cv = document.createElement('canvas'); cv.width = vp.width; cv.height = vp.height;
      await page.render({ canvasContext: cv.getContext('2d'), viewport: vp }).promise; out.push(cv);
    }
    return out;
  }
  // Résultat commun PDF / photo : mentions + montants lus
  function textResult(txt, source) {
    const m = findHtMentions(txt), a = parseInvoiceAmounts(txt);
    const isHt = m.length > 0;
    let htBase = null;
    if (isHt) {
      if (a.htBase) htBase = a.htBase;                                    // ligne « non imponibile … 1.700,00 »
      else if (a.total != null && (!a.vat || a.vat === 0)) htBase = a.total; // toute la facture sans TVA
      else if (a.ht != null && a.vat === 0) htBase = a.ht;
    }
    const cap = a.total != null ? r2(a.total - (a.vat || 0)) : null;   // la part HT ne peut dépasser total − TVA
    if (htBase != null && cap != null && htBase > cap) htBase = cap;
    const fill = { total: a.total, vat: a.vat, amount_ht: a.ht };
    const found = [a.total != null && `Total ${money(a.total)}`, a.ht != null && `HT ${money(a.ht)}`, a.vat != null && `TVA ${money(a.vat)}`].filter(Boolean).join(' · ');
    return { mentions: m, htBase, fill, exact: !!(isHt && a.htBase && htBase === a.htBase), amountsFound: !!found,
      summary: `${source}${found ? ` — montants lus : ${found}` : ' — montants non trouvés, saisissez-les'}${isHt ? (htBase != null ? `. <b>Part payée HT : ${money(htBase)}</b>` : '. Indiquez la part payée HT') : '. Aucune mention de facturation HT'}.` };
  }
  const ocrResult = txt => textResult(txt, 'Document lu (photo / scan)');

  async function analyseInvoiceFile(file, onProgress) {
    const name = file.name.toLowerCase();
    if (/^image\//.test(file.type) || /\.(jpe?g|png|heic|webp)$/.test(name)) return ocrResult(await ocr([file], onProgress));
    if (name.endsWith('.xml') || /xml/.test(file.type)) {
      const r = parseFatturaPA(await file.text());
      if (r) return r;
      const m = findHtMentions(await file.text()); return m.length ? { mentions: m, summary: 'Mention de facturation HT détectée dans le fichier.' } : null;
    }
    if (name.endsWith('.pdf') || file.type === 'application/pdf') {
      const txt = await pdfExtractText(file);
      if (txt.replace(/\s/g, '').length < 20) return ocrResult(await ocr(await pdfPagesAsCanvas(file), onProgress));
      return textResult(txt, 'Facture PDF lue');
    }
    return null;
  }

  // ------------------------------------------------------------------
  //  Export Excel
  // ------------------------------------------------------------------
  async function exportCharterXlsx(B, asArray = false) {
    try { await loadScript(XLSX_JS); } catch (e) { toast(errMsg(e), 'err'); return null; }
    const st = computeStatement(B), c = B.charter;
    const rows = [[c.yacht || YACHT], [c.title], [c.client_name ? `Client : ${c.client_name}` : ''],
      [`${fdate(c.start_date)} → ${fdate(c.end_date)}`], [],
      ['Number', 'Date', 'Company', 'Invoice', 'Category', 'Description', 'Excl taxes', 'VAT', 'Total', 'Method']];
    if (st.fuelLines.length) rows.push([1, fdate(c.end_date), c.fuel_company, '', 'Fuel', `${c.fuel_description} — ${st.fuelLiters} L`, '', '', st.fuelTotal, c.fuel_method]);
    B.expenses.forEach(e => rows.push([e.num, fdate(e.date), e.company, e.invoice_ref, e.category, e.description, e.amount_ht, e.vat, toNum(e.total), e.method]));
    if (st.tipsAuto) rows.push(['', fdate((c.client_signed_at || '').slice(0, 10)), 'Crew', '', 'TIPS', `Crew gratuity${st.tipsAutoPct != null ? ` ${st.tipsAutoPct}% of charter fee` : ''}`, '', '', st.tipsAuto, 'APA']);
    rows.push([], ['', '', '', '', '', 'TOTAL', '', '', st.total], [], ['', '', 'APA paid', st.apa], ['', '', 'APA Left', st.balance]);
    if (st.vatEst) {
      rows.push(['', '', `VAT estimated (${charterVatRate(c)}%${c.vat_country ? ' ' + c.vat_country : ''})`, st.vatEst], ['', '', 'APA Left including VAT estimated', st.balanceInclVat]);
      rows.push([], ['VAT estimated detail'], ['Number', 'Company', 'Description', 'Paid excl. VAT', 'VAT estimated']);
      st.fuelLines.filter(f => f.vatEst).forEach(f => rows.push(['Fuel', `Bunkering #${f.bunker_num ?? ''}`, `${f.liters} L paid excl. VAT`, f.amount, f.vatEst]));
      B.expenses.filter(e => vatEstOf(e, c)).forEach(e => rows.push([e.num, e.company, e.description, toNum(e.vat_est_base), vatEstOf(e, c)]));
    }
    if (c.charter_fee) rows.push(['', '', 'Charter fee', toNum(c.charter_fee)]);
    if (B.voyages.length) {
      rows.push([], ['Fuel'], ['Voyage', 'ROB before (L)', 'ROB after (L)', 'Consumption (L)']);
      B.voyages.forEach(v => rows.push([v.description, v.rob_before, v.rob_after, consumption(v)]));
      rows.push([], ['Bunkering', 'Date', 'Port', 'Fuel type', 'Liters', 'Price/L excl. VAT', 'VAT/L', 'Price/L billed', 'Paid', 'Total', 'VAT estimated']);
      st.fuelLines.forEach(f => rows.push([f.bunker_num, fdate(f.bunker_date), f.town, f.fuel_type, f.liters, toNum(f.price_ht_l), toNum(f.vat_l) || 0, f.price, f.paidHt ? 'excl. VAT' : 'incl. VAT', f.amount, f.vatEst || '']));
    }
    if (c.client_signed_at) rows.push([], [`Signed by ${c.client_signed_name} on ${fdatetime(c.client_signed_at)}`]);
    const ws = window.XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{ wch: 8 }, { wch: 11 }, { wch: 22 }, { wch: 16 }, { wch: 14 }, { wch: 40 }, { wch: 12 }, { wch: 10 }, { wch: 12 }, { wch: 22 }];
    const wb = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, ws, 'Accounting');
    if (asArray) return window.XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    saveBlob(new Blob([window.XLSX.write(wb, { type: 'array', bookType: 'xlsx' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `${safeName(c.title || 'charter')}.xlsx`);
    return null;
  }

  const safeName = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\w\-. ]+/g, '_').replace(/\s+/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '').slice(0, 80);

  // Rapport complet en .zip : Excel de la compta + toutes les factures (dépenses + soutages du charter)
  async function downloadReportZip(B, btn) {
    busy(btn, true);
    try {
      await loadScript(JSZIP_JS);
      const zip = new window.JSZip(), c = B.charter, st = computeStatement(B);
      const xlsx = await exportCharterXlsx(B, true);
      if (xlsx) zip.file(`${safeName(c.title)}.xlsx`, xlsx);
      const items = [];
      B.expenses.filter(e => e.file_path).forEach(e => items.push({ path: e.file_path,
        name: `Factures/${String(e.num ?? '').padStart(2, '0')}_${safeName(e.company || 'facture')}_${safeName(e.file_name || e.file_path.split('/').pop())}` }));
      const seen = new Set();
      st.fuelLines.filter(f => f.file_path && !seen.has(f.file_path) && seen.add(f.file_path)).forEach(f => items.push({ path: f.file_path,
        name: `Fuel/Soutage_${f.bunker_num ?? ''}_${safeName(f.file_name || f.file_path.split('/').pop())}` }));
      for (const inv of (B.invoices || [])) {
        try { const doc = await invoicePdf(inv); zip.file(`Facture_bateau_${inv.number}.pdf`, doc.output('arraybuffer')); } catch { /* */ }
      }
      let missing = 0;
      for (const it of items) {
        try { const url = api.fileUrl(it.path); const r = await fetch(url); if (!r.ok) throw 0; zip.file(it.name, await r.blob()); }
        catch { missing++; }
      }
      const blob = await zip.generateAsync({ type: 'blob' });
      saveBlob(blob, `${safeName(c.title)}_rapport.zip`);
      if (missing) toast(`${missing} facture(s) n’ont pas pu être ajoutées`, 'err');
    } catch (e) { toast(errMsg(e), 'err'); }
    finally { busy(btn, false); }
  }

  async function exportFuelXlsx(S) {
    try { await loadScript(XLSX_JS); } catch (e) { toast(errMsg(e), 'err'); return; }
    const X = window.XLSX, used = lotUsage(S), wb = X.utils.book_new();
    const bl = S.bunkers.slice().sort(byLot);
    X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([
      ['Number', 'Date of Bunkering', 'Country', 'Town', 'Quantity (L)', 'Type of Fuel', 'Price HT / L', 'VAT / L', 'Total Price / L', 'Total Bunkering Price', 'Invoice', 'Remaining (L)'],
      ...bl.map(b => [b.num, fdate(b.date), b.country, b.town, b.quantity, b.fuel_type, b.price_ht_l, b.vat_l, r3(ttcL(b)),
        b.total_price ?? r2(b.quantity * ttcL(b)), b.invoice_ref, r2(b.quantity - (used[b.id] || 0))])]), 'Bunkering');
    const lotNum = id => S.bunkers.find(b => b.id === id)?.num;
    X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([
      ['Voyage', 'Description', 'Start Date', 'End Date', 'ROB before (L)', 'ROB after (L)', 'Consumption (L)', 'Distribution (L)', 'Fuel used (lot)', 'Hours', 'L/h'],
      ...S.voyages.slice().sort(byPos).flatMap(v => {
        const al = S.allocations.filter(a => a.voyage_id === v.id);
        const base = [v.num, v.description, fdate(v.start_date), fdate(v.end_date), v.rob_before, v.rob_after, consumption(v)];
        const lh = v.hours ? r2(consumption(v) / v.hours) : '';
        if (!al.length) return [[...base, '', '', v.hours, lh]];
        return al.map((a, i) => i === 0 ? [...base, a.liters, lotNum(a.bunker_id), v.hours, lh] : ['', '', '', '', '', '', '', a.liters, lotNum(a.bunker_id)]);
      })]), 'Voyages');
    saveBlob(new Blob([X.write(wb, { type: 'array', bookType: 'xlsx' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `Fuel_booklet_${today()}.xlsx`);
  }

  // ==================================================================
  //  VUE CLIENT (lien secret ?c=…)
  // ==================================================================
  async function runClient(token, mode = 'client') {
    const REPORT = mode === 'report';
    LANG = store.get('ca-lang') || ((navigator.language || 'en').toLowerCase().startsWith('fr') ? 'fr' : 'en');
    document.documentElement.lang = LANG;
    let B = null, lastAt = null, inflight = false, prevStatus = null;

    async function refresh() {
      if (inflight) return; inflight = true;
      try {
        B = REPORT ? await api.reportView(token) : await api.clientView(token); lastAt = new Date();
        if (!$('#modal-root').innerHTML) render();
        if (!REPORT) maybeAnnounceClosed();
      } catch (e) { if (!B) renderError(errMsg(e)); }
      finally { inflight = false; }
    }

    // Message « compta clôturée » : à l'ouverture du lien, ou en direct dès que le capitaine clôture
    function maybeAnnounceClosed() {
      if (!B) return;
      const c = B.charter, key = 'ca-closed-seen-' + token + '-' + (c.closed_at || '');
      const justClosed = prevStatus && prevStatus !== 'closed' && c.status === 'closed';
      prevStatus = c.status;
      if (c.status !== 'closed' || $('#modal-root').innerHTML) return;
      if (!justClosed && store.get(key)) return;
      store.set(key, '1');
      const m = openModal({
        title: t('closedPopupTitle'),
        body: `<div class="closed-popup"><div class="closed-icon">${ICON.check}</div><p>${t('closedPopupText')}</p>
          <div class="sign-summary"><div><span>${t('totalExpenses')}</span><b>${money(computeStatement(B).total, c.currency)}</b></div>
          <div><span>${t('apaBalance')}</span><b>${money(computeStatement(B).balance, c.currency)}</b></div></div>
          ${computeStatement(B).vatEst ? `<div class="sign-summary vat-summary"><div><span>${t('vatEstimated')}</span><b>${money(computeStatement(B).vatEst, c.currency)}</b></div>
            <div><span>${t('apaLeftInclVat')}</span><b>${money(computeStatement(B).balanceInclVat, c.currency)}</b></div></div>` : ''}</div>`,
        footer: `<button class="btn btn-ghost" data-close>${t('later')}</button><button class="btn btn-gold" data-ok>${ICON.pen}${t('signNow')}</button>`,
      });
      $('[data-ok]', m).onclick = () => { closeModal(); render(); openSign(); };
      $$('[data-close]', m).forEach(x => x.addEventListener('click', () => setTimeout(render, 0)));
      if (document.hidden) document.title = '● ' + t('closedPopupTitle');
    }

    function renderError(msg) {
      $('#app').innerHTML = `<div class="center-screen"><div class="card narrow">
        <div class="brand-mark">${ICON.anchor}</div><h2>${t('invalidLink')}</h2><p class="muted">${esc(msg || t('invalidLinkHelp'))}</p></div></div>`;
    }

    function render() {
      if (!B) return renderError(REPORT ? t('reportUnavailable') : t('invalidLinkHelp'));
      const c = B.charter, st = computeStatement(B), signed = c.status === 'signed', closed = c.status === 'closed';
      document.title = `${c.title} — ${c.yacht || YACHT}`;
      $('#app').innerHTML = `
      <header class="topbar topbar-client">
        <div class="topbar-inner">
          <div class="brand"><span class="brand-mark sm">${ICON.anchor}</span><span class="brand-name">${esc(c.yacht || YACHT)}</span></div>
          <div class="topbar-right">
            <span class="live"><span class="live-dot"></span>${t('live')} · ${t('updated')} ${lastAt ? lastAt.toLocaleTimeString(locale(), { hour: '2-digit', minute: '2-digit' }) : ''}</span>
            <div class="lang"><button data-lang="fr" class="${LANG === 'fr' ? 'on' : ''}">FR</button><button data-lang="en" class="${LANG === 'en' ? 'on' : ''}">EN</button></div>
          </div>
        </div>
      </header>
      <main class="container">
        <section class="hero">
          <div class="eyebrow">${REPORT ? t('reportTitle') : t('statement')}</div>
          <h1>${esc(c.title)}</h1>
          <p class="muted">${esc(c.client_name || '')}${c.client_name ? ' · ' : ''}${fdate(c.start_date)} → ${fdate(c.end_date)}</p>
          <div class="hero-actions no-print">
            <button class="btn btn-ghost" data-act="print">${ICON.print}${t('print')}</button>
            <button class="btn btn-ghost" data-act="xlsx">${ICON.sheet}${t('excel')}</button>
            ${REPORT || signed ? `<button class="btn btn-ghost" data-act="zip">${ICON.file}${t('zipAll')}</button>` : ''}
            ${closed ? `<button class="btn btn-gold" data-act="sign">${ICON.pen}${t('signCta')}</button>` : ''}
          </div>
        </section>
        ${signed ? `<div class="banner banner-ok">${ICON.check}<div><strong>${t('signedBanner')}</strong><br>
          <span>${t('signedOn')} ${fdatetime(c.client_signed_at)} ${t('by')} ${esc(c.client_signed_name)}</span></div></div>`
        : closed ? `<div class="banner banner-gold">${ICON.lock}<div><strong>${t('closedBanner')}</strong></div>
            <button class="btn btn-gold btn-sm no-print" data-act="sign">${t('signNow')}</button></div>`
        : `<div class="banner banner-info">${ICON.anchor}<div>${t('openBanner')}</div></div>`}
        ${renderKpis(B, st)}
        <section class="panel"><div class="panel-head"><h2>${t('expenses')}</h2></div>${renderExpensesTable(B, st)}</section>
        <div class="grid-2">
          ${B.voyages.length ? `<section class="panel"><div class="panel-head"><h2>${t('fuelDetail')}</h2></div>${renderFuelDetail(B, st)}</section>` : ''}
          <section class="panel"><div class="panel-head"><h2>${t('byCategory')}</h2></div>${renderCategories(B, st)}
            ${B.cash.length || st.cashSpent ? `<div class="cash-sum"><h3>${t('cash')}</h3>
              <div class="kv"><span>${t('cashWithdrawn')}</span><b>${money(st.cashWithdrawn, c.currency)}</b></div>
              <div class="kv"><span>${t('cashSpent')}</span><b>${money(st.cashSpent, c.currency)}</b></div>
              <div class="kv"><span>${t('cashInHand')}</span><b>${money(st.cashInHand, c.currency)}</b></div></div>` : ''}
          </section>
        </div>
        <section class="panel"><div class="panel-head"><h2>${t('signatures')}</h2></div>
          <div class="sig-grid">
            ${renderSignatureBox(t('charterer'), c.client_signature, c.client_signed_name, c.client_signed_at)}
            ${renderSignatureBox(t('master'), c.master_signature, c.master_signed_name, c.master_signed_at)}
          </div>
          ${closed ? `<div class="center no-print"><button class="btn btn-gold btn-lg" data-act="sign">${ICON.pen}${t('signCta')}</button></div>` : ''}
        </section>
        <footer class="foot muted small">${esc(c.yacht || YACHT)} · ${t('statement')} · ${new Date().toLocaleDateString(locale())}</footer>
      </main>`;
    }

    // Signature en 2 étapes : 1) tips  2) signature
    function openSign() {
      if (!B || B.charter.status !== 'closed') return;
      const c = B.charter, cur = c.currency, fee = toNum(c.charter_fee), withTips = tipsEnabled(c);
      let choice = fee ? '10' : 'custom', customVal = '', customUnit = fee ? 'pct' : 'amt';
      const tipsNow = () => {
        if (choice !== 'custom') return { pct: +choice, amount: r2(fee * +choice / 100) };
        const v = toNum(customVal);
        if (v === null || v < 0) return { pct: null, amount: null };
        return customUnit === 'pct' && fee ? { pct: v, amount: r2(fee * v / 100) } : { pct: fee ? r2(v / fee * 100) : null, amount: r2(v) };
      };

      function stepTips() {
        const m = openModal({
          title: t('tipsTitle'), wide: true,
          body: `<p class="step-label">${t('step')} 1 / 2</p>
            <p>${fee ? `${t('tipsIntro')} (${t('charterFee')} : <b>${money(fee, cur)}</b>).` : t('tipsIntroNoFee')}</p>
            <div class="tips-grid">
              ${fee ? ['10', '15', '20'].map(p => `<button type="button" class="tip-opt ${choice === p ? 'on' : ''}" data-tip="${p}">
                <span class="tip-pct">${p} %</span><span class="tip-amt">${money(fee * p / 100, cur)}</span></button>`).join('') : ''}
              <button type="button" class="tip-opt ${choice === 'custom' ? 'on' : ''}" data-tip="custom"><span class="tip-pct">${t('tipsCustom')}</span>
                <span class="tip-amt">${t('tipsCustomHint')}</span></button>
            </div>
            <div class="tip-custom ${choice === 'custom' ? '' : 'hidden'}">
              <input type="number" min="0" step="0.01" inputmode="decimal" name="custom" value="${esc(customVal)}" placeholder="0">
              ${fee ? `<div class="seg"><button type="button" data-unit="pct" class="${customUnit === 'pct' ? 'on' : ''}">%</button>
                <button type="button" data-unit="amt" class="${customUnit === 'amt' ? 'on' : ''}">${esc(cur)}</button></div>` : `<span class="muted">${esc(cur)}</span>`}
            </div>
            <div class="tip-result" id="tip-result"></div>`,
          footer: `<button class="btn btn-ghost" data-close>${t('cancel')}</button><button class="btn btn-gold" data-ok>${t('continue')}</button>`,
        });
        const update = () => {
          const tp = tipsNow(), el = $('#tip-result', m);
          el.innerHTML = tp.amount === null ? '' : `${t('tipsChosen')} : <b>${money(tp.amount, cur)}</b>${tp.pct != null && fee ? ` <span class="muted">(${fmt(tp.pct, tp.pct % 1 ? 2 : 0)} % ${t('ofCharterFee')})</span>` : ''}`;
        };
        $$('[data-tip]', m).forEach(btn => btn.onclick = () => {
          choice = btn.dataset.tip;
          $$('[data-tip]', m).forEach(x => x.classList.toggle('on', x === btn));
          $('.tip-custom', m).classList.toggle('hidden', choice !== 'custom');
          if (choice === 'custom') $('[name=custom]', m).focus();
          update();
        });
        $$('[data-unit]', m).forEach(btn => btn.onclick = () => {
          customUnit = btn.dataset.unit; $$('[data-unit]', m).forEach(x => x.classList.toggle('on', x === btn)); update();
        });
        $('[name=custom]', m).addEventListener('input', e => { customVal = e.target.value; update(); });
        update();
        $('[data-ok]', m).onclick = () => {
          const tp = tipsNow();
          if (tp.amount === null) return toast(t('tipsInvalid'), 'err');
          stepSign(tp);
        };
      }

      function stepSign(tp) {
        const st = computeStatement(B, { tipsAmount: tp.amount, tipsPct: tp.pct });
        const m = openModal({
          title: t('signTitle'), wide: true,
          body: `${withTips ? `<p class="step-label">${t('step')} 2 / 2</p>` : ''}
            <div class="sign-summary ${withTips ? 'three' : ''}">
              ${withTips ? `<div><span>${t('tipsChosen')}</span><b>${money(tp.amount, cur)}</b></div>` : ''}
              <div><span>${t('totalToSign')}</span><b>${money(st.total, cur)}</b></div>
              <div><span>${t('apaLeft')}</span><b class="${st.balance < 0 ? 'neg' : ''}">${money(st.balance, cur)}</b></div></div>
            ${st.vatEst ? `<div class="sign-summary vat-summary"><div><span>${t('vatEstimated')} (${fmt(charterVatRate(c), charterVatRate(c) % 1 ? 1 : 0)} %)</span><b>${money(st.vatEst, cur)}</b></div>
              <div><span>${t('apaLeftInclVat')}</span><b class="${st.balanceInclVat < 0 ? 'neg' : ''}">${money(st.balanceInclVat, cur)}</b></div></div>
              <p class="small muted">${t('vatEstNote')}.</p>` : ''}
            <label class="field"><span>${t('yourName')}</span><input name="name" autocomplete="name" required></label>
            <div class="field"><span>${t('signHere')}</span><canvas class="sig-pad" aria-label="${t('signHere')}"></canvas>
              <button type="button" class="btn btn-ghost btn-sm" data-clear>${t('clear')}</button></div>
            <label class="check"><input type="checkbox" name="approve"> <span>${t('approve')} (${money(st.total, cur)}${st.vatEst ? ` · ${t('apaLeftInclVat')} ${money(st.balanceInclVat, cur)}` : ''}).</span></label>`,
          footer: `${withTips ? `<button class="btn btn-ghost mr-auto" data-back>${ICON.back}${t('back')}</button>` : ''}
            <button class="btn btn-ghost" data-close>${t('cancel')}</button><button class="btn btn-gold" data-ok>${ICON.pen}${t('sign')}</button>`,
        });
        const pad = signaturePad($('.sig-pad', m));
        $('[data-clear]', m).onclick = () => pad.clear();
        if (withTips) $('[data-back]', m).onclick = () => stepTips();
        $('[data-ok]', m).onclick = async ev => {
          const name = $('[name=name]', m).value.trim();
          if (!name) return toast(t('nameNeeded'), 'err');
          if (pad.isEmpty()) return toast(t('sigNeeded'), 'err');
          if (!$('[name=approve]', m).checked) return toast(t('approveNeeded'), 'err');
          busy(ev.currentTarget, true);
          try { B = await api.clientSign(token, name, pad.toDataURL(), tp.pct, tp.amount); prevStatus = 'signed'; closeModal(); render(); toast(t('signedThanks')); }
          catch (e) { toast(errMsg(e), 'err'); busy(ev.currentTarget, false); }
        };
      }
      if (withTips) stepTips(); else stepSign({ pct: null, amount: 0 });
    }

    $('#app').addEventListener('click', e => {
      const b = e.target.closest('[data-act],[data-lang]'); if (!b) return;
      if (b.dataset.lang) { LANG = b.dataset.lang; store.set('ca-lang', LANG); document.documentElement.lang = LANG; render(); return; }
      if (b.dataset.act === 'print') window.print();
      if (b.dataset.act === 'xlsx') exportCharterXlsx(B);
      if (b.dataset.act === 'zip') downloadReportZip(B, b);
      if (b.dataset.act === 'sign') openSign();
    });

    await refresh();
    if (REPORT) return;
    api.subscribe(token, () => refresh());
    setInterval(() => { if (!document.hidden && !$('#modal-root').innerHTML) refresh(); }, 30000);
    document.addEventListener('visibilitychange', () => { if (!document.hidden) { if (B) document.title = `${B.charter.title} — ${B.charter.yacht || YACHT}`; refresh(); } });
  }

  // ==================================================================
  //  ESPACE CAPITAINE
  // ==================================================================
  let S = null;          // toutes les données
  let USER = '';

  async function reload() {
    S = await api.loadAll();
    PROFILE = { ...PROFILE_DEFAULTS, ...((S.boat_profile || [])[0]?.data || {}) };
  }
  const charterById = id => S.charters.find(c => c.id === id);
  const clientUrl = c => `${location.origin}${location.pathname}?c=${c.client_token}`;

  function notifyCharter(charterId) {
    const c = charterId && charterById(charterId);
    if (c) api.notify(c.client_token);
  }

  async function runAdmin() {
    LANG = 'fr';
    const session = await api.session();
    if (!session) return renderLogin();
    try {
      if (!(await api.isAdmin())) return renderLogin('Ce compte n’est pas autorisé. Ajoutez son e-mail dans la table « admins » de Supabase.');
      USER = await api.userEmail();
      await reload();
      await autoCreateDueVoyages();
    } catch (e) { return renderLogin('Erreur de connexion à la base : ' + errMsg(e)); }
    window.addEventListener('hashchange', route);
    route();
  }

  function renderLogin(error = '') {
    $('#app').innerHTML = `<div class="center-screen login-bg"><form class="card narrow login" autocomplete="on">
      <div class="brand-mark">${ICON.anchor}</div>
      <h1>${esc(YACHT)}</h1><p class="muted">Comptabilité charter & fuel booklet</p>
      ${DEMO ? `<div class="banner banner-info small">Mode démo : l’application n’est pas encore reliée à Supabase. Les données (fictives) restent dans ce navigateur.</div>
        <button class="btn btn-primary btn-block" type="submit">Entrer en mode démo</button>`
      : `${error ? `<div class="banner banner-err small">${esc(error)}</div>` : ''}
        <label class="field"><span>E-mail</span><input name="email" type="email" autocomplete="username" required></label>
        <label class="field"><span>Mot de passe</span><input name="password" type="password" autocomplete="current-password" required></label>
        <button class="btn btn-primary btn-block" type="submit">Se connecter</button>`}
    </form></div>`;
    $('form.login').addEventListener('submit', async e => {
      e.preventDefault();
      const btn = $('button[type=submit]', e.target); busy(btn, true);
      try {
        if (DEMO) await api.signIn();
        else await api.signIn(e.target.email.value.trim(), e.target.password.value);
        await runAdmin();
      } catch (err) { busy(btn, false); renderLogin(errMsg(err)); }
    });
  }

  function shell(active, inner) {
    return `<header class="topbar">
      <div class="topbar-inner">
        <a class="brand" href="#/charters"><span class="brand-mark sm">${ICON.anchor}</span><span class="brand-name">${esc(YACHT)}</span></a>
        <nav class="nav">
          <a href="#/charters" class="${active === 'charters' ? 'on' : ''}">Charters</a>
          <a href="#/fuel/voyages" class="${active === 'fuel' ? 'on' : ''}">Fuel booklet</a>
          <a href="#/invoices" class="${active === 'invoices' ? 'on' : ''}">Factures</a>
          <a href="#/settings" class="${active === 'settings' ? 'on' : ''}">Paramètres</a>
        </nav>
        <div class="topbar-right"><span class="user small">${esc(USER)}</span>
          <button class="icon-btn light" data-act="logout" title="Se déconnecter">${ICON.logout}</button></div>
      </div>
    </header>
    ${DEMO ? `<div class="demo-strip">Mode démo — données fictives stockées dans ce navigateur. <button data-act="demo-reset">Réinitialiser</button></div>` : ''}
    <main class="container">${inner}</main>`;
  }

  async function route() {
    closeModal();
    const parts = (location.hash.replace(/^#\/?/, '') || 'charters').split('/');
    if (parts[0] === 'charter' && parts[1]) return viewCharter(parts[1]);
    if (parts[0] === 'fuel') return viewFuel(parts[1] || 'voyages');
    if (parts[0] === 'settings') return viewSettings();
    if (parts[0] === 'invoices') return viewInvoices(parts[1]);
    return viewCharters();
  }

  // Délégation globale des clics de l'espace capitaine
  const ACTIONS = {};
  document.addEventListener('click', e => {
    const b = e.target.closest('[data-act]');
    if (!b || !ACTIONS[b.dataset.act] || IS_CLIENT) return;
    e.preventDefault();
    ACTIONS[b.dataset.act](b, e);
  });
  ACTIONS.logout = async () => { await api.signOut(); location.hash = ''; renderLogin(); };
  ACTIONS['demo-reset'] = async () => { if (await confirmDialog('Remettre les données de démonstration à zéro ?', 'Réinitialiser')) { api.reset(); await reload(); route(); } };

  // ---------------- Voyages : ordre chronologique & voyage charter automatique ----------------
  // ROB au départ estimé à une date : ROB à l'arrivée du dernier voyage terminé avant cette date + soutages faits depuis
  function suggestRobAt(date) {
    const done = S.voyages.filter(v => toNum(v.rob_after) !== null && (!date || !v.end_date || v.end_date <= date)).sort(byPos);
    const last = done[done.length - 1];
    if (!last) return null;
    const loaded = sum(S.bunkers.filter(b => b.date && last.end_date && b.date > last.end_date && (!date || b.date <= date)), b => b.quantity);
    return r2(last.rob_after + loaded);
  }
  // Position pour qu'un voyage daté s'insère au bon endroit (les voyages suivants sont décalés)
  async function placeVoyage(date) {
    const list = S.voyages.slice().sort(byPos);
    const maxPos = Math.max(0, ...list.map(v => v.position || 0));
    const next = date ? list.find(v => v.start_date && v.start_date > date) : null;
    if (!next) return maxPos + 1;
    const pos = next.position;
    for (const v of list.filter(x => (x.position || 0) >= pos).reverse()) {
      await api.update('voyages', v.id, { position: (v.position || 0) + 1 });
      v.position = (v.position || 0) + 1;
    }
    return pos;
  }
  async function createCharterVoyage(c) {
    const position = await placeVoyage(c.start_date);
    const v = await api.insert('voyages', {
      num: Math.max(0, ...S.voyages.map(x => x.num || 0)) + 1, position,
      description: `Charter ${c.client_name || c.title}`, start_date: c.start_date, end_date: c.end_date,
      rob_before: suggestRobAt(c.start_date), rob_after: null, hours: null, is_charter: true, charter_id: c.id });
    await api.update('charters', c.id, { auto_voyage: false });
    return v;
  }
  // Au chargement : crée le voyage des charters qui ont commencé (une seule fois par charter)
  async function autoCreateDueVoyages() {
    const due = S.charters.filter(c => c.auto_voyage && c.start_date && c.start_date <= today()
      && !S.voyages.some(v => v.charter_id === c.id)).sort((a, b) => a.start_date.localeCompare(b.start_date));
    const created = [];
    for (const c of due) { try { await createCharterVoyage(c); created.push(c.title); await reload(); } catch { /* réessaiera au prochain chargement */ } }
    if (created.length) toast(`Voyage fuel créé pour : ${created.join(', ')}`);
  }

  // ---------------- Paramètres : fiche bateau & fournisseurs ----------------
  function viewSettings() {
    const P = PROFILE;
    $('#app').innerHTML = shell('settings', `
      <div class="page-head"><div><div class="eyebrow">Paramètres</div><h1>Bateau & facturation</h1></div></div>
      <form id="f-settings">
      <section class="panel">
        <div class="panel-head"><h2>Société qui facture (le bateau)</h2></div>
        <p class="muted small">Ces informations apparaissent sur les factures émises automatiquement par le bateau (fuel et lignes dont le fournisseur est le bateau).</p>
        <div class="form-grid">
          <label class="field"><span>Nom dans la liste des fournisseurs</span><input name="short_name" required value="${esc(P.short_name)}" placeholder="ex. TDB"></label>
          <label class="field"><span>Raison sociale <small>(sur la facture)</small></span><input name="legal_name" value="${esc(P.legal_name)}" placeholder="ex. TDB SAS"></label>
          <label class="field span-2"><span>Adresse de facturation</span><textarea name="address" rows="3" placeholder="Rue&#10;Code postal Ville&#10;Pays">${esc(P.address)}</textarea></label>
          <label class="field"><span>N° de TVA intracommunautaire</span><input name="vat_number" value="${esc(P.vat_number)}"></label>
          <label class="field"><span>Immatriculation <small>(RCS / SIRET…)</small></span><input name="registration" value="${esc(P.registration)}"></label>
          <label class="field"><span>E-mail</span><input name="email" type="email" value="${esc(P.email)}"></label>
          <label class="field"><span>Téléphone</span><input name="phone" value="${esc(P.phone)}"></label>
          <label class="field"><span>Banque</span><input name="bank" value="${esc(P.bank)}"></label>
          <label class="field"><span>IBAN</span><input name="iban" value="${esc(P.iban)}"></label>
          <label class="field"><span>BIC</span><input name="bic" value="${esc(P.bic)}"></label>
        </div>
      </section>
      <section class="panel">
        <div class="panel-head"><h2>Factures</h2></div>
        <div class="form-grid">
          <label class="field"><span>Préfixe de numérotation</span><input name="invoice_prefix" value="${esc(P.invoice_prefix)}" placeholder="TDB"><small class="muted">Ex. TDB → TDB-2026-0001, TDB-2026-0002…</small></label>
          <label class="field"><span>Taux de TVA par défaut sur la refacturation (%)</span><input type="number" step="0.01" min="0" name="vat_rate" value="${esc(P.vat_rate)}"></label>
          <label class="field span-2"><span>Mention TVA <small>(ex. régime d’exonération applicable — à valider avec votre représentant fiscal)</small></span><input name="vat_mention" value="${esc(P.vat_mention)}"></label>
          <label class="field span-2"><span>Conditions de règlement</span><input name="payment_terms" value="${esc(P.payment_terms)}"></label>
          <label class="field span-2"><span>Pied de facture <small>(facultatif)</small></span><input name="footer" value="${esc(P.footer)}"></label>
        </div>
      </section>
      <section class="panel">
        <div class="panel-head"><h2>Liste des fournisseurs</h2></div>
        <p class="muted small">Un fournisseur par ligne. Ils sont proposés lors de la saisie d’une dépense. Le bateau (« ${esc(boatName())} ») est toujours en tête de liste : le choisir marque la ligne comme facturée par le bateau.</p>
        <label class="field"><textarea name="suppliers" rows="8" placeholder="BWA&#10;Cavavin&#10;Metro">${esc((P.suppliers || []).join('\n'))}</textarea></label>
      </section>
      <div class="row-gap"><button class="btn btn-primary" type="submit">${ICON.check}Enregistrer les paramètres</button></div>
      </form>`);
    $('#f-settings').addEventListener('submit', async e => {
      e.preventDefault();
      const f = e.target; if (!f.reportValidity()) return;
      const d = Object.fromEntries(new FormData(f));
      const data = { ...PROFILE, ...d, vat_rate: toNum(d.vat_rate) || 0,
        invoice_prefix: (d.invoice_prefix || d.short_name || 'BOAT').replace(/[^A-Za-z0-9]/g, '').toUpperCase(),
        suppliers: [...new Set(String(d.suppliers || '').split('\n').map(s => s.trim()).filter(s => s && !isBoatName(s)))] };
      const btn = $('button[type=submit]', f); busy(btn, true);
      try { await api.saveProfile(data); await reload(); viewSettings(); toast('Paramètres enregistrés'); }
      catch (err) { toast(errMsg(err), 'err'); busy(btn, false); }
    });
  }

  // ---------------- Registre des factures émises par le bateau ----------------
  function viewInvoices(yearSel) {
    const all = S.invoices.slice().sort((a, b) => b.year - a.year || b.seq - a.seq);
    const years = [...new Set(all.map(i => i.year))];
    const year = yearSel ? +yearSel : years[0];
    const list = all.filter(i => !year || i.year === year);
    const ok = list.filter(i => i.status === 'issued');
    $('#app').innerHTML = shell('invoices', `
      <div class="page-head"><div><div class="eyebrow">${esc(boatName())}</div><h1>Factures émises</h1></div>
        <div class="head-actions">
          ${years.length > 1 ? `<select id="inv-year">${years.map(y => `<option ${y === year ? 'selected' : ''}>${y}</option>`).join('')}</select>` : ''}
          ${list.length ? `<button class="btn btn-ghost" data-act="inv-register">${ICON.sheet}Registre (Excel)</button>` : ''}
        </div></div>
      ${list.length ? `<div class="kpis kpis-3">
          <div class="kpi"><div class="kpi-label">Factures ${year || ''}</div><div class="kpi-value">${ok.length}</div></div>
          <div class="kpi"><div class="kpi-label">Total HT</div><div class="kpi-value">${money(sum(ok, i => i.total_ht))}</div></div>
          <div class="kpi"><div class="kpi-label">Total TTC</div><div class="kpi-value">${money(sum(ok, i => i.total_ttc))}</div></div></div>
        <section class="panel"><div class="table-wrap"><table class="table table-cards">
          <thead><tr><th>N°</th><th>Date</th><th>Client</th><th>Charter</th><th class="num">HT</th><th class="num">TVA</th><th class="num">TTC</th><th>Statut</th><th></th></tr></thead>
          <tbody>${list.map(i => { INV_CACHE[i.id] = i; const ch = charterById(i.charter_id); return `<tr class="${i.status === 'cancelled' ? 'row-cancelled' : ''}">
            <td data-label="N°" class="strong">${esc(i.number)}</td><td data-label="Date">${fdate(i.issue_date)}</td>
            <td data-label="Client">${esc(i.buyer?.name || '')}</td>
            <td data-label="Charter">${ch ? `<a href="#/charter/${ch.id}">${esc(ch.title)}</a>` : esc(i.buyer?.charter || '')}</td>
            <td data-label="HT" class="num">${money(i.total_ht, i.currency)}</td><td data-label="TVA" class="num">${money(i.total_vat, i.currency)}</td>
            <td data-label="TTC" class="num strong">${money(i.total_ttc, i.currency)}</td>
            <td data-label="Statut">${i.status === 'cancelled' ? '<span class="badge badge-open">Annulée</span>' : '<span class="badge badge-ok">Émise</span>'}</td>
            <td class="actions"><button class="btn btn-primary btn-sm" data-inv="${esc(i.id)}">${ICON.file}PDF</button></td></tr>`; }).join('')}</tbody>
        </table></div></section>`
      : `<div class="empty-state card"><p>Aucune facture émise pour le moment.</p><p class="muted small">Les factures de ${esc(boatName())} sont émises automatiquement à la clôture d’un charter qui contient du fuel ou des lignes facturées par le bateau.</p></div>`}`);
    const sel = $('#inv-year'); if (sel) sel.onchange = () => { location.hash = `#/invoices/${sel.value}`; };
  }
  ACTIONS['inv-register'] = async () => {
    try { await loadScript(XLSX_JS); } catch (e) { return toast(errMsg(e), 'err'); }
    const X = window.XLSX, y = +(location.hash.split('/')[2] || 0);
    const list = S.invoices.slice().sort((a, b) => a.year - b.year || a.seq - b.seq).filter(i => !y || i.year === y);
    const rows = [['N°', 'Date', 'Client', 'Charter', 'Total HT', 'TVA', 'Total TTC', 'Devise', 'Statut', 'Annulée le'],
      ...list.map(i => [i.number, fdate(i.issue_date), i.buyer?.name || '', i.buyer?.charter || '', toNum(i.total_ht), toNum(i.total_vat), toNum(i.total_ttc), i.currency,
        i.status === 'cancelled' ? 'Annulée' : 'Émise', i.cancelled_at ? fdate(i.cancelled_at.slice(0, 10)) : ''])];
    const ws = X.utils.aoa_to_sheet(rows); ws['!cols'] = [{ wch: 16 }, { wch: 11 }, { wch: 22 }, { wch: 34 }, { wch: 12 }, { wch: 10 }, { wch: 12 }, { wch: 7 }, { wch: 9 }, { wch: 11 }];
    const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, ws, 'Factures');
    saveBlob(new Blob([X.write(wb, { type: 'array', bookType: 'xlsx' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `Registre_factures_${boatName()}_${y || 'tout'}.xlsx`);
  };

  // ---------------- Liste des charters ----------------
  function viewCharters() {
    const list = S.charters.slice().sort((a, b) => String(b.start_date || '').localeCompare(String(a.start_date || '')));
    const cards = list.map(c => {
      const st = computeStatement(buildBundle(S, c));
      return `<a class="charter-card" href="#/charter/${c.id}">
        <div class="cc-top">${statusBadge(c)}
          <span class="muted small">${fdate(c.start_date)} → ${fdate(c.end_date)}</span></div>
        <h3>${esc(c.title)}</h3><p class="muted">${esc(c.client_name || '')}</p>
        <div class="cc-figs"><div><span>Dépenses</span><b>${money(st.total, c.currency)}</b></div>
          <div><span>Solde APA</span><b class="${st.balance < 0 ? 'neg' : 'pos'}">${money(st.balance, c.currency)}</b></div></div>
      </a>`;
    }).join('');
    $('#app').innerHTML = shell('charters', `
      <div class="page-head"><div><div class="eyebrow">Comptabilité</div><h1>Charters</h1></div>
        <button class="btn btn-primary" data-act="new-charter">${ICON.plus}Nouveau charter</button></div>
      ${list.length ? `<div class="cards">${cards}</div>` : `<div class="empty-state card"><p>Aucun charter pour le moment.</p>
        <button class="btn btn-primary" data-act="new-charter">${ICON.plus}Créer le premier charter</button></div>`}`);
  }

  // ---------------- Formulaire charter ----------------
  function charterForm(c = {}) {
    const isNew = !c.id;
    const opt = (list, v) => list.map(x => `<option ${x === v ? 'selected' : ''}>${esc(x)}</option>`).join('');
    const m = openModal({
      title: isNew ? 'Nouveau charter' : 'Modifier le charter', wide: true,
      body: `<form class="form-grid" id="f-charter">
        <label class="field span-2"><span>Intitulé</span><input name="title" required value="${esc(c.title || '')}" placeholder="Charter Dupont — Juillet 2026"></label>
        <label class="field"><span>Client (charterer)</span><input name="client_name" value="${esc(c.client_name || '')}"></label>
        <label class="field"><span>N° TVA du client <small>(facultatif)</small></span><input name="client_vat" value="${esc(c.client_vat || '')}"></label>
        <label class="field span-2"><span>Adresse de facturation du client <small>(pour la facture du bateau)</small></span><textarea name="client_address" rows="2">${esc(c.client_address || '')}</textarea></label>
        <label class="field"><span>Yacht</span><input name="yacht" value="${esc(c.yacht || YACHT)}"></label>
        <label class="field"><span>Début</span><input type="date" name="start_date" value="${esc(c.start_date || '')}"></label>
        <label class="field"><span>Fin</span><input type="date" name="end_date" value="${esc(c.end_date || '')}"></label>
        <label class="field"><span>APA reçu</span><input type="number" step="0.01" name="apa_paid" value="${esc(c.apa_paid ?? '')}"></label>
        <label class="field field-fee"><span>Charter fee <small>(sert uniquement au calcul des tips)</small></span><input type="number" step="0.01" name="charter_fee" value="${esc(c.charter_fee ?? '')}" placeholder="ex. 155000"></label>
        ${!S.voyages.some(v => v.charter_id === c.id) ? `<label class="check span-2 boat-toggle"><input type="checkbox" name="auto_voyage" ${c.auto_voyage !== false ? 'checked' : ''}>
          <span><b>Créer automatiquement le voyage dans le fuel booklet</b> le jour du début du charter (nom, dates, ROB au départ calculé ce jour-là)</span></label>` : ''}
        <label class="check span-2 tips-toggle"><input type="checkbox" name="tips_prompt" ${tipsEnabled(c) ? 'checked' : ''}>
          <span><b>Proposer les tips au client avant sa signature</b><br><small class="muted">Fenêtre 10 % / 15 % / 20 % du charter fee ou montant libre. Décochez si vous préférez ne pas aborder les tips dans l’app.</small></span></label>
        <label class="field"><span>Pays de début du charter <small>(TVA)</small></span><select name="vat_country">
          ${VAT_COUNTRIES.map(([n, r]) => `<option value="${n}" data-rate="${r ?? ''}" ${n === (c.vat_country || 'France') ? 'selected' : ''}>${n}${r != null ? ` — ${r} %` : ''}</option>`).join('')}</select></label>
        <label class="field"><span>Taux de TVA appliqué (%) <small>(montants payés HT)</small></span><input type="number" step="0.01" min="0" name="vat_rate" value="${esc(c.vat_rate ?? 20)}"></label>
        <label class="field"><span>Devise</span><select name="currency">${opt(['EUR', 'USD', 'GBP', 'CHF'], c.currency || CFG.currency || 'EUR')}</select></label>
        <div class="field-sep span-2">Ligne fuel (calculée automatiquement)</div>
        <label class="field"><span>Société</span><input name="fuel_company" value="${esc(c.fuel_company ?? boatName())}" list="dl-fuel-co"><datalist id="dl-fuel-co"><option value="${esc(boatName())}"></datalist></label>
        <label class="check span-2 boat-toggle"><input type="checkbox" name="fuel_by_boat" ${c.fuel_by_boat !== false ? 'checked' : ''}>
          <span><b>Fuel facturé par le bateau (${esc(boatName())})</b> — inclus dans la facture émise automatiquement à la clôture</span></label>
        <label class="field"><span>Description</span><input name="fuel_description" value="${esc(c.fuel_description ?? 'MEs & DGs')}"></label>
        <label class="field"><span>Paiement</span><select name="fuel_method">${opt(METHODS, c.fuel_method || 'To be paid')}</select></label>
        <div class="field-sep span-2">Envoi du rapport final (après signature)</div>
        <label class="field"><span>E-mail du représentant fiscal <small>(destinataire)</small></span><input name="report_to" value="${esc(c.report_to || '')}" placeholder="fiscal@exemple.com" inputmode="email"></label>
        <label class="field"><span>E-mails des brokers <small>(en copie, séparés par une virgule)</small></span><input name="report_cc" value="${esc(c.report_cc || '')}" placeholder="broker1@…, broker2@…" inputmode="email"></label>
        <label class="field span-2"><span>Notes internes <small>(non visibles par le client)</small></span><textarea name="notes" rows="2">${esc(c.notes || '')}</textarea></label>
      </form>`,
      footer: `${!isNew ? `<button class="btn btn-ghost danger mr-auto" data-del>${ICON.trash}Supprimer</button>` : ''}
        <button class="btn btn-ghost" data-close>Annuler</button><button class="btn btn-primary" data-ok>Enregistrer</button>`,
    });
    const fc = $('#f-charter', m);
    fc.vat_country.addEventListener('change', () => { const r = fc.vat_country.selectedOptions[0].dataset.rate; if (r !== '') fc.vat_rate.value = r; });
    $('[data-ok]', m).onclick = async ev => {
      const f = $('#f-charter', m);
      if (!f.reportValidity()) return;
      const d = Object.fromEntries(new FormData(f));
      const row = { title: d.title.trim(), client_name: d.client_name || null, yacht: d.yacht || YACHT, start_date: d.start_date || null,
        end_date: d.end_date || null, apa_paid: toNum(d.apa_paid) || 0, charter_fee: toNum(d.charter_fee), currency: d.currency,
        vat_country: d.vat_country || null, vat_rate: toNum(d.vat_rate) ?? 0,
        fuel_company: d.fuel_company || null, fuel_description: d.fuel_description || null, fuel_method: d.fuel_method, notes: d.notes || null, tips_prompt: !!d.tips_prompt, fuel_by_boat: !!d.fuel_by_boat,
        ...(S.voyages.some(v => v.charter_id === c.id) ? {} : { auto_voyage: !!d.auto_voyage }),
        client_address: d.client_address || null, client_vat: d.client_vat || null,
        report_to: cleanEmails(d.report_to) || null, report_cc: cleanEmails(d.report_cc) || null };
      const bad = [...splitEmails(d.report_to), ...splitEmails(d.report_cc)].filter(x => !EMAIL_RE.test(x));
      if (bad.length) return toast('E-mail invalide : ' + bad.join(', '), 'err');
      busy(ev.currentTarget, true);
      try {
        const saved = isNew ? await api.insert('charters', row) : await api.update('charters', c.id, row);
        await reload();
        // voyage charter pas encore terminé : suit les dates du charter
        const lv = S.voyages.find(v => v.charter_id === saved.id && toNum(v.rob_after) === null);
        if (!isNew && lv && (lv.start_date !== row.start_date || lv.end_date !== row.end_date)) { await api.update('voyages', lv.id, { start_date: row.start_date, end_date: row.end_date }); await reload(); }
        await autoCreateDueVoyages();
        closeModal(); notifyCharter(saved.id);
        if (isNew) location.hash = `#/charter/${saved.id}`; else route();
        toast('Charter enregistré');
      } catch (e) { toast(errMsg(e), 'err'); busy(ev.currentTarget, false); }
    };
    const del = $('[data-del]', m);
    if (del) del.onclick = async () => {
      if (!(await confirmDialog(`Supprimer définitivement « ${esc(c.title)} » et toutes ses dépenses ?`, 'Supprimer'))) return;
      try {
        const files = S.expenses.filter(e => e.charter_id === c.id && e.file_path).map(e => e.file_path);
        await api.remove('charters', c.id);
        for (const p of files) await api.removeFile(p).catch(() => {});
        await reload(); location.hash = '#/charters'; toast('Charter supprimé');
      } catch (e) { toast(errMsg(e), 'err'); }
    };
  }
  ACTIONS['new-charter'] = () => {
    const last = S.charters.slice().sort((a, b) => String(b.created_at || b.start_date || '').localeCompare(String(a.created_at || a.start_date || '')))[0];
    charterForm({ tips_prompt: last ? tipsEnabled(last) : true, report_to: last?.report_to || '', fuel_company: boatName(), fuel_by_boat: true,
      vat_country: last?.vat_country || 'France', vat_rate: last?.vat_rate ?? 20, auto_voyage: true });
  };
  ACTIONS['edit-charter'] = b => charterForm(charterById(b.dataset.id));

  // ---------------- Détail d'un charter ----------------
  let CURRENT = null;
  function viewCharter(id) {
    const c = charterById(id);
    if (!c) { location.hash = '#/charters'; return; }
    CURRENT = id;
    const B = buildBundle(S, c), st = computeStatement(B);
    const signed = c.status === 'signed', closed = c.status === 'closed', locked = signed || closed;
    const freeVoyages = S.voyages.filter(v => !v.charter_id);
    $('#app').innerHTML = shell('charters', `
      <a class="back" href="#/charters">${ICON.back}Charters</a>
      <div class="page-head">
        <div><div class="eyebrow">${esc(c.yacht || YACHT)}</div><h1>${esc(c.title)}</h1>
          <p class="muted">${esc(c.client_name || '')}${c.client_name ? ' · ' : ''}${fdate(c.start_date)} → ${fdate(c.end_date)}
          · ${statusBadge(c)}${c.charter_fee ? ` · Charter fee ${money(c.charter_fee, c.currency)}` : ''}</p></div>
        <div class="head-actions">
          ${!locked ? `<button class="btn btn-gold" data-act="close-charter" data-id="${c.id}">${ICON.lock}Clôturer la compta</button>` : ''}
          ${closed ? `<button class="btn btn-gold" data-act="notify-closed" data-id="${c.id}">${ICON.mail}Prévenir le client</button>` : ''}
          ${signed ? `<button class="btn btn-gold" data-act="send-report" data-id="${c.id}">${ICON.mail}Envoyer le rapport</button>` : ''}
          <button class="btn btn-ghost" data-act="share" data-id="${c.id}">${ICON.share}Partager</button>
          <a class="btn btn-ghost" href="${esc(clientUrl(c))}" target="_blank" rel="noopener">${ICON.eye}Vue client</a>
          <button class="btn btn-ghost" data-act="print">${ICON.print}PDF</button>
          <button class="btn btn-ghost" data-act="xlsx-charter">${ICON.sheet}Excel</button>
          ${B.invoices.length ? B.invoices.map(i => { INV_CACHE[i.id] = i; return `<button class="btn btn-ghost" data-inv="${esc(i.id)}">${ICON.file}Facture ${esc(i.number)}</button>`; }).join('') : ''}
          <button class="btn btn-ghost" data-act="zip-charter">${ICON.file}Factures (.zip)</button>
          <button class="btn btn-ghost" data-act="edit-charter" data-id="${c.id}">${ICON.edit}Modifier</button>
        </div>
      </div>
      ${signed ? `<div class="banner banner-ok">${ICON.check}<div><strong>Comptabilité signée par ${esc(c.client_signed_name)}</strong> le ${fdatetime(c.client_signed_at)}${st.tipsAuto ? ` — tips choisis : <b>${money(st.tipsAuto, c.currency)}</b>${c.tips_pct != null ? ` (${fmt(c.tips_pct, c.tips_pct % 1 ? 2 : 0)} % du charter fee)` : ''}` : ''}.
        Les dépenses sont verrouillées. <button class="link-btn" data-act="reopen" data-id="${c.id}">Rouvrir la compta</button>
        <br>${c.report_sent_at ? `Rapport envoyé le ${fdatetime(c.report_sent_at)}${c.report_to ? ` à ${esc(c.report_to)}` : ''}. <button class="link-btn" data-act="send-report" data-id="${c.id}">Renvoyer</button>`
          : `<b>Prochaine étape :</b> <button class="link-btn" data-act="send-report" data-id="${c.id}">envoyer le rapport</button> au représentant fiscal${c.report_cc ? ' et aux brokers' : ''}.`}</div></div>`
      : closed ? `<div class="banner banner-gold">${ICON.lock}<div><strong>Compta clôturée le ${fdatetime(c.closed_at)}</strong> — en attente de la signature du client.
        Les dépenses sont verrouillées. <button class="link-btn" data-act="reopen" data-id="${c.id}">Rouvrir la compta</button></div></div>`
      : !c.charter_fee && tipsEnabled(c) ? `<div class="banner banner-info">${ICON.edit}<div>Renseignez le <b>charter fee</b> (bouton Modifier) : il sert au calcul des tips proposés au client lors de la signature.</div></div>` : ''}
      ${renderKpis(B, st)}
      <section class="panel">
        <div class="panel-head"><h2>Dépenses</h2>${!locked ? `<button class="btn btn-primary btn-sm" data-act="new-expense">${ICON.plus}Ajouter une dépense</button>` : ''}</div>
        ${renderExpensesTable(B, st, { admin: true, locked })}
      </section>
      <div class="grid-2">
        <section class="panel">
          <div class="panel-head"><h2>Fuel du charter</h2>
            <div class="row-gap">
              ${freeVoyages.length ? `<button class="btn btn-ghost btn-sm" data-act="link-voyage">${ICON.link}Lier un voyage</button>` : ''}
              <button class="btn btn-primary btn-sm" data-act="new-voyage-charter">${ICON.plus}Voyage charter</button></div></div>
          ${B.voyages.length ? renderFuelDetail(B, st) + `<div class="row-gap wrap">${B.voyages.map(v => `
            <button class="btn btn-ghost btn-sm" data-act="edit-voyage" data-id="${v.id}">${ICON.edit}Modifier « ${esc(v.description || 'voyage')} »</button>
            <button class="btn btn-ghost btn-sm danger" data-act="unlink-voyage" data-id="${v.id}">Délier</button>`).join('')}</div>`
            : c.auto_voyage !== false && c.start_date
              ? `<div class="banner banner-info small">${ICON.drop}<div>Le voyage fuel de ce charter sera <b>créé automatiquement le ${fdate(c.start_date)}</b> (début du charter), avec le ROB au départ calculé ce jour-là.
                  <br><button class="link-btn" data-act="voyage-now">Créer maintenant</button></div></div>`
              : `<p class="muted">Aucun voyage lié. Créez le voyage du charter (ROB départ / arrivée) : la consommation est déduite de vos soutes et facturée au prix du lot utilisé.</p>`}
          ${B.voyages.some(v => toNum(v.rob_after) === null) && c.end_date && c.end_date <= today() ? `<div class="banner banner-gold small">${ICON.drop}<div>Charter terminé : renseignez le <b>ROB à l’arrivée</b> et la répartition par soutage. ${B.voyages.filter(v => toNum(v.rob_after) === null).map(v => `<button class="link-btn" data-act="edit-voyage" data-id="${v.id}">Compléter « ${esc(v.description || 'voyage')} »</button>`).join(' ')}</div></div>` : ''}
        </section>
        <section class="panel">
          <div class="panel-head"><h2>Répartition</h2></div>
          ${renderCategories(B, st)}
          <div class="cash-sum"><div class="panel-head sub"><h3>Espèces (Cash ATM)</h3>${!locked ? `<button class="btn btn-ghost btn-sm" data-act="new-cash">${ICON.plus}Retrait</button>` : ''}</div>
            ${B.cash.map(w => `<div class="kv"><span>${fdate(w.date)} · ${esc(w.location || '')}</span><b>${money(w.amount, c.currency)}
              ${!locked ? `<button class="icon-btn danger xs" data-act="del-cash" data-id="${w.id}" title="Supprimer">${ICON.trash}</button>` : ''}</b></div>`).join('')}
            <div class="kv"><span>Total retraits</span><b>${money(st.cashWithdrawn, c.currency)}</b></div>
            <div class="kv"><span>Dépenses payées « Cash APA »</span><b>${money(st.cashSpent, c.currency)}</b></div>
            <div class="kv strong"><span>Espèces restantes</span><b>${money(st.cashInHand, c.currency)}</b></div>
          </div>
        </section>
      </div>
      ${invoicePanel(c, B, locked)}
      <section class="panel">
        <div class="panel-head"><h2>Signatures</h2></div>
        <div class="sig-grid">
          ${renderSignatureBox('Charterer', c.client_signature, c.client_signed_name, c.client_signed_at)}
          ${renderSignatureBox('Capitaine', c.master_signature, c.master_signed_name, c.master_signed_at)}
        </div>
        <div class="row-gap center no-print">
          <button class="btn btn-ghost" data-act="master-sign">${ICON.pen}${c.master_signature ? 'Re-signer (capitaine)' : 'Signer (capitaine)'}</button>
          ${!locked ? `<span class="muted small">Le client pourra signer${tipsEnabled(c) ? ' (et choisir ses tips)' : ''} une fois la compta clôturée.</span>` : ''}
        </div>
      </section>`);
  }

  function invoicePanel(c, B, locked) {
    const draft = boatInvoiceDraft(B), issued = B.invoices[0];
    const cancelled = S.invoices.filter(i => i.charter_id === c.id && i.status === 'cancelled');
    const profileOk = !!(PROFILE.legal_name || PROFILE.short_name !== 'Boat') && !!PROFILE.address;
    if (!draft.lines.length && !issued && !cancelled.length) return `<section class="panel"><div class="panel-head"><h2>Facture du bateau</h2></div>
      <p class="muted small">Aucune ligne facturée par le bateau (${esc(boatName())}). Cochez « Fuel facturé par le bateau » dans <i>Modifier</i>, ou choisissez « ${esc(boatName())} » comme fournisseur d’une dépense.</p></section>`;
    const changed = issued && draft.fingerprint !== issued.fingerprint;
    return `<section class="panel">
      <div class="panel-head"><h2>Facture du bateau (${esc(boatName())})</h2>
        <div class="row-gap">
          ${!issued ? `<button class="btn btn-ghost btn-sm" data-act="inv-draft">${ICON.eye}Aperçu (brouillon)</button>` : ''}
          ${locked && (!issued || changed) && draft.lines.length ? `<button class="btn btn-primary btn-sm" data-act="inv-issue">${ICON.file}${issued ? 'Ré-émettre' : 'Émettre la facture'}</button>` : ''}
        </div></div>
      ${!profileOk ? `<div class="banner banner-info small">${ICON.edit}<div>Complétez la fiche du bateau (raison sociale, adresse…) dans <a href="#/settings">Paramètres</a> avant la clôture : elle figure sur la facture.</div></div>` : ''}
      ${issued ? `<div class="inv-card"><div><div class="inv-num">N° ${esc(issued.number)}</div>
          <div class="muted small">Émise le ${fdate(issued.issue_date)} · ${issued.lines.length} ligne(s)</div></div>
          <div class="inv-total">${money(issued.total_ttc, issued.currency)}</div>
          <div>${invoiceLink(issued)}</div></div>
          ${changed ? `<p class="small warn">La compta a changé depuis l’émission : ${locked ? 'cliquez sur « Ré-émettre » pour annuler cette facture et en émettre une nouvelle.' : 'une nouvelle facture remplacera celle-ci à la prochaine clôture.'}</p>` : ''}`
        : `<p class="small muted">${draft.lines.length} ligne(s) · total ${money(draft.total_ttc, c.currency)} TTC. ${locked ? 'Cliquez sur « Émettre la facture ».' : 'La facture sera numérotée et émise automatiquement à la clôture de la compta, puis visible par le client.'}</p>`}
      ${cancelled.length ? `<p class="small muted">Factures annulées : ${cancelled.map(i => `${invoiceLink(i)}`).join(' ')}</p>` : ''}
    </section>`;
  }
  ACTIONS['inv-draft'] = () => downloadInvoice({ ...boatInvoiceDraft(buildBundle(S, charterById(CURRENT))), issue_date: today() }, true);
  ACTIONS['inv-issue'] = async b => {
    busy(b, true);
    try { const r = await ensureBoatInvoice(charterById(CURRENT)); await reload(); notifyCharter(CURRENT); route();
      toast(r.inv ? `Facture ${r.inv.number} émise${r.replaced ? ` (remplace ${r.replaced})` : ''}` : 'Aucune ligne à facturer'); }
    catch (e) { toast(errMsg(e), 'err'); busy(b, false); }
  };

  ACTIONS.print = () => window.print();
  ACTIONS['zip-charter'] = b => downloadReportZip(buildBundle(S, charterById(CURRENT)), b);
  ACTIONS['xlsx-charter'] = () => { const c = charterById(CURRENT); exportCharterXlsx(buildBundle(S, c)); };

  const EMAIL_RE = /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/;
  const splitEmails = s => String(s || '').split(/[\s,;]+/).map(x => x.trim()).filter(Boolean);
  const cleanEmails = s => splitEmails(s).join(', ');
  const reportUrl = c => `${location.origin}${location.pathname}?r=${c.report_token}`;

  // Liens de composition : messagerie par défaut (mailto) ou Gmail
  function composeUrls({ to = '', cc = '', subject = '', body = '' }) {
    const clean = s => String(s || '').replace(/\s+/g, '');
    const q = [];
    if (cc) q.push('cc=' + encodeURIComponent(clean(cc)));
    q.push('subject=' + encodeURIComponent(subject), 'body=' + encodeURIComponent(body));
    const g = new URLSearchParams({ view: 'cm', fs: '1', to: clean(to), su: subject, body });
    if (cc) g.set('cc', clean(cc));
    return { mailto: `mailto:${clean(to)}?${q.join('&')}`, gmail: `https://mail.google.com/mail/?${g.toString()}` };
  }

  function statusBadge(c) {
    if (c.status === 'signed') return '<span class="badge badge-ok">Signée</span>';
    if (c.status === 'closed') return '<span class="badge badge-gold">Clôturée · à signer</span>';
    return '<span class="badge badge-open">En cours</span>';
  }

  function messageModal(c, kind) {
    const url = clientUrl(c), yacht = c.yacht || YACHT;
    const msg = kind === 'closed'
      ? `Bonjour,\n\nLa comptabilité de votre charter à bord de ${yacht} est désormais clôturée. Vous pouvez la vérifier et la signer ici :\n${url}\n\n—\n\nHello,\n\nThe accounts of your charter on board ${yacht} are now closed. You can review and sign them here:\n${url}`
      : `Bonjour, voici le lien pour suivre en direct la comptabilité de votre charter à bord de ${yacht} :\n${url}\n\nHello, here is the link to follow the accounts of your charter on board ${yacht} live:\n${url}`;
    const subject = kind === 'closed' ? `${c.title} — compta clôturée / accounts closed` : c.title;
    const m = openModal({
      title: kind === 'closed' ? 'Prévenir le client : compta clôturée' : 'Partager au client', wide: true,
      body: `${kind === 'closed'
        ? `<p>Envoyez ce message au client. S’il a déjà le lien ouvert, il reçoit aussi <b>immédiatement</b> une fenêtre « Comptabilité clôturée » l’invitant à signer.</p>`
        : `<p class="muted">Ce lien privé donne accès en lecture à la compta de ce charter, à ses factures et uniquement aux factures fuel des soutages consommés pendant ce charter. Le client pourra signer une fois la compta clôturée.</p>`}
        <label class="field"><span>Message</span><textarea rows="${kind === 'closed' ? 9 : 5}" id="share-msg">${esc(msg)}</textarea></label>
        <div class="row-gap wrap">
          <button class="btn btn-primary" data-copy>${ICON.copy}Copier le message</button>
          <a class="btn btn-ghost" data-gmail href="#">${ICON.mail}Gmail</a>
          <a class="btn btn-ghost" data-mail href="#">${ICON.mail}Autre messagerie</a>
          <a class="btn btn-ghost" data-wa href="#" target="_blank" rel="noopener">WhatsApp</a>
          <a class="btn btn-ghost" href="${esc(url)}" target="_blank" rel="noopener">${ICON.eye}Ouvrir le lien</a>
        </div>
        ${DEMO ? `<p class="small muted">En mode démo, le lien ne fonctionne que dans ce navigateur.</p>` : ''}
        ${kind !== 'closed' ? `<hr><button class="link-btn danger" data-regen>Générer un nouveau lien (l’ancien ne fonctionnera plus)</button>` : ''}`,
    });
    const text = () => $('#share-msg', m).value;
    $('[data-mail]', m).onclick = e => { e.preventDefault(); location.href = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(text())}`; };
    $('[data-gmail]', m).onclick = e => { e.preventDefault(); window.open(composeUrls({ subject, body: text() }).gmail, '_blank', 'noopener'); };
    $('[data-wa]', m).onclick = e => { e.preventDefault(); window.open(`https://wa.me/?text=${encodeURIComponent(text())}`, '_blank', 'noopener'); };
    $('[data-copy]', m).onclick = async () => {
      try { await navigator.clipboard.writeText(text()); toast('Message copié'); }
      catch { $('#share-msg', m).select(); document.execCommand('copy'); toast('Message copié'); }
    };
    const regen = $('[data-regen]', m);
    if (regen) regen.onclick = async () => {
      if (!(await confirmDialog('Générer un nouveau lien ? L’ancien lien cessera de fonctionner.', 'Générer'))) return;
      try { await api.update('charters', c.id, { client_token: uuid() }); await reload(); route(); messageModal(charterById(c.id), kind); }
      catch (e) { toast(errMsg(e), 'err'); }
    };
  }
  ACTIONS.share = b => messageModal(charterById(b.dataset.id), 'share');
  ACTIONS['notify-closed'] = b => messageModal(charterById(b.dataset.id), 'closed');

  // Envoi du rapport final au représentant fiscal (+ brokers en copie) — uniquement si signé
  ACTIONS['send-report'] = b => {
    const c = charterById(b.dataset.id);
    if (c.status !== 'signed') return toast('Le rapport ne peut être envoyé qu’une fois la compta clôturée et signée.', 'err');
    const B = buildBundle(S, c), st = computeStatement(B), cur = c.currency, url = reportUrl(c), yacht = c.yacht || YACHT;
    const nFiles = B.expenses.filter(e => e.file_path).length + new Set(st.fuelLines.filter(f => f.file_path).map(f => f.file_path)).size;
    const subject = `${yacht} — ${c.title} — Rapport de charter signé`;
    const body = `Bonjour,

Veuillez trouver ci-dessous le rapport complet du charter « ${c.title} » (${fdate(c.start_date)} → ${fdate(c.end_date)}) à bord de ${yacht}, clôturé et signé par le client${c.client_signed_name ? ` (${c.client_signed_name})` : ''} le ${fdatetime(c.client_signed_at)}.

• APA reçu : ${money(st.apa, cur)}
• Total des dépenses : ${money(st.total, cur)}${st.fuelTotal ? `\n  dont fuel : ${money(st.fuelTotal, cur)} (${fmt(st.fuelLiters)} L)` : ''}${st.tipsAuto ? `\n  dont tips équipage : ${money(st.tipsAuto, cur)}` : ''}
• Solde APA (APA left) : ${money(st.balance, cur)} (${st.balance >= 0 ? 'à rembourser au client' : 'à régler par le client'})${st.vatEst ? `\n• TVA estimée (VAT estimated, ${charterVatRate(c)} %) sur montants payés HT : ${money(st.vatEst, cur)}\n• Solde APA après TVA estimée (APA left including VAT estimated) : ${money(st.balanceInclVat, cur)}` : ''}${B.invoices.length ? `\n• Facture émise par ${boatName()} : n° ${B.invoices.map(i => `${i.number} (${money(i.total_ttc, i.currency)})`).join(', ')}` : ''}

Rapport complet (compta signée, détail fuel et ${nFiles} facture(s) téléchargeables, export Excel et PDF) :
${url}

Bien cordialement,
${c.master_signed_name || 'Le Capitaine'}
${yacht}`;
    const m = openModal({
      title: 'Envoyer le rapport', wide: true,
      body: `<p class="muted">Le rapport est accessible uniquement par ce lien, et seulement parce que la compta est signée. Le message s’ouvre prêt à partir dans Gmail ou votre messagerie habituelle : il ne vous reste qu’à appuyer sur Envoyer.</p>
        <div class="form-grid">
          <label class="field"><span>À : représentant fiscal</span><input id="rp-to" value="${esc(c.report_to || '')}" inputmode="email" placeholder="fiscal@exemple.com"></label>
          <label class="field"><span>Cc : brokers</span><input id="rp-cc" value="${esc(c.report_cc || '')}" inputmode="email" placeholder="broker1@…, broker2@…"></label>
          <label class="field span-2"><span>Objet</span><input id="rp-subject" value="${esc(subject)}"></label>
          <label class="field span-2"><span>Message</span><textarea id="rp-body" rows="12">${esc(body)}</textarea></label>
        </div>
        <div class="row-gap wrap">
          <button class="btn btn-gold" data-send="gmail">${ICON.mail}Envoyer avec Gmail</button>
          <button class="btn btn-ghost" data-send="mailto">${ICON.mail}Autre messagerie</button>
          <button class="btn btn-ghost" data-copy>${ICON.copy}Copier le message</button>
          <button class="btn btn-ghost" data-zip>${ICON.file}Télécharger le rapport (.zip)</button>
          <a class="btn btn-ghost" href="${esc(url)}" target="_blank" rel="noopener">${ICON.eye}Voir le rapport</a>
        </div>
        <p class="small muted">Le .zip (Excel + ${nFiles} facture(s)) peut aussi être joint à l’e-mail si votre représentant le préfère.
        ${c.report_sent_at ? `<br><b>Dernier envoi préparé le ${fdatetime(c.report_sent_at)}.</b>` : ''}</p>`,
    });
    const val = id => $('#' + id, m).value;
    const saveEmails = async () => {
      const to = cleanEmails(val('rp-to')), cc = cleanEmails(val('rp-cc'));
      const bad = [...splitEmails(to), ...splitEmails(cc)].filter(x => !EMAIL_RE.test(x));
      if (bad.length) { toast('E-mail invalide : ' + bad.join(', '), 'err'); return null; }
      if (!to) { toast('Indiquez l’e-mail du représentant fiscal', 'err'); return null; }
      if (to !== (c.report_to || '') || cc !== (c.report_cc || '')) {
        try { await api.update('charters', c.id, { report_to: to, report_cc: cc || null }); } catch (e) { toast(errMsg(e), 'err'); }
      }
      return { to, cc };
    };
    $$('[data-send]', m).forEach(btn => btn.onclick = async () => {
      const r = await saveEmails(); if (!r) return;
      const u = composeUrls({ to: r.to, cc: r.cc, subject: val('rp-subject'), body: val('rp-body') });
      if (btn.dataset.send === 'gmail') window.open(u.gmail, '_blank', 'noopener'); else location.href = u.mailto;
      try { await api.update('charters', c.id, { report_sent_at: new Date().toISOString() }); } catch { /* */ }
      await reload(); setTimeout(() => { closeModal(); route(); toast('Message préparé dans votre messagerie'); }, 600);
    });
    $('[data-copy]', m).onclick = async () => {
      const txt = `À : ${cleanEmails(val('rp-to'))}\n${val('rp-cc') ? `Cc : ${cleanEmails(val('rp-cc'))}\n` : ''}Objet : ${val('rp-subject')}\n\n${val('rp-body')}`;
      try { await navigator.clipboard.writeText(txt); } catch { $('#rp-body', m).select(); document.execCommand('copy'); }
      toast('Message copié');
    };
    $('[data-zip]', m).onclick = ev => downloadReportZip(B, ev.currentTarget);
  };

  ACTIONS['close-charter'] = async b => {
    const c = charterById(b.dataset.id), st = computeStatement(buildBundle(S, c));
    const warn = [];
    if (tipsEnabled(c) && !c.charter_fee) warn.push('Le <b>charter fee</b> n’est pas renseigné : le client ne pourra saisir qu’un montant de tips personnalisé.');
    if (!S.voyages.some(v => v.charter_id === c.id)) warn.push('Aucun voyage fuel n’est lié à ce charter : aucun fuel ne sera facturé.');
    const draft = boatInvoiceDraft(buildBundle(S, c));
    if (draft.lines.length && !PROFILE.address) warn.push('La fiche du bateau (adresse de facturation) n’est pas complétée dans <b>Paramètres</b> : la facture sera incomplète.');
    if (draft.lines.length && !c.client_address) warn.push('L’adresse de facturation du client n’est pas renseignée (bouton Modifier).');
    const ok = await confirmDialog(`Clôturer la compta de « ${esc(c.title)} » ?<br><br>
      Total : <b>${money(st.total, c.currency)}</b> · Solde APA : <b>${money(st.balance, c.currency)}</b><br><br>
      Les dépenses seront verrouillées et le client pourra signer${tipsEnabled(c) ? ' (en choisissant ses tips)' : ' (sans fenêtre de tips)'}.
      ${draft.lines.length ? `<br>Une facture de <b>${esc(boatName())}</b> (${money(draft.total_ttc, c.currency)}) sera émise automatiquement.` : ''}
      ${warn.length ? `<br><br><span class="warn">⚠ ${warn.join('<br>⚠ ')}</span>` : ''}`, 'Clôturer', false);
    if (!ok) return;
    try {
      await api.update('charters', c.id, { status: 'closed', closed_at: new Date().toISOString() });
      await reload();
      let invMsg = '';
      try { const r = await ensureBoatInvoice(charterById(c.id)); if (r.inv && r.changed) invMsg = ` · facture ${r.inv.number} émise`; await reload(); }
      catch (e) { toast('Facture non émise : ' + errMsg(e), 'err'); }
      notifyCharter(c.id); route(); toast('Compta clôturée' + invMsg);
      messageModal(charterById(c.id), 'closed');
    } catch (e) { toast(errMsg(e), 'err'); }
  };

  ACTIONS.reopen = async b => {
    const c = charterById(b.dataset.id);
    const msg = c.status === 'signed'
      ? 'Rouvrir la compta ? La signature du client et les tips qu’il a choisis seront effacés : il devra signer à nouveau après une nouvelle clôture.'
      : 'Rouvrir la compta ? Le client ne pourra plus signer tant que vous ne l’aurez pas clôturée à nouveau.';
    if (!(await confirmDialog(msg, 'Rouvrir'))) return;
    try {
      await api.update('charters', c.id, { status: 'open', closed_at: null, client_signature: null, client_signed_name: null,
        client_signed_at: null, tips_pct: null, tips_amount: null, report_sent_at: null });
      await reload(); notifyCharter(c.id); route(); toast('Compta rouverte');
      return true;
    } catch (e) { toast(errMsg(e), 'err'); return false; }
  };
  // Dépense d'une compta clôturée / signée : proposer de rouvrir puis ouvrir la modification
  ACTIONS['locked-edit'] = async b => {
    const c = charterById(CURRENT), id = b.dataset.id;
    if (await ACTIONS.reopen({ dataset: { id: c.id } })) expenseForm(S.expenses.find(x => x.id === id));
  };
  // Clic sur une ligne de dépense = modifier (sauf clic sur un lien / bouton)
  document.addEventListener('click', e => {
    if (IS_CLIENT) return;
    const tr = e.target.closest('tr[data-row-exp]');
    if (!tr || e.target.closest('a,button,input,select,label')) return;
    const c = charterById(CURRENT); if (!c) return;
    if (c.status === 'open') expenseForm(S.expenses.find(x => x.id === tr.dataset.rowExp));
    else ACTIONS['locked-edit']({ dataset: { id: tr.dataset.rowExp } });
  });

  ACTIONS['master-sign'] = () => {
    const c = charterById(CURRENT);
    const m = openModal({
      title: 'Signature du capitaine', wide: true,
      body: `<label class="field"><span>Nom</span><input name="name" value="${esc(c.master_signed_name || '')}" autocomplete="name"></label>
        <div class="field"><span>Signez dans le cadre</span><canvas class="sig-pad"></canvas>
        <button type="button" class="btn btn-ghost btn-sm" data-clear>Effacer</button></div>`,
      footer: `<button class="btn btn-ghost" data-close>Annuler</button><button class="btn btn-primary" data-ok>${ICON.pen}Signer</button>`,
    });
    const pad = signaturePad($('.sig-pad', m));
    $('[data-clear]', m).onclick = () => pad.clear();
    $('[data-ok]', m).onclick = async ev => {
      const name = $('[name=name]', m).value.trim();
      if (!name || pad.isEmpty()) return toast('Nom et signature requis', 'err');
      busy(ev.currentTarget, true);
      try {
        await api.update('charters', c.id, { master_signature: pad.toDataURL(), master_signed_name: name, master_signed_at: new Date().toISOString() });
        await reload(); closeModal(); notifyCharter(c.id); route(); toast('Signature enregistrée');
      } catch (e) { toast(errMsg(e), 'err'); busy(ev.currentTarget, false); }
    };
  };

  // ---------------- Dépenses ----------------
  function expenseForm(e = {}) {
    const isNew = !e.id;
    const c = charterById(e.charter_id || CURRENT);
    const nextNum = Math.max(1, ...S.expenses.filter(x => x.charter_id === c.id).map(x => x.num || 0)) + 1;
    const opts = (list, v) => {
      const all = v && !list.includes(v) ? [v, ...list] : list;
      return all.map(x => `<option ${x === v ? 'selected' : ''}>${esc(x)}</option>`).join('');
    };
    const m = openModal({
      title: isNew ? 'Nouvelle dépense' : 'Modifier la dépense', wide: true,
      body: `<form class="form-grid" id="f-exp">
        <label class="field"><span>N°</span><input type="number" name="num" value="${esc(e.num ?? nextNum)}"></label>
        <label class="field"><span>Date</span><input type="date" name="date" value="${esc(e.date || today())}"></label>
        <label class="field"><span>Fournisseur</span><input name="company" value="${esc(e.company || '')}" list="dl-companies" autocomplete="off"></label>
        <label class="field"><span>Réf. facture</span><input name="invoice_ref" value="${esc(e.invoice_ref || '')}" placeholder="Invoice / Receipt / n°"></label>
        <label class="field"><span>Catégorie</span><select name="category">${opts(CATEGORIES, e.category || 'Food')}</select></label>
        <label class="field"><span>Paiement</span><select name="method">${opts(METHODS, e.method || 'Credit Card APA')}</select></label>
        <label class="field span-2"><span>Description</span><input name="description" value="${esc(e.description || '')}"></label>
        <label class="field"><span>Montant HT</span><input type="number" step="0.01" name="amount_ht" value="${esc(e.amount_ht ?? '')}"></label>
        <label class="field"><span>TVA</span><input type="number" step="0.01" name="vat" value="${esc(e.vat ?? '')}"></label>
        <label class="field"><span>Total TTC</span><input type="number" step="0.01" name="total" required value="${esc(e.total ?? '')}"></label>
        <div class="field"><span>Facture (PDF / photo)</span>
          ${e.file_path ? `<div class="current-file">${fileLink(e.file_path, e.file_name || 'Facture')}<label class="check small"><input type="checkbox" name="remove_file"> retirer</label></div>` : ''}
          <input type="file" name="file" accept="application/pdf,image/*,.xml,text/xml,application/xml"></div>
        <div class="span-2 ht-box">
          <label class="check"><input type="checkbox" name="ht_paid" ${toNum(e.vat_est_base) ? 'checked' : ''}>
            <span><b>Tout ou partie payé HT</b> — TVA à régulariser par le représentant fiscal (ex. facture italienne « art. 7 », autoliquidation)</span></label>
          ${e.file_path ? `<button type="button" class="btn btn-ghost btn-sm" data-analyse>${ICON.wand}Analyser la facture jointe</button>` : ''}
          <div id="ht-detect" class="small"></div>
          <div class="form-grid ht-fields ${toNum(e.vat_est_base) ? '' : 'hidden'}">
            <label class="field"><span>Montant payé HT concerné</span><input type="number" step="0.01" name="vat_est_base" value="${esc(e.vat_est_base ?? '')}"></label>
            <label class="field"><span>TVA estimée <small>(vide = calcul auto à ${fmt(charterVatRate(c), charterVatRate(c) % 1 ? 1 : 0)} %${c.vat_country ? ' — ' + esc(c.vat_country) : ''})</small></span><input type="number" step="0.01" name="vat_est" value="${esc(e.vat_est ?? '')}" placeholder="auto"></label>
          </div>
          <div id="ht-result" class="small ht-result"></div>
        </div>
        <label class="check span-2 boat-toggle"><input type="checkbox" name="billed_by_boat" ${e.billed_by_boat ? 'checked' : ''}>
          <span><b>Facturée par le bateau (${esc(boatName())})</b> — incluse dans la facture émise automatiquement à la clôture</span></label>
        <datalist id="dl-companies"><option value="${esc(boatName())}">${esc(boatName())} — le bateau (facture auto)</option>${[...new Set([...(PROFILE.suppliers || []), ...S.expenses.map(x => x.company)].filter(x => x && !isBoatName(x)))].sort().map(x => `<option value="${esc(x)}">`).join('')}</datalist>
      </form>`,
      footer: `${!isNew ? `<button class="btn btn-ghost danger mr-auto" data-del-exp>${ICON.trash}Supprimer</button>` : ''}
        <button class="btn btn-ghost" data-close>Annuler</button><button class="btn btn-primary" data-ok>Enregistrer</button>`,
    });
    const f = $('#f-exp', m);
    const delBtn = $('[data-del-exp]', m);
    if (delBtn) delBtn.onclick = () => ACTIONS['del-expense']({ dataset: { id: e.id } });
    let totalTouched = !isNew && e.total != null && (e.amount_ht == null || e.vat == null);
    f.total.addEventListener('input', () => { totalTouched = true; });
    const syncTotal = () => {
      const ht = toNum(f.amount_ht.value), vat = toNum(f.vat.value);
      if (!totalTouched && ht !== null) f.total.value = r2(ht + (vat || 0));
    };
    f.amount_ht.addEventListener('input', syncTotal); f.vat.addEventListener('input', syncTotal);
    // --- partie payée HT / TVA estimée
    const rate = charterVatRate(c);
    let baseAuto = !toNum(e.vat_est_base);   // la base suit le montant HT / TTC tant qu'on ne la modifie pas à la main
    const htResult = () => {
      const box = $('#ht-result', m); if (!box) return;
      if (!f.ht_paid.checked) { box.innerHTML = ''; return; }
      const base = toNum(f.vat_est_base.value), man = toNum(f.vat_est.value);
      const est = man != null ? man : (base ? r2(base * rate / 100) : 0);
      box.innerHTML = !base ? '<span class="warn">Indiquez le montant payé HT concerné.</span>'
        : man != null ? `TVA estimée saisie : <b>${money(est, c.currency)}</b>`
        : rate ? `TVA estimée : <b>${money(est, c.currency)}</b> <span class="muted">(${fmt(base, 2)} × ${fmt(rate, rate % 1 ? 1 : 0)} %${c.vat_country ? ' — ' + esc(c.vat_country) : ''})</span>`
        : `<span class="warn">Le taux de TVA du charter est 0 %${c.vat_country ? ' (' + esc(c.vat_country) + ')' : ''} : aucune TVA estimée. Changez le pays / taux du charter (bouton Modifier) ou saisissez la TVA estimée ci-dessus.</span>`;
      f.vat_est.placeholder = base && rate ? `auto : ${fmt(r2(base * rate / 100), 2)}` : 'auto';
    };
    const followBase = () => { if (f.ht_paid.checked && baseAuto) f.vat_est_base.value = f.amount_ht.value || f.total.value || ''; htResult(); };
    const htToggle = () => { $('.ht-fields', m).classList.toggle('hidden', !f.ht_paid.checked); followBase(); };
    f.ht_paid.addEventListener('change', htToggle);
    f.vat_est_base.addEventListener('input', () => { baseAuto = !f.vat_est_base.value; htResult(); });
    f.vat_est.addEventListener('input', htResult);
    f.total.addEventListener('input', followBase); f.amount_ht.addEventListener('input', followBase);
    htResult();
    const detectBox = $('#ht-detect', m);
    const runAnalysis = async file => {
      detectBox.innerHTML = '<span class="muted">Analyse de la facture…</span>';
      try {
        const r = await analyseInvoiceFile(file, pr => { detectBox.innerHTML = `<span class="muted">Lecture du document… ${Math.round(pr * 100)} %</span>`; });
        if (!r) { detectBox.innerHTML = '<span class="muted">Aucune mention de facturation HT détectée.</span>'; return; }
        if (r.fill) {
          const set = (name, v) => { if (v != null && v !== '' && (!f[name].value || name === 'total' || name === 'amount_ht' || name === 'vat')) f[name].value = v; };
          set('company', r.fill.company); set('invoice_ref', r.fill.invoice_ref); if (r.fill.date) f.date.value = r.fill.date;
          set('amount_ht', r.fill.amount_ht); set('vat', r.fill.vat); set('total', r.fill.total); totalTouched = true;
        }
        const isHt = r.htBase > 0 || !!r.mentions?.length;
        if (isHt) {
          f.ht_paid.checked = true; $('.ht-fields', m).classList.remove('hidden');
          if (r.htBase > 0) { f.vat_est_base.value = r.htBase; baseAuto = false; } else { baseAuto = true; followBase(); }
        }
        htResult();
        detectBox.innerHTML = `<div class="detect ${isHt ? 'is-ht' : ''}">${ICON.wand}<div>${r.summary}${r.mentions?.length ? `<br>Mention(s) : <b>${r.mentions.map(esc).join(' · ')}</b>` : ''}
          ${isHt ? `<br><span class="muted">${r.exact ? 'Part HT lue sur la facture — vérifiez-la.' : 'La part HT reprend le total hors TVA : si la facture mélange lignes HT et TTC, ne gardez que les lignes HT.'}</span>` : ''}</div></div>`;
      } catch (err) { detectBox.innerHTML = `<span class="muted">Lecture impossible : ${esc(errMsg(err))}. Cochez « payé HT » vous-même si besoin.</span>`; }
    };
    f.file.addEventListener('change', () => { const file = f.file.files[0]; detectBox.innerHTML = ''; if (file) runAnalysis(file); });
    const anaBtn = $('[data-analyse]', m);
    if (anaBtn) anaBtn.onclick = async () => {
      try {
        const url = api.fileUrl(e.file_path); const r = await fetch(url); if (!r.ok) throw new Error('fichier introuvable');
        const blob = await r.blob();
        runAnalysis(new File([blob], e.file_name || e.file_path.split('/').pop(), { type: blob.type }));
      } catch (err) { detectBox.innerHTML = `<span class="muted">Lecture impossible : ${esc(errMsg(err))}</span>`; }
    };
    f.company.addEventListener('input', () => { if (isBoatName(f.company.value)) f.billed_by_boat.checked = true; });
    f.company.addEventListener('change', () => { f.billed_by_boat.checked = isBoatName(f.company.value) || (f.billed_by_boat.checked && !f.company.value); });
    $('[data-ok]', m).onclick = async ev => {
      if (!f.reportValidity()) return;
      const d = new FormData(f);
      const row = { charter_id: c.id, num: toNum(d.get('num')), date: d.get('date') || null, company: d.get('company') || null,
        invoice_ref: d.get('invoice_ref') || null, category: d.get('category'), description: d.get('description') || null,
        amount_ht: toNum(d.get('amount_ht')), vat: toNum(d.get('vat')), total: r2(toNum(d.get('total')) || 0), method: d.get('method'),
        billed_by_boat: !!d.get('billed_by_boat'),
        vat_est_base: d.get('ht_paid') ? toNum(d.get('vat_est_base')) : null,
        vat_est: d.get('ht_paid') ? toNum(d.get('vat_est')) : null };
      if (d.get('ht_paid') && !row.vat_est_base) return toast('Indiquez le montant payé HT concerné', 'err');
      busy(ev.currentTarget, true);
      try {
        const file = f.file.files[0];
        let oldPath = null;
        if (file) { Object.assign(row, await api.upload(file, `expenses/${c.id}`)); oldPath = e.file_path; }
        else if (d.get('remove_file')) { row.file_path = null; row.file_name = null; oldPath = e.file_path; }
        if (isNew) await api.insert('expenses', row); else await api.update('expenses', e.id, row);
        if (oldPath) api.removeFile(oldPath).catch(() => {});
        await reload(); closeModal(); notifyCharter(c.id); route(); toast('Dépense enregistrée');
      } catch (err) { toast(errMsg(err), 'err'); busy(ev.currentTarget, false); }
    };
  }
  ACTIONS['new-expense'] = () => expenseForm({ charter_id: CURRENT });
  ACTIONS['edit-expense'] = b => expenseForm(S.expenses.find(x => x.id === b.dataset.id));
  ACTIONS['del-expense'] = async b => {
    const e = S.expenses.find(x => x.id === b.dataset.id);
    if (!(await confirmDialog(`Supprimer la dépense « ${esc(e.company || '')} ${esc(e.description || '')} » (${money(e.total)}) ?`, 'Supprimer'))) return;
    try { await api.remove('expenses', e.id); if (e.file_path) api.removeFile(e.file_path).catch(() => {});
      await reload(); notifyCharter(e.charter_id); route(); toast('Dépense supprimée'); }
    catch (err) { toast(errMsg(err), 'err'); }
  };

  // ---------------- Cash ----------------
  ACTIONS['new-cash'] = () => {
    const m = openModal({
      title: 'Retrait d’espèces (APA)',
      body: `<form class="form-grid" id="f-cash"><label class="field"><span>Date</span><input type="date" name="date" value="${today()}"></label>
        <label class="field"><span>Lieu</span><input name="location"></label>
        <label class="field span-2"><span>Montant</span><input type="number" step="0.01" name="amount" required></label></form>`,
      footer: `<button class="btn btn-ghost" data-close>Annuler</button><button class="btn btn-primary" data-ok>Enregistrer</button>`,
    });
    $('[data-ok]', m).onclick = async ev => {
      const f = $('#f-cash', m); if (!f.reportValidity()) return;
      busy(ev.currentTarget, true);
      try { await api.insert('cash_withdrawals', { charter_id: CURRENT, date: f.date.value || null, location: f.location.value || null, amount: toNum(f.amount.value) });
        await reload(); closeModal(); notifyCharter(CURRENT); route(); }
      catch (e) { toast(errMsg(e), 'err'); busy(ev.currentTarget, false); }
    };
  };
  ACTIONS['del-cash'] = async b => {
    if (!(await confirmDialog('Supprimer ce retrait ?', 'Supprimer'))) return;
    try { await api.remove('cash_withdrawals', b.dataset.id); await reload(); notifyCharter(CURRENT); route(); } catch (e) { toast(errMsg(e), 'err'); }
  };

  // ---------------- Lier un voyage au charter ----------------
  ACTIONS['link-voyage'] = () => {
    const free = S.voyages.filter(v => !v.charter_id).sort(byPos).reverse();
    const m = openModal({
      title: 'Lier un voyage du fuel booklet',
      body: `<label class="field"><span>Voyage</span><select name="v">${free.map(v => `<option value="${v.id}">${esc(v.num ?? '')} · ${esc(v.description || '')} (${fdate(v.start_date)}) — ${fmt(consumption(v))} L</option>`).join('')}</select></label>`,
      footer: `<button class="btn btn-ghost" data-close>Annuler</button><button class="btn btn-primary" data-ok>Lier</button>`,
    });
    $('[data-ok]', m).onclick = async () => {
      try { await api.update('voyages', $('[name=v]', m).value, { charter_id: CURRENT, is_charter: true }); await reload(); closeModal(); notifyCharter(CURRENT); route(); }
      catch (e) { toast(errMsg(e), 'err'); }
    };
  };
  ACTIONS['unlink-voyage'] = async b => {
    if (!(await confirmDialog('Délier ce voyage du charter ? Le fuel ne sera plus facturé dans cette compta (le voyage reste dans le fuel booklet).', 'Délier'))) return;
    try { await api.update('voyages', b.dataset.id, { charter_id: null }); await reload(); notifyCharter(CURRENT); route(); } catch (e) { toast(errMsg(e), 'err'); }
  };
  ACTIONS['voyage-now'] = async b => {
    busy(b, true);
    try { const c = charterById(CURRENT); await createCharterVoyage(c); await reload(); route(); toast('Voyage fuel créé'); }
    catch (e) { toast(errMsg(e), 'err'); busy(b, false); }
  };
  ACTIONS['new-voyage-charter'] = () => {
    const c = charterById(CURRENT);
    voyageForm({ charter_id: c.id, is_charter: true, description: `Charter ${c.client_name || c.title}`, start_date: c.start_date, end_date: c.end_date });
  };

  // ==================================================================
  //  FUEL BOOKLET
  // ==================================================================
  function viewFuel(tab) {
    const used = lotUsage(S);
    const tabs = [['voyages', 'Voyages & consommation'], ['bunkers', 'Soutages'], ['stock', 'État des soutes']];
    let inner = '';
    if (tab === 'bunkers') inner = fuelBunkers(used);
    else if (tab === 'stock') inner = fuelStock(used);
    else inner = fuelVoyages();
    $('#app').innerHTML = shell('fuel', `
      <div class="page-head"><div><div class="eyebrow">Fuel booklet</div><h1>Gestion du fuel</h1></div>
        <div class="head-actions">
          <button class="btn btn-ghost" data-act="xlsx-fuel">${ICON.sheet}Excel</button>
          <button class="btn btn-ghost" data-act="new-bunker">${ICON.plus}Soutage</button>
          <button class="btn btn-primary" data-act="new-voyage">${ICON.plus}Voyage</button></div></div>
      ${stockStrip(used)}
      <div class="tabs">${tabs.map(([k, l]) => `<a href="#/fuel/${k}" class="${tab === k ? 'on' : ''}">${l}</a>`).join('')}</div>
      ${inner}`);
  }
  ACTIONS['xlsx-fuel'] = () => exportFuelXlsx(S);

  function stockFigures(used) {
    const theo = sum(S.bunkers, b => (toNum(b.quantity) || 0) - (used[b.id] || 0));
    const lastV = S.voyages.slice().sort(byPos).reverse().find(v => toNum(v.rob_after) !== null);
    const lastB = lastV ? S.bunkers.filter(b => b.date && lastV.end_date && b.date > lastV.end_date) : [];
    const rob = lastV ? lastV.rob_after + sum(lastB, b => b.quantity) : null;
    const valHT = sum(S.bunkers, b => Math.max(0, (toNum(b.quantity) || 0) - (used[b.id] || 0)) * (toNum(b.price_ht_l) || 0));
    const valTTC = sum(S.bunkers, b => Math.max(0, (toNum(b.quantity) || 0) - (used[b.id] || 0)) * ttcL(b));
    return { theo, lastV, lastB, rob, valHT, valTTC };
  }

  function stockStrip(used) {
    const f = stockFigures(used);
    const gap = f.rob != null ? f.rob - f.theo : null;
    return `<div class="kpis">
      <div class="kpi"><div class="kpi-label">Stock théorique (lots)</div><div class="kpi-value">${fmt(f.theo)} L</div></div>
      <div class="kpi"><div class="kpi-label">Dernier ROB relevé</div><div class="kpi-value">${f.rob != null ? fmt(f.rob) + ' L' : '—'}</div>
        <div class="kpi-sub">${f.lastV ? `${esc(f.lastV.description || '')} · ${fdate(f.lastV.end_date)}${f.lastB.length ? ` + ${f.lastB.length} soutage(s)` : ''}` : ''}</div></div>
      <div class="kpi ${gap != null && Math.abs(gap) > 200 ? 'kpi-warn' : ''}"><div class="kpi-label">Écart ROB / théorique</div><div class="kpi-value">${gap != null ? (gap > 0 ? '+' : '') + fmt(gap) + ' L' : '—'}</div></div>
      <div class="kpi"><div class="kpi-label">Valeur du stock</div><div class="kpi-value">${money(f.valHT)}</div><div class="kpi-sub">HT · ${money(f.valTTC)} TTC</div></div>
    </div>`;
  }

  function fuelBunkers(used) {
    const list = S.bunkers.slice().sort(byLot).reverse();
    return `<section class="panel"><div class="table-wrap"><table class="table table-cards">
      <thead><tr><th>N°</th><th>Date</th><th>Lieu</th><th class="num">Quantité</th><th>Type</th><th class="num">HT / L</th><th class="num">TVA / L</th>
        <th class="num">TTC / L</th><th class="num">Total</th><th>Facture</th><th>Restant</th><th></th></tr></thead>
      <tbody>${list.map(b => {
        const rem = (toNum(b.quantity) || 0) - (used[b.id] || 0), pct = b.quantity ? Math.max(0, Math.min(100, rem / b.quantity * 100)) : 0;
        return `<tr>
        <td data-label="N°">${esc(b.num ?? '')}</td><td data-label="Date">${fdate(b.date)}</td>
        <td data-label="Lieu">${esc([b.town, b.country].filter(Boolean).join(', '))}</td>
        <td data-label="Quantité" class="num">${fmt(b.quantity)} L</td><td data-label="Type" class="small">${esc(b.fuel_type || '')}</td>
        <td data-label="HT / L" class="num">${fmt(b.price_ht_l, 4)}</td><td data-label="TVA / L" class="num">${fmt(b.vat_l, 4)}</td>
        <td data-label="TTC / L" class="num strong">${fmt(ttcL(b), 4)}<br>${toNum(b.vat_l) > 0 ? '<span class="chip">TTC</span>' : '<span class="chip chip-warn" title="Payé HT : le client paiera la TVA estimée au taux du charter">HT</span>'}</td>
        <td data-label="Total" class="num">${money(b.total_price ?? r2(b.quantity * ttcL(b)))}</td>
        <td data-label="Facture" class="files">${b.file_path ? fileLink(b.file_path, b.invoice_ref || 'Facture') : `<span class="muted">${esc(b.invoice_ref || '')}</span>`}</td>
        <td data-label="Restant"><div class="meter ${rem < -0.5 ? 'is-neg' : ''}"><span style="width:${pct}%"></span></div><span class="small ${rem < -0.5 ? 'neg' : ''}">${fmt(rem)} L</span></td>
        <td class="actions"><button class="icon-btn" data-act="edit-bunker" data-id="${b.id}" title="Modifier">${ICON.edit}</button></td></tr>`;
      }).join('')}</tbody></table></div></section>`;
  }

  function fuelStock(used) {
    const open = S.bunkers.slice().sort(byLot).filter(b => Math.abs((toNum(b.quantity) || 0) - (used[b.id] || 0)) > 0.5);
    return `<section class="panel"><div class="panel-head"><h2>Lots encore en soute</h2></div>
      ${open.length ? open.map(b => {
        const rem = b.quantity - (used[b.id] || 0);
        return `<div class="lot-row"><div><strong>Soutage #${esc(b.num ?? '')}</strong> <span class="muted">· ${fdate(b.date)} · ${esc(b.town || '')}</span>
          <div class="small muted">${fmt(ttcL(b), 4)} €/L TTC · ${fmt(b.price_ht_l, 4)} €/L HT</div></div>
          <div class="lot-meter"><div class="meter ${rem < 0 ? 'is-neg' : ''}"><span style="width:${Math.max(0, Math.min(100, rem / b.quantity * 100))}%"></span></div>
          <b class="${rem < 0 ? 'neg' : ''}">${fmt(rem)} L</b> <span class="muted small">/ ${fmt(b.quantity)} L</span></div></div>`;
      }).join('') : '<p class="muted">Tous les lots sont consommés.</p>'}
      ${open.some(b => b.quantity - (used[b.id] || 0) < 0) ? `<p class="small muted">Un restant négatif signifie que plus de litres ont été répartis sur ce lot qu’il n’en contenait : corrigez la répartition des voyages concernés.</p>` : ''}
      <p class="small muted">Règle de valorisation : consommation charter au prix TTC, hors charter au prix HT.</p></section>`;
  }

  function fuelVoyages() {
    const lotNum = id => S.bunkers.find(b => b.id === id)?.num;
    const list = S.voyages.slice().sort(byPos).reverse();
    return `<section class="panel"><div class="table-wrap"><table class="table table-cards">
      <thead><tr><th>N°</th><th>Description</th><th>Dates</th><th class="num">ROB départ</th><th class="num">ROB arrivée</th>
        <th class="num">Conso.</th><th>Lots utilisés</th><th class="num">L/h</th><th class="num">Valeur</th><th></th></tr></thead>
      <tbody>${list.map(v => {
        const cons = consumption(v), al = S.allocations.filter(a => a.voyage_id === v.id);
        const dist = sum(al, a => a.liters);
        const ch = v.charter_id && charterById(v.charter_id);
        const val = sum(al, a => { const b = S.bunkers.find(x => x.id === a.bunker_id); return b ? a.liters * (v.is_charter ? ttcL(b) : (toNum(b.price_ht_l) || 0)) : 0; });
        const mismatch = cons != null && Math.abs(dist - cons) > 0.5;
        return `<tr class="${v.is_charter ? 'row-charter' : ''}">
          <td data-label="N°">${esc(v.num ?? '')}</td>
          <td data-label="Description">${esc(v.description || '—')}${ch ? `<br><a class="small" href="#/charter/${ch.id}">${ICON.link} ${esc(ch.title)}</a>` : (v.is_charter ? ' <span class="chip chip-gold">Charter</span>' : '')}</td>
          <td data-label="Dates" class="small">${fdate(v.start_date)}${v.end_date && v.end_date !== v.start_date ? ' → ' + fdate(v.end_date) : ''}</td>
          <td data-label="ROB départ" class="num">${fmt(v.rob_before)}</td><td data-label="ROB arrivée" class="num">${fmt(v.rob_after)}</td>
          <td data-label="Conso." class="num strong">${cons != null ? fmt(cons) + ' L' : ''}</td>
          <td data-label="Lots utilisés">${al.map(a => `<span class="chip">#${esc(lotNum(a.bunker_id))} · ${fmt(a.liters)}</span>`).join(' ')}
            ${mismatch ? `<span class="chip chip-warn" title="La répartition (${fmt(dist)} L) ne correspond pas à la consommation">≠ ${fmt(dist)}</span>` : ''}
            ${cons != null && cons > 0 && !al.length ? '<span class="chip chip-warn">à répartir</span>' : ''}</td>
          <td data-label="L/h" class="num">${v.hours && cons ? fmt(cons / v.hours, 1) : ''}</td>
          <td data-label="Valeur" class="num">${val ? money(val) : ''}</td>
          <td class="actions"><button class="icon-btn" data-act="edit-voyage" data-id="${v.id}" title="Modifier">${ICON.edit}</button></td></tr>`;
      }).join('')}</tbody></table></div></section>`;
  }

  // ---------------- Formulaire soutage ----------------
  function bunkerForm(b = {}) {
    const isNew = !b.id;
    const next = Math.max(0, ...S.bunkers.map(x => x.num || 0)) + 1;
    const rate = b.price_ht_l ? r2((b.vat_l || 0) / b.price_ht_l * 100) : '';
    const m = openModal({
      title: isNew ? 'Nouveau soutage' : `Soutage #${b.num ?? ''}`, wide: true,
      body: `<form class="form-grid" id="f-bunk">
        <label class="field"><span>N°</span><input type="number" name="num" value="${esc(b.num ?? next)}"></label>
        <label class="field"><span>Date</span><input type="date" name="date" value="${esc(b.date || today())}"></label>
        <label class="field"><span>Pays</span><input name="country" value="${esc(b.country || '')}"></label>
        <label class="field"><span>Port</span><input name="town" value="${esc(b.town || '')}"></label>
        <label class="field"><span>Quantité (L)</span><input type="number" step="1" name="quantity" required value="${esc(b.quantity ?? '')}"></label>
        <label class="field"><span>Type de fuel</span><input name="fuel_type" value="${esc(b.fuel_type || 'GASOIL B7 25μ / KATHON')}"></label>
        <label class="field"><span>Prix HT / L</span><input type="number" step="0.0001" name="price_ht_l" required value="${esc(b.price_ht_l ?? '')}"></label>
        <label class="field"><span>Taux TVA %</span><input type="number" step="0.01" name="vat_rate" value="${esc(rate)}" placeholder="0, 20, 21, 22, 24…"></label>
        <label class="field"><span>TVA / L</span><input type="number" step="0.0001" name="vat_l" value="${esc(b.vat_l ?? 0)}"></label>
        <label class="field"><span>Total facturé <small>(vide = auto)</small></span><input type="number" step="0.01" name="total_price" value="${esc(b.total_price ?? '')}"></label>
        <label class="field"><span>N° de facture</span><input name="invoice_ref" value="${esc(b.invoice_ref || '')}"></label>
        <div class="field"><span>Facture (PDF / photo)</span>
          ${b.file_path ? `<div class="current-file">${fileLink(b.file_path, b.file_name || 'Facture')}</div>` : ''}
          <input type="file" name="file" accept="application/pdf,image/*,.xml,text/xml,application/xml"></div>
        <div class="span-2 small" id="bunk-detect"></div>
        <p class="small muted span-2" id="bunk-preview"></p>
      </form>`,
      footer: `${!isNew ? `<button class="btn btn-ghost danger mr-auto" data-del>${ICON.trash}Supprimer</button>` : ''}
        <button class="btn btn-ghost" data-close>Annuler</button><button class="btn btn-primary" data-ok>Enregistrer</button>`,
    });
    const f = $('#f-bunk', m);
    const preview = () => {
      const ht = toNum(f.price_ht_l.value) || 0, vat = toNum(f.vat_l.value) || 0, q = toNum(f.quantity.value) || 0;
      $('#bunk-preview', m).innerHTML = (vat > 0
        ? `<span class="chip">Payé TTC</span> Prix TTC / L : ${fmt(ht + vat, 4)} € · facturé au client ${fmt(r3(ht + vat), 3)} €/L`
        : `<span class="chip chip-warn">Payé HT</span> facturé au client ${fmt(r3(ht), 3)} €/L HT + <b>TVA estimée</b> au taux du pays du charter`)
        + ` · Total calculé : ${money(r2(q * (ht + vat)))}`;
    };
    const applyRate = () => { const ht = toNum(f.price_ht_l.value), rt = toNum(f.vat_rate.value); if (ht != null && rt != null) f.vat_l.value = Math.round(ht * rt / 100 * 1e5) / 1e5; preview(); };
    f.vat_rate.addEventListener('input', applyRate);
    f.price_ht_l.addEventListener('input', applyRate);
    f.vat_l.addEventListener('input', preview); f.quantity.addEventListener('input', preview); preview();
    f.file.addEventListener('change', async () => {
      const file = f.file.files[0], box = $('#bunk-detect', m); box.innerHTML = ''; if (!file) return;
      box.innerHTML = '<span class="muted">Analyse de la facture…</span>';
      try {
        const r = await analyseInvoiceFile(file);
        if (!r) { box.innerHTML = /\.(pdf|xml)$/i.test(file.name) ? '<span class="muted">Aucune mention HT détectée : vérifiez la TVA par litre.</span>' : '<span class="muted">Photo : pas de lecture automatique.</span>'; return; }
        const fuel = (r.lines || []).filter(l => l.qty && l.unitPrice).sort((a, b) => b.amount - a.amount)[0];
        if (fuel) { f.quantity.value = fuel.qty; f.price_ht_l.value = fuel.unitPrice; f.vat_rate.value = fuel.rate || 0; }
        if (r.fill) { if (r.fill.invoice_ref) f.invoice_ref.value = r.fill.invoice_ref; if (r.fill.date) f.date.value = r.fill.date; if (r.fill.total) f.total_price.value = r.fill.total; }
        const ht = !!(r.mentions?.length || r.htBase > 0) && !(fuel && fuel.rate > 0);
        if (ht) { f.vat_rate.value = 0; f.vat_l.value = 0; }
        applyRate();
        box.innerHTML = `<div class="detect ${ht ? 'is-ht' : ''}">${ICON.wand}<div>${r.summary}${r.mentions?.length ? `<br>Mention(s) : <b>${r.mentions.map(esc).join(' · ')}</b>` : ''}
          <br><b>${ht ? 'Soutage payé HT' : 'Soutage payé TTC'}</b>${ht ? ' : le client paiera la TVA estimée au taux du pays du charter.' : ''} Vérifiez les champs remplis.</div></div>`;
      } catch (err) { box.innerHTML = `<span class="muted">Lecture impossible : ${esc(errMsg(err))}</span>`; }
    });
    $('[data-ok]', m).onclick = async ev => {
      if (!f.reportValidity()) return;
      const d = new FormData(f);
      const row = { num: toNum(d.get('num')), date: d.get('date') || null, country: d.get('country') || null, town: d.get('town') || null,
        quantity: toNum(d.get('quantity')), fuel_type: d.get('fuel_type') || null, price_ht_l: toNum(d.get('price_ht_l')) || 0,
        vat_l: toNum(d.get('vat_l')) || 0, total_price: toNum(d.get('total_price')), invoice_ref: d.get('invoice_ref') || null };
      busy(ev.currentTarget, true);
      try {
        const file = f.file.files[0];
        if (file) Object.assign(row, await api.upload(file, 'bunkers'));
        if (isNew) await api.insert('bunkers', row); else await api.update('bunkers', b.id, row);
        if (file && b.file_path) api.removeFile(b.file_path).catch(() => {});
        await reload(); closeModal(); notifyAffectedCharters(b.id); route(); toast('Soutage enregistré');
      } catch (e) { toast(errMsg(e), 'err'); busy(ev.currentTarget, false); }
    };
    const del = $('[data-del]', m);
    if (del) del.onclick = async () => {
      if (S.allocations.some(a => a.bunker_id === b.id)) return toast('Ce soutage est utilisé dans des voyages : retirez d’abord ses répartitions.', 'err');
      if (!(await confirmDialog(`Supprimer le soutage #${esc(b.num ?? '')} ?`, 'Supprimer'))) return;
      try { await api.remove('bunkers', b.id); if (b.file_path) api.removeFile(b.file_path).catch(() => {}); await reload(); route(); }
      catch (e) { toast(errMsg(e), 'err'); }
    };
  }
  function notifyAffectedCharters(bunkerId) {
    const vids = new Set(S.allocations.filter(a => a.bunker_id === bunkerId).map(a => a.voyage_id));
    new Set(S.voyages.filter(v => vids.has(v.id) && v.charter_id).map(v => v.charter_id)).forEach(notifyCharter);
  }
  ACTIONS['new-bunker'] = () => bunkerForm();
  ACTIONS['edit-bunker'] = b => bunkerForm(S.bunkers.find(x => x.id === b.dataset.id));

  // ---------------- Formulaire voyage + répartition par lot ----------------
  function voyageForm(v = {}) {
    const isNew = !v.id;
    const sorted = S.voyages.slice().sort(byPos);
    const last = sorted.slice().reverse().find(x => toNum(x.rob_after) !== null);
    if (isNew && v.rob_before == null && last) v.rob_before = suggestRobAt(v.start_date || today());
    let rows = isNew ? [] : S.allocations.filter(a => a.voyage_id === v.id).map(a => ({ bunker_id: a.bunker_id, liters: a.liters }));
    const othersUsed = {};
    S.allocations.filter(a => a.voyage_id !== v.id).forEach(a => { othersUsed[a.bunker_id] = (othersUsed[a.bunker_id] || 0) + (toNum(a.liters) || 0); });
    const lots = S.bunkers.slice().sort(byLot);
    const lotLabel = b => `#${b.num ?? ''} · ${fdate(b.date)} · ${b.town || ''} — reste ${fmt(b.quantity - (othersUsed[b.id] || 0))} L`;
    const charters = S.charters.slice().sort((a, b) => String(b.start_date || '').localeCompare(String(a.start_date || '')));
    const m = openModal({
      title: isNew ? 'Nouveau voyage' : `Voyage ${v.num ?? ''}`, wide: true,
      body: `<form class="form-grid" id="f-voy">
        <label class="field"><span>N°</span><input type="number" name="num" value="${esc(v.num ?? (Math.max(0, ...S.voyages.map(x => x.num || 0)) + 1))}"></label>
        <label class="field"><span>Heures de route</span><input type="number" step="0.1" name="hours" value="${esc(v.hours ?? '')}"></label>
        <label class="field span-2"><span>Description</span><input name="description" required value="${esc(v.description || '')}" placeholder="Positioning : Toulon (France) > Genova (Italy)"></label>
        <label class="field"><span>Début</span><input type="date" name="start_date" value="${esc(v.start_date || '')}"></label>
        <label class="field"><span>Fin</span><input type="date" name="end_date" value="${esc(v.end_date || '')}"></label>
        <label class="field"><span>ROB au départ (L)</span><input type="number" step="1" name="rob_before" value="${esc(v.rob_before ?? '')}"></label>
        <label class="field"><span>ROB à l’arrivée (L)</span><input type="number" step="1" name="rob_after" value="${esc(v.rob_after ?? '')}"></label>
        <label class="check span-2"><input type="checkbox" name="is_charter" ${v.is_charter ? 'checked' : ''}> <span>Voyage charter (valorisé TTC — sinon HT)</span></label>
        <label class="field span-2"><span>Facturer ce fuel dans la compta du charter</span><select name="charter_id"><option value="">— Aucun —</option>
          ${charters.map(c => `<option value="${c.id}" ${c.id === v.charter_id ? 'selected' : ''}>${esc(c.title)}</option>`).join('')}</select></label>
      </form>
      <div class="alloc">
        <div class="alloc-head"><div><strong>Répartition par soutage</strong><div class="small muted" id="alloc-status"></div></div>
          <div class="row-gap"><button class="btn btn-ghost btn-sm" data-auto>${ICON.wand}Répartir auto (FIFO)</button>
          <button class="btn btn-ghost btn-sm" data-add>${ICON.plus}Ligne</button></div></div>
        <div id="alloc-rows"></div>
      </div>`,
      footer: `${!isNew ? `<button class="btn btn-ghost danger mr-auto" data-del>${ICON.trash}Supprimer</button>` : ''}
        <button class="btn btn-ghost" data-close>Annuler</button><button class="btn btn-primary" data-ok>Enregistrer</button>`,
    });
    const f = $('#f-voy', m);
    let robTouched = !isNew;
    f.rob_before.addEventListener('input', () => { robTouched = true; });
    f.start_date.addEventListener('change', () => { if (!robTouched) { const r = suggestRobAt(f.start_date.value || today()); if (r != null) { f.rob_before.value = r; status(); } } });
    const cons = () => { const a = toNum(f.rob_before.value), b = toNum(f.rob_after.value); return a != null && b != null ? a - b : null; };
    const drawRows = () => {
      $('#alloc-rows', m).innerHTML = rows.map((r, i) => `<div class="alloc-row">
        <select data-i="${i}" data-k="bunker_id">${lots.map(b => `<option value="${b.id}" ${b.id === r.bunker_id ? 'selected' : ''}>${esc(lotLabel(b))}</option>`).join('')}</select>
        <input type="number" step="1" data-i="${i}" data-k="liters" value="${esc(r.liters)}" aria-label="Litres"><span class="muted small">L</span>
        <button class="icon-btn danger" data-rm="${i}" title="Retirer">${ICON.trash}</button></div>`).join('') || '<p class="muted small">Aucune répartition.</p>';
      status();
    };
    const status = () => {
      const c = cons(), d = sum(rows, r => r.liters);
      const el = $('#alloc-status', m);
      if (c == null) { el.textContent = 'Renseignez les ROB pour calculer la consommation.'; el.className = 'small muted'; return; }
      const ok = Math.abs(c - d) < 0.5;
      el.textContent = `Consommation ${fmt(c)} L · réparti ${fmt(d)} L${ok ? ' ✓' : ` · reste ${fmt(c - d)} L à répartir`}`;
      el.className = 'small ' + (ok ? 'pos' : 'warn');
    };
    $('#alloc-rows', m).addEventListener('input', e => { const i = e.target.dataset.i; if (i == null) return;
      rows[i][e.target.dataset.k] = e.target.dataset.k === 'liters' ? toNum(e.target.value) : e.target.value; status(); });
    $('#alloc-rows', m).addEventListener('change', e => { const i = e.target.dataset.i; if (i != null && e.target.dataset.k === 'bunker_id') rows[i].bunker_id = e.target.value; });
    $('#alloc-rows', m).addEventListener('click', e => { const b = e.target.closest('[data-rm]'); if (b) { rows.splice(+b.dataset.rm, 1); drawRows(); } });
    $('[data-add]', m).onclick = () => {
      const c = cons(), d = sum(rows, r => r.liters);
      const firstOpen = lots.find(b => b.quantity - (othersUsed[b.id] || 0) > 0) || lots[0];
      if (!firstOpen) return toast('Ajoutez d’abord un soutage', 'err');
      rows.push({ bunker_id: firstOpen.id, liters: c != null ? Math.max(0, c - d) : 0 }); drawRows();
    };
    $('[data-auto]', m).onclick = () => {
      const c = cons(); if (c == null || c <= 0) return toast('Renseignez ROB départ et arrivée', 'err');
      const res = autoAllocate(S, v.id, f.end_date.value || null, c);
      rows = res.out; drawRows();
      if (res.left > 0.5) toast(`Stock insuffisant dans les lots : ${fmt(res.left)} L non répartis`, 'err');
    };
    f.rob_before.addEventListener('input', status); f.rob_after.addEventListener('input', status);
    f.charter_id.addEventListener('change', () => { if (f.charter_id.value) f.is_charter.checked = true; });
    drawRows();

    $('[data-ok]', m).onclick = async ev => {
      if (!f.reportValidity()) return;
      const d = new FormData(f);
      const row = { num: toNum(d.get('num')), description: d.get('description').trim(), start_date: d.get('start_date') || null,
        end_date: d.get('end_date') || null, rob_before: toNum(d.get('rob_before')), rob_after: toNum(d.get('rob_after')),
        hours: toNum(d.get('hours')), is_charter: !!d.get('is_charter'), charter_id: d.get('charter_id') || null };
      const clean = rows.filter(r => r.bunker_id && toNum(r.liters));
      const c = cons();
      if (c != null && clean.length && Math.abs(c - sum(clean, r => r.liters)) > 0.5 &&
        !(await confirmDialog(`La répartition (${fmt(sum(clean, r => r.liters))} L) ne correspond pas à la consommation (${fmt(c)} L). Enregistrer quand même ?`, 'Enregistrer', false))) return;
      busy(ev.currentTarget, true);
      try {
        if (isNew) row.position = await placeVoyage(row.start_date);
        const saved = isNew ? await api.insert('voyages', row) : await api.update('voyages', v.id, row);
        await api.removeWhere('allocations', 'voyage_id', saved.id);
        await api.insertMany('allocations', clean.map(r => ({ voyage_id: saved.id, bunker_id: r.bunker_id, liters: toNum(r.liters) })));
        await reload(); closeModal();
        [v.charter_id, row.charter_id].filter(Boolean).forEach(notifyCharter);
        route(); toast('Voyage enregistré');
      } catch (e) { toast(errMsg(e), 'err'); busy(ev.currentTarget, false); }
    };
    const del = $('[data-del]', m);
    if (del) del.onclick = async () => {
      if (!(await confirmDialog(`Supprimer le voyage « ${esc(v.description || '')} » et sa répartition ?`, 'Supprimer'))) return;
      try { await api.remove('voyages', v.id); await reload(); notifyCharter(v.charter_id); route(); } catch (e) { toast(errMsg(e), 'err'); }
    };
  }
  ACTIONS['new-voyage'] = () => voyageForm();
  ACTIONS['edit-voyage'] = b => voyageForm(S.voyages.find(x => x.id === b.dataset.id));

  // ==================================================================
  //  Démarrage
  // ==================================================================
  const QS = new URLSearchParams(location.search);
  const TOKEN = QS.get('c'), RTOKEN = QS.get('r');
  const IS_CLIENT = !!(TOKEN || RTOKEN);

  // Installation sur l'écran d'accueil (PWA)
  if (IS_CLIENT) {
    // Le client garde SON lien (avec le code) s'il ajoute la page à son écran d'accueil
    document.getElementById('app-manifest')?.remove();
  } else if ('serviceWorker' in navigator && location.protocol === 'https:') {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }

  (async () => {
    try {
      await api.init();
      if (RTOKEN) await runClient(RTOKEN, 'report'); else if (TOKEN) await runClient(TOKEN); else await runAdmin();
    } catch (e) {
      const offline = !navigator.onLine || /Chargement impossible|Failed to fetch|NetworkError/i.test(errMsg(e));
      $('#app').innerHTML = `<div class="center-screen"><div class="card narrow"><div class="brand-mark">${ICON.anchor}</div>
        <h2>${offline ? 'Pas de connexion' : 'Erreur'}</h2>
        <p class="muted">${offline ? 'Impossible de joindre le serveur. Vérifiez la connexion internet du bord (4G / Wi-Fi) puis réessayez. — No connection: please check your internet access and retry.' : esc(errMsg(e))}</p>
        <button class="btn btn-primary btn-block" onclick="location.reload()">Réessayer / Retry</button></div></div>`;
    }
  })();
})();
