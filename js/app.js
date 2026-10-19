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
  const VAT_COUNTRIES = [['France', 20], ['Italie', 22], ['Espagne', 21], ['Croatie', 13], ['Grèce', 0], ['Monaco', 20]];
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
      vatEstimated: 'TVA estimée', apaLeftInclVat: 'Solde APA après TVA estimée', toValidate: 'À valider par le représentant fiscal', vatEstNote: 'Montant payé HT : la TVA sera facturée par le représentant fiscal après envoi du rapport',
      htChip: 'HT', apaLeft: 'Solde APA',
      boatInvoice: 'Facture', boatInvoicePending: 'Facture émise à la clôture',
      reportUnavailable: 'Ce rapport n’est pas disponible (compta non signée ou lien invalide).', preparingZip: 'Préparation du zip…',
      tabHome: 'Aujourd’hui', homeDay: 'Jour', homeOf: 'sur', programmeOfDay: 'Programme du jour', meals: 'Repas', lunch: 'Déjeuner', dinner: 'Dîner',
      captainNote: 'Le mot du capitaine', freeDay: 'Journée au mouillage / à quai — profitez !', wxTitle: 'Météo', wxWind: 'Vent', wxGust: 'rafales', wxSea: 'Mer', wxWaves: 'vagues',
      wxRain: 'Pluie', wxWater: 'Eau', wxSun: 'Soleil', wxUv: 'UV', wxNA: 'Prévisions disponibles environ 15 jours avant la date.', wxLoading: 'Chargement de la météo…',
      startsIn: 'Votre charter commence dans {n} jour(s) — aperçu du premier jour', menuSoon: 'Menu communiqué prochainement',
      wx_sunny: 'Ensoleillé', wx_partly: 'Éclaircies', wx_cloudy: 'Nuageux', wx_fog: 'Brouillard', wx_drizzle: 'Bruine', wx_rain: 'Pluie', wx_snow: 'Neige', wx_showers: 'Averses', wx_storm: 'Orages',
      tabAccounts: 'Comptes', tabItinerary: 'Itinéraire', programme: 'Programme', trip: 'Trajet', dayShort: 'Jour', allDays: 'Tout', playRoute: '▶ Voir le parcours',
      sailingTime: 'de navigation', mapPlan: 'Plan', routeNote: 'Routes indicatives selon le programme prévu (peuvent évoluer selon la météo et les conditions de mer).',
      nowSailing: 'Aujourd’hui : en navigation (position prévue)', nowBefore: 'Aujourd’hui : avant le départ du jour', nowArrived: 'Aujourd’hui : arrivé à l’escale du jour',
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
      vatEstimated: 'VAT estimated', apaLeftInclVat: 'APA left including VAT estimated', toValidate: 'To be validated by the fiscal representative', vatEstNote: 'Paid excl. VAT: this VAT will be invoiced by the fiscal representative once the report is sent',
      htChip: 'excl. VAT', apaLeft: 'APA left',
      boatInvoice: 'Invoice', boatInvoicePending: 'Invoice issued at closing',
      reportUnavailable: 'This report is not available (accounts not signed or invalid link).', preparingZip: 'Preparing zip…',
      tabHome: 'Today', homeDay: 'Day', homeOf: 'of', programmeOfDay: 'Today’s programme', meals: 'Meals', lunch: 'Lunch', dinner: 'Dinner',
      captainNote: 'Captain’s note', freeDay: 'Day at anchor / alongside — enjoy!', wxTitle: 'Weather', wxWind: 'Wind', wxGust: 'gusts', wxSea: 'Sea', wxWaves: 'waves',
      wxRain: 'Rain', wxWater: 'Water', wxSun: 'Sun', wxUv: 'UV', wxNA: 'Forecast available about 15 days before the date.', wxLoading: 'Loading weather…',
      startsIn: 'Your charter starts in {n} day(s) — preview of day one', menuSoon: 'Menu to be announced',
      wx_sunny: 'Sunny', wx_partly: 'Partly cloudy', wx_cloudy: 'Cloudy', wx_fog: 'Fog', wx_drizzle: 'Drizzle', wx_rain: 'Rain', wx_snow: 'Snow', wx_showers: 'Showers', wx_storm: 'Thunderstorms',
      tabAccounts: 'Accounts', tabItinerary: 'Itinerary', programme: 'Programme', trip: 'Trip', dayShort: 'Day', allDays: 'All', playRoute: '▶ Play the route',
      sailingTime: 'sailing', mapPlan: 'Map', routeNote: 'Indicative routes as per the planned programme (subject to weather and sea conditions).',
      nowSailing: 'Today: under way (planned position)', nowBefore: 'Today: before the day’s departure', nowArrived: 'Today: arrived at the day’s stop',
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
  const TABLES = ['charters', 'expenses', 'cash_withdrawals', 'bunkers', 'voyages', 'allocations', 'boat_profile', 'invoices', 'charter_documents'];
  const OPTIONAL_TABLES = ['charter_documents']; // table absente tant que la mise à jour SQL n'est pas faite → liste vide

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
      res.forEach((r, i) => { if (r.error) { if (OPTIONAL_TABLES.includes(TABLES[i])) { S[TABLES[i]] = []; S._missing = [...(S._missing || []), TABLES[i]]; return; } throw r.error; } S[TABLES[i]] = r.data; });
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
    async clientItinerary(token) {
      const { data, error } = await this.sb.rpc('client_itinerary', { p_token: token });
      if (error) return null; return data;
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
        ['expenses', 'cash_withdrawals', 'charter_documents'].forEach(tb => { this.S[tb] = (this.S[tb] || []).filter(x => x.charter_id !== id); });
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
    async clientItinerary(token) {
      const c = this.S.charters.find(x => x.client_token === token); if (!c) return null;
      const d = (this.S.charter_documents || []).find(x => x.charter_id === c.id && x.folder === 'itinerary_plan');
      let p = null; try { p = d?.note ? JSON.parse(d.note) : null; } catch { p = null; }
      return p && p.share_client ? p : null;
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
    fork: svg('<circle cx="6" cy="5" r="2"/><circle cx="18" cy="5" r="2"/><circle cx="12" cy="19" r="2"/><path d="M6 7v2a3 3 0 0 0 3 3h6a3 3 0 0 0 3-3V7M12 12v5"/>'),
    sun: svg('<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>'),
    map: svg('<path d="M9 4 3 6v14l6-2 6 2 6-2V4l-6 2z"/><path d="M9 4v14M15 6v14"/>'),
    folder: svg('<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>'),
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
  const errMsg = e => {
    const m = (e && (e.message || e.error_description || e.details)) || String(e);
    const col = /(broker_fee_pct|auto_kind|auto_ref|'fee')/.exec(m);
    return col ? `Base de données à mettre à jour : exécutez le fichier « maj-frais.sql » dans Supabase (SQL Editor). (${m})` : m;
  };

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
      <div class="kpi kpi-vatbal"><div class="kpi-label">${t('apaLeftInclVat')}</div>
        <div class="kpi-value">${money(st.balanceInclVat, cur)}</div><div class="kpi-sub">${st.balanceInclVat >= 0 ? t('toRefund') : t('toPay')}</div>
        <div class="kpi-validate">⚠ ${t('toValidate')}</div></div>` : ''}
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
    let PLAN = null, planStr = '', CTAB = 'accounts', firstLoad = true, HOME_DATE = null;

    async function refresh() {
      if (inflight) return; inflight = true;
      try {
        const prevB = B ? JSON.stringify(B) : '';
        B = REPORT ? await api.reportView(token) : await api.clientView(token); lastAt = new Date();
        let newPlan = null; if (!REPORT && api.clientItinerary) { try { newPlan = await api.clientItinerary(token); } catch { newPlan = null; } }
        const np = newPlan ? JSON.stringify(newPlan) : '', planChanged = np !== planStr;
        PLAN = newPlan && ((newPlan.legs || []).length || Object.keys(newPlan.days || {}).length) ? newPlan : null; planStr = np;
        if (firstLoad && B) { firstLoad = false; // à l'ouverture : page d'accueil du jour pendant le charter (et juste avant)
          const ds = PLAN ? planDates(B.charter, PLAN) : [], td = localToday();
          if (ds.length && td <= ds[ds.length - 1] && dayDiff(ds[0], td) <= 7) CTAB = 'home';
          else { const saved = store.get('ca-tab-' + token); if (saved && saved !== 'home' && (saved !== 'itin' || PLAN)) CTAB = saved; } }
        if (CTAB === 'itin' && !(PLAN && (PLAN.legs || []).length)) CTAB = 'accounts';
        if (CTAB === 'home' && !PLAN) CTAB = 'accounts';
        const quiet = CTAB !== 'accounts' && $('#client-tab-body');
        if (!$('#modal-root').innerHTML && !(quiet && !planChanged && prevB === JSON.stringify(B))) render();
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

    const TX = () => ({ day: t('dayShort'), all: t('allDays'), play: t('playRoute'), sailing: t('sailingTime'), satellite: 'Satellite', plan: t('mapPlan'),
      nowSailing: t('nowSailing'), nowBefore: t('nowBefore'), nowArrived: t('nowArrived') });
    function clientItinHtml(c) {
      const P = computePlan(c, PLAN);
      return `<section class="panel route-panel"><div class="panel-head"><h2>${t('tabItinerary')}</h2>
          <span class="muted small">${fmt(P.totalD)} NM · ${hhmm(P.totalH)} ${t('sailingTime')}</span></div>
        <div class="route-map" id="client-map"></div><div id="client-route-ctl">${routeDaysHtml(P, TX())}</div>
        <p class="muted small">${t('routeNote')}</p></section>
        <section class="panel"><div class="panel-head"><h2>${t('programme')}</h2></div>
        <div class="table-wrap"><table class="table table-cards table-compact itin-table"><thead><tr><th>${t('dayShort')}</th><th>Date</th><th>${t('trip')}</th><th class="num">NM</th><th class="num">ETD</th><th class="num">ETA</th></tr></thead>
        <tbody>${P.rows.filter(r => !r.alt).map(r => `<tr class="${r.sameDay ? 'same-day' : ''} ${r.date === today() ? 'is-today' : ''}"><td data-label="${t('dayShort')}" class="num">${r.sameDay ? '' : r.day ?? ''}</td><td data-label="Date">${r.sameDay ? '' : fdate(r.date)}</td>
          <td data-label="${t('trip')}" class="itin-trip"><b>${esc(r.trip || '')}</b>${r.remarks ? `<div class="itin-rem">${esc(r.remarks)}</div>` : ''}</td>
          <td data-label="NM" class="num">${r.dist ? fmt(r.dist) : ''}</td><td data-label="ETD" class="num">${r.dist ? esc(r.etd || '') : ''}</td><td data-label="ETA" class="num">${esc(r.eta || '')}</td></tr>`).join('')}</tbody></table></div></section>`;
    }
    function homeDates(c) { return planDates(c, PLAN); }
    function curHomeDate(c) { const ds = homeDates(c), td = localToday();
      if (HOME_DATE && ds.includes(HOME_DATE)) return HOME_DATE;
      return ds.includes(td) ? td : td < ds[0] ? ds[0] : ds[ds.length - 1]; }
    function clientHomeHtml(c) {
      const ds = homeDates(c); if (!ds.length) return '';
      const d = curHomeDate(c), k = ds.indexOf(d), n = dayDiff(d, ds[0]) + 1, td = localToday();
      const P = computePlan(c, PLAN), rows = P.rows.filter(r => !r.alt && r.date === d), x = (PLAN.days || {})[d] || {};
      const long = new Date(d + 'T12:00:00').toLocaleDateString(locale(), { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
      const meal = (label, icon, time, menu) => `<div class="meal"><div class="meal-head"><span>${icon} ${label}</span>${time ? `<b>${esc(time)}</b>` : ''}</div>
        ${menu ? `<ul>${menu.split('\n').filter(s => s.trim()).map(s => `<li>${esc(s.trim())}</li>`).join('')}</ul>` : `<p class="muted small">${t('menuSoon')}</p>`}</div>`;
      return `${td < ds[0] ? `<div class="banner banner-info">${ICON.anchor}<div>${t('startsIn').replace('{n}', dayDiff(ds[0], td))}</div></div>` : ''}
      <section class="home-hero">
        <button class="home-nav" data-home-nav="${k > 0 ? ds[k - 1] : ''}" ${k > 0 ? '' : 'disabled'} aria-label="prev">‹</button>
        <div class="home-date"><div class="home-day">${t('homeDay')} ${n} <span>${t('homeOf')} ${dayDiff(ds[ds.length - 1], ds[0]) + 1}</span></div>
          <h2>${esc(long.charAt(0).toUpperCase() + long.slice(1))}</h2>
          ${rows.length ? `<div class="home-route">${esc(rows.map(r => r.trip).filter(Boolean).join(' · '))}</div>` : ''}</div>
        <button class="home-nav" data-home-nav="${k < ds.length - 1 ? ds[k + 1] : ''}" ${k < ds.length - 1 ? '' : 'disabled'} aria-label="next">›</button>
      </section>
      <div class="home-grid">
        <section class="panel home-wx" id="home-wx"><div class="panel-head"><h2>${t('wxTitle')}</h2><span class="muted small" id="wx-place"></span></div><p class="muted small">${t('wxLoading')}</p></section>
        <section class="panel"><div class="panel-head"><h2>${t('programmeOfDay')}</h2></div>
          ${rows.some(r => r.dist) ? rows.map(r => `<div class="home-leg">${r.dist ? `<div class="hl-time">${esc(r.etd || '')}${r.eta ? `<span>→ ${esc(r.eta)}</span>` : ''}</div>` : '<div class="hl-time"></div>'}
            <div><b>${esc(r.trip || '')}</b>${r.dist ? ` <span class="muted small">· ${fmt(r.dist)} NM</span>` : ''}${r.remarks ? `<div class="small">${esc(r.remarks)}</div>` : ''}</div></div>`).join('')
          : `<p>${rows[0]?.trip ? `<b>${esc(rows[0].trip)}</b> — ` : ''}${t('freeDay')}</p>${rows.filter(r => r.remarks).map(r => `<p class="small">${esc(r.remarks)}</p>`).join('')}`}
        </section>
      </div>
      <section class="panel"><div class="panel-head"><h2>${t('meals')}</h2></div>
        <div class="meals">${meal(t('lunch'), '🍽', x.lunch_time, x.lunch_menu)}${meal(t('dinner'), '🍷', x.dinner_time, x.dinner_menu)}</div>
        ${x.note ? `<div class="home-note"><b>${t('captainNote')}</b><p>${esc(x.note)}</p></div>` : ''}
      </section>`;
    }
    async function loadHomeWeather(c) {
      const box = $('#home-wx'); if (!box) return;
      const d = curHomeDate(c), pos = dayPosition(c, PLAN, d);
      const body = html => { box.innerHTML = `<div class="panel-head"><h2>${t('wxTitle')}</h2><span class="muted small">${pos ? esc(pos.name || '') : ''}</span></div>${html}`; };
      if (!pos) return body(`<p class="muted small">${t('wxNA')}</p>`);
      const w = await fetchWeather(pos, d);
      if (curHomeDate(c) !== d || !$('#home-wx')) return;
      if (!w) return body(`<p class="muted small">${t('wxNA')}</p>`);
      const [ic, key] = WMO(w.code);
      body(`<div class="wx-main"><span class="wx-ic">${ic}</span><div><div class="wx-t">${Math.round(w.tmax)}° <span>/ ${Math.round(w.tmin)}°</span></div><div class="muted">${t('wx_' + key)}</div></div></div>
        <div class="wx-grid">
          <div><span>${t('wxWind')}</span><b>${compass(w.dir)} ${Math.round(w.wind)} kn</b><small>${w.gust != null ? `${t('wxGust')} ${Math.round(w.gust)} kn · ` : ''}F${beaufort(w.wind)}</small></div>
          ${w.wave != null ? `<div><span>${t('wxSea')}</span><b>${fmt(w.wave, 1)} m</b><small>${t('wxWaves')}</small></div>` : ''}
          ${w.sst != null ? `<div><span>${t('wxWater')}</span><b>${Math.round(w.sst)}°</b></div>` : ''}
          ${w.rain != null ? `<div><span>${t('wxRain')}</span><b>${Math.round(w.rain)} %</b></div>` : ''}
          ${w.sunrise ? `<div><span>${t('wxSun')}</span><b>${w.sunrise} – ${w.sunset}</b></div>` : ''}
          ${w.uv != null ? `<div><span>${t('wxUv')}</span><b>${Math.round(w.uv)}</b></div>` : ''}
        </div>`);
    }
    async function mountClientMap(c) {
      const el = $('#client-map'); if (!el) return;
      try { const ctl = await mountRouteMap(el, c, PLAN, { labels: TX() }); bindRouteControls($('.route-panel'), ctl, ctl.P, TX(), false); }
      catch (e) { el.innerHTML = `<div class="muted center small" style="padding:30px">${esc(errMsg(e))}</div>`; }
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
          <div class="eyebrow">${REPORT ? t('reportTitle') : PLAN && CTAB !== 'accounts' ? esc(c.yacht || YACHT) : t('statement')}</div>
          <h1>${esc(c.title)}</h1>
          <p class="muted">${esc(c.client_name || '')}${c.client_name ? ' · ' : ''}${fdate(c.start_date)} → ${fdate(c.end_date)}</p>
          <div class="hero-actions no-print">
            <button class="btn btn-ghost" data-act="print">${ICON.print}${t('print')}</button>
            <button class="btn btn-ghost" data-act="xlsx">${ICON.sheet}${t('excel')}</button>
            ${REPORT || signed ? `<button class="btn btn-ghost" data-act="zip">${ICON.file}${t('zipAll')}</button>` : ''}
            ${closed ? `<button class="btn btn-gold" data-act="sign">${ICON.pen}${t('signCta')}</button>` : ''}
          </div>
        </section>
        ${PLAN ? `<nav class="folder-tabs client-tabs no-print">
          <a href="#" data-ctab="home" class="folder-tab ${CTAB === 'home' ? 'on' : ''}">${ICON.sun}<span><b>${t('tabHome')}</b></span></a>
          <a href="#" data-ctab="accounts" class="folder-tab ${CTAB === 'accounts' ? 'on' : ''}">${ICON.sheet}<span><b>${t('tabAccounts')}</b></span></a>
          ${(PLAN.legs || []).length ? `<a href="#" data-ctab="itin" class="folder-tab ${CTAB === 'itin' ? 'on' : ''}">${ICON.map}<span><b>${t('tabItinerary')}</b></span></a>` : ''}</nav>` : ''}
        ${PLAN && CTAB === 'home' ? `<div id="client-tab-body">${clientHomeHtml(c)}</div>` : PLAN && CTAB === 'itin' ? `<div id="client-tab-body">${clientItinHtml(c)}</div>` : `
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
        </section>`}
        <footer class="foot muted small">${esc(c.yacht || YACHT)} · ${t('statement')} · ${new Date().toLocaleDateString(locale())}</footer>
      </main>`;
      if (PLAN && CTAB === 'itin') mountClientMap(c);
      if (PLAN && CTAB === 'home') loadHomeWeather(c);
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
      const hn = e.target.closest('[data-home-nav]');
      if (hn) { if (hn.dataset.homeNav) { HOME_DATE = hn.dataset.homeNav; render(); } return; }
      const tb = e.target.closest('[data-ctab]');
      if (tb) { e.preventDefault(); CTAB = tb.dataset.ctab; store.set('ca-tab-' + token, CTAB); render(); window.scrollTo(0, 0); return; }
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
    if (parts[0] === 'charter' && parts[1]) return viewCharter(parts[1], parts[2]);
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
  const charterYear = c => String(c.start_date || c.end_date || c.created_at || '').slice(0, 4) || 'Sans date';
  const OPEN_YEARS = new Set();
  function viewCharters() {
    const list = S.charters.slice().sort((a, b) => String(b.start_date || '').localeCompare(String(a.start_date || '')));
    const card = c => {
      const st = computeStatement(buildBundle(S, c));
      return `<a class="charter-card" href="#/charter/${c.id}">
        <div class="cc-top">${statusBadge(c)}
          <span class="muted small">${fdate(c.start_date)} → ${fdate(c.end_date)}</span></div>
        <h3>${esc(c.title)}</h3><p class="muted">${esc(c.client_name || '')}</p>
        <div class="cc-figs"><div><span>Dépenses</span><b>${money(st.total, c.currency)}</b></div>
          <div><span>Solde APA</span><b class="${st.balance < 0 ? 'neg' : 'pos'}">${money(st.balance, c.currency)}</b></div></div>
      </a>`;
    };
    const years = [...new Set(list.map(charterYear))].sort((a, b) => b.localeCompare(a));
    if (!OPEN_YEARS.size && years.length) OPEN_YEARS.add(years.includes(String(new Date().getFullYear())) ? String(new Date().getFullYear()) : years[0]);
    const folders = years.map(y => {
      const cs = list.filter(c => charterYear(c) === y);
      const open = OPEN_YEARS.has(y);
      return `<details class="year-folder" data-year="${esc(y)}" ${open ? 'open' : ''}>
        <summary>${ICON.folder}<span class="yf-name">${esc(y)}</span><span class="yf-count">${cs.length} charter${cs.length > 1 ? 's' : ''}</span>
          ${cs.some(c => c.status !== 'signed') ? `<span class="badge badge-open">${cs.filter(c => c.status !== 'signed').length} en cours</span>` : ''}<span class="yf-chev">›</span></summary>
        <div class="cards">${cs.map(card).join('')}</div></details>`;
    }).join('');
    $('#app').innerHTML = shell('charters', `
      <div class="page-head"><div><div class="eyebrow">Comptabilité</div><h1>Charters</h1></div>
        <button class="btn btn-primary" data-act="new-charter">${ICON.plus}Nouveau charter</button></div>
      ${list.length ? `<div class="year-folders">${folders}</div>` : `<div class="empty-state card"><p>Aucun charter pour le moment.</p>
        <button class="btn btn-primary" data-act="new-charter">${ICON.plus}Créer le premier charter</button></div>`}`);
    $$('.year-folder').forEach(d => d.addEventListener('toggle', () => { d.open ? OPEN_YEARS.add(d.dataset.year) : OPEN_YEARS.delete(d.dataset.year); }));
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
        <label class="field"><span>Bank fee sur l’APA (%) <small>(ligne de dépense auto, 0 = aucune)</small></span><input type="number" step="0.01" min="0" name="broker_fee_pct" value="${esc(c.broker_fee_pct ?? 0.4)}"></label>
        <label class="field field-fee"><span>Charter fee <small>(sert uniquement au calcul des tips)</small></span><input type="number" step="0.01" name="charter_fee" value="${esc(c.charter_fee ?? '')}" placeholder="ex. 155000"></label>
        ${!S.voyages.some(v => v.charter_id === c.id) ? `<label class="check span-2 boat-toggle"><input type="checkbox" name="auto_voyage" ${c.auto_voyage !== false ? 'checked' : ''}>
          <span><b>Créer automatiquement le voyage dans le fuel booklet</b> le jour du début du charter (nom, dates, ROB au départ calculé ce jour-là)</span></label>` : ''}
        ${(() => {
          const cur = c.vat_country || 'France';
          const known = VAT_COUNTRIES.map(x => x[0]);
          const used = [...new Set(S.charters.map(x => x.vat_country).filter(x => x && !known.includes(x)))].sort();
          const all = [...known, ...used];
          const other = !all.includes(cur);
          return `<label class="field"><span>Pays de début du charter <small>(TVA)</small></span><select name="vat_country_sel">
            ${all.map(n => { const r = VAT_COUNTRIES.find(x => x[0] === n)?.[1]; return `<option value="${esc(n)}" data-rate="${r ?? ''}" ${!other && n === cur ? 'selected' : ''}>${esc(n)}</option>`; }).join('')}
            <option value="__other" data-rate="" ${other ? 'selected' : ''}>Autre pays (saisir)…</option></select>
            <input name="vat_country_other" class="vat-other mt-6" placeholder="Nom du pays, ex. Malte" value="${other ? esc(cur) : ''}" ${other ? '' : 'hidden'}></label>`;
        })()}
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
    fc.vat_country_sel.addEventListener('change', () => {
      const o = fc.vat_country_sel.selectedOptions[0], isOther = o.value === '__other';
      fc.vat_country_other.hidden = !isOther; fc.vat_country_other.required = isOther;
      if (isOther) { fc.vat_country_other.focus(); fc.vat_rate.select?.(); }
      else if (o.dataset.rate !== '') fc.vat_rate.value = o.dataset.rate;
      else { const prev = S.charters.find(x => x.vat_country === o.value); if (prev) fc.vat_rate.value = prev.vat_rate ?? 0; }
    });
    fc.vat_country_other.required = fc.vat_country_sel.value === '__other';
    $('[data-ok]', m).onclick = async ev => {
      const f = $('#f-charter', m);
      if (!f.reportValidity()) return;
      const d = Object.fromEntries(new FormData(f));
      const row = { title: d.title.trim(), client_name: d.client_name || null, yacht: d.yacht || YACHT, start_date: d.start_date || null,
        end_date: d.end_date || null, apa_paid: toNum(d.apa_paid) || 0, broker_fee_pct: toNum(d.broker_fee_pct) ?? 0, charter_fee: toNum(d.charter_fee), currency: d.currency,
        vat_country: (d.vat_country_sel === '__other' ? (d.vat_country_other || '').trim() : d.vat_country_sel) || null, vat_rate: toNum(d.vat_rate) ?? 0,
        fuel_company: d.fuel_company || null, fuel_description: d.fuel_description || null, fuel_method: d.fuel_method, notes: d.notes || null, fuel_by_boat: !!d.fuel_by_boat,
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
        try { if (await syncBrokerFee(charterById(saved.id))) await reload(); } catch (e) { toast('Bank fee : ' + errMsg(e), 'err'); }
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
        const files = [...S.expenses, ...(S.charter_documents || [])].filter(e => e.charter_id === c.id && e.file_path).map(e => e.file_path);
        await api.remove('charters', c.id);
        for (const p of files) await api.removeFile(p).catch(() => {});
        await reload(); location.hash = '#/charters'; toast('Charter supprimé');
      } catch (e) { toast(errMsg(e), 'err'); }
    };
  }
  // ---------------- Dossiers Preference list / Itinerary ----------------
  function viewCharterFolder(c, tab) {
    const [, label, sub] = FOLDERS.find(f => f[0] === tab);
    const docs = (S.charter_documents || []).filter(d => d.charter_id === c.id && d.folder === tab)
      .sort((a, b) => String(b.created_at || '').localeCompare(String(a.created_at || '')));
    const missing = (S._missing || []).includes('charter_documents');
    const item = d => d.file_path
      ? `<div class="doc-item"><div class="doc-main">${fileLink(d.file_path, d.title || d.file_name || 'Fichier')}
          <small class="muted">${d.title && d.file_name && d.title !== d.file_name ? esc(d.file_name) + ' · ' : ''}ajouté le ${fdatetime(d.created_at)}</small></div>
          <button class="icon-btn danger" data-act="doc-del" data-id="${d.id}" title="Supprimer">${ICON.trash}</button></div>`
      : `<div class="doc-item doc-note"><div class="doc-ico">${ICON.edit}</div><div class="doc-main"><b>${esc(d.title || 'Note')}</b>
          <div class="doc-text">${esc(d.note || '')}</div><small class="muted">${fdatetime(d.created_at)}</small></div>
          <div class="row-gap"><button class="icon-btn" data-act="doc-note" data-id="${d.id}" title="Modifier">${ICON.edit}</button>
          <button class="icon-btn danger" data-act="doc-del" data-id="${d.id}" title="Supprimer">${ICON.trash}</button></div></div>`;
    $('#app').innerHTML = shell('charters', `
      <a class="back" href="#/charters">${ICON.back}Charters</a>
      <div class="page-head">
        <div><div class="eyebrow">${esc(c.yacht || YACHT)}</div><h1>${esc(c.title)}</h1>
          <p class="muted">${esc(c.client_name || '')}${c.client_name ? ' · ' : ''}${fdate(c.start_date)} → ${fdate(c.end_date)} · ${statusBadge(c)}</p></div>
        <div class="head-actions"><button class="btn btn-ghost" data-act="edit-charter" data-id="${c.id}">${ICON.edit}Modifier</button></div>
      </div>
      ${folderTabs(c, tab)}
      ${missing ? `<div class="banner banner-gold">${ICON.lock}<div>Pour utiliser ce dossier, exécutez une fois le fichier <b>maj-dossiers.sql</b> dans Supabase (SQL Editor), puis rechargez la page.</div></div>` : ''}
      ${tab === 'itinerary' && !missing ? itineraryPanel(c) + (getPlan(c).legs.length ? mapPanelHtml(c, getPlan(c)) : '') + daysPanelHtml(c, getPlan(c)) : ''}
      <section class="panel">
        <div class="panel-head"><h2>${tab === 'itinerary' ? 'Documents & notes' : esc(label)}</h2>
          ${missing ? '' : `<div class="row-gap"><button class="btn btn-ghost btn-sm" data-act="doc-note" data-folder="${tab}">${ICON.edit}Ajouter une note</button>
            <label class="btn btn-primary btn-sm file-btn">${ICON.plus}Ajouter des fichiers<input type="file" multiple class="sr-file" data-doc-upload="${tab}"
              accept="application/pdf,image/*,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.txt,.csv,.heic,.pages,.numbers,.key"></label></div>`}</div>
        <p class="muted small">${esc(sub)} — dossier privé : visible uniquement par vous (pas dans la vue client).</p>
        ${docs.length ? `<div class="doc-list">${docs.map(item).join('')}</div>`
          : `<div class="empty-state"><p class="muted">Dossier vide. Ajoutez des fichiers (PDF, photos, Word, Excel…) ou une note.</p></div>`}
      </section>`);
    if (tab === 'itinerary' && !missing) { bindItinerary(c); bindMapPanel(c); }
    const inp = $('[data-doc-upload]');
    if (inp) inp.addEventListener('change', async () => {
      const files = [...inp.files]; if (!files.length) return;
      let ok = 0;
      for (const f of files) {
        try { toast(`Envoi de « ${f.name} »…`);
          const up = await api.upload(f, `docs/${c.id}/${tab}`);
          await api.insert('charter_documents', { charter_id: c.id, folder: tab, title: f.name, ...up }); ok++; }
        catch (e) { toast(`${f.name} : ${errMsg(e)}`, 'err'); }
      }
      await reload(); route(); if (ok) toast(ok > 1 ? `${ok} fichiers ajoutés` : 'Fichier ajouté');
    });
  }
  ACTIONS['doc-note'] = b => {
    const d = b.dataset.id ? S.charter_documents.find(x => x.id === b.dataset.id) : null;
    const folder = d ? d.folder : b.dataset.folder;
    const m = openModal({ title: d ? 'Modifier la note' : 'Nouvelle note', wide: true,
      body: `<form class="form-grid" id="f-note"><label class="field span-2"><span>Titre</span><input name="title" value="${esc(d?.title || '')}" placeholder="${folder === 'itinerary' ? 'ex. Jour 1 — Cannes → Saint-Tropez' : 'ex. Allergies / boissons'}"></label>
        <label class="field span-2"><span>Note</span><textarea name="note" rows="10" required>${esc(d?.note || '')}</textarea></label></form>`,
      footer: `<button class="btn btn-ghost" data-close>Annuler</button><button class="btn btn-primary" data-ok>Enregistrer</button>` });
    $('[data-ok]', m).onclick = async ev => {
      const f = $('#f-note', m); if (!f.reportValidity()) return; busy(ev.currentTarget, true);
      const row = { title: f.title.value.trim() || null, note: f.note.value };
      try { d ? await api.update('charter_documents', d.id, row) : await api.insert('charter_documents', { charter_id: CURRENT, folder, ...row });
        await reload(); closeModal(); route(); toast('Note enregistrée'); }
      catch (e) { toast(errMsg(e), 'err'); busy(ev.currentTarget, false); }
    };
  };
  ACTIONS['doc-del'] = async b => {
    const d = S.charter_documents.find(x => x.id === b.dataset.id); if (!d) return;
    if (!(await confirmDialog(`Supprimer « ${esc(d.title || d.file_name || 'Note')} » ?`, 'Supprimer'))) return;
    try { await api.remove('charter_documents', d.id); if (d.file_path) api.removeFile(d.file_path).catch(() => {}); await reload(); route(); toast('Supprimé'); }
    catch (e) { toast(errMsg(e), 'err'); }
  };

  // ---------------- Itinéraire (programme) avec calcul de consommation ----------------
  // Stocké dans charter_documents (folder 'itinerary_plan', note = JSON) : pas de table supplémentaire.
  const PLAN_DEFAULTS = { eng_lph: 100, gen_lph: 20, speed: 12 };
  const planRow = c => (S.charter_documents || []).find(d => d.charter_id === c.id && d.folder === 'itinerary_plan');
  function getPlan(c) {
    const row = planRow(c);
    let p = null; try { p = row?.note ? JSON.parse(row.note) : null; } catch { p = null; }
    if (!p) { // reprend les réglages du dernier programme
      const last = (S.charter_documents || []).filter(d => d.folder === 'itinerary_plan' && d.note).sort((a, b) => String(b.created_at).localeCompare(String(a.created_at)))[0];
      let lp = null; try { lp = last ? JSON.parse(last.note) : null; } catch { lp = null; }
      p = { eng_lph: lp?.eng_lph ?? PLAN_DEFAULTS.eng_lph, gen_lph: lp?.gen_lph ?? PLAN_DEFAULTS.gen_lph, speed: lp?.speed ?? PLAN_DEFAULTS.speed, legs: [] };
    }
    p.legs = (p.legs || []).map((l, k) => l.id ? l : { ...l, id: 'leg' + k }); // id stable tant que non enregistré
    return p;
  }
  async function savePlan(c, p) {
    p.legs = p.legs.map((l, i) => ({ ...l, _i: i })).sort((a, b) => String(a.date || '9999').localeCompare(String(b.date || '9999')) || a._i - b._i).map(({ _i, ...l }) => l);
    const row = planRow(c), note = JSON.stringify(p);
    if (row) await api.update('charter_documents', row.id, { note });
    else await api.insert('charter_documents', { charter_id: c.id, folder: 'itinerary_plan', title: 'Programme', note });
    await reload();
  }
  const addDays = (d, n) => { const x = new Date(d + 'T12:00:00Z'); x.setUTCDate(x.getUTCDate() + n); return x.toISOString().slice(0, 10); };
  const dayDiff = (a, b) => Math.round((new Date(a + 'T12:00:00Z') - new Date(b + 'T12:00:00Z')) / 864e5);
  const hhmm = h => { if (h == null || !isFinite(h)) return ''; const m = Math.round(h * 60); return `${Math.floor(m / 60)} h ${String(m % 60).padStart(2, '0')}`; };
  const clock = (etd, h) => { if (!etd) return ''; const [H, M] = etd.split(':').map(Number); const t = H * 60 + M + Math.round((h || 0) * 60);
    const d = Math.floor(t / 1440), r = t % 1440; return `${String(Math.floor(r / 60)).padStart(2, '0')}:${String(r % 60).padStart(2, '0')}${d ? ` (+${d} j)` : ''}`; };
  function computePlan(c, p) {
    const first = p.legs.map(l => l.date).filter(Boolean).sort()[0] || c.start_date;
    let cumD = 0, cumL = 0, cumH = 0, eng = 0, gen = 0, prevDate = null;
    // Générateur = par JOURNÉE (on peut bouger plusieurs fois le même jour) : heures du jour = somme des lignes du jour, max 24 h,
    // comptées une seule fois sur le premier trajet du jour.
    const dayKey = (l, i) => l.date || 'row' + i, dayGen = {}, firstOfDay = {};
    const shoreDay = {}; // à quai, branché sur le quai : pas de générateur ce jour-là
    p.legs.forEach((l, i) => { if (l.alt) return; const k = dayKey(l, i); dayGen[k] = Math.min(24, (dayGen[k] || 0) + (toNum(l.gen_h) || 0)); if (l.shore) shoreDay[k] = true; if (firstOfDay[k] == null) firstOfDay[k] = i; });
    Object.keys(shoreDay).forEach(k => { dayGen[k] = 0; });
    const rows = p.legs.map((l, i) => {
      const k = dayKey(l, i), isFirst = firstOfDay[k] === i;
      const dist = toNum(l.dist) || 0, speed = toNum(l.speed) || toNum(p.speed) || 0, genH = l.alt ? toNum(l.gen_h) || 0 : isFirst ? dayGen[k] : 0;
      const hours = dist > 0 && speed > 0 ? dist / speed : 0;
      const lEng = r2(hours * (toNum(p.eng_lph) || 0)), lGen = r2(genH * (toNum(p.gen_lph) || 0)), cons = r2(lEng + lGen);
      if (!l.alt) { cumD = r2(cumD + dist); cumL = r2(cumL + cons); cumH += hours; eng += lEng; gen += lGen; }
      const sameDay = l.date && l.date === prevDate; prevDate = l.date;
      const grp = l.group ? p.legs.filter(x => x.group === l.group) : [];
      const optLetter = grp.length > 1 ? String.fromCharCode(65 + grp.indexOf(l)) : '';
      return { ...l, i, shoreDay: !!shoreDay[k] && isFirst, day: l.date && first ? dayDiff(l.date, first) + 1 : null, sameDay, dist, speed, genH, hours, cons, cumD, cumL, optLetter, grpN: grp.length,
        eta: hours ? clock(l.etd, hours) : '' };
    });
    rows.forEach(r => { if (r.alt && r.group) { const ch = rows.find(x => x.group === r.group && !x.alt); r.diff = ch ? r2(r.cons - ch.cons) : null; r.diffD = ch ? r2(r.dist - ch.dist) : null; } });
    const voy = ((S && S.voyages) || []).filter(v => v.charter_id === c.id).sort(byPos)[0];
    const rob = toNum(voy?.rob_before);
    return { rows, totalD: cumD, totalH: cumH, totalL: cumL, eng: r2(eng), gen: r2(gen), rob, robEnd: rob != null ? r2(rob - cumL) : null, voy };
  }
  function itineraryPanel(c) {
    const p = getPlan(c), P = computePlan(c, p);
    const hasGps = p.legs.some(l => l.gps || l.gps_link);
    const gps = s => !s ? '' : /^https?:\/\//.test(s) ? `<a href="${esc(s)}" target="_blank" rel="noopener">Carte ↗</a>` : esc(s);
    const rows = P.rows.map(r => `<tr class="row-click ${r.sameDay ? 'same-day' : ''} ${r.alt ? 'leg-alt' : ''} ${r.optLetter && !r.alt ? 'leg-chosen' : ''}" data-leg="${r.i}">
        <td data-label="Day" class="num">${r.sameDay ? '' : (r.day ?? '')}</td>
        <td data-label="Date">${r.sameDay ? '' : fdate(r.date)}</td>
        <td data-label="Trip" class="itin-trip">${r.optLetter ? `<span class="opt-badge ${r.alt ? '' : 'on'}">Option ${r.optLetter}${r.alt ? '' : ' · retenue'}</span>` : ''}<b>${esc(r.trip || '') || '<span class="muted">— à compléter —</span>'}</b>
          ${r.alt ? `<div class="opt-diff">${r.diff != null ? `${r.diff > 0 ? '+' : ''}${fmt(r.diff)} L · ${r.diffD > 0 ? '+' : ''}${fmt(r.diffD)} NM vs option retenue` : ''} <button class="link-btn" data-act="leg-choose" data-i="${r.i}">Retenir cette option</button></div>` : ''}${r.remarks ? `<div class="itin-rem">${esc(r.remarks)}</div>` : ''}</td>
        ${hasGps ? `<td data-label="Position GPS">${gps(r.gps_link)}${r.gps_link && r.gps ? ' · ' : ''}${esc(r.gps || '')}</td>` : ''}
        <td data-label="Distance NM" class="num">${r.dist ? fmt(r.dist) : ''}</td>
        <td data-label="Cumul NM" class="num">${r.alt ? '—' : fmt(r.cumD)}</td>
        <td data-label="Speed kn" class="num">${r.dist ? fmt(r.speed, r.speed % 1 ? 1 : 0) : ''}</td>
        <td data-label="Sailing time" class="num">${r.hours ? hhmm(r.hours) : ''}</td>
        <td data-label="Générateur" class="num">${r.shoreDay ? '<span class="shore-badge" title="À quai, branché : pas de générateur">⚡ À quai</span>' : r.genH ? r.genH + ' h' : r.sameDay ? '<span class="muted">—</span>' : ''}</td>
        <td data-label="Conso (L)" class="num"><b>${fmt(r.cons)}</b></td>
        <td data-label="Conso cumul (L)" class="num">${r.alt ? '—' : fmt(r.cumL)}</td>
        <td data-label="ETD" class="num">${esc(r.etd || '')}</td>
        <td data-label="ETA" class="num">${esc(r.eta)}</td>
        <td class="actions"><button class="icon-btn" data-act="leg-edit" data-i="${r.i}" title="Modifier">${ICON.edit}</button>
          ${r.alt ? '' : `<button class="btn btn-ghost btn-xs" data-act="leg-same" data-i="${r.i}" title="Ajouter un autre trajet ce jour-là">+ trajet</button>`}
          <button class="icon-btn danger" data-act="leg-del" data-i="${r.i}" title="Supprimer">${ICON.trash}</button></td></tr>`).join('');
    return `<section class="panel itin">
      <div class="panel-head"><h2>Programme & consommation</h2>
        <div class="row-gap wrap">
          ${!p.legs.length && c.start_date && c.end_date ? `<button class="btn btn-ghost btn-sm" data-act="plan-days">${ICON.plus}Créer les jours du charter</button>` : ''}
          <label class="btn btn-ghost btn-sm file-btn">${ICON.sheet}Importer un Excel<input type="file" class="sr-file" data-plan-import accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"></label>
          ${p.legs.length ? `<button class="btn btn-ghost btn-sm" data-act="plan-xlsx">${ICON.sheet}Excel</button>` : ''}
          <button class="btn btn-primary btn-sm" data-act="leg-new">${ICON.plus}Ajouter une étape</button></div></div>
      <form class="plan-params" id="f-plan">
        <label>Moteurs <input type="number" step="1" min="0" name="eng_lph" value="${esc(p.eng_lph)}"> L/h de navigation</label>
        <label>Générateur <input type="number" step="1" min="0" name="gen_lph" value="${esc(p.gen_lph)}"> L/h</label>
        <label>Vitesse par défaut <input type="number" step="0.5" min="0" name="speed" value="${esc(p.speed)}"> nœuds</label>
        <span class="muted small">Conso d’une étape = temps de navigation × L/h moteurs + heures générateur × L/h générateur</span>
      </form>
      ${p.legs.length ? `<div class="table-wrap"><table class="table table-cards table-compact itin-table">
        <thead><tr><th>Day</th><th>Date</th><th>Trip / Remarks</th>${hasGps ? '<th>Position GPS</th>' : ''}<th class="num">Distance NM</th><th class="num">Cumul NM</th><th class="num">Speed kn</th>
          <th class="num">Sailing time</th><th class="num">Géné.</th><th class="num">Conso L</th><th class="num">Cumul L</th><th class="num">ETD</th><th class="num">ETA</th><th></th></tr></thead>
        <tbody>${rows}</tbody></table></div>
        <div class="itin-totals">
          <div><span>Distance totale</span><b>${fmt(P.totalD)} NM</b></div>
          <div><span>Temps de navigation</span><b>${hhmm(P.totalH)}</b></div>
          <div><span>Consommation estimée</span><b>${fmt(P.totalL)} L</b><small>moteurs ${fmt(P.eng)} L · générateur ${fmt(P.gen)} L</small></div>
          ${P.rob != null ? `<div><span>ROB au départ (fuel booklet)</span><b>${fmt(P.rob)} L</b></div>
            <div class="${P.robEnd < 0 ? 'neg' : ''}"><span>ROB estimé à l’arrivée</span><b>${fmt(P.robEnd)} L</b>${toNum(P.voy?.rob_after) != null ? `<small>réel : ${fmt(P.voy.rob_after)} L (conso réelle ${fmt(r2(P.rob - P.voy.rob_after))} L)</small>` : ''}</div>` : ''}
        </div>`
      : `<div class="empty-state"><p class="muted">Aucune étape. ${c.start_date && c.end_date ? 'Créez les jours du charter, ' : ''}ajoutez une étape ou importez votre fichier Excel de programme.</p></div>`}
    </section>`;
  }
  function bindItinerary(c) {
    const f = $('#f-plan'); if (!f) return;
    let tm;
    f.addEventListener('input', () => { clearTimeout(tm); tm = setTimeout(async () => {
      const p = getPlan(c); ['eng_lph', 'gen_lph', 'speed'].forEach(k => { p[k] = toNum(f[k].value) ?? 0; });
      try { await savePlan(c, p); const a = document.activeElement?.name; route(); if (a) { const el = $(`#f-plan [name=${a}]`); el?.focus(); const v = el?.value; if (el) { el.value = ''; el.value = v; } } }
      catch (e) { toast(errMsg(e), 'err'); }
    }, 700); });
    $$('.itin-table tr[data-leg]').forEach(tr => tr.addEventListener('click', e => { if (e.target.closest('button,a')) return; legForm(c, +tr.dataset.leg); }));
    const imp = $('[data-plan-import]');
    if (imp) imp.addEventListener('change', async () => { const file = imp.files[0]; if (file) await importPlanXlsx(c, file); imp.value = ''; });
  }
  function legForm(c, i, preset) {
    const p = getPlan(c), isNew = i == null, l = isNew ? preset : p.legs[i];
    const sameDayExtra = isNew ? !!preset?._sameDay : p.legs.findIndex(x => !x.alt && x.date && x.date === l.date) !== i && !!l.date;
    const dayShoreNow = !isNew && (l.date ? p.legs.some(x => !x.alt && x.date === l.date && x.shore) : !!l.shore);
    const dayGenNow = !isNew && l.date ? Math.min(24, sum(p.legs.filter(x => !x.alt && x.date === l.date), x => toNum(x.gen_h) || 0)) : null;
    const m = openModal({ title: isNew ? (preset?._group ? 'Autre trajet possible (option)' : 'Nouvelle étape') : (l.group ? 'Modifier l’option' : 'Modifier l’étape'), wide: true,
      body: `<form class="form-grid" id="f-leg">
        <label class="field"><span>Date</span><input type="date" name="date" value="${esc(l.date || '')}"></label>
        <label class="field"><span>Trip <small>(ex. Hydra/Agistri)</small></span><input name="trip" value="${esc(l.trip || '')}" required></label>
        <label class="field"><span>Position GPS Google <small>(lien)</small></span><input name="gps_link" value="${esc(l.gps_link || '')}" placeholder="https://maps.google.com/…"></label>
        <label class="field"><span>Position GPS</span><input name="gps" value="${esc(l.gps || '')}" placeholder="37°20'N 23°28'E"></label>
        <label class="field"><span>Distance (NM) <small>(0 = à l’ancre / au port)</small></span><input type="number" step="0.1" min="0" name="dist" value="${esc(l.dist ?? '')}"></label>
        <label class="field"><span>Vitesse (nœuds)</span><input type="number" step="0.5" min="0" name="speed" value="${esc(l.speed ?? p.speed)}"></label>
        ${sameDayExtra ? `<input type="hidden" name="gen_h" value="0"><div class="field muted small">Générateur : compté une seule fois pour la journée (sur le 1<sup>er</sup> trajet du jour).</div>`
          : `<label class="field"><span>Heures générateur de la journée <small>(12 = ½ journée, 24 = journée)</small></span><input type="number" step="0.5" min="0" max="24" name="gen_h" value="${esc(dayGenNow ?? l.gen_h ?? 24)}"></label>
            <label class="check span-2 boat-toggle"><input type="checkbox" name="shore" ${dayShoreNow ? 'checked' : ''}>
              <span><b>⚡ À quai — branché sur le quai</b> : pas de consommation générateur ce jour-là</span></label>`}
        <label class="field"><span>ETD (heure de départ)</span><input type="time" name="etd" value="${esc(l.etd || '')}"></label>
        <label class="field span-2"><span>Remarks / Suggestions</span><textarea name="remarks" rows="3">${esc(l.remarks || '')}</textarea></label>
        <div class="span-2 leg-preview muted small" id="leg-prev"></div></form>`,
      footer: `${!isNew ? `<button class="btn btn-ghost danger mr-auto" data-del>${ICON.trash}Supprimer</button>` : ''}<button class="btn btn-ghost" data-close>Annuler</button><button class="btn btn-primary" data-ok>Enregistrer</button>` });
    const f = $('#f-leg', m);
    const genOn = () => !(f.shore && f.shore.checked);
    if (f.shore) { const sync = () => { f.gen_h.disabled = f.shore.checked; }; f.shore.addEventListener('change', sync); sync(); }
    const prev = () => { const d = toNum(f.dist.value) || 0, s = toNum(f.speed.value) || 0, g = genOn() ? toNum(f.gen_h.value) || 0 : 0, h = d && s ? d / s : 0;
      $('#leg-prev', m).innerHTML = `Temps de navigation <b>${h ? hhmm(h) : '—'}</b> · ETA <b>${h && f.etd.value ? clock(f.etd.value, h) : '—'}</b> · Consommation estimée <b>${fmt(r2(h * (toNum(p.eng_lph) || 0) + g * (toNum(p.gen_lph) || 0)))} L</b>`; };
    f.addEventListener('input', prev); prev();
    $('[data-ok]', m).onclick = async ev => {
      if (!f.reportValidity()) return; busy(ev.currentTarget, true);
      const leg = { date: f.date.value || null, trip: f.trip.value.trim(), gps_link: f.gps_link.value.trim() || null, gps: f.gps.value.trim() || null,
        dist: toNum(f.dist.value) ?? 0, speed: toNum(f.speed.value), gen_h: toNum(f.gen_h.value) ?? 0, etd: f.etd.value || null, remarks: f.remarks.value.trim() || null,
        ...(f.shore ? { shore: f.shore.checked } : {}) };
      const q = getPlan(c);
      if (!sameDayExtra && leg.date) q.legs.forEach((x, j) => { if (j !== i && !x.alt && x.date === leg.date) { x.gen_h = 0; delete x.shore; } }); // générateur porté par le 1er trajet du jour
      if (isNew) {
        leg.id = uuid();
        if (preset?._group) { leg.group = preset._group; leg.alt = true; const base = q.legs.find(x => x.id === preset._base); if (base && !base.group) base.group = preset._group; }
        if (preset?._after != null) q.legs.splice(preset._after + 1, 0, leg); else q.legs.push(leg);
      } else q.legs[i] = { ...q.legs[i], ...leg };
      try { await savePlan(c, q); closeModal(); route(); } catch (e) { toast(errMsg(e), 'err'); busy(ev.currentTarget, false); }
    };
    const del = $('[data-del]', m); if (del) del.onclick = () => { closeModal(); delLeg(c, i); };
  }
  async function delLeg(c, i) {
    const p = getPlan(c); if (!(await confirmDialog(`Supprimer l’étape « ${esc(p.legs[i]?.trip || '')} » ?`, 'Supprimer'))) return;
    const [gone] = p.legs.splice(i, 1);
    if (gone && !gone.alt && gone.date && toNum(gone.gen_h)) { const nx = p.legs.find(x => !x.alt && x.date === gone.date); if (nx) nx.gen_h = Math.min(24, (toNum(nx.gen_h) || 0) + toNum(gone.gen_h)); }
    if (gone?.group) { const rest = p.legs.filter(x => x.group === gone.group);
      if (rest.length === 1) { delete rest[0].group; rest[0].alt = false; } else if (rest.length && !gone.alt) rest[0].alt = false; }
    try { await savePlan(c, p); route(); } catch (e) { toast(errMsg(e), 'err'); }
  }
  ACTIONS['leg-new'] = () => { const c = charterById(CURRENT), p = getPlan(c), last = p.legs[p.legs.length - 1];
    legForm(c, null, { date: last?.date ? addDays(last.date, 1) : (c.start_date || today()), speed: p.speed, gen_h: 24, etd: '08:00' }); };
  ACTIONS['leg-edit'] = b => legForm(charterById(CURRENT), +b.dataset.i);
  ACTIONS['leg-del'] = b => delLeg(charterById(CURRENT), +b.dataset.i);
  ACTIONS['leg-same'] = async b => { // trajet supplémentaire le même jour (on peut bouger plusieurs fois)
    const c = charterById(CURRENT), p = getPlan(c), i = +b.dataset.i, l = p.legs[i];
    const lastOfDay = l.date ? p.legs.reduce((k, x, j) => (!x.alt && x.date === l.date ? j : k), i) : i;
    legForm(c, null, { date: l.date, speed: p.speed, gen_h: 0, etd: '', _after: lastOfDay, _sameDay: !!l.date });
  };
  ACTIONS['leg-option'] = b => { // trajet alternatif pour la même étape (une seule option compte dans les totaux)
    const c = charterById(CURRENT), p = getPlan(c), i = +b.dataset.i, l = p.legs[i];
    const group = l.group || l.id;
    const last = p.legs.reduce((k, x, j) => (x.group === group || x.id === l.id ? j : k), i);
    legForm(c, null, { date: l.date, speed: l.speed ?? p.speed, gen_h: l.gen_h, etd: l.etd, _after: last, _group: group, _base: l.id });
  };
  ACTIONS['leg-choose'] = async b => {
    const c = charterById(CURRENT), p = getPlan(c), l = p.legs[+b.dataset.i];
    p.legs.forEach(x => { if (x.group === l.group) x.alt = x.id !== l.id; });
    try { await savePlan(c, p); route(); toast('Option retenue : totaux recalculés'); } catch (e) { toast(errMsg(e), 'err'); }
  };
  ACTIONS['plan-days'] = async () => {
    const c = charterById(CURRENT), p = getPlan(c), n = dayDiff(c.end_date, c.start_date);
    for (let k = 0; k <= n; k++) p.legs.push({ date: addDays(c.start_date, k), trip: '', dist: 0, speed: p.speed, gen_h: k === 0 || k === n ? 12 : 24, etd: '08:00' });
    try { await savePlan(c, p); route(); toast(`${n + 1} jours créés : cliquez sur une ligne pour la compléter`); } catch (e) { toast(errMsg(e), 'err'); }
  };
  async function importPlanXlsx(c, file) {
    try { await loadScript(XLSX_JS); } catch (e) { return toast(errMsg(e), 'err'); }
    try {
      const X = window.XLSX, wb = X.read(await file.arrayBuffer(), { cellDates: false, cellFormula: true });
      const ws = wb.Sheets[wb.SheetNames[0]], rg = X.utils.decode_range(ws['!ref']);
      const cell = (r, col) => ws[X.utils.encode_cell({ r, c: col })];
      let hr = -1, col = {};
      for (let r = rg.s.r; r <= rg.e.r && hr < 0; r++) for (let k = rg.s.c; k <= rg.e.c; k++) if (/^day$/i.test(String(cell(r, k)?.v || '').trim())) { hr = r; break; }
      if (hr < 0) throw new Error('En-tête « Day » introuvable dans le fichier.');
      for (let k = rg.s.c; k <= rg.e.c; k++) { const h = String(cell(hr, k)?.v || '').toLowerCase();
        if (/^day/.test(h)) col.day = k; else if (/period|date/.test(h)) col.date = k; else if (/^trip/.test(h)) col.trip = k;
        else if (/gps google/.test(h)) col.gps_link = k; else if (/gps/.test(h)) col.gps = k; else if (/distance cumul/.test(h)) col.cum = k;
        else if (/distance/.test(h)) col.dist = k; else if (/speed/.test(h)) col.speed = k; else if (/consumption cumul/.test(h)) col.cumc = k;
        else if (/consumption/.test(h)) col.cons = k; else if (/etd|departure/.test(h)) col.etd = k; else if (/remark/.test(h)) col.rem = k; }
      const p = getPlan(c), legs = []; let lastDate = null;
      const toDate = v => {
        if (typeof v === 'number' && v > 20000) return new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 864e5).toISOString().slice(0, 10);
        const s = String(v ?? '').match(/(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})/);
        return s ? `${s[3].length === 2 ? '20' + s[3] : s[3]}-${s[2].padStart(2, '0')}-${s[1].padStart(2, '0')}` : null; };
      const toTime = v => { if (v instanceof Date) return `${String(v.getHours()).padStart(2, '0')}:${String(v.getMinutes()).padStart(2, '0')}`;
        if (typeof v === 'number') { const m = Math.round((v % 1) * 1440); return `${String(Math.floor(m / 60)).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`; }
        const s = String(v || '').match(/(\d{1,2})[:h](\d{2})/); return s ? `${s[1].padStart(2, '0')}:${s[2]}` : null; };
      for (let r = hr + 1; r <= rg.e.r; r++) {
        const trip = String(cell(r, col.trip)?.v ?? '').trim(); if (!trip) continue;
        const d = toDate(cell(r, col.date)?.v); if (d) lastDate = d;
        const f = String(cell(r, col.cons)?.f || '');
        const mg = f.match(/\(\s*(\d+(?:\.\d+)?)\s*\*\s*(\d+(?:\.\d+)?)\s*\)\s*$/), me = f.match(/\(\s*(\d+(?:\.\d+)?)\s*\*/);
        if (me) p.eng_lph = +me[1]; if (mg) p.gen_lph = +mg[2];
        legs.push({ date: d || lastDate, trip, gps_link: cell(r, col.gps_link)?.v ? String(cell(r, col.gps_link).v) : null, gps: cell(r, col.gps)?.v ? String(cell(r, col.gps).v) : null,
          dist: toNum(cell(r, col.dist)?.v) || 0, speed: toNum(cell(r, col.speed)?.v) ?? p.speed, gen_h: mg ? +mg[1] : 24,
          etd: toTime(cell(r, col.etd)?.v), remarks: cell(r, col.rem)?.v ? String(cell(r, col.rem).v).trim() : null });
      }
      if (!legs.length) throw new Error('Aucune étape trouvée sous l’en-tête.');
      if (p.legs.length && !(await confirmDialog(`Remplacer les ${p.legs.length} étapes actuelles par les ${legs.length} étapes du fichier ?`, 'Remplacer', false))) return;
      p.legs = legs; await savePlan(c, p); route(); toast(`${legs.length} étapes importées`);
    } catch (e) { toast('Import : ' + errMsg(e), 'err'); }
  }
  ACTIONS['plan-xlsx'] = async () => {
    try { await loadScript(XLSX_JS); } catch (e) { return toast(errMsg(e), 'err'); }
    const c = charterById(CURRENT), p = getPlan(c), P = computePlan(c, p), X = window.XLSX;
    const title = `${c.title}${c.start_date ? ` from ${fdate(c.start_date)} to ${fdate(c.end_date)}` : ''}`;
    const head = ['Day', 'Period', 'Trip', 'Position GPS Google', 'Position GPS', 'Distance in NM', 'Distance Cumulée en NM', 'Speed (knots)',
      'Approximate sailing time (hour)', 'Generator (hours)', 'Consumption (liters)', 'Consumption cumulated (liters)', 'ETD (departure hour)', 'ETA (arrival hour)', 'Remarks/Suggestions'];
    const aoa = [[c.yacht || YACHT], [title], [`Engines ${p.eng_lph} L/h · Generator ${p.gen_lph} L/h`], [], head,
      ...P.rows.map(r => [r.sameDay ? '' : r.day ?? '', r.sameDay ? '' : fdate(r.date), (r.optLetter ? `[Option ${r.optLetter}${r.alt ? '' : ' - retenue'}] ` : '') + (r.trip || ''), r.gps_link || '', r.gps || '', r.dist || 0, r.alt ? '' : r.cumD, r.dist ? r.speed : '',
        r.hours ? hhmm(r.hours) : '', r.shoreDay ? 'À quai (0)' : r.genH, r.cons, r.alt ? '' : r.cumL, r.etd || '', r.eta, r.remarks || '']),
      [], ['', '', 'TOTAL', '', '', P.totalD, '', '', hhmm(P.totalH), '', P.totalL]];
    const ws = X.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [6, 11, 24, 22, 18, 10, 12, 9, 12, 10, 12, 13, 10, 10, 70].map(w => ({ wch: w }));
    const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, ws, 'Programme');
    saveBlob(new Blob([X.write(wb, { type: 'array', bookType: 'xlsx' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
      `Programme_${c.title.replace(/[^\w-]+/g, '_')}.xlsx`);
  };

  // ---------------- Routes maritimes (évite la terre) ----------------
  // Trait de côte Natural Earth 1:10m (world-atlas) rastérisé sur un canvas, puis recherche A* sur la grille d'eau.
  const LAND_JSON = 'https://cdn.jsdelivr.net/npm/world-atlas@2.0.2/land-10m.json';
  const TOPOJSON_JS = 'https://cdn.jsdelivr.net/npm/topojson-client@3.1.0/dist/topojson-client.min.js';
  let LAND = null;
  // Petites îles (Capri, Antipaxos…) absentes de land-10m : couche « minor islands » de Natural Earth
  const MINOR_ISLANDS_JSON = 'https://cdn.jsdelivr.net/gh/nvkelso/natural-earth-vector@v5.1.2/geojson/ne_10m_minor_islands.geojson';
  async function loadLand() {
    if (LAND) return LAND;
    await loadScript(TOPOJSON_JS);
    const polys = [];
    const add = f => { const g = f.geometry || f; const list = g.type === 'Polygon' ? [g.coordinates] : g.type === 'MultiPolygon' ? g.coordinates : [];
      list.forEach(rings => { let a = 180, b = 90, c = -180, d = -90; rings[0].forEach(([x, y]) => { if (x < a) a = x; if (x > c) c = x; if (y < b) b = y; if (y > d) d = y; }); polys.push({ rings, bb: [a, b, c, d] }); }); };
    const topo = await (await fetch(LAND_JSON)).json();
    const geo = window.topojson.feature(topo, topo.objects.land);
    (geo.features || [geo]).forEach(add);
    try { const mi = await (await fetch(MINOR_ISLANDS_JSON)).json(); (mi.features || []).forEach(add); } catch { /* facultatif */ }
    return (LAND = polys);
  }
  // Trait de côte détaillé OpenStreetMap (Overpass) pour les trajets courts : petites îles, caps, ports
  const COAST_CACHE = {};
  async function fetchCoast(s, w, n, e) {
    const key = [s, w, n, e].map(v => v.toFixed(2)).join(',');
    if (COAST_CACHE[key]) return COAST_CACHE[key];
    if ((n - s) * (e - w) > 1.2) return (COAST_CACHE[key] = []); // trop grand : on garde Natural Earth
    try {
      const q = `[out:json][timeout:25];way["natural"="coastline"](${s.toFixed(4)},${w.toFixed(4)},${n.toFixed(4)},${e.toFixed(4)});out geom;`;
      const ctrl = new AbortController(), tm = setTimeout(() => ctrl.abort(), 30000);
      const r = await fetch('https://overpass-api.de/api/interpreter', { method: 'POST', body: 'data=' + encodeURIComponent(q),
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, signal: ctrl.signal });
      clearTimeout(tm); if (!r.ok) throw new Error('HTTP ' + r.status);
      const j = await r.json();
      return (COAST_CACHE[key] = (j.elements || []).filter(x => x.type === 'way' && x.geometry).map(x => x.geometry.map(g => [g.lon, g.lat])));
    } catch { return (COAST_CACHE[key] = []); }
  }
  const nmBetween = (a, b) => { const R = 3440.065, t = Math.PI / 180, dLa = (b.lat - a.lat) * t, dLo = (b.lng - a.lng) * t;
    const h = Math.sin(dLa / 2) ** 2 + Math.cos(a.lat * t) * Math.cos(b.lat * t) * Math.sin(dLo / 2) ** 2; return 2 * R * Math.asin(Math.sqrt(h)); };
  const pathNm = pts => pts.reduce((s, p, i) => i ? s + nmBetween(pts[i - 1], p) : 0, 0);
  // Rastérise la terre d'une zone sur un canvas → grille terre / mer
  function landGrid(minLat, maxLat, minLng, maxLng, cell, land, buffer = 0, coast = []) {
    const kx = Math.cos(((minLat + maxLat) / 2) * Math.PI / 180);
    const H = Math.max(2, Math.ceil((maxLat - minLat) / cell)), W = Math.max(2, Math.ceil((maxLng - minLng) * kx / cell));
    const cv = document.createElement('canvas'); cv.width = W; cv.height = H;
    const g = cv.getContext('2d'); g.fillStyle = '#000';
    const X = lng => (lng - minLng) * kx / cell, Y = lat => (maxLat - lat) / cell;
    g.beginPath();
    land.forEach(p => { if (p.bb[2] < minLng || p.bb[0] > maxLng || p.bb[3] < minLat || p.bb[1] > maxLat) return;
      p.rings.forEach(r => { r.forEach(([x, y], i) => i ? g.lineTo(X(x), Y(y)) : g.moveTo(X(x), Y(y))); g.closePath(); }); });
    g.fill('evenodd');
    if (coast.length) { // côtes OSM : îles fermées remplies, autres côtes = murs infranchissables
      g.strokeStyle = '#000'; g.lineWidth = 1.6; g.lineJoin = 'round';
      coast.forEach(l => { if (l.length < 2) return; g.beginPath(); l.forEach(([x, y], i) => i ? g.lineTo(X(x), Y(y)) : g.moveTo(X(x), Y(y)));
        const closed = l[0][0] === l[l.length - 1][0] && l[0][1] === l[l.length - 1][1]; if (closed) { g.closePath(); g.fill(); } g.stroke(); });
    }
    const px = g.getImageData(0, 0, W, H).data, N = W * H; let isLand = new Uint8Array(N);
    for (let i = 0; i < N; i++) isLand[i] = px[i * 4 + 3] > 40 ? 1 : 0;
    if (coast.length) { // la mer = l'eau reliée aux bords de la zone ; l'eau enfermée par une côte = terre
      const sea = new Uint8Array(N), st = [];
      for (let x = 0; x < W; x++) { st.push(x, (H - 1) * W + x); } for (let y = 0; y < H; y++) { st.push(y * W, y * W + W - 1); }
      while (st.length) { const i = st.pop(); if (sea[i] || isLand[i]) continue; sea[i] = 1; const x = i % W, y = (i / W) | 0;
        if (x) st.push(i - 1); if (x < W - 1) st.push(i + 1); if (y) st.push(i - W); if (y < H - 1) st.push(i + W); }
      let n = 0; for (let i = 0; i < N; i++) if (sea[i]) n++;
      if (n > N * 0.05) for (let i = 0; i < N; i++) isLand[i] = sea[i] ? 0 : 1;
    }
    for (let r = 0; r < buffer; r++) { // marge de sécurité le long des côtes (le trait de côte est simplifié)
      const nx = isLand.slice();
      for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) { const i = y * W + x; if (isLand[i]) continue;
        if ((x && isLand[i - 1]) || (x < W - 1 && isLand[i + 1]) || (y && isLand[i - W]) || (y < H - 1 && isLand[i + W])) nx[i] = 1; }
      isLand = nx;
    }
    const toLL = i => ({ lat: maxLat - (((i / W) | 0) + 0.5) * cell, lng: minLng + ((i % W) + 0.5) * cell / kx });
    const near = (pt, maxR = 80) => { // case de mer la plus proche
      const x0 = Math.max(0, Math.min(W - 1, Math.floor(X(pt.lng)))), y0 = Math.max(0, Math.min(H - 1, Math.floor(Y(pt.lat))));
      if (!isLand[y0 * W + x0]) return y0 * W + x0;
      for (let r = 1; r < maxR; r++) { let best = -1, bd = Infinity;
        for (let dy = -r; dy <= r; dy++) for (let dx = -r; dx <= r; dx++) { if (Math.max(Math.abs(dx), Math.abs(dy)) !== r) continue;
          const x = x0 + dx, y = y0 + dy; if (x < 0 || y < 0 || x >= W || y >= H || isLand[y * W + x]) continue; const d = dx * dx + dy * dy; if (d < bd) { bd = d; best = y * W + x; } }
        if (best >= 0) return best; }
      return -1; };
    return { W, H, isLand, toLL, near };
  }
  // Ramène une escale « à terre » (centre-ville, port mal tracé) sur l'eau la plus proche
  async function snapToSea(pt, land) {
    const d = 0.08, coast = await fetchCoast(pt.lat - d, pt.lng - d * 1.4, pt.lat + d, pt.lng + d * 1.4);
    const G = landGrid(pt.lat - d, pt.lat + d, pt.lng - d * 1.4, pt.lng + d * 1.4, 0.0006, land, 1, coast);
    const i = G.near(pt, 250); if (i < 0) return pt;
    const q = G.toLL(i); return nmBetween(pt, q) < 0.05 ? pt : { lat: r6(q.lat), lng: r6(q.lng) };
  }
  async function seaRoute(a, b, land, padF = 0.6) {
    const span = Math.max(Math.abs(a.lat - b.lat), Math.abs(a.lng - b.lng));
    const pad = Math.max(0.15, span * padF);
    const minLat = Math.min(a.lat, b.lat) - pad, maxLat = Math.max(a.lat, b.lat) + pad;
    const kx = Math.cos(((minLat + maxLat) / 2) * Math.PI / 180);
    const minLng = Math.min(a.lng, b.lng) - pad / kx, maxLng = Math.max(a.lng, b.lng) + pad / kx;
    const cell = Math.max(0.0015, (maxLat - minLat) / 600); // ≈ 170 m pour les trajets courts
    const buffer = Math.max(1, Math.round(0.003 / cell));   // ≈ 300 m de marge le long des côtes
    const coast = await fetchCoast(minLat, minLng, maxLat, maxLng);
    const G = landGrid(minLat, maxLat, minLng, maxLng, cell, land, buffer, coast), { W, H, isLand } = G, N = W * H;
    const S0 = G.near(a, 120), E0 = G.near(b, 120); if (S0 < 0 || E0 < 0) return null;
    const ex = E0 % W, ey = (E0 / W) | 0;
    const gS = new Float32Array(N).fill(Infinity), came = new Int32Array(N).fill(-1), closed = new Uint8Array(N);
    const heap = [], hv = [];
    const push = (i, f) => { heap.push(i); hv.push(f); let k = heap.length - 1; while (k) { const q = (k - 1) >> 1; if (hv[q] <= hv[k]) break; [heap[q], heap[k]] = [heap[k], heap[q]]; [hv[q], hv[k]] = [hv[k], hv[q]]; k = q; } };
    const pop = () => { const top = heap[0], li = heap.pop(), lv = hv.pop(); if (heap.length) { heap[0] = li; hv[0] = lv; let k = 0;
      for (;;) { const l = 2 * k + 1, r = l + 1; let m = k; if (l < heap.length && hv[l] < hv[m]) m = l; if (r < heap.length && hv[r] < hv[m]) m = r; if (m === k) break;
        [heap[m], heap[k]] = [heap[k], heap[m]]; [hv[m], hv[k]] = [hv[k], hv[m]]; k = m; } } return top; };
    const hx = i => { const dx = Math.abs(i % W - ex), dy = Math.abs(((i / W) | 0) - ey); return Math.max(dx, dy) + 0.4142 * Math.min(dx, dy); };
    gS[S0] = 0; push(S0, hx(S0));
    const D = [[1, 0, 1], [-1, 0, 1], [0, 1, 1], [0, -1, 1], [1, 1, 1.4142], [1, -1, 1.4142], [-1, 1, 1.4142], [-1, -1, 1.4142]];
    let found = false, guard = 0;
    while (heap.length && guard++ < 4e6) {
      const cur = pop(); if (cur === E0) { found = true; break; } if (closed[cur]) continue; closed[cur] = 1;
      const x = cur % W, y = (cur / W) | 0;
      for (const [dx, dy, c] of D) { const nx = x + dx, ny = y + dy; if (nx < 0 || ny < 0 || nx >= W || ny >= H) continue;
        const n = ny * W + nx; if (isLand[n] || closed[n]) continue;
        if (dx && dy && (isLand[y * W + x + dx] || isLand[(y + dy) * W + x])) continue;
        const ng = gS[cur] + c; if (ng < gS[n]) { gS[n] = ng; came[n] = cur; push(n, ng + hx(n)); } }
    }
    if (!found) return null;
    const cells = []; for (let k = E0; k !== -1; k = came[k]) cells.push(k); cells.reverse();
    const clear = (i, j) => { let x0 = i % W, y0 = (i / W) | 0; const x1 = j % W, y1 = (j / W) | 0, dx = Math.abs(x1 - x0), dy = -Math.abs(y1 - y0), sx = x0 < x1 ? 1 : -1, sy = y0 < y1 ? 1 : -1; let err = dx + dy;
      for (;;) { if (isLand[y0 * W + x0]) return false; if (x0 === x1 && y0 === y1) return true; const e2 = 2 * err; if (e2 >= dy) { err += dy; x0 += sx; } if (e2 <= dx) { err += dx; y0 += sy; } } };
    const keep = [cells[0]]; let k = 0;
    while (k < cells.length - 1) { let j = cells.length - 1; while (j > k + 1 && !clear(cells[k], cells[j])) j--; keep.push(cells[j]); k = j; }
    const ll = i => { const q = G.toLL(i); return { lat: r4(q.lat), lng: r4(q.lng) }; };
    const mid = keep.map(ll);
    // escale → première case de mer (hors marge côtière) → … → dernière case de mer → escale
    const out = [a]; mid.forEach(q => { if (nmBetween(out[out.length - 1], q) > 0.05) out.push(q); });
    if (nmBetween(out[out.length - 1], b) > 0.05) out.push(b); else out[out.length - 1] = b;
    return out;
  }
  const r4 = v => Math.round(v * 1e4) / 1e4;
  // Route d'un trajet : départ → points de passage → arrivée ; segments vers une escale = route maritime auto,
  // segment court entre 2 points de passage (canal, chenal) = ligne droite.
  async function legRoutePoints(a, b, via, land) {
    const pts = [a, ...(via || []).map(([lat, lng]) => ({ lat, lng })), b], out = [a];
    for (let k = 1; k < pts.length; k++) {
      const p = pts[k - 1], q = pts[k], straight = nmBetween(p, q), bothVia = k > 1 && k < pts.length - 1;
      let seg = null;
      if (land && !(bothVia && straight < 12)) { seg = await seaRoute(p, q, land); if (!seg) seg = await seaRoute(p, q, land, 1.5);
        if (seg && pathNm(seg) > Math.max(4 * straight, straight + 40)) seg = null; }
      out.push(...(seg ? seg.slice(1) : [q]));
    }
    return out;
  }

  // ---------------- Carte de l'itinéraire (Leaflet + vue satellite) ----------------
  const LEAFLET_JS = 'https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.js';
  const LEAFLET_CSS = 'https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.css';
  async function ensureLeaflet() {
    if (!$(`link[href="${LEAFLET_CSS}"]`)) { const l = document.createElement('link'); l.rel = 'stylesheet'; l.href = LEAFLET_CSS; document.head.appendChild(l); }
    await loadScript(LEAFLET_JS);
    return window.L;
  }
  // « 40.55, 14.24 » · « 37°20'N 23°28'E » · « 37°20.5'N » · lien Google Maps (@lat,lng / q=lat,lng / !3dlat!4dlng)
  function parseGps(s) {
    if (!s) return null; s = String(s).trim();
    let m = s.match(/!3d(-?\d+(?:\.\d+)?)!4d(-?\d+(?:\.\d+)?)/) || s.match(/@(-?\d+(?:\.\d+)?),(-?\d+(?:\.\d+)?)/) || s.match(/[?&](?:q|query|ll)=(-?\d+(?:\.\d+)?),\s*(-?\d+(?:\.\d+)?)/);
    if (m) return { lat: +m[1], lng: +m[2] };
    const dms = [...s.matchAll(/(\d+(?:[.,]\d+)?)\s*°\s*(?:(\d+(?:[.,]\d+)?)\s*['’′]\s*)?(?:(\d+(?:[.,]\d+)?)\s*["”″]\s*)?([NSEW])/gi)];
    if (dms.length >= 2) {
      const v = d => { const n = x => x ? +String(x).replace(',', '.') : 0; let x = n(d[1]) + n(d[2]) / 60 + n(d[3]) / 3600; return /[SW]/i.test(d[4]) ? -x : x; };
      const la = dms.find(d => /[NS]/i.test(d[4])), lo = dms.find(d => /[EW]/i.test(d[4]));
      if (la && lo) return { lat: v(la), lng: v(lo) };
    }
    m = s.match(/^(-?\d{1,2}(?:\.\d+)?)\s*[,; ]\s*(-?\d{1,3}(?:\.\d+)?)$/);
    if (m && Math.abs(+m[1]) <= 90 && Math.abs(+m[2]) <= 180) return { lat: +m[1], lng: +m[2] };
    return null;
  }
  const tripParts = trip => String(trip || '').split('/').map(x => x.trim()).filter(Boolean);
  const normName = s => String(s || '').toLowerCase().replace(/\(.*?\)/g, '').replace(/[^a-z0-9à-ÿ]+/g, ' ').trim();
  // Escales dans l'ordre : départ du 1er trajet puis destination de chaque trajet (une escale = un point)
  function planNodes(c, p) {
    const P = computePlan(c, p), nodes = [];
    P.rows.filter(r => !r.alt).forEach(r => {
      const parts = tripParts(r.trip), leg = p.legs[r.i];
      if (!nodes.length && parts.length > 1) {
        const g = leg.from_lat != null ? { lat: leg.from_lat, lng: leg.from_lng } : null;
        nodes.push({ name: parts[0], lat: g?.lat, lng: g?.lng, legIndex: r.i, kind: 'from', day: r.day, date: r.date, row: null });
      }
      const name = parts[parts.length - 1] || r.trip || '';
      const prev = nodes[nodes.length - 1];
      let g = leg.to_lat != null ? { lat: leg.to_lat, lng: leg.to_lng } : parseGps(leg.gps_link) || parseGps(leg.gps);
      if (!g && prev && prev.lat != null && (r.dist === 0 || normName(prev.name) === normName(name))) g = { lat: prev.lat, lng: prev.lng, inherited: true };
      nodes.push({ name, lat: g?.lat, lng: g?.lng, inherited: !!g?.inherited, legIndex: r.i, kind: 'to', day: r.day, date: r.date, row: r });
    });
    return { nodes, P };
  }
  async function geocodeName(name, near) {
    const tries = [];
    const inPar = (name.match(/\(([^)]+)\)/) || [])[1]; if (inPar) tries.push(inPar);
    tries.push(name.replace(/\(.*?\)/g, '').trim());
    for (const q of [...new Set(tries)].filter(Boolean)) {
      const vb = near ? `&viewbox=${near.lng - 3},${near.lat + 3},${near.lng + 3},${near.lat - 3}` : '';
      const r = await fetch(`https://nominatim.openstreetmap.org/search?format=json&limit=1&accept-language=en${vb}&q=${encodeURIComponent(q)}`);
      if (r.ok) { const j = await r.json(); if (j[0]) return { lat: +j[0].lat, lng: +j[0].lon }; }
      await new Promise(res => setTimeout(res, 1100)); // 1 requête / seconde (règle Nominatim)
    }
    return null;
  }
  const lerp = (a, b, t) => ({ lat: a.lat + (b.lat - a.lat) * t, lng: a.lng + (b.lng - a.lng) * t });
  const alongPath = (pts, f) => { const tot = pathNm(pts); let d = Math.max(0, Math.min(1, f)) * tot;
    for (let k = 1; k < pts.length; k++) { const s = nmBetween(pts[k - 1], pts[k]); if (d <= s) return lerp(pts[k - 1], pts[k], s ? d / s : 0); d -= s; } return pts[pts.length - 1]; };
  const routeKey = (a, b, via) => 'v2|' + [a.lat, a.lng, b.lat, b.lng].map(r4).join(',') + '|' + JSON.stringify((via || []).map(v => v.map(r4)));
  // Calcule (et met en cache dans le programme) les routes maritimes manquantes ; renvoie true si quelque chose a changé
  async function computeSeaRoutes(c, p, onStatus) {
    const { nodes } = planNodes(c, p); const todo = [];
    for (let k = 1; k < nodes.length; k++) { const a = nodes[k - 1], b = nodes[k];
      if (a.lat == null || b.lat == null || !b.row || !(b.row.dist > 0)) continue;
      const leg = p.legs[b.row.i], key = routeKey(a, b, leg.via); if (!leg.route || leg.route.k !== key) todo.push({ a, b, leg, key }); }
    // escales trouvées automatiquement (centre-ville…) : ramenées sur l'eau la plus proche, une fois
    const toSnap = nodes.filter(n => { const leg = p.legs[n.legIndex]; return n.lat != null && !n.inherited && (n.kind === 'from' ? leg.from_lat != null && !leg.from_fixed : leg.to_lat != null && !leg.to_fixed); });
    if (!todo.length && !toSnap.length) return false;
    onStatus?.('Calcul des routes maritimes…');
    const land = await loadLand();
    let snapped = false;
    for (const n of toSnap) { const leg = p.legs[n.legIndex], q = await snapToSea({ lat: n.lat, lng: n.lng }, land);
      if (n.kind === 'from') { leg.from_lat = q.lat; leg.from_lng = q.lng; leg.from_fixed = true; } else { leg.to_lat = q.lat; leg.to_lng = q.lng; leg.to_fixed = true; }
      // les jours « à quai » suivants héritent de la même escale
      if (q.lat !== n.lat || q.lng !== n.lng) snapped = true; }
    if (snapped || toSnap.length) { todo.length = 0; const re = planNodes(c, p).nodes;
      for (let k = 1; k < re.length; k++) { const a2 = re[k - 1], b2 = re[k]; if (a2.lat == null || b2.lat == null || !b2.row || !(b2.row.dist > 0)) continue;
        const leg = p.legs[b2.row.i], key = routeKey(a2, b2, leg.via); if (!leg.route || leg.route.k !== key) todo.push({ a: a2, b: b2, leg, key }); } }
    let k2 = 0;
    for (const x of todo) { onStatus?.(`Calcul des routes maritimes… ${++k2}/${todo.length}`); await new Promise(r => setTimeout(r, 0));
      const pts = await legRoutePoints({ lat: x.a.lat, lng: x.a.lng }, { lat: x.b.lat, lng: x.b.lng }, x.leg.via, land);
      x.leg.route = { k: x.key, pts: pts.map(q => [r4(q.lat), r4(q.lng)]) }; }
    return true;
  }
  const CORINTH = [[37.9155, 23.0080], [37.9475, 22.9585]]; // entrées est / ouest du canal de Corinthe
  const clockMin = s => { if (!s) return null; const [h, m] = s.split(':').map(Number); return h * 60 + m; };
  // Carte réutilisable (capitaine & client). opts: { editable, onMoveNode(node, latlng), labels }
  const MAPVIEW = {};
  async function mountRouteMap(el, c, p, opts = {}) {
    const L = await ensureLeaflet();
    const tx = opts.labels || {};
    const { nodes, P } = planNodes(c, p);
    const located = nodes.filter(n => n.lat != null);
    el.innerHTML = '';
    const map = L.map(el, { zoomControl: true, scrollWheelZoom: !!opts.editable, tap: true });
    const sat = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', { maxZoom: 18, attribution: 'Imagery © Esri, Maxar, Earthstar Geographics' });
    const labelsL = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}', { maxZoom: 18 });
    const plan = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 18, attribution: '© OpenStreetMap' });
    L.layerGroup([sat, labelsL]).addTo(map);
    L.control.layers({ [tx.satellite || 'Satellite']: L.layerGroup([sat, labelsL]), [tx.plan || 'Plan']: plan }, null, { position: 'topright' }).addTo(map);
    const key = c.id;
    if (MAPVIEW[key]) map.setView(MAPVIEW[key].center, MAPVIEW[key].zoom);
    else if (located.length) map.fitBounds(L.latLngBounds(located.map(n => [n.lat, n.lng])).pad(0.25));
    else map.setView([41, 15], 5);
    map.on('moveend', () => { MAPVIEW[key] = { center: map.getCenter(), zoom: map.getZoom() }; });
    // Tracé complet
    const segs = []; // segments navigués : {day, a, b, row, pts (route maritime)}
    for (let k = 1; k < nodes.length; k++) {
      const a = nodes[k - 1], b = nodes[k];
      if (a.lat == null || b.lat == null || !b.row || !(b.row.dist > 0)) continue;
      const leg = p.legs[b.row.i], key = routeKey(a, b, leg.via);
      const pts = leg.route && leg.route.k === key ? leg.route.pts.map(([lat, lng]) => ({ lat, lng })) : [a, ...(leg.via || []).map(([lat, lng]) => ({ lat, lng })), b];
      segs.push({ day: b.day, a, b, row: b.row, leg, legIndex: b.row.i, pts, nm: pathNm(pts), routed: !!(leg.route && leg.route.k === key) });
    }
    const ll = s => s.pts.map(q => [q.lat, q.lng]);
    L.polyline(segs.map(ll), { color: '#ffffff', weight: 2, opacity: .65, dashArray: '6 6' }).addTo(map);
    if (opts.editable) { // clic sur un trajet = ajouter un point de passage ; points de passage déplaçables
      segs.forEach(s => {
        L.polyline(ll(s), { color: '#000', weight: 16, opacity: 0.001 }).addTo(map).on('click', e => { L.DomEvent.stop(e); opts.onAddVia?.(s, e.latlng); });
        (s.leg.via || []).forEach(([lat, lng], vi) => {
          const vm = L.marker([lat, lng], { draggable: true, icon: L.divIcon({ className: 'via-pin', html: '<span></span>', iconSize: [12, 12], iconAnchor: [6, 6] }) }).addTo(map);
          vm.on('dragend', () => opts.onMoveVia?.(s, vi, vm.getLatLng()));
          vm.bindPopup(() => { const d = document.createElement('div'); d.innerHTML = `<b>Point de passage</b><br><button class="btn btn-ghost btn-xs" style="margin-top:6px">Supprimer</button>`;
            d.querySelector('button').onclick = () => opts.onDelVia?.(s, vi); return d; });
        });
      });
    }
    const dayLayer = L.layerGroup().addTo(map);
    // Escales
    const seen = new Map();
    nodes.forEach(n => {
      if (n.lat == null) return;
      const k = n.lat.toFixed(4) + ',' + n.lng.toFixed(4);
      if (seen.has(k)) { seen.get(k).days.add(n.day); return; }
      const mk = L.marker([n.lat, n.lng], { draggable: !!opts.editable, icon: L.divIcon({ className: 'stop-pin', html: `<span></span>`, iconSize: [16, 16], iconAnchor: [8, 8] }) }).addTo(map);
      const ent = { mk, days: new Set([n.day]), node: n };
      mk.bindTooltip(() => `<b>${esc(n.name)}</b><br>${[...ent.days].filter(Boolean).map(d => (tx.day || 'Day') + ' ' + d).join(', ')}`, { direction: 'top', offset: [0, -8] });
      if (opts.editable) mk.on('dragend', () => opts.onMoveNode?.(n, mk.getLatLng()));
      seen.set(k, ent);
    });
    if (opts.editable && opts.onMapClick) map.on('click', e => opts.onMapClick(e.latlng));
    // Bateau
    const boat = L.marker(located[0] ? [located[0].lat, located[0].lng] : [0, 0], { icon: L.divIcon({ className: 'boat-pin', html: '<span>⛴</span>', iconSize: [34, 34], iconAnchor: [17, 17] }), zIndexOffset: 1000, interactive: false });
    if (located.length) boat.addTo(map);
    let anim = null;
    const stop = () => { if (anim) cancelAnimationFrame(anim.raf); anim = null; };
    // chemin d'un jour (ou de tout le charter)
    const pathFor = day => {
      const ss = day == null ? segs : segs.filter(s => s.day === day);
      if (!ss.length) { const n = [...nodes].reverse().find(x => x.lat != null && (day == null || x.day <= day)); return n ? [n] : []; }
      const pts = [ss[0].pts[0]]; ss.forEach(s => pts.push(...s.pts.slice(1))); return pts;
    };
    function highlight(day) {
      dayLayer.clearLayers();
      const ss = day == null ? segs : segs.filter(s => s.day === day);
      ss.forEach(s => L.polyline(ll(s), { color: '#e3b35c', weight: 5, opacity: .95, interactive: false }).addTo(dayLayer));
      seen.forEach(e => e.mk.getElement()?.classList.toggle('on', day == null || e.days.has(day)));
      if (ss.length && opts.fit !== false) map.fitBounds(L.latLngBounds(ss.flatMap(ll)).pad(0.3), { maxZoom: 11 });
    }
    function animate(day, done) {
      stop(); const pts = pathFor(day); if (!pts.length) return;
      if (pts.length === 1) { boat.setLatLng([pts[0].lat, pts[0].lng]); return done?.(); }
      const dists = [0]; for (let k = 1; k < pts.length; k++) dists.push(dists[k - 1] + map.distance([pts[k - 1].lat, pts[k - 1].lng], [pts[k].lat, pts[k].lng]));
      const total = dists[dists.length - 1] || 1, dur = Math.min(9000, Math.max(2500, pts.length * 1400));
      const t0 = performance.now();
      const step = now => {
        const f = Math.min(1, (now - t0) / dur), d = f * total;
        let k = 1; while (k < dists.length - 1 && dists[k] < d) k++;
        const seg = (d - dists[k - 1]) / ((dists[k] - dists[k - 1]) || 1);
        const q = lerp(pts[k - 1], pts[k], Math.max(0, Math.min(1, seg))); boat.setLatLng([q.lat, q.lng]);
        if (f < 1) anim.raf = requestAnimationFrame(step); else { anim = null; done?.(); }
      };
      anim = { raf: requestAnimationFrame(step) };
    }
    // Position prévue maintenant (selon les ETD/ETA du programme)
    function plannedNow() {
      const td = today(), nowMin = new Date().getHours() * 60 + new Date().getMinutes();
      const daySegs = segs.filter(s => s.row.date === td);
      const dayRow = P.rows.find(r => r.date === td && !r.alt);
      if (!dayRow) return null;
      for (const s of daySegs) {
        const a = clockMin(s.row.etd), dur = (s.row.hours || 0) * 60;
        if (a == null) continue;
        if (nowMin < a) return { latlng: s.a, day: dayRow.day, state: 'before' };
        if (nowMin <= a + dur) return { latlng: alongPath(s.pts, (nowMin - a) / (dur || 1)), day: dayRow.day, state: 'sailing' };
      }
      const last = [...nodes].reverse().find(n => n.lat != null && n.date && n.date <= td);
      return last ? { latlng: last, day: dayRow.day, state: 'arrived' } : null;
    }
    setTimeout(() => map.invalidateSize(), 50);
    return { map, highlight, animate, stop, plannedNow, setBoat: ll => boat.setLatLng([ll.lat, ll.lng]), nodes, segs, P,
      days: [...new Set(P.rows.filter(r => !r.alt && r.day != null).map(r => r.day))] };
  }
  // Barre des jours + fiche du jour (commune capitaine / client)
  function routeDaysHtml(P, tx, opts = {}) {
    const days = [...new Set(P.rows.filter(r => !r.alt && r.day != null).map(r => r.day))];
    return `<div class="route-days">
      <button class="day-chip on" data-day="">${tx.all || 'Tout'}</button>
      ${days.map(d => { const r = P.rows.find(x => x.day === d); return `<button class="day-chip ${r.date === today() ? 'today' : ''}" data-day="${d}"><b>${tx.day || 'J'} ${d}</b><small>${fdate(r.date).slice(0, 5)}</small></button>`; }).join('')}
      <button class="btn btn-gold btn-sm" data-route-play>${tx.play || '▶ Parcours'}</button></div>
      <div class="route-info" data-route-info></div>`;
  }
  function routeDayInfo(P, day, tx, showCons, segNm) {
    const rows = P.rows.filter(r => !r.alt && (day == null || r.day === day));
    if (day == null) return `<div class="ri-head"><b>${fmt(P.totalD)} NM</b> · ${hhmm(P.totalH)} ${tx.sailing || 'de navigation'}${showCons ? ` · ${fmt(P.totalL)} L` : ''}</div>`;
    const r0 = rows[0];
    return `<div class="ri-head"><b>${tx.day || 'Jour'} ${day}</b> — ${fdate(r0?.date)}</div>
      ${rows.map(r => `<div class="ri-leg"><div><b>${esc(r.trip || '')}</b>${r.dist ? ` · ${fmt(r.dist)} NM` : ''}${showCons && segNm && segNm[r.i] != null ? ` <span class="muted">(tracé ≈ ${fmt(segNm[r.i], 0)} NM)</span>` : ''}${r.etd && r.dist ? ` · ${esc(r.etd)} → ${esc(r.eta)}` : ''}${showCons ? ` · ${fmt(r.cons)} L` : ''}</div>
        ${r.remarks ? `<div class="muted small">${esc(r.remarks)}</div>` : ''}</div>`).join('')}`;
  }
  function bindRouteControls(root, ctl, P, tx, showCons) {
    const info = $('[data-route-info]', root); let cur = null;
    const segNm = showCons ? Object.fromEntries(ctl.segs.map(s => [s.legIndex, s.nm])) : null;
    const select = (day, play = true) => {
      cur = day; $$('.day-chip', root).forEach(b => b.classList.toggle('on', String(b.dataset.day) === String(day ?? '')));
      info.innerHTML = routeDayInfo(P, day, tx, showCons, segNm); ctl.highlight(day); if (play) ctl.animate(day);
    };
    $$('.day-chip', root).forEach(b => b.addEventListener('click', () => select(b.dataset.day === '' ? null : +b.dataset.day)));
    $('[data-route-play]', root)?.addEventListener('click', () => { // tous les jours, l'un après l'autre
      const days = ctl.days.slice(); let k = 0;
      const next = () => { if (k >= days.length) return; const d = days[k++]; cur = d;
        $$('.day-chip', root).forEach(b => b.classList.toggle('on', String(b.dataset.day) === String(d)));
        info.innerHTML = routeDayInfo(P, d, tx, showCons, segNm); ctl.highlight(d); ctl.animate(d, () => setTimeout(next, 500)); };
      next();
    });
    const now = ctl.plannedNow();
    if (now) { select(now.day, false); ctl.setBoat(now.latlng);
      info.insertAdjacentHTML('afterbegin', `<div class="ri-now">⛴ ${now.state === 'sailing' ? (tx.nowSailing || 'En navigation (position prévue)') : now.state === 'before' ? (tx.nowBefore || 'Au mouillage avant le départ du jour') : (tx.nowArrived || 'Arrivé à l’escale du jour')}</div>`); }
    else select(null, false);
  }

  // --- côté capitaine : panneau carte dans l'onglet Itinerary ---
  function mapPanelHtml(c, p) {
    const { nodes } = planNodes(c, p);
    const missing = nodes.filter(n => n.lat == null);
    return `<section class="panel route-panel">
      <div class="panel-head"><h2>Carte de l’itinéraire</h2>
        <div class="row-gap wrap">
          <label class="check"><input type="checkbox" data-share-itin ${p.share_client ? 'checked' : ''}> <span>Visible par le client <small class="muted">(accueil du jour + itinéraire)</small></span></label>
          <button class="btn btn-ghost btn-sm" data-act="route-geocode">📍 Localiser les escales</button></div></div>
      ${missing.length ? `<div class="banner banner-gold small">📍<div><b>Escales non localisées</b> :
        ${missing.filter((n, k, arr) => arr.findIndex(x => normName(x.name) === normName(n.name)) === k).map(n => `<button class="link-btn" data-place-node="${n.legIndex}:${n.kind}">${esc(n.name || '?')}</button>`).join(' · ')}
        <br><small>Cliquez « Localiser les escales » (recherche automatique), ou cliquez un nom puis cliquez sur la carte. Vous pouvez aussi déplacer chaque point à la main.</small></div></div>` : ''}
      <div class="route-map" id="route-map"><div class="muted center small" style="padding:40px">Chargement de la carte…</div></div>
      <div id="route-ctl"></div>
      <p class="muted small">Routes maritimes calculées automatiquement en évitant la terre (trait de côte 1:10 000 000, indicatif). Cliquez sur un trajet pour ajouter un point de passage (chenal, canal…), glissez-le pour l’ajuster, touchez-le pour le supprimer. Glissez une escale pour corriger sa position.</p>
    </section>`;
  }
  let PLACING = null;
  async function bindMapPanel(c) {
    const el = $('#route-map'); if (!el) return;
    const p = getPlan(c);
    const share = $('[data-share-itin]');
    share?.addEventListener('change', async () => { const q = getPlan(c); q.share_client = share.checked;
      try { await savePlan(c, q); notifyCharter(c.id); toast(share.checked ? 'Itinéraire visible par le client' : 'Itinéraire masqué au client'); } catch (e) { toast(errMsg(e), 'err'); } });
    $$('[data-place-node]').forEach(b => b.addEventListener('click', () => { PLACING = b.dataset.placeNode; $$('[data-place-node]').forEach(x => x.classList.toggle('on', x === b)); toast(`Cliquez sur la carte pour placer « ${b.textContent} »`); }));
    const setNode = async (legIndex, kind, ll) => {
      const q = getPlan(c), leg = q.legs[legIndex]; if (!leg) return;
      if (kind === 'from') { leg.from_lat = r6(ll.lat); leg.from_lng = r6(ll.lng); leg.from_fixed = true; } else { leg.to_lat = r6(ll.lat); leg.to_lng = r6(ll.lng); leg.to_fixed = true; }
      try { await savePlan(c, q); notifyCharter(c.id); route(); } catch (e) { toast(errMsg(e), 'err'); }
    };
    const saveVia = async (fn) => { const q = getPlan(c); fn(q); try { await savePlan(c, q); notifyCharter(c.id); route(); } catch (e) { toast(errMsg(e), 'err'); } };
    try {
      try { if (await computeSeaRoutes(c, p, msg => { el.innerHTML = `<div class="muted center small" style="padding:40px">${msg}</div>`; })) { await savePlan(c, p); notifyCharter(c.id); } }
      catch (e) { toast('Routes maritimes indisponibles (lignes droites) : ' + errMsg(e), 'err'); }
      const ctl = await mountRouteMap(el, c, p, { editable: true,
        onAddVia: (s, ll) => PLACING ? (() => { const [i, k] = PLACING.split(':'); PLACING = null; setNode(+i, k, ll); })() : saveVia(q => { const leg = q.legs[s.legIndex], via = (leg.via || []).slice(), pts = [s.a, ...via.map(([lat, lng]) => ({ lat, lng })), s.b];
          let best = 0, bestCost = Infinity; // insertion qui allonge le moins le trajet
          for (let k = 0; k < pts.length - 1; k++) { const cost = nmBetween(pts[k], ll) + nmBetween(ll, pts[k + 1]) - nmBetween(pts[k], pts[k + 1]); if (cost < bestCost) { bestCost = cost; best = k; } }
          via.splice(best, 0, [r6(ll.lat), r6(ll.lng)]); leg.via = via; }),
        onMoveVia: (s, vi, ll) => saveVia(q => { q.legs[s.legIndex].via[vi] = [r6(ll.lat), r6(ll.lng)]; }),
        onDelVia: (s, vi) => saveVia(q => { q.legs[s.legIndex].via.splice(vi, 1); if (!q.legs[s.legIndex].via.length) delete q.legs[s.legIndex].via; }),
        onMoveNode: (n, ll) => setNode(n.legIndex, n.kind, ll),
        onMapClick: ll => { if (!PLACING) return; const [i, k] = PLACING.split(':'); PLACING = null; setNode(+i, k, ll); } });
      $('#route-ctl').innerHTML = routeDaysHtml(ctl.P, {});
      bindRouteControls($('.route-panel'), ctl, ctl.P, {}, true, ctl);
      // trajets dont la route tracée est bien plus longue que la distance saisie (ex. tour du Péloponnèse au lieu du canal)
      const odd = ctl.segs.filter(s => s.routed && s.row.dist > 0 && s.nm > s.row.dist * 1.4);
      if (odd.length) $('#route-ctl').insertAdjacentHTML('beforeend', `<div class="banner banner-gold small">🧭<div>${odd.map(s => {
        const nearCanal = Math.min(s.a.lat, s.b.lat) - 1 < 37.93 && Math.max(s.a.lat, s.b.lat) + 1 > 37.93 && Math.min(s.a.lng, s.b.lng) - 1 < 22.98 && Math.max(s.a.lng, s.b.lng) + 1 > 22.98;
        return `<b>${esc(s.row.trip)}</b> : route tracée ${fmt(s.nm, 0)} NM pour ${fmt(s.row.dist)} NM saisis — ajoutez un point de passage en cliquant sur le trajet${nearCanal ? ` ou <button class="link-btn" data-canal="${s.legIndex}">passer par le canal de Corinthe</button>` : ''}.`; }).join('<br>')}</div></div>`);
      $$('[data-canal]').forEach(btn => btn.addEventListener('click', () => saveVia(q => { const leg = q.legs[+btn.dataset.canal]; const s = ctl.segs.find(x => x.legIndex === +btn.dataset.canal);
        const e = { lat: CORINTH[0][0], lng: CORINTH[0][1] }, w = { lat: CORINTH[1][0], lng: CORINTH[1][1] };
        leg.via = nmBetween(s.a, e) < nmBetween(s.a, w) ? [CORINTH[0], CORINTH[1]] : [CORINTH[1], CORINTH[0]]; })));
    } catch (e) { el.innerHTML = `<div class="muted center small" style="padding:30px">Carte indisponible : ${esc(errMsg(e))}</div>`; }
  }
  const r6 = v => Math.round(v * 1e6) / 1e6;
  ACTIONS['route-geocode'] = async b => {
    const c = charterById(CURRENT), p = getPlan(c), { nodes } = planNodes(c, p);
    const todo = nodes.filter(n => n.lat == null);
    if (!todo.length) return toast('Toutes les escales sont déjà localisées');
    busy(b, true); let ok = 0, near = nodes.find(n => n.lat != null);
    const cache = {};
    for (const n of nodes) {
      if (n.lat != null) { near = n; continue; }
      const key = normName(n.name);
      let g = cache[key];
      if (g === undefined) { toast(`Recherche de « ${n.name} »…`); try { g = await geocodeName(n.name, near); } catch { g = null; } cache[key] = g; await new Promise(r => setTimeout(r, 1100)); }
      if (g) { const leg = p.legs[n.legIndex]; if (n.kind === 'from') { leg.from_lat = r6(g.lat); leg.from_lng = r6(g.lng); } else { leg.to_lat = r6(g.lat); leg.to_lng = r6(g.lng); } n.lat = g.lat; n.lng = g.lng; near = n; ok++; }
    }
    try { await savePlan(c, p); notifyCharter(c.id); } catch (e) { toast(errMsg(e), 'err'); }
    busy(b, false); route();
    const left = planNodes(charterById(CURRENT), getPlan(charterById(CURRENT))).nodes.filter(n => n.lat == null).length;
    toast(`${ok} escale${ok > 1 ? 's' : ''} localisée${ok > 1 ? 's' : ''}${left ? ` · ${left} à placer à la main` : ''}`, left ? 'err' : undefined);
  };

  // ---------------- Journées : repas & accueil client ----------------
  const localToday = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; };
  // Jours du charter : dates du charter + dates du programme
  function planDates(c, p) {
    const set = new Set((p.legs || []).filter(l => !l.alt && l.date).map(l => l.date));
    if (c.start_date && c.end_date && c.end_date >= c.start_date) for (let d = c.start_date, k = 0; d <= c.end_date && k < 60; d = addDays(d, 1), k++) set.add(d);
    return [...set].sort();
  }
  // Position d'un jour (escale d'arrivée du dernier trajet du jour, sinon la dernière escale connue)
  function dayPosition(c, p, date) {
    const { nodes } = planNodes(c, p);
    const n = [...nodes].reverse().find(x => x.lat != null && x.date && x.date <= date) || nodes.find(x => x.lat != null);
    return n ? { lat: n.lat, lng: n.lng, name: n.name } : null;
  }
  // Météo Open-Meteo (gratuit, sans clé) : prévisions du jour + état de la mer
  const WX_CACHE = {};
  async function fetchWeather(pos, date) {
    const key = `${pos.lat.toFixed(2)},${pos.lng.toFixed(2)},${date}`;
    if (WX_CACHE[key]) return WX_CACHE[key];
    const qs = `latitude=${pos.lat.toFixed(4)}&longitude=${pos.lng.toFixed(4)}&timezone=auto&start_date=${date}&end_date=${date}`;
    const [w, m] = await Promise.all([
      fetch(`https://api.open-meteo.com/v1/forecast?${qs}&daily=weather_code,temperature_2m_max,temperature_2m_min,wind_speed_10m_max,wind_gusts_10m_max,wind_direction_10m_dominant,precipitation_probability_max,sunrise,sunset,uv_index_max&wind_speed_unit=kn`).then(r => r.ok ? r.json() : null).catch(() => null),
      fetch(`https://marine-api.open-meteo.com/v1/marine?${qs}&daily=wave_height_max,sea_surface_temperature_max`).then(r => r.ok ? r.json() : null).catch(() => null)]);
    const d = w?.daily, s = m?.daily;
    if (!d || d.weather_code?.[0] == null) return (WX_CACHE[key] = null);
    return (WX_CACHE[key] = { code: d.weather_code[0], tmax: d.temperature_2m_max[0], tmin: d.temperature_2m_min[0], wind: d.wind_speed_10m_max[0], gust: d.wind_gusts_10m_max?.[0],
      dir: d.wind_direction_10m_dominant[0], rain: d.precipitation_probability_max?.[0], sunrise: d.sunrise?.[0]?.slice(11, 16), sunset: d.sunset?.[0]?.slice(11, 16), uv: d.uv_index_max?.[0],
      wave: s?.wave_height_max?.[0], sst: s?.sea_surface_temperature_max?.[0] });
  }
  const WMO = c => c === 0 ? ['☀️', 'sunny'] : c <= 2 ? ['🌤️', 'partly'] : c === 3 ? ['☁️', 'cloudy'] : c <= 48 ? ['🌫️', 'fog'] : c <= 57 ? ['🌦️', 'drizzle']
    : c <= 67 ? ['🌧️', 'rain'] : c <= 77 ? ['🌨️', 'snow'] : c <= 82 ? ['🌦️', 'showers'] : c <= 86 ? ['🌨️', 'snow'] : ['⛈️', 'storm'];
  const compass = deg => ['N', 'NE', 'E', 'SE', 'S', 'SW', 'W', 'NW'][Math.round(((deg % 360) + 360) % 360 / 45) % 8];
  const beaufort = kn => [1, 4, 7, 11, 17, 22, 28, 34, 41, 48, 56, 64].findIndex(x => kn < x) === -1 ? 12 : [1, 4, 7, 11, 17, 22, 28, 34, 41, 48, 56, 64].findIndex(x => kn < x);

  // --- côté capitaine : panneau « Journées (accueil client) » dans l'onglet Itinerary ---
  function daysPanelHtml(c, p) {
    const dates = planDates(c, p); if (!dates.length) return '';
    const P = computePlan(c, p), days = p.days || {};
    const first = dates[0];
    return `<section class="panel days-panel">
      <div class="panel-head"><h2>Journées — page d’accueil du client</h2></div>
      <p class="muted small">Chaque jour, le client voit à l’ouverture de son lien : la date, le jour, la météo du lieu, le programme du jour et les repas ci-dessous.
        Les menus n’apparaissent que sur cette page d’accueil (pas dans l’itinéraire). ${p.share_client ? '' : '<b>Cochez « Visible par le client » sur la carte pour l’activer.</b>'}</p>
      <div class="days-grid">${dates.map(d => { const x = days[d] || {}, rows = P.rows.filter(r => !r.alt && r.date === d);
        const n = dayDiff(d, first) + 1;
        return `<div class="day-card ${d === localToday() ? 'today' : ''}" data-act="day-edit" data-date="${d}">
          <div class="dc-head"><b>Jour ${n}</b><span>${fdate(d)}</span></div>
          <div class="dc-trip">${rows.length ? rows.map(r => esc(r.trip || '')).join(' · ') : '<span class="muted">—</span>'}</div>
          <div class="dc-meal"><span>🍽 Lunch${x.lunch_time ? ' ' + esc(x.lunch_time) : ''}</span>${x.lunch_menu ? `<small>${esc(x.lunch_menu.split('\n')[0])}${x.lunch_menu.includes('\n') ? '…' : ''}</small>` : '<small class="muted">à compléter</small>'}</div>
          <div class="dc-meal"><span>🍷 Dinner${x.dinner_time ? ' ' + esc(x.dinner_time) : ''}</span>${x.dinner_menu ? `<small>${esc(x.dinner_menu.split('\n')[0])}${x.dinner_menu.includes('\n') ? '…' : ''}</small>` : '<small class="muted">à compléter</small>'}</div>
          ${x.note ? `<div class="dc-note small">✉️ ${esc(x.note.slice(0, 60))}${x.note.length > 60 ? '…' : ''}</div>` : ''}
        </div>`; }).join('')}</div>
    </section>`;
  }
  ACTIONS['day-edit'] = b => {
    const c = charterById(CURRENT), p = getPlan(c), d = b.dataset.date, x = (p.days || {})[d] || {};
    const dates = planDates(c, p), n = dayDiff(d, dates[0]) + 1;
    const m = openModal({ title: `Jour ${n} — ${fdate(d)}`, wide: true,
      body: `<form class="form-grid" id="f-day">
        <label class="field"><span>🍽 Heure du lunch</span><input type="time" name="lunch_time" value="${esc(x.lunch_time || '')}"></label>
        <label class="field"><span>🍷 Heure du dinner</span><input type="time" name="dinner_time" value="${esc(x.dinner_time || '')}"></label>
        <label class="field"><span>Menu du lunch <small>(une ligne par plat)</small></span><textarea name="lunch_menu" rows="6" placeholder="Burrata, tomates anciennes&#10;Linguine aux palourdes&#10;Tiramisu">${esc(x.lunch_menu || '')}</textarea></label>
        <label class="field"><span>Menu du dinner <small>(une ligne par plat)</small></span><textarea name="dinner_menu" rows="6" placeholder="Carpaccio de daurade&#10;Filet de bœuf, légumes grillés&#10;Sorbet citron">${esc(x.dinner_menu || '')}</textarea></label>
        <label class="field span-2"><span>Mot du jour <small>(facultatif — ex. « Tenue de soirée », « Réveil 7h pour la randonnée »)</small></span><textarea name="note" rows="2">${esc(x.note || '')}</textarea></label>
      </form>`,
      footer: `<button class="btn btn-ghost" data-close>Annuler</button><button class="btn btn-primary" data-ok>Enregistrer</button>` });
    $('[data-ok]', m).onclick = async ev => {
      const f = $('#f-day', m); busy(ev.currentTarget, true);
      const q = getPlan(c); q.days = q.days || {};
      const v = { lunch_time: f.lunch_time.value || null, lunch_menu: f.lunch_menu.value.trim() || null, dinner_time: f.dinner_time.value || null, dinner_menu: f.dinner_menu.value.trim() || null, note: f.note.value.trim() || null };
      if (Object.values(v).some(Boolean)) q.days[d] = v; else delete q.days[d];
      try { await savePlan(c, q); notifyCharter(c.id); closeModal(); route(); toast('Journée enregistrée'); } catch (e) { toast(errMsg(e), 'err'); busy(ev.currentTarget, false); }
    };
  };

  // Ligne de dépense automatique « Bank fee » = APA reçu × % (suit l'APA et le taux du charter)
  const nextExpenseNum = cid => Math.max(1, ...S.expenses.filter(x => x.charter_id === cid).map(x => x.num || 0)) + 1;
  async function syncBrokerFee(c) {
    if (!c || (c.status && c.status !== 'open')) return false;
    const pct = toNum(c.broker_fee_pct) ?? 0, apa = toNum(c.apa_paid) || 0, amount = r2(apa * pct / 100);
    const ex = S.expenses.find(e => e.charter_id === c.id && e.auto_kind === 'broker_fee');
    if (!(amount > 0)) { if (ex) { await api.remove('expenses', ex.id); return true; } return false; }
    const description = `Bank fee ${String(pct).replace(".", ",")} % on APA transfer (${money(apa, c.currency)})`;
    if (ex) {
      const date = c.start_date || ex.date;
      if (r2(ex.total) === amount && ex.description === description && ex.date === date) return false;
      await api.update('expenses', ex.id, { total: amount, description, date }); return true;
    }
    await api.insert('expenses', { charter_id: c.id, num: nextExpenseNum(c.id), date: c.start_date || today(),
      company: S.expenses.filter(e => e.auto_kind === 'broker_fee' && e.company).sort((a, b) => String(b.created_at || '').localeCompare(String(a.created_at || '')))[0]?.company || 'Bank',
      category: 'Bank Fee', description, amount_ht: null, vat: null, total: amount, method: 'Bank Fee', auto_kind: 'broker_fee' });
    return true;
  }

  ACTIONS['new-charter'] = () => {
    const last = S.charters.slice().sort((a, b) => String(b.created_at || b.start_date || '').localeCompare(String(a.created_at || a.start_date || '')))[0];
    charterForm({ tips_prompt: last ? tipsEnabled(last) : true, report_to: last?.report_to || '', fuel_company: boatName(), fuel_by_boat: true,
      vat_country: last?.vat_country || 'France', vat_rate: last?.vat_rate ?? 20, auto_voyage: true, broker_fee_pct: last?.broker_fee_pct ?? 0.4 });
  };
  ACTIONS['edit-charter'] = b => charterForm(charterById(b.dataset.id));

  // ---------------- Détail d'un charter ----------------
  let CURRENT = null;
  const FOLDERS = [['apa', 'APA', 'Comptabilité'], ['preferences', 'Preference list', 'Préférences des invités'], ['itinerary', 'Itinerary', 'Itinéraire']];
  function folderTabs(c, tab) {
    const n = f => (S.charter_documents || []).filter(d => d.charter_id === c.id && d.folder === f).length;
    return `<nav class="folder-tabs no-print">${FOLDERS.map(([k, l, sub]) => {
      const cnt = k === 'apa' ? S.expenses.filter(e => e.charter_id === c.id).length : n(k);
      return `<a href="#/charter/${c.id}${k === 'apa' ? '' : '/' + k}" class="folder-tab ${tab === k ? 'on' : ''}">${ICON.folder}
        <span><b>${l}</b><small><span class="ft-sub">${sub}</span>${cnt ? `<span class="ft-sub"> · </span>${cnt} ${k === 'apa' ? 'dép.' : 'élém.'}` : ''}</small></span></a>`; }).join('')}</nav>`;
  }
  function viewCharter(id, tab = 'apa') {
    const c = charterById(id);
    if (!c) { location.hash = '#/charters'; return; }
    CURRENT = id;
    if (!FOLDERS.some(f => f[0] === tab)) tab = 'apa';
    if (tab !== 'apa') return viewCharterFolder(c, tab);
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
      ${folderTabs(c, 'apa')}
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
            ${B.cash.map(w => `<div class="kv"><span>${fdate(w.date)} · ${esc(w.location || '')}${toNum(w.fee) ? ` <small class="muted">(frais ATM ${money(w.fee, c.currency)})</small>` : ''}</span><b>${money(w.amount, c.currency)}
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
          ${!locked ? `<span class="muted small">Le client pourra signer une fois la compta clôturée (la question des tips vous sera posée à la clôture).</span>` : ''}
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

  function tipsQuestion(c) {
    return new Promise(res => {
      const fee = c.charter_fee ? `10 % = ${money(c.charter_fee * 0.1, c.currency)} · 15 % = ${money(c.charter_fee * 0.15, c.currency)} · 20 % = ${money(c.charter_fee * 0.2, c.currency)}`
        : '<span class="warn">Charter fee non renseigné : le client ne pourra saisir qu’un montant libre.</span>';
      const m = openModal({ title: 'Fenêtre tips',
        body: `<p class="tips-reminder">🔔 <b>Voulez-vous que la fenêtre des tips apparaisse au client au moment de sa signature ?</b></p>
          <p class="muted">Elle lui proposera 10 %, 15 % ou 20 % du charter fee, ou un montant libre.<br>${fee}</p>`,
        footer: `<button class="btn btn-ghost mr-auto" data-close>Annuler</button><button class="btn btn-ghost" data-no>Non</button><button class="btn btn-primary" data-yes>Oui</button>` });
      let done = false; const fin = v => { done = true; closeModal(); res(v); };
      $('[data-yes]', m).addEventListener('click', () => fin(true));
      $('[data-no]', m).addEventListener('click', () => fin(false));
      const obs = new MutationObserver(() => { if (!$('#modal-root').innerHTML) { obs.disconnect(); if (!done) res(null); } });
      obs.observe($('#modal-root'), { childList: true });
    });
  }

  ACTIONS['close-charter'] = async b => {
    const c = charterById(b.dataset.id), st = computeStatement(buildBundle(S, c));
    const warn = [];
    if (!S.voyages.some(v => v.charter_id === c.id)) warn.push('Aucun voyage fuel n’est lié à ce charter : aucun fuel ne sera facturé.');
    const draft = boatInvoiceDraft(buildBundle(S, c));
    if (draft.lines.length && !PROFILE.address) warn.push('La fiche du bateau (adresse de facturation) n’est pas complétée dans <b>Paramètres</b> : la facture sera incomplète.');
    if (draft.lines.length && !c.client_address) warn.push('L’adresse de facturation du client n’est pas renseignée (bouton Modifier).');
    const ok = await confirmDialog(`Clôturer la compta de « ${esc(c.title)} » ?<br><br>
      Total : <b>${money(st.total, c.currency)}</b> · Solde APA : <b>${money(st.balance, c.currency)}</b><br><br>
      Les dépenses seront verrouillées et le client pourra signer.
      ${draft.lines.length ? `<br>Une facture de <b>${esc(boatName())}</b> (${money(draft.total_ttc, c.currency)}) sera émise automatiquement.` : ''}
      ${warn.length ? `<br><br><span class="warn">⚠ ${warn.join('<br>⚠ ')}</span>` : ''}`, 'Clôturer', false);
    if (!ok) return;
    const tips = await tipsQuestion(c);
    if (tips === null) return;
    try {
      await api.update('charters', c.id, { status: 'closed', closed_at: new Date().toISOString(), tips_prompt: tips });
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
        <label class="field"><span>Montant retiré</span><input type="number" step="0.01" name="amount" required></label>
        <label class="field"><span>Frais ATM <small>(ajoutés en dépense « ATM fee »)</small></span><input type="number" step="0.01" min="0" name="fee" placeholder="0,00"></label></form>`,
      footer: `<button class="btn btn-ghost" data-close>Annuler</button><button class="btn btn-primary" data-ok>Enregistrer</button>`,
    });
    $('[data-ok]', m).onclick = async ev => {
      const f = $('#f-cash', m); if (!f.reportValidity()) return;
      busy(ev.currentTarget, true);
      try {
        const fee = r2(toNum(f.fee.value) || 0), date = f.date.value || null, location = f.location.value || null;
        const w = await api.insert('cash_withdrawals', { charter_id: CURRENT, date, location, amount: toNum(f.amount.value), ...(fee > 0 ? { fee } : {}) });
        if (fee > 0) await api.insert('expenses', { charter_id: CURRENT, num: nextExpenseNum(CURRENT), date, company: location ? `ATM ${location}` : 'ATM',
          category: 'ATM fee', description: `ATM fee — cash withdrawal ${money(w.amount, charterById(CURRENT).currency)}`, amount_ht: null, vat: null,
          total: fee, method: 'ATM Fee', auto_kind: 'atm_fee', auto_ref: w.id });
        await reload(); closeModal(); notifyCharter(CURRENT); route(); if (fee > 0) toast('Retrait et frais ATM ajoutés'); }
      catch (e) { toast(errMsg(e), 'err'); busy(ev.currentTarget, false); }
    };
  };
  ACTIONS['del-cash'] = async b => {
    const fees = S.expenses.filter(e => e.auto_kind === 'atm_fee' && e.auto_ref === b.dataset.id);
    if (!(await confirmDialog(`Supprimer ce retrait${fees.length ? ' et sa dépense « ATM fee »' : ''} ?`, 'Supprimer'))) return;
    try { for (const e of fees) await api.remove('expenses', e.id);
      await api.remove('cash_withdrawals', b.dataset.id); await reload(); notifyCharter(CURRENT); route(); } catch (e) { toast(errMsg(e), 'err'); }
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
