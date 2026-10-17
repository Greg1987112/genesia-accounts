// Données FICTIVES utilisées uniquement en mode démo (avant de configurer Supabase).
// Vos vraies données ne sont jamais dans ce fichier : elles vivent dans Supabase.
window.DEMO_DATA = {
  charters: [{
    id: "demo-charter-1", title: "Charter Martin — Août 2026 (démo)", yacht: "Motor Yacht GENESIA",
    client_name: "Famille Martin", start_date: "2026-08-08", end_date: "2026-08-15", currency: "EUR",
    apa_paid: 40000, charter_fee: 120000, fuel_company: "TDB", fuel_description: "MEs & DGs",
    fuel_method: "To be paid", client_token: "demo-token-martin", status: "open", notes: null
  }],
  expenses: [
    { id: "d-e1", charter_id: "demo-charter-1", num: 2, date: "2026-08-01", company: "Banque", invoice_ref: null, category: "Bank Fee", description: "Frais de virement APA", amount_ht: null, vat: null, total: 45, method: "Bank Fee" },
    { id: "d-e2", charter_id: "demo-charter-1", num: 3, date: "2026-08-06", company: "Cave du Port", invoice_ref: "F-2026-118", category: "Drinks", description: "Vins & champagne", amount_ht: 3250, vat: 650, total: 3900, method: "Bank Transfer from APA" },
    { id: "d-e3", charter_id: "demo-charter-1", num: 4, date: "2026-08-07", company: "Primeur Marin", invoice_ref: "2241", category: "Food", description: "Viandes, poissons, fruits & légumes", amount_ht: null, vat: null, total: 6480.5, method: "Bank Transfer from APA" },
    { id: "d-e4", charter_id: "demo-charter-1", num: 5, date: "2026-08-10", company: "Marina di Portisco", invoice_ref: "MP-5521", category: "Berth", description: "Place de port 1 nuit", amount_ht: 1500, vat: 330, total: 1830, method: "Credit Card APA" },
    { id: "d-e5", charter_id: "demo-charter-1", num: 6, date: "2026-08-11", company: "Supermarket", invoice_ref: "Receipt", category: "Food", description: "Pain, fromages", amount_ht: null, vat: null, total: 186.4, method: "Cash APA" },
    { id: "d-e6", charter_id: "demo-charter-1", num: 7, date: "2026-08-12", company: "Station Tender", invoice_ref: "Receipt", category: "Fuel", description: "Tender", amount_ht: 210, vat: 46.2, total: 256.2, method: "Credit Card APA" }
  ],
  cash_withdrawals: [
    { id: "d-c1", charter_id: "demo-charter-1", date: "2026-08-08", location: "Olbia", amount: 1000 }
  ],
  bunkers: [
    { id: "d-b1", num: 1, date: "2026-05-29", country: "France", town: "La Seyne-sur-mer", quantity: 20000, fuel_type: "GASOIL B7", price_ht_l: 1.76, vat_l: 0, total_price: 35200, invoice_ref: "H0001" },
    { id: "d-b2", num: 2, date: "2026-07-20", country: "Italy", town: "Olbia", quantity: 25000, fuel_type: "GASOIL B7", price_ht_l: 1.55, vat_l: 0.341, total_price: 47275, invoice_ref: "IT-778" }
  ],
  voyages: [
    { id: "d-v1", position: 1, num: 1, description: "Positioning Toulon (France) > Genova (Italy)", start_date: "2026-06-01", end_date: "2026-06-02", rob_before: 20000, rob_after: 17500, hours: 16, is_charter: false, charter_id: null },
    { id: "d-v2", position: 2, num: 2, description: "Positioning Genova (Italy) > Olbia (Italy)", start_date: "2026-07-18", end_date: "2026-07-19", rob_before: 17500, rob_after: 13000, hours: 22, is_charter: false, charter_id: null },
    { id: "d-v3", position: 3, num: 3, description: "Charter Martin", start_date: "2026-08-08", end_date: "2026-08-15", rob_before: 38000, rob_after: 31200, hours: null, is_charter: true, charter_id: "demo-charter-1" }
  ],
  allocations: [
    { id: "d-a1", voyage_id: "d-v1", bunker_id: "d-b1", liters: 2500 },
    { id: "d-a2", voyage_id: "d-v2", bunker_id: "d-b1", liters: 4500 },
    { id: "d-a3", voyage_id: "d-v3", bunker_id: "d-b1", liters: 6800 }
  ]
};
