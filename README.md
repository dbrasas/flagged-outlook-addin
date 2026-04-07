# 🚩 Flagged by Due Date — Outlook Add-in

Rodo visus flagged laiškus iš **visų aplankų**, surūšiuotus pagal Due Date (ankščiausias viršuje), suskirstytus į kategorijas: **Vėluoja · Šiandien · Artimiausi · Be termino**.

---

## ✅ Reikalavimai

- Windows 10/11 su naujuoju Outlook
- Node.js 18+ (https://nodejs.org)
- Microsoft 365 paskyra (darbo arba asmeninė)

---

## 1️⃣ Azure App registracija (VIENAS KARTAS)

Add-in naudoja Microsoft Graph API skaityti laiškams — tam reikia Azure App ID.

1. Eik į https://portal.azure.com → **Azure Active Directory** → **App registrations** → **New registration**
2. Pavadinimas: `FlaggedSorterAddin`
3. Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**
4. Redirect URI: `https://dbrasas.github.io/flagged-outlook-addin/src/taskpane.html` (tipas: **Single-page application**)
5. Spusk **Register**
6. Nukopijuok **Application (client) ID**
7. Eik į **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated** → pridėk:
   - `Mail.Read`
8. Spusk **Grant admin consent** (arba paprašyk IT admino)

### Tada įdėk savo Client ID į kodą:

Atsidaryk `src/taskpane.html`, rask šią eilutę (~367):
```javascript
const CLIENT_ID = 'YOUR_CLIENT_ID_HERE';
```
Pakeisk `YOUR_CLIENT_ID_HERE` į savo **Application (client) ID**.

---

### ✅ SVARBU: Saugumas
**Sertifikatai niekada neturi būti keliami į Git (GitHub)!**  
Projektas automatiškai ignoruoja `certs/` aplanką per `.gitignore`. Jei netyčia įkeltumėte raktus, juos būtina nedelsiant pergeneruoti.

---

## 2️⃣ SSL sertifikatai (lokaliam serveriui)

Outlook reikalauja HTTPS net dirbant lokaliai. Naudosime oficialius Microsoft dev sertifikatus:

```bash
# Sugeneruok ir patikimuose įdiek sertifikatus
npx office-addin-dev-certs install

# Sukurk aplanką sertifikatams (jis nebus keliamas į Git)
mkdir certs

# Nukopijuok iš sistemos į projektą (Windows pavyzdys):
copy "%USERPROFILE%\.office-addin-dev-certs\localhost.key" certs\server.key
copy "%USERPROFILE%\.office-addin-dev-certs\localhost.crt" certs\server.crt
```

---

## 3️⃣ Paleisk serverį

```bash
# Projekto aplanke:
npm install
node server.js
```

Turėtum matyti:
```
✅ Add-in serveris veikia: https://localhost:3000
📋 Manifest: https://dbrasas.github.io/flagged-outlook-addin/manifest.xml
```

Patikrink naršyklėje: https://dbrasas.github.io/flagged-outlook-addin/src/taskpane.html — turi atsidaryti UI.

---

## 4️⃣ Įdiek Add-in į Outlook

### Greičiausias būdas (Outlook Web / New Outlook):
1. Naršyklėje atsidaryk: [https://aka.ms/olksideload](https://aka.ms/olksideload) (arba [https://outlook.office.com/mail/addins](https://outlook.office.com/mail/addins))
2. Pasirink **My add-ins** skirtuką.
3. Apačioje spusk **+ Add a custom add-in** → **Add from File...**
4. Pasirink `manifest.xml` failą.

### Naujasis Outlook (Windows Rankiniu būdu):
1. Spusk ⚙️ → **View all Outlook settings** → **Mail** → **Customize actions** → **Add-ins**
2. **My add-ins** → **Add a custom add-in** → **Add from URL**
3. Įvesk: `https://dbrasas.github.io/flagged-outlook-addin/manifest.xml`

---

## 5️⃣ Naudojimas

1. Atidaryk bet kurį laišką Outlook
2. Toolbar viršuje pamatyk mygtuką **📋 Flagged by Date**
3. Spusk — dešinėje atsidaro panel su visais flagged laiškais
4. Laiškus skirsto automatiškai: 🔴 Vėluoja → 🟡 Šiandien → 🟢 Artimiausi → ⬜ Be termino
5. Spusk ant laiško — jis atsidaro naršyklėje (Outlook on the web)

---

## ❓ Dažnos klaidos

| Klaida | Sprendimas |
|--------|-----------|
| SSL klaida naršyklėje | Paleisk `npx office-addin-dev-certs install` iš naujo |
| "Nepavyko prisijungti" | Patikrink CLIENT_ID ir Graph permissions |
| Add-in neatsiranda | Iš naujo paleisk Outlook po manifest instaliavimo |
| "Graph klaida: 403" | Azure App neturi `Mail.Read` leidimo arba trūksta admin consent |

---

## 📁 Projekto struktūra

```
flagged-addin/
├── manifest.xml          ← Outlook add-in aprašas
├── server.js             ← Lokalus HTTPS serveris
├── package.json
├── certs/                ← SSL sertifikatai (**IGNORUOJAMA GIT**)
│   ├── server.key        ← Jūsų privatus raktas
│   └── server.crt
├── .gitignore            ← Apsaugo jautrius duomenis
└── src/
    └── taskpane.html     ← Visas UI + logika
```

---

## 🚀 Gamybinė versija (cloud hosting)

Jei nori naudoti be lokalaus serverio, galima hostinti nemokamais servisais:
- **GitHub Pages** (tik statiniam HTML)
- **Vercel** / **Netlify** — nemokamas HTTPS

Tada manifest.xml URL pakeisk į savo cloud adresą.
