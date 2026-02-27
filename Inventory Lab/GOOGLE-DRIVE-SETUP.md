# Paano I-Set Up ang Google Drive (Upload PC Forms)

Sundin ang mga hakbang na ito para gumana ang **Connect Google Drive** at **Upload PC forms to Drive** sa Inventory app.

---

## 1. Google Cloud Console — Gumawa ng Project (kung wala pa)

1. Pumunta sa: **https://console.cloud.google.com/**
2. Mag-sign in sa Google account (yung email na gagamitin para sa Drive).
3. Sa taas, click ang **project dropdown** (malapit sa "Google Cloud").
4. Click **New Project** → lagyan ng name (hal. "Lab Inventory") → **Create**.

---

## 2. I-Enable ang Google Drive API

1. Sa left sidebar: **APIs & Services** → **Library** (o **Enable APIs and Services**).
2. Sa search box, type: **Google Drive API**.
3. Click **Google Drive API** → click **Enable**.

---

## 3. Gumawa ng OAuth Consent Screen (kailangan bago ang Credentials)

1. **APIs & Services** → **OAuth consent screen**.
2. Piliin **External** (para kahit anong Google account) → **Create**.
3. Punuan:
   - **App name:** Lab Inventory (o kahit ano)
   - **User support email:** piliin ang iyong email
   - **Developer contact:** iyong email
4. Click **Save and Continue**.
5. Sa **Scopes** → **Add or Remove Scopes** → hanapin at i-check:
   - `https://www.googleapis.com/auth/drive.file`
   - `https://www.googleapis.com/auth/userinfo.email`
6. **Save and Continue** → **Back to Dashboard**.
7. Kung may "Publishing status" → **Publish App** (o iwan muna sa Testing at idagdag ang iyong email sa Test users).

---

## 4. Gumawa ng OAuth 2.0 Client ID (Credentials)

1. **APIs & Services** → **Credentials**.
2. Click **+ Create Credentials** → **OAuth client ID**.
3. **Application type:** **Web application**.
4. **Name:** hal. "Lab Inventory Web".
5. Sa **Authorized JavaScript origins** click **+ ADD URI** at idagdag **lahat** ng host kung saan mo binuksan ang app:
   - **Kung Vercel:** `https://<iyong-app>.vercel.app` (hal. `https://ojt-project.vercel.app`)
   - **Kung GitHub Pages:** `https://21codeme.github.io`
   Walang trailing slash, walang path.

6. Sa **Authorized redirect URIs** click **+ ADD URI** at idagdag **lahat** ng callback URL:
   - **Kung Vercel:** `https://<iyong-app>.vercel.app/oauth-callback.html`
   - **Kung GitHub Pages:** `https://21codeme.github.io/OJT-Project/oauth-callback.html`
   Dapat **exact match**. Ang app ay awtomatikong gumagamit ng tamang redirect depende sa kung saan naka-open (Vercel o GitHub).

7. Click **Create**.
8. Lalabas ang **Client ID** at **Client secret**. **I-copy ang Client ID** (hindi kailangan ang secret para sa flow na ginagamit ng app).

---

## 4b. Idagdag ang Vercel URLs (kung naka-Vercel ang app)

Kung naka-deploy ang app sa **Vercel**, kailangan idagdag ang Vercel URLs sa OAuth client:

1. **Alamin ang Vercel URL mo**  
   Buksan ang app sa Vercel at tingnan ang address bar. Hal.: `https://lab-inventory-abc123.vercel.app` o `https://ojt-project.vercel.app`. Yan ang **origin** mo (walang path, walang trailing slash).

2. Pumunta sa **Google Cloud Console** → **APIs & Services** → **Credentials**.

3. Sa **OAuth 2.0 Client IDs**, hanapin ang **"Lab Inventory Web"** at i-click ang **icon na lapis (Edit)** sa kanan.

4. **Authorized JavaScript origins**
   - I-click **+ ADD URI**.
   - Ilagay: `https://<iyong-vercel-app>.vercel.app`  
     Hal.: `https://ojt-project.vercel.app`  
     Dapat **exact** — walang `/` sa dulo, walang path.

5. **Authorized redirect URIs**
   - I-click **+ ADD URI**.
   - Ilagay: `https://<iyong-vercel-app>.vercel.app/oauth-callback.html`  
     Hal.: `https://ojt-project.vercel.app/oauth-callback.html`  
     Dapat **exact** — may `/oauth-callback.html` sa dulo.

6. I-click **SAVE** sa baba ng page.

7. Maghintay ng 1–2 minuto, tapos subukan ulit ang **Connect Google Drive** sa app na naka-open sa **Vercel URL** mo.

Pwedeng naka-add na rin ang GitHub Pages URLs; okay lang na pareho naka-list kung gagamitin mo both.

---

## 5. Ilagay ang Client ID sa App

1. Buksan ang file: **`Inventory Lab/config.js`**.
2. Hanapin ang linya:
   ```js
   const GOOGLE_DRIVE_CLIENT_ID = ''; // e.g. '123456789-xxx.apps.googleusercontent.com'
   ```
3. Ilagay ang Client ID sa loob ng quotes:
   ```js
   const GOOGLE_DRIVE_CLIENT_ID = '123456789012-xxxxxxxxxxxxxxxxxx.apps.googleusercontent.com';
   ```
4. I-save ang file. I-commit at i-push kung gusto mo (wag ilagay ang Client ID kung public repo at ayaw mo ipakita; pwede naman i-keep sa local lang).

---

## 6. Dapat Naka-HTTPS ang App

- **Hindi gagana** kung binuksan ang app bilang **file://** (doble-click sa HTML).
- **Kailangan** naka-host sa **HTTPS**, hal.:
  - **GitHub Pages:** i-push ang repo, then **Settings** → **Pages** → Source: main branch → Save. Ang URL ay `https://<username>.github.io/<repo>/` — doon mo kunin ang base para sa redirect URI.
  - O kahit anong web server na naka-HTTPS (Vercel, Netlify, etc.).

---

## 7. Gamitin sa App

1. Buksan ang **Inventory** app sa **HTTPS** URL (e.g. GitHub Pages).
2. Sa **Google Drive folder name:** ilagay ang gusto mong name ng folder sa Drive (hal. "Lab Inventory PC Forms").
3. Click **🔗 Connect Google Drive**.
   - Bubukas ang popup ng Google → piliin ang account → **Allow**.
   - Pag success, sasara ang popup at lalabas na "Connected".
4. Pag may data na sa table (at may PC sections at items), click **📤 Upload PC forms to Drive**.
5. Sa Google Drive mo, makikita ang:
   - **Root folder** (yung name na nilagay mo)
   - Sa loob: **folder per sheet** (Sheet 1, Sheet 2, …)
   - Sa loob ng bawat sheet: **folder per PC section** (PC 1, PC 2, …)
   - Sa loob ng bawat PC section: **HTML file per item** (nakapicture na form, same layout ng PC Location view).

---

## Madalas na Problema

| Problema | Gawin |
|----------|--------|
| "Redirect URI mismatch" | Tiyaking **eksakto** ang redirect URI sa Google Console (may `Inventory%20Lab/oauth-callback.html`, tama ang domain at path). |
| "Connect" walang nangyayari / alert na lagay Client ID | Naka-lagay na ba ang `GOOGLE_DRIVE_CLIENT_ID` sa `config.js`? Naka-HTTPS ba ang page (hindi file://)? |
| Popup na-block ng browser | I-allow ang popup para sa site na ginagamit mo. |
| Token expired | Mag-reconnect (click ulit **Connect Google Drive**); normal na mag-expire ang token pagkatapos ng ilang oras. |

---

## Quick checklist

- [ ] Google Drive API naka-**Enable**
- [ ] OAuth consent screen na-**configure** (External, scopes: drive.file + userinfo.email)
- [ ] **OAuth 2.0 Client ID** (Web application) na-create
- [ ] **Redirect URI** na-add: `https://<domain>/Inventory%20Lab/oauth-callback.html`
- [ ] **Client ID** na-copy at nilagay sa `config.js` → `GOOGLE_DRIVE_CLIENT_ID`
- [ ] App na-**host sa HTTPS** (e.g. GitHub Pages)
- [ ] Sa app: folder name → Connect → Upload PC forms to Drive
