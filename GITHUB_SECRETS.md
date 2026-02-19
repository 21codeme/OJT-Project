# GitHub Actions Secrets – Supabase

Para magamit ang Supabase sa GitHub Actions (o para may record ka ng config), idagdag ang mga secret sa repo.

## Paano idagdag sa GitHub

1. **Punta sa repo:** https://github.com/21codeme/OJT-Project  
2. **Settings** (tab sa taas) → sa left sidebar, **Code and automation** → **Actions** → **Secrets and variables** → **Actions**.  
3. **New repository secret.**  
4. Idagdag **dalawa** (isang beses bawat isa):

---

### Secret 1: `SUPABASE_URL`

| Field   | Value |
|--------|--------|
| **Name**   | `SUPABASE_URL` |
| **Secret** | `https://bferfkrkejwccvfsigze.supabase.co` |

Tapos **Add secret**.

---

### Secret 2: `SUPABASE_ANON_KEY`

| Field   | Value |
|--------|--------|
| **Name**   | `SUPABASE_ANON_KEY` |
| **Secret** | `eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJmZXJma3JrZWp3Y2N2ZnNpZ3plIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzAyNzM1NTUsImV4cCI6MjA4NTg0OTU1NX0.4nc1SgH-lXD4GvZ6XSbfzyCp-Swf6Mon-O3dA_mEpXE` |

Tapos **Add secret**.

---

## Saan makikita sa GitHub

- **Settings** → **Actions** → **Secrets and variables** → **Actions**  
- Dapat makita mo: **SUPABASE_URL** at **SUPABASE_ANON_KEY**.

Kung may workflow na gumagamit ng `secrets.SUPABASE_URL` o `secrets.SUPABASE_ANON_KEY`, gagana na iyon pagkatapos mong idagdag ang dalawang secret na ito.
