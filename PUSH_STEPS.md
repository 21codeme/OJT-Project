# Push the keep-alive fix (one-time)

Sa terminal (PowerShell o CMD), pumunta sa project folder at patakbuhin:

```powershell
cd "c:\OJT Project\OJT-Project"
```

Kung first time mo mag-commit dito, set muna identity (palitan ng sariling name/email kung gusto):

```powershell
git config user.email "your-email@example.com"
git config user.name "Your Name"
```

Tapos i-stage, commit, at push:

```powershell
git add .github/workflows/keep-alive.yml
git commit -m "fix: ping Supabase API directly for keep-alive (not just Vercel)"
git push
```

Kung may prompt para sa GitHub login, mag-sign in ka (browser o token). Pagkatapos ng push, idagdag sa repo ang **Secrets**: `SUPABASE_URL` at `SUPABASE_ANON_KEY` (Settings → Secrets and variables → Actions).
