# GitHub Actions Secrets – Supabase

**Huwag ilagay dito ang tunay na URL o keys.** Ang file na ito ay template lang; ang totoong halaga ay nasa GitHub repo settings lamang.

Para magamit ang Supabase sa GitHub Actions, idagdag ang mga secret sa repo (hindi sa markdown file).

## Paano idagdag sa GitHub

1. **Punta sa repo** → **Settings** → **Secrets and variables** → **Actions**.
2. **New repository secret** — idagdag ang dalawa:

| Name | Halaga (ilagay sa GitHub UI, hindi sa repo) |
|------|---------------------------------------------|
| `SUPABASE_URL` | Project URL mula sa Supabase → Settings → API (hal. `https://xxxxx.supabase.co`) |
| `SUPABASE_ANON_KEY` | **anon public** key mula sa parehong page |

3. **Add secret** para sa bawat isa.

## Bakit hindi sa git?

- Ang sinumang may access sa repo (o public clone) ay makakakita ng anumang naka-commit na key.
- Kahit “anon” key, mas mabuting hindi i-publish sa docs; gamitin ang GitHub **Encrypted secrets** para sa Actions.

## Kung na-push na dati ang tunay na keys

- Pwede kang mag-**rotate** ng API keys sa Supabase → Settings → API (kung available), o bumuo ng bagong project kung kailangan.
- Tandaan: mananatili pa rin sa **git history** ang lumang commit hanggang hindi nire-write ang history (advanced).
