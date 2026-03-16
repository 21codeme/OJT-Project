# Inventory Pictures

Ang mga larawan ng inventory ay naka-store sa **Supabase Storage** (bucket: `inventory-pictures`). Pwede mong i-sync ang mga iyon dito sa local folder para makita sa project.

## I-sync ang mga picture dito (automatic copy mula Supabase)

1. **Lagyan ng credentials** ang `.env` sa **project root** (same level ng `package.json`):
   ```
   SUPABASE_URL=https://xxxx.supabase.co
   SUPABASE_ANON_KEY=your-anon-key
   ```
   (Makikita sa Supabase Dashboard → Settings → API.)

2. **I-run ang sync** (sa project root):
   ```bash
   npm install
   npm run sync-pictures
   ```
   Lalabas dito sa folder na ito ang lahat ng picture mula sa Supabase, naka-folder pa rin ayon sa PC section (PC 1, PC 2, etc.).

3. **Para parang “automatic”:** Pwede mong i-schedule (hal. araw-araw) ang `npm run sync-pictures`:
   - **Windows:** Task Scheduler → Create Task → Program: `npm`, Arguments: `run sync-pictures`, Start in: `C:\OJT Project\OJT-Project`

## Structure

```
inventory-pictures/
├── PC 1/
│   └── sheet-2/5/1734567890123.jpg
├── PC 2/
│   └── ...
└── uncategorized/
    └── ...
```

- Bawat **PC section** = isang folder. Sa loob: `{sheetId}/{rowIndex}/{timestamp}.jpg`.
- Kapag na-export at na-import ang Excel, gumagana pa rin ang picture sa link dahil naka-save ang URL.
