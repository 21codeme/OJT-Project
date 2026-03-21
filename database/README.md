# Database Folder

This folder contains all database-related files for the Lab Inventory Management System.

## Files

- **`supabase-setup.sql`** - SQL script for **Lab Inventory** (inventory_items, sheets)
- **`class-schedule-supabase.sql`** - SQL script for **Class Schedule** (class_schedule_sheets, class_schedule_entries)
- **`SUPABASE_SETUP.md`** - Step-by-step instructions for setting up and connecting to Supabase

## Quick Setup

1. **Get Supabase Credentials**
   - Go to your Supabase project dashboard
   - Settings > API
   - Copy Project URL and anon key

2. **Update Config**
   - Edit `../config.js` in the root directory
   - Add your Supabase URL and anon key

3. **Run SQL Setup**
   - Open Supabase SQL Editor
   - For **Lab Inventory**: run `supabase-setup.sql`
   - For **Class Schedule**: run `class-schedule-supabase.sql`

4. **Verify Connection**
   - Open the application
   - Check browser console for "Supabase connected successfully"

5. **Storage bucket (para lumabas ang picture sa PC Location link)**
   - Sa Supabase Dashboard: **Storage** → **New bucket**
   - Name: `inventory-pictures`
   - Public bucket: **ON**
   - Create bucket.
   - Pag may error na **"new row violates row-level security policy"**, i-run sa **SQL Editor** ang file na **`database/storage-policies.sql`** para payagan ang upload at read sa bucket na iyon.
   - **Folder structure:** Sa loob ng bucket, ang mga larawan ay naka-organize bilang: **{PC Section} / {sheetId} / {rowIndex} / {timestamp}.jpg** (hal. `PC 1/sheet-2/5/1234567890.jpg`). Lahat ng picture na na-upload sa isang PC section (e.g. PC 1) ay nasa iisang folder; kapag na-import ang Excel, gumagana pa rin ang picture dahil naka-save ang URL sa link.

6. **OJT Trainee documents (Edit Profile → Required Documents)**
   - Sa **Storage** → **New bucket**: name `ojt-trainee-documents`, **Public: ON**
   - I-run ang **`database/ojt-documents-storage.sql`** sa SQL Editor (tulad ng `storage-policies.sql` para sa inventory)
   - Ang mga file ay naka-upload sa path na `{trainee_id}/{doc_type}{extension}`; ang metadata at public URL ay nasa table na `ojt_trainee_documents.file_data` (JSON).

7. **OJT Daily Logs photos (Dashboard → Daily Logs)**
   - **Mabilis na fix (404 sa REST, 400 sa Storage):** i-run ang **`database/setup-ojt-daily-logs-one-shot.sql`** sa SQL Editor — gumagawa ng table `ojt_daily_logs`, bucket `ojt-daily-logs`, storage policies, at `NOTIFY` para sa schema cache.
   - Kailangan umiiral na ang `ojt_trainees`. Kung hindi pa, i-run muna ang buong **`database/ojt-tables.sql`**.
   - **Manwal (kung ayaw mo ng one-shot):** Storage → New bucket `ojt-daily-logs` (Public ON) → run **`ojt-daily-logs-storage.sql`** → siguraduhing may `ojt_daily_logs` mula sa **`ojt-tables.sql`**
   - Photo path: `{trainee_id}/{log_date}/{timestamp}.{ext}`

8. **Trainee forgot password (Attendance login)**
   - I-run ang **`database/ojt-trainee-password-reset.sql`** sa SQL Editor (kailangan umiiral na ang `ojt_trainees`).
   - Sa login page: **Nakalimutan ang password?** → email → 6-digit code → bagong password.
   - **Email (opsyonal):** Sa `config.js`, punan ang `OJT_PASSWORD_RESET_EMAILJS`:
     - **publicKey** — EmailJS → Account → API Keys → Public Key (i-copy; dapat tumugma byte-per-byte).
     - **serviceId** — Email Services → ang service (hal. `service_...`).
     - **templateId** — mula sa URL ng template editor (hal. `.../templates/3td75xs` → `3td75xs`). Hindi palaging `template_xxxxx`.
   - Sa EmailJS template: **To email** = `{{to_email}}`; sa content: `{{reset_code}}`, at `{{email}}` kung kailangan sa Reply-To.
   - **Handa nang i-paste na HTML:** `database/emailjs-template-password-reset.html`.
   - Account → **Allowed domains:** idagdag ang production domain (hal. `ojt-project-laboratory.vercel.app`) at `localhost`.
   - Kung walang EmailJS o mali ang config, lalabas ang code sa page pagkatapos mag-request (**fallback**).

## Database Schema

### Tables

- **`sheets`** - Stores sheet information
  - `id` (TEXT, PRIMARY KEY)
  - `name` (TEXT)
  - `created_at` (TIMESTAMP)
  - `updated_at` (TIMESTAMP)

- **`inventory_items`** - Stores all inventory data
  - `id` (UUID, PRIMARY KEY)
  - `sheet_id` (TEXT, FOREIGN KEY)
  - `sheet_name` (TEXT)
  - `row_index` (INTEGER)
  - `article` (TEXT)
  - `description` (TEXT)
  - `old_property_n_assigned` (TEXT)
  - `unit_of_meas` (TEXT)
  - `unit_value` (TEXT)
  - `quantity` (TEXT)
  - `location` (TEXT)
  - `condition` (TEXT)
  - `remarks` (TEXT)
  - `user` (TEXT)
  - `picture_url` (TEXT)
  - `is_pc_header` (BOOLEAN)
  - `is_highlighted` (BOOLEAN)
  - `created_at` (TIMESTAMP)
  - `updated_at` (TIMESTAMP)

## Features

- ✅ Auto-sync to Supabase
- ✅ Multi-sheet support
- ✅ Picture storage (base64)
- ✅ Highlight states
- ✅ PC section support

### Class Schedule tables

- **`class_schedule_sheets`** - Lab sheets (e.g. COMPUTER LABORATORY, MULTIMEDIA AND SPEECH LABORATORY)
  - `id` (BIGSERIAL, PRIMARY KEY)
  - `name` (TEXT)
  - `created_at`, `updated_at`

- **`class_schedule_entries`** - Schedule entries per sheet
  - `id` (BIGSERIAL, PRIMARY KEY)
  - `sheet_id` (BIGINT, FK to class_schedule_sheets, ON DELETE CASCADE)
  - `day`, `time_slot`, `type`, `instructor`, `course`, `code`
  - `created_at`, `updated_at`
