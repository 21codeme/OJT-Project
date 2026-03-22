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

5. **Storage bucket (so pictures show in the PC Location link)**
   - In Supabase Dashboard: **Storage** → **New bucket**
   - Name: `inventory-pictures`
   - Public bucket: **ON**
   - Create bucket.
   - If you see **"new row violates row-level security policy"**, run **`database/storage-policies.sql`** in the **SQL Editor** to allow uploads and reads on that bucket.
   - **Folder structure:** Inside the bucket, images are organized as **{PC Section} / {sheetId} / {rowIndex} / {timestamp}.jpg** (e.g. `PC 1/sheet-2/5/1234567890.jpg`). All pictures uploaded for a PC section (e.g. PC 1) share one folder; after Excel import, pictures still work because the URL is saved in the link.

6. **OJT Trainee documents (Edit Profile → Required Documents)**
   - In **Storage** → **New bucket**: name `ojt-trainee-documents`, **Public: ON**
   - Run **`database/ojt-documents-storage.sql`** in the SQL Editor (same pattern as `storage-policies.sql` for inventory)
   - Files are uploaded to path `{trainee_id}/{doc_type}{extension}`; metadata and public URL are in `ojt_trainee_documents.file_data` (JSON).

7. **OJT Daily Logs photos (Dashboard → Daily Logs)**
   - **Quick fix (404 on REST, 400 on Storage):** run **`database/setup-ojt-daily-logs-one-shot.sql`** in the SQL Editor — creates `ojt_daily_logs` table, `ojt-daily-logs` bucket, storage policies, and `NOTIFY` for schema cache.
   - `ojt_trainees` must already exist. If not, run **`database/ojt-tables.sql`** first.
   - **Manual (if you skip the one-shot):** Storage → New bucket `ojt-daily-logs` (Public ON) → run **`ojt-daily-logs-storage.sql`** → ensure `ojt_daily_logs` exists from **`ojt-tables.sql`**
   - Photo path: `{trainee_id}/{log_date}/{timestamp}.{ext}`

8. **Trainee forgot password (Attendance login)**
   - Run **`database/ojt-trainee-password-reset.sql`** in the SQL Editor (`ojt_trainees` must exist).
   - On the login page: **Forgot password?** → email → 6-digit code → new password.
   - **Email (optional):** In `config.js`, set `OJT_PASSWORD_RESET_EMAILJS`:
     - **publicKey** — EmailJS → Account → API Keys → Public Key (copy exactly; must match byte-for-byte).
     - **serviceId** — Email Services → your service (e.g. `service_...`).
     - **templateId** — Email Templates → template → **Settings** → **Template ID** (e.g. `template_05sxotw`). Do not use the URL slug if it differs from the Template ID.
   - In the EmailJS template: **To email** = `{{to_email}}`; in content: `{{reset_code}}`, and `{{email}}` if needed for Reply-To.
   - **Ready-to-paste HTML:** `database/emailjs-template-password-reset.html`.
   - Account → **Allowed domains:** add your production domain (e.g. `ojt-project-laboratory.vercel.app`) and `localhost`.
   - If EmailJS is missing or misconfigured, the code is shown on the page after requesting (**fallback**).

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

## Free tier: Supabase auto-pause

Free projects can **auto-pause** after about a week without **project** activity. That is controlled by Supabase, not by this app’s code.

- **Paid plan:** projects stay active without pings.
- **Stay on free:** add GitHub repository secrets `SUPABASE_URL` and `SUPABASE_ANON_KEY`, then rely on the scheduled workflow [`.github/workflows/keep-alive.yml`](../.github/workflows/keep-alive.yml) (or run it manually under **Actions**). It hits the REST API for `ojt_trainees` so Supabase sees activity. Opening the site alone may not be enough if no API calls reach Supabase.
