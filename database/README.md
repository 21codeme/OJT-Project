# Database Folder

This folder contains all database-related files for the Lab Inventory Management System.

## Files

- **`supabase-setup.sql`** - SQL script to set up the Supabase database tables, indexes, and security policies
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
   - Copy and paste contents of `supabase-setup.sql`
   - Run the script

4. **Verify Connection**
   - Open the application
   - Check browser console for "Supabase connected successfully"

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
