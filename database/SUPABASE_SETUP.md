# Supabase Setup Instructions

## Step 1: Get Your Supabase Credentials

1. Go to your Supabase project dashboard: https://supabase.com/dashboard
2. Select your project (OJT Project)
3. Go to **Settings** > **API**
4. Copy the following:
   - **Project URL** (e.g., `https://xxxxx.supabase.co`)
   - **anon/public key** (under "Project API keys")

## Step 2: Update config.js

Open `config.js` and replace the placeholder values:

```javascript
const SUPABASE_CONFIG = {
    url: 'https://your-project.supabase.co', // Your Project URL
    anonKey: 'your-anon-key-here' // Your anon/public key
};
```

## Step 3: Run SQL Setup

1. Go to your Supabase project dashboard
2. Click on **SQL Editor** in the left sidebar
3. Click **New Query**
4. Copy and paste the contents of `database/supabase-setup.sql` (or open the file from this folder)
5. Click **Run** to execute the SQL

This will create:
- `inventory_items` table - stores all inventory data
- `sheets` table - stores sheet information
- Indexes for better performance
- Row Level Security policies
- Auto-update triggers

## Step 4: Verify Connection

1. Open the application in your browser
2. Open browser console (F12)
3. You should see: "Supabase connected successfully"
4. If you see "Supabase not configured", check your `config.js` file

## Step 5: Test the Integration

1. Add some inventory items
2. Check your Supabase dashboard > **Table Editor** > `inventory_items`
3. You should see the data being saved automatically

## Features

- ✅ **Auto-sync**: All changes are automatically saved to Supabase
- ✅ **Multi-sheet support**: Each sheet is stored separately
- ✅ **Picture storage**: Images are stored as base64 in the database
- ✅ **Highlight states**: Row highlights are preserved
- ✅ **PC sections**: PC header rows are properly stored

## Security Note

The current setup uses public access. For production, you may want to:
1. Enable Row Level Security (RLS) with proper policies
2. Add authentication if needed
3. Restrict access based on user roles

## Troubleshooting

- **"Supabase not configured"**: Check that you've updated `config.js` with your credentials
- **"Error loading from Supabase"**: Check that you've run the SQL setup script
- **Data not saving**: Check browser console for errors
- **Connection issues**: Verify your Supabase project is active
