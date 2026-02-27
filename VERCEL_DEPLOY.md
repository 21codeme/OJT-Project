# Vercel Deployment Guide

## Step 1: Prepare Your Project

Your project is already ready for Vercel! It's a static site (HTML/CSS/JS) so no build configuration is needed.

## Step 2: Deploy to Vercel

### Option A: Deploy via Vercel Dashboard (Recommended)

1. **Go to Vercel**
   - Visit: https://vercel.com
   - Sign in with your GitHub account (same account as your repository)

2. **Import Your Project**
   - Click "Add New Project" or "Import Project"
   - Select "Import Git Repository"
   - Find and select `21codeme/OJT-Project`
   - Click "Import"

3. **Configure Project**
   - **Framework Preset**: Select "Other" or leave as default
   - **Root Directory**: Leave as `./` (root)
   - **Build Command**: Leave empty (no build needed)
   - **Output Directory**: Leave empty (serves from root)
   - **Install Command**: Leave empty (no dependencies)

4. **Environment Variables** (Optional)
   - You don't need to add environment variables since Supabase config is in `config.js`
   - But if you want to use environment variables instead, you can add:
     - `VITE_SUPABASE_URL` (if using Vite)
     - `VITE_SUPABASE_ANON_KEY` (if using Vite)

5. **Deploy**
   - Click "Deploy"
   - Wait for deployment to complete (usually 1-2 minutes)

6. **Access Your Site**
   - Vercel will provide a URL like: `https://ojt-project.vercel.app`
   - Your site is now live!

7. **Excel links (PC Location) — walang security warning**
   - Buksan ang app mula sa **Vercel URL** (hal. `https://your-project.vercel.app`)
   - Pumasok sa **Inventory** → mag-export to Excel doon
   - Ang links sa Excel ay magiging **https** (hindi `file://`), kaya kapag pinindot sa Excel **hindi na lalabas** ang "Microsoft Excel Security Notice"
   - Kung binuksan ang app mula sa file (Downloads/folder), ang links ay `file://` at may warning pa rin

### Option B: Deploy via Vercel CLI

1. **Install Vercel CLI**
   ```bash
   npm install -g vercel
   ```

2. **Login to Vercel**
   ```bash
   vercel login
   ```

3. **Deploy**
   ```bash
   cd "C:\Users\Lenovo\Downloads\lab inventory"
   vercel
   ```

4. **Follow the prompts**
   - Link to existing project or create new
   - Confirm settings
   - Deploy!

## Step 3: Custom Domain (Optional)

1. Go to your project in Vercel Dashboard
2. Click "Settings" > "Domains"
3. Add your custom domain
4. Follow DNS configuration instructions

## Step 4: Verify Deployment

1. Open your Vercel URL
2. Check browser console (F12) for:
   - "Supabase connected successfully" (if configured)
   - No errors

3. Test Features:
   - Import Excel file
   - Add items
   - Export to Excel
   - Check if data syncs to Supabase (if configured)

## Important Notes

### ✅ What Works Automatically:
- Static files (HTML, CSS, JS) are served
- Images in `images/` folder are accessible
- Excel import/export works (client-side)
- All features work as expected
- **Excel “PC Location” links**: kung nag-export ka mula sa Vercel URL (https://…), ang link sa cell ay https at hindi na magpapakita ng Excel security warning kapag pinindot

### ⚠️ Things to Remember:
- **Supabase Config**: Make sure `config.js` has your credentials
- **CORS**: Supabase should handle CORS automatically
- **File Uploads**: Excel file uploads work (client-side only)
- **Database**: Make sure you've run the SQL setup in Supabase

### 🔒 Security Note:
- Your `config.js` with Supabase keys will be public in the deployed site
- This is okay for the anon key (it's meant to be public)
- Never commit service_role keys to public repos

## Troubleshooting

### Issue: Site not loading
- Check Vercel deployment logs
- Verify all files are in the repository
- Check browser console for errors

### Issue: Supabase not connecting
- Verify `config.js` has correct credentials
- Check browser console for connection errors
- Make sure Supabase project is active

### Issue: Images not loading
- Verify `images/` folder is in the repository
- Check image paths in HTML/CSS
- Ensure images are committed to Git

## Continuous Deployment

Vercel automatically deploys when you push to GitHub:
- Push to `main` branch → Production deployment
- Create pull request → Preview deployment

## Need Help?

- Vercel Docs: https://vercel.com/docs
- Vercel Support: https://vercel.com/support
