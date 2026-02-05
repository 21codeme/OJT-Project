# Lab Inventory Management System

A web-based inventory management system for Occidental Mindoro State College - Multimedia and Speech Laboratory. This application allows you to import Excel inventory files, view them in a user-friendly interface, and export them back to Excel with the same layout.

## Features

- 📁 **Import Excel Files** - Upload and read Excel (.xlsx, .xls) inventory files
- ➕ **Add Items** - Create new inventory items directly on the web interface
- 🖥️ **Add PC Sections** - Add new PC sections (like "PC USED BY: NAME" or "SERVER")
- ✏️ **Edit Items** - Click on any cell to edit inventory data inline
- 🗑️ **Delete Items** - Remove items with a single click
- 👁️ **View Inventory** - Display inventory data in a clean, organized table format
- 📥 **Export to Excel** - Export the inventory back to Excel with the same layout and structure
- 🎨 **Excel-like Layout** - Maintains the original Excel formatting and structure
- 📱 **Responsive Design** - Works on desktop and mobile devices

## How to Use

1. **Open the Application**
   - Simply open `index.html` in your web browser
   - No installation or server required!

2. **Import Excel File** (Optional)
   - Click the "📁 Import Excel File" button
   - Select your Excel inventory file (e.g., `INVENTORY-AS-OF-AUGUST-2026.xlsx`)
   - The data will automatically load and display in the table

3. **Add Items Manually**
   - Click "➕ Add Item" to add a new inventory row
   - Click "🖥️ Add PC Section" to add a new PC section (like "PC USED BY: NAME" or "SERVER")
   - Fill in the details by clicking on any cell and typing

4. **Edit Items**
   - Click on any cell in the table to edit it
   - Press Enter or click outside to save changes
   - All cells are editable

5. **Delete Items**
   - Click the "🗑️ Delete" button on any row to remove it
   - Confirm the deletion when prompted

6. **View Inventory**
   - Scroll through the inventory items
   - The table displays all columns: Article/It, Description, Old Property N Assigned, Unit of meas, Unit Value, Quantity per Physical count, Location/Whereabout, Condition, Remarks, and User
   - PC header rows are highlighted in green

7. **Export to Excel**
   - Click the "📥 Export to Excel" button
   - The file will be downloaded with all your data (imported and manually added)
   - The filename will include the current month and year
   - The exported file maintains the same layout as the original Excel format

8. **Clear Data**
   - Click "🗑️ Clear Data" to remove all loaded data and start fresh

## File Structure

```
lab inventory/
├── index.html              # Main HTML file
├── styles.css              # Styling and layout
├── script.js               # JavaScript functionality
├── config.js               # Supabase configuration
├── database/               # Database setup files
│   ├── supabase-setup.sql  # SQL script for database setup
│   └── SUPABASE_SETUP.md   # Database setup instructions
├── images/                 # Logo and image files
├── README.md               # This file
└── VERCEL_DEPLOY.md        # Vercel deployment guide
```

## Deployment

This application can be easily deployed to Vercel. See `VERCEL_DEPLOY.md` for detailed deployment instructions.

Quick steps:
1. Push code to GitHub
2. Import project in Vercel
3. Deploy (no build needed - it's a static site)

## Technical Details

- Uses **SheetJS (xlsx.js)** library for Excel file processing
- Pure HTML, CSS, and JavaScript - no backend required
- Works entirely in the browser
- Maintains original Excel structure and formatting

## Browser Compatibility

- Chrome (recommended)
- Firefox
- Edge
- Safari

## Notes

- The application preserves the original Excel layout and structure
- All data is processed locally in your browser - no data is sent to any server
- Large Excel files may take a few seconds to load

## Support

For issues or questions, please check that:
- Your Excel file is in .xlsx or .xls format
- The file is not corrupted
- Your browser supports modern JavaScript features
