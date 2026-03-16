/**
 * Sync pictures from Supabase Storage (bucket: inventory-pictures) to local folder
 * inventory-pictures/ so you can see uploaded images in your project.
 *
 * Setup: Create .env in project root with:
 *   SUPABASE_URL=https://xxxx.supabase.co
 *   SUPABASE_ANON_KEY=your-anon-key
 *
 * Run: npm run sync-pictures   (or: node scripts/sync-pictures-to-local.js)
 * Optional: Schedule (e.g. Task Scheduler on Windows) to run daily for "automatic" sync.
 */

const path = require('path');
const fs = require('fs');

// Load .env from project root (parent of scripts/)
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_ANON_KEY = process.env.SUPABASE_ANON_KEY;
const BUCKET = 'inventory-pictures';
const LOCAL_BASE = path.join(__dirname, '..', 'inventory-pictures');

if (!SUPABASE_URL || !SUPABASE_ANON_KEY) {
  console.error('Missing SUPABASE_URL or SUPABASE_ANON_KEY. Add them to .env in project root.');
  process.exit(1);
}

async function main() {
  const { createClient } = require('@supabase/supabase-js');
  const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

  let total = 0;
  let errors = 0;

  async function listAndDownload(prefix) {
    const { data: items, error } = await supabase.storage.from(BUCKET).list(prefix, { limit: 1000 });
    if (error) {
      console.error('List error at', prefix || '(root)', error.message);
      return;
    }
    if (!items || items.length === 0) return;

    for (const item of items) {
      const fullPath = prefix ? `${prefix}/${item.name}` : item.name;
      // Folders have no 'id'; files do
      if (item.id != null) {
        const localPath = path.join(LOCAL_BASE, fullPath);
        const dir = path.dirname(localPath);
        if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
        const { data: blob, error: downError } = await supabase.storage.from(BUCKET).download(fullPath);
        if (downError) {
          console.error('Download error', fullPath, downError.message);
          errors++;
          continue;
        }
        const buf = Buffer.from(await blob.arrayBuffer());
        fs.writeFileSync(localPath, buf);
        console.log('Saved', fullPath);
        total++;
      } else {
        await listAndDownload(fullPath);
      }
    }
  }

  if (!fs.existsSync(LOCAL_BASE)) fs.mkdirSync(LOCAL_BASE, { recursive: true });
  console.log('Syncing Supabase Storage →', LOCAL_BASE);
  await listAndDownload('');
  console.log('Done. Files saved:', total, errors ? `Errors: ${errors}` : '');
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
