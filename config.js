// Supabase Configuration
// Replace these with your actual Supabase project credentials
// You can find these in your Supabase project settings: Settings > API

const SUPABASE_CONFIG = {
    url: 'YOUR_SUPABASE_URL', // e.g., 'https://xxxxx.supabase.co'
    anonKey: 'YOUR_SUPABASE_ANON_KEY' // Your Supabase anon/public key
};

// Initialize Supabase client
let supabase = null;

// Function to initialize Supabase (called after library loads)
function initSupabase() {
    // Check if Supabase credentials are configured
    if (SUPABASE_CONFIG.url && SUPABASE_CONFIG.url !== 'YOUR_SUPABASE_URL' && 
        SUPABASE_CONFIG.anonKey && SUPABASE_CONFIG.anonKey !== 'YOUR_SUPABASE_ANON_KEY') {
        try {
            if (typeof supabaseClient !== 'undefined') {
                supabase = supabaseClient.createClient(SUPABASE_CONFIG.url, SUPABASE_CONFIG.anonKey);
                console.log('Supabase connected successfully');
                return true;
            } else {
                console.warn('Supabase library not loaded yet');
                return false;
            }
        } catch (error) {
            console.error('Error initializing Supabase:', error);
            return false;
        }
    } else {
        console.warn('Supabase not configured. Please update config.js with your credentials.');
        return false;
    }
}

// Try to initialize when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initSupabase);
} else {
    // DOM already loaded, try to initialize after a short delay
    setTimeout(initSupabase, 100);
}
