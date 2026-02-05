// Supabase Configuration
// Replace these with your actual Supabase project credentials
// You can find these in your Supabase project settings: Settings > API

const SUPABASE_CONFIG = {
    url: 'https://bferfkrkejwccvfsigze.supabase.co',
    anonKey: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJmZXJma3JrZWp3Y2N2ZnNpZ3plIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzAyNzM1NTUsImV4cCI6MjA4NTg0OTU1NX0.4nc1SgH-lXD4GvZ6XSbfzyCp-Swf6Mon-O3dA_mEpXE'
};

// Initialize Supabase client (use window object to avoid conflicts)
window.supabaseClient = null;

// Function to initialize Supabase (called after library loads)
function initSupabase() {
    // Check if Supabase credentials are configured
    if (SUPABASE_CONFIG.url && SUPABASE_CONFIG.url !== 'YOUR_SUPABASE_URL' && 
        SUPABASE_CONFIG.anonKey && SUPABASE_CONFIG.anonKey !== 'YOUR_SUPABASE_ANON_KEY') {
        try {
            // Supabase JS v2 exposes 'supabase' in global scope
            // Check if the library is loaded
            if (typeof window.supabase !== 'undefined' && window.supabase.createClient) {
                window.supabaseClient = window.supabase.createClient(SUPABASE_CONFIG.url, SUPABASE_CONFIG.anonKey);
                console.log('Supabase connected successfully');
                return true;
            } else {
                console.warn('Supabase library not loaded yet. Waiting...');
                // Try again after a delay
                setTimeout(initSupabase, 500);
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

// Function to check Supabase connection
function checkSupabaseConnection() {
    if (typeof window.supabaseClient !== 'undefined' && window.supabaseClient !== null) {
        return true;
    }
    return false;
}

// Create a getter function for easier access
function getSupabaseClient() {
    return window.supabaseClient;
}

// Try to initialize when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initSupabase);
} else {
    // DOM already loaded, try to initialize after a short delay
    setTimeout(initSupabase, 100);
}
