import { createClient } from '@supabase/supabase-js';

// Public project URL + anon/publishable key. Safe to ship in the client:
// Row Level Security (owner = auth.uid()) is what actually protects the data.
// Override via Vite env vars (VITE_SUPABASE_URL / VITE_SUPABASE_KEY) if needed.
const url = import.meta.env.VITE_SUPABASE_URL || 'https://kxefekyuukmoyjcfoqhe.supabase.co';
const key = import.meta.env.VITE_SUPABASE_KEY || 'sb_publishable_zjALXihtEoD3i1WDGL8MFQ_yo9dSnwo';

export const supabase = createClient(url, key, {
  auth: {
    persistSession: true,
    autoRefreshToken: true,
    detectSessionInUrl: true, // picks up the magic-link token on redirect
  },
});
