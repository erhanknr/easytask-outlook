import { createClient } from '@supabase/supabase-js';

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL as string;
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY as string;

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: {
    // Use localStorage so the session survives task pane reloads.
    // Outlook re-renders the task pane on every open, so persistence matters.
    persistSession: true,
    storage: window.localStorage,
  },
});

export type SupabaseClient = typeof supabase;
