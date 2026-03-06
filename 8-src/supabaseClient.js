// 8-src/supabaseClient.js
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const supabaseUrl = window.APP_CONFIG.SUPABASE_URL;
const supabaseAnonKey = window.APP_CONFIG.SUPABASE_ANON_KEY;

export const supabase = createClient(supabaseUrl, supabaseAnonKey);