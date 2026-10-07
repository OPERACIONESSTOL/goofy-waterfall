import { createClient } from "@supabase/supabase-js";

const supabaseUrl = "https://iuxkupyelfrqhbbylger.supabase.co";
const supabaseKey = "sb_publishable_LYeLeIv8ocXsx6_I1-mM7g_v_A5zIr1";

export const supabase = createClient(supabaseUrl, supabaseKey);
