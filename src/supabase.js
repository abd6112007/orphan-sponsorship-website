import { createClient } from '@supabase/supabase-js'

// هذا هو الرابط الخاص بمشروعك (تأكد من نسخه من صفحة API في نفس المكان الذي أخذت منه المفاتيح)
const SUPABASE_URL = 'https://jpurkbladftnchsqvxjl.supabase.co'

// هذا هو المفتاح الجديد (Publishable key)
const SUPABASE_ANON_KEY = 'sb_publishable_G9JMTgePnLq7f7M0qxvWGw_J_iH-wIi'

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY)