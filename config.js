// ============================================================
//  إعدادات التطبيق — عدّل الملف ده بس
// ============================================================

// 1) إعدادات Firebase: انسخها من
//    Firebase Console > Project settings > Your apps > Web app > firebaseConfig
export const firebaseConfig = {
  apiKey: "AIzaSyCuSPvtDK3733_Gb41WZN53KJdUxQ1RdUg",
  authDomain: "trucks-5a49d.firebaseapp.com",
  projectId: "trucks-5a49d",
  storageBucket: "trucks-5a49d.firebasestorage.app",
  messagingSenderId: "223915690540",
  appId: "1:223915690540:web:f942f0cdcde555520bdf89"
};

// 2) المستخدمين. اسم الدخول بيتحوّل لإيميل شكله  zizo@trucks.app
//    لازم تعمل نفس الإيميلات في Firebase > Authentication > Users
//    ولو غيّرت الأسماء هنا، غيّرها كمان في ملف firestore.rules
export const EMAIL_DOMAIN = "trucks.app";

export const USERS = {
  zizo:  { name: "Zizo",  role: "admin" },
  ussef: { name: "Ussef", role: "admin" },
  user:  { name: "User",  role: "admin" }
};
