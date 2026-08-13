import { initializeApp, deleteApp, getApps, getApp } from "firebase/app";
import { getFirestore, enableIndexedDbPersistence } from "firebase/firestore";
import { getAuth, createUserWithEmailAndPassword, signOut } from "firebase/auth";

const firebaseConfig = {
  apiKey: "AIzaSyA9pvMee-kkVFIUXE4islcwlkdkNwFcEq4",
  authDomain: "stok-takip-f84a2.firebaseapp.com",
  projectId: "stok-takip-f84a2",
  storageBucket: "stok-takip-f84a2.firebasestorage.app",
  messagingSenderId: "572028161303",
  appId: "1:572028161303:web:dd9a75fac084a1ee65bb52",
  measurementId: "G-N4HS6D87BC",
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const auth = getAuth(app);

// Yazma işlemleri tarayıcıda kalıcı olarak (IndexedDB) kuyruklanır; bağlantı
// kesilse veya sayfa yenilense bile bekleyen yazılar kaybolmaz, bağlantı
// geri gelince otomatik gönderilir. Birden fazla sekme açıksa bu devre dışı
// kalabilir, sorun değil - o zaman normal (bellek içi) modda çalışır.
enableIndexedDbPersistence(db).catch(() => {});

// Yönetici, oturumu bozulmadan (kendi girişinden çıkmadan) yeni kullanıcı
// oluşturabilsin diye geçici, ikincil bir Firebase bağlantısı açıp orada
// hesabı oluşturuyoruz, sonra hemen kapatıyoruz.
export async function digerKullaniciOlustur(email, sifre) {
  const ikincilApp = initializeApp(firebaseConfig, "ikincil-" + Date.now());
  const ikincilAuth = getAuth(ikincilApp);
  try {
    const sonuc = await createUserWithEmailAndPassword(ikincilAuth, email, sifre);
    await signOut(ikincilAuth);
    return sonuc.user;
  } finally {
    await deleteApp(ikincilApp);
  }
}

// Eski SAKLAZ-METALERP programının Firebase projesi. Bu sadece bir kerelik
// veri taşıma (migrasyon) için kullanılıyor - eski projedeki talep/malzeme
// verilerini okuyup bu programın kendi veritabanına kopyalamak için.
const ESKI_METALERP_CONFIG = {
  apiKey: "AIzaSyDOytXQM1rXaL7BEqdpptfkdq0lRU13mkE",
  authDomain: "metalerp-a86bf.firebaseapp.com",
  projectId: "metalerp-a86bf",
  storageBucket: "metalerp-a86bf.firebasestorage.app",
  messagingSenderId: "277080695066",
  appId: "1:277080695066:web:4f118819b36300ccd40ca3",
};

export function eskiMetalErpDb() {
  const adi = "eski-metalerp";
  const mevcutApp = getApps().find((a) => a.name === adi);
  const eskiApp = mevcutApp || initializeApp(ESKI_METALERP_CONFIG, adi);
  return getFirestore(eskiApp);
}
