import React, { useState, useEffect, useMemo, useRef } from "react";
import { Plus, Trash2, ClipboardList, Users, Cog, BarChart3, Factory, X, Lock, Upload, Download, Search } from "lucide-react";
import { db } from "./firebase";
import { APP_PASSWORD } from "./config";
import {
  collection, onSnapshot, addDoc, deleteDoc, doc,
} from "firebase/firestore";
import * as XLSX from "xlsx";

const todayISO = () => new Date().toISOString().slice(0, 10);

const TABS = [
  { id: "kayit", label: "Kayıt Ekle", icon: ClipboardList },
  { id: "raporlar", label: "Raporlar", icon: BarChart3 },
  { id: "takimlar", label: "Takımlar", icon: Users },
  { id: "makineler", label: "Makineler", icon: Cog },
];

// ---------- Excel yardımcıları ----------
function excelDenIsimOku(dosya) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "binary" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        const baslikKelimeler = ["takım", "takim", "makine", "ad", "isim", "i̇sim", "name"];
        const isimler = rows
          .flat()
          .map((v) => (v === undefined || v === null ? "" : String(v).trim()))
          .filter((v) => v && !baslikKelimeler.includes(v.toLowerCase()));
        resolve([...new Set(isimler)]);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsBinaryString(dosya);
  });
}

function excelIndir(veri, dosyaAdi, sayfaAdi) {
  const ws = XLSX.utils.json_to_sheet(veri);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sayfaAdi);
  XLSX.writeFile(wb, dosyaAdi);
}

// ---------- Şifre Kapısı ----------
function GirisEkrani({ onGiris }) {
  const [val, setVal] = useState("");
  const [hata, setHata] = useState(false);

  const dene = () => {
    if (val === APP_PASSWORD) {
      localStorage.setItem("uretim_takip_auth", "1");
      onGiris();
    } else {
      setHata(true);
      setTimeout(() => setHata(false), 1600);
    }
  };

  return (
    <div style={{ minHeight: "100vh", background: "#14181c", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Inter', system-ui, sans-serif" }}>
      <div style={{ background: "#1b2127", border: "1px solid #2a3138", borderRadius: 12, padding: 32, width: 320 }}>
        <div style={{ width: 40, height: 40, borderRadius: 10, background: "#e8a33d", display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 16 }}>
          <Lock size={20} color="#14181c" />
        </div>
        <div style={{ color: "#e7e5e0", fontWeight: 700, fontSize: 16, marginBottom: 4 }}>Üretim Takip</div>
        <div style={{ color: "#8b929a", fontSize: 12.5, marginBottom: 18 }}>Devam etmek için ortak şifreyi girin</div>
        <input
          type="password"
          value={val}
          onChange={(e) => setVal(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && dene()}
          placeholder="Şifre"
          style={{ width: "100%", background: "#14181c", border: `1px solid ${hata ? "#c0392b" : "#2f3740"}`, borderRadius: 7, padding: "10px 12px", color: "#e7e5e0", fontSize: 14, outline: "none", marginBottom: 12 }}
        />
        <button onClick={dene} style={{ width: "100%", background: "#e8a33d", color: "#14181c", border: "none", borderRadius: 7, padding: "10px 0", fontWeight: 700, fontSize: 13.5, cursor: "pointer" }}>
          Giriş Yap
        </button>
        {hata && <div style={{ color: "#e07a6b", fontSize: 12, marginTop: 10 }}>Şifre yanlış, tekrar deneyin.</div>}
      </div>
    </div>
  );
}

export default function App() {
  const [authed, setAuthed] = useState(localStorage.getItem("uretim_takip_auth") === "1");
  if (!authed) return <GirisEkrani onGiris={() => setAuthed(true)} />;
  return <Panel onCikis={() => { localStorage.removeItem("uretim_takip_auth"); setAuthed(false); }} />;
}

function Panel({ onCikis }) {
  const [tab, setTab] = useState("kayit");
  const [teams, setTeams] = useState([]);
  const [machines, setMachines] = useState([]);
  const [records, setRecords] = useState([]);

  // Firestore canlı dinleme - herkes aynı anda güncel veriyi görür
  useEffect(() => {
    const unsub1 = onSnapshot(collection(db, "teams"), (snap) =>
      setTeams(snap.docs.map((d) => ({ id: d.id, ...d.data() })))
    );
    const unsub2 = onSnapshot(collection(db, "machines"), (snap) =>
      setMachines(snap.docs.map((d) => ({ id: d.id, ...d.data() })))
    );
    const unsub3 = onSnapshot(collection(db, "records"), (snap) =>
      setRecords(snap.docs.map((d) => ({ id: d.id, ...d.data() })))
    );
    return () => { unsub1(); unsub2(); unsub3(); };
  }, []);

  return (
    <div style={{ minHeight: "100vh", background: "#14181c", color: "#e7e5e0", fontFamily: "'Inter', system-ui, sans-serif" }}>
      <style>{`
        * { box-sizing: border-box; }
        input, select, button { font-family: inherit; }
        ::placeholder { color: #6b7178; }
        .card { background: #1b2127; border: 1px solid #2a3138; border-radius: 10px; }
        .field-label { font-size: 11px; text-transform: uppercase; letter-spacing: 0.08em; color: #8b929a; margin-bottom: 6px; display: block; font-weight: 600; }
        .input { width: 100%; background: #14181c; border: 1px solid #2f3740; border-radius: 7px; padding: 10px 12px; color: #e7e5e0; font-size: 14px; outline: none; transition: border-color .15s; }
        .input:focus { border-color: #e8a33d; }
        table { border-collapse: collapse; width: 100%; }
        th { text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.06em; color: #8b929a; padding: 10px 12px; border-bottom: 1px solid #2a3138; font-weight: 600; white-space: nowrap; }
        td { padding: 10px 12px; border-bottom: 1px solid #21272d; font-size: 13.5px; }
        tr:hover td { background: #1f262c; }
        .btn-ghost { display: flex; align-items: center; gap: 6px; background: transparent; border: 1px solid #2f3740; color: #c7cbd1; border-radius: 7px; padding: 8px 13px; font-size: 12.5px; font-weight: 600; cursor: pointer; white-space: nowrap; }
        .btn-ghost:hover { border-color: #e8a33d; color: #e8a33d; }
      `}</style>

      <header style={{ borderBottom: "1px solid #2a3138", padding: "18px 24px", display: "flex", alignItems: "center", gap: 12 }}>
        <div style={{ width: 34, height: 34, borderRadius: 8, background: "#e8a33d", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <Factory size={19} color="#14181c" />
        </div>
        <div style={{ flex: 1 }}>
          <div style={{ fontWeight: 700, fontSize: 16, letterSpacing: -0.2 }}>Üretim Takip</div>
          <div style={{ fontSize: 11.5, color: "#8b929a" }}>Kesim / Dikim Takımları · Günlük Çıkış Takibi</div>
        </div>
        <button onClick={onCikis} style={{ background: "none", border: "1px solid #2f3740", color: "#8b929a", borderRadius: 7, padding: "7px 12px", fontSize: 12.5, cursor: "pointer" }}>
          Çıkış Yap
        </button>
      </header>

      <nav style={{ display: "flex", gap: 4, padding: "14px 24px 0", borderBottom: "1px solid #2a3138", overflowX: "auto" }}>
        {TABS.map(({ id, label, icon: Icon }) => (
          <button
            key={id}
            onClick={() => setTab(id)}
            style={{
              display: "flex", alignItems: "center", gap: 7, padding: "9px 14px",
              background: "transparent", border: "none", cursor: "pointer",
              color: tab === id ? "#e8a33d" : "#8b929a",
              borderBottom: tab === id ? "2px solid #e8a33d" : "2px solid transparent",
              fontSize: 13.5, fontWeight: 600, whiteSpace: "nowrap",
            }}
          >
            <Icon size={15} /> {label}
          </button>
        ))}
      </nav>

      <main style={{ padding: 24, maxWidth: 1100, margin: "0 auto" }}>
        {tab === "kayit" && <KayitEkle teams={teams} machines={machines} records={records} />}
        {tab === "raporlar" && <Raporlar teams={teams} machines={machines} records={records} />}
        {tab === "takimlar" && <ListeYonetimi title="Takım" baslikCogul="Takımlar" koleksiyon="teams" placeholder="Örn: Kesim Takım 1" items={teams} icon={Users} />}
        {tab === "makineler" && <ListeYonetimi title="Makine" baslikCogul="Makineler" koleksiyon="machines" placeholder="Makine listesini buradan ekleyin" items={machines} icon={Cog} />}
      </main>
    </div>
  );
}

// ---------- Kayıt Ekle ----------
function KayitEkle({ teams, machines, records }) {
  const [form, setForm] = useState({ tarih: todayISO(), takim: "", magaza: "", makine: "", urun: "", adet: "" });
  const [msg, setMsg] = useState("");
  const [arama, setArama] = useState("");
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  const submit = async () => {
    if (!form.takim || !form.makine || !form.adet) {
      setMsg("Takım, makine ve adet zorunlu.");
      setTimeout(() => setMsg(""), 2500);
      return;
    }
    await addDoc(collection(db, "records"), { ...form, adet: Number(form.adet), olusturma: Date.now() });
    setForm((f) => ({ tarih: f.tarih, takim: "", magaza: "", makine: "", urun: "", adet: "" }));
    setMsg("Kayıt eklendi.");
    setTimeout(() => setMsg(""), 1800);
  };

  const sil = async (id) => { await deleteDoc(doc(db, "records", id)); };

  const listelenecek = useMemo(() => {
    const sirali = [...records].sort((a, b) => (b.olusturma || 0) - (a.olusturma || 0));
    if (!arama.trim()) return sirali.slice(0, 10);
    const q = arama.trim().toLowerCase();
    return sirali.filter((r) =>
      (r.takim || "").toLowerCase().includes(q) ||
      (r.magaza || "").toLowerCase().includes(q) ||
      (r.makine || "").toLowerCase().includes(q) ||
      (r.urun || "").toLowerCase().includes(q) ||
      (r.tarih || "").includes(q)
    ).slice(0, 50);
  }, [records, arama]);

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 16 }}>Günlük Çıkış Kaydı</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
          <div><label className="field-label">Tarih</label><input className="input" type="date" value={form.tarih} onChange={set("tarih")} /></div>
          <div>
            <label className="field-label">Takım</label>
            <select className="input" value={form.takim} onChange={set("takim")}>
              <option value="">Seçiniz</option>
              {teams.map((t) => <option key={t.id} value={t.name}>{t.name}</option>)}
            </select>
          </div>
          <div><label className="field-label">Mağaza</label><input className="input" placeholder="Mağaza / müşteri adı" value={form.magaza} onChange={set("magaza")} /></div>
          <div>
            <label className="field-label">Makine</label>
            <select className="input" value={form.makine} onChange={set("makine")}>
              <option value="">Seçiniz</option>
              {machines.map((m) => <option key={m.id} value={m.name}>{m.name}</option>)}
            </select>
          </div>
          <div><label className="field-label">Ürün / Model</label><input className="input" placeholder="Opsiyonel" value={form.urun} onChange={set("urun")} /></div>
          <div><label className="field-label">Adet</label><input className="input" type="number" min="0" placeholder="0" value={form.adet} onChange={set("adet")} /></div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 16 }}>
          <button onClick={submit} style={{ display: "flex", alignItems: "center", gap: 7, background: "#e8a33d", color: "#14181c", border: "none", borderRadius: 7, padding: "10px 16px", fontWeight: 700, fontSize: 13.5, cursor: "pointer" }}>
            <Plus size={16} /> Kaydı Ekle
          </button>
          {msg && <span style={{ fontSize: 12.5, color: "#8b929a" }}>{msg}</span>}
        </div>
        {(teams.length === 0 || machines.length === 0) && (
          <div style={{ marginTop: 14, fontSize: 12.5, color: "#c98a2e", background: "#2a2115", border: "1px solid #4a3a1a", borderRadius: 7, padding: "9px 12px" }}>
            {teams.length === 0 && machines.length === 0 ? "Önce Takımlar ve Makineler sekmelerinden liste oluşturun."
              : teams.length === 0 ? "Önce Takımlar sekmesinden takım ekleyin."
              : "Önce Makineler sekmesinden makine listesini ekleyin."}
          </div>
        )}
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a3138", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
          <div style={{ fontWeight: 700, fontSize: 14 }}>{arama ? "Arama Sonuçları" : "Son Kayıtlar"}</div>
          <div style={{ position: "relative", minWidth: 220 }}>
            <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
            <input
              className="input"
              style={{ paddingLeft: 30 }}
              placeholder="Takım, mağaza, makine, tarih ara…"
              value={arama}
              onChange={(e) => setArama(e.target.value)}
            />
          </div>
        </div>
        <div style={{ overflowX: "auto" }}>
          <table>
            <thead><tr><th>Tarih</th><th>Takım</th><th>Mağaza</th><th>Makine</th><th>Ürün</th><th>Adet</th><th></th></tr></thead>
            <tbody>
              {listelenecek.length === 0 && <tr><td colSpan={7} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>{arama ? "Sonuç bulunamadı." : "Henüz kayıt yok."}</td></tr>}
              {listelenecek.map((r) => (
                <tr key={r.id}>
                  <td style={{ fontFamily: "monospace" }}>{r.tarih}</td>
                  <td>{r.takim}</td>
                  <td>{r.magaza || "—"}</td>
                  <td>{r.makine}</td>
                  <td>{r.urun || "—"}</td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: "#e8a33d" }}>{r.adet}</td>
                  <td><button onClick={() => sil(r.id)} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", padding: 4 }}><Trash2 size={14} /></button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ---------- Raporlar ----------
function Raporlar({ teams, machines, records }) {
  const [f, setF] = useState({ start: "", end: "", takim: "", makine: "", magaza: "", arama: "" });
  const set = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));

  const filtered = useMemo(() => {
    const q = f.arama.trim().toLowerCase();
    return records.filter((r) => {
      if (f.start && r.tarih < f.start) return false;
      if (f.end && r.tarih > f.end) return false;
      if (f.takim && r.takim !== f.takim) return false;
      if (f.makine && r.makine !== f.makine) return false;
      if (f.magaza && !(r.magaza || "").toLowerCase().includes(f.magaza.toLowerCase())) return false;
      if (q && !(
        (r.takim || "").toLowerCase().includes(q) ||
        (r.magaza || "").toLowerCase().includes(q) ||
        (r.makine || "").toLowerCase().includes(q) ||
        (r.urun || "").toLowerCase().includes(q)
      )) return false;
      return true;
    }).sort((a, b) => (a.tarih < b.tarih ? 1 : -1));
  }, [records, f]);

  const sil = async (id) => { await deleteDoc(doc(db, "records", id)); };
  const toplam = filtered.reduce((s, r) => s + (Number(r.adet) || 0), 0);
  const magazalar = [...new Set(records.map((r) => r.magaza).filter(Boolean))];

  const disaAktar = () => {
    excelIndir(
      filtered.map((r) => ({ Tarih: r.tarih, Takım: r.takim, Mağaza: r.magaza || "", Makine: r.makine, Ürün: r.urun || "", Adet: r.adet })),
      "uretim-raporu.xlsx",
      "Rapor"
    );
  };

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15 }}>Filtrele</div>
          <button className="btn-ghost" onClick={disaAktar}><Download size={14} /> Excel'e Aktar</button>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 14 }}>
          <div style={{ gridColumn: "1 / -1" }}>
            <label className="field-label">Serbest Arama</label>
            <div style={{ position: "relative" }}>
              <Search size={14} color="#6b7178" style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
              <input className="input" style={{ paddingLeft: 30 }} placeholder="Takım, mağaza, makine, ürün ara…" value={f.arama} onChange={set("arama")} />
            </div>
          </div>
          <div><label className="field-label">Başlangıç</label><input className="input" type="date" value={f.start} onChange={set("start")} /></div>
          <div><label className="field-label">Bitiş</label><input className="input" type="date" value={f.end} onChange={set("end")} /></div>
          <div>
            <label className="field-label">Takım</label>
            <select className="input" value={f.takim} onChange={set("takim")}>
              <option value="">Tümü</option>
              {teams.map((t) => <option key={t.id} value={t.name}>{t.name}</option>)}
            </select>
          </div>
          <div>
            <label className="field-label">Makine</label>
            <select className="input" value={f.makine} onChange={set("makine")}>
              <option value="">Tümü</option>
              {machines.map((m) => <option key={m.id} value={m.name}>{m.name}</option>)}
            </select>
          </div>
          <div>
            <label className="field-label">Mağaza</label>
            <input className="input" list="magaza-list" placeholder="Ara / seç" value={f.magaza} onChange={set("magaza")} />
            <datalist id="magaza-list">{magazalar.map((m) => <option key={m} value={m} />)}</datalist>
          </div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
        <Stat label="Kayıt Sayısı" value={filtered.length} />
        <Stat label="Toplam Çıkış (Adet)" value={toplam.toLocaleString("tr-TR")} highlight />
        <Stat label="Aktif Mağaza" value={new Set(filtered.map((r) => r.magaza).filter(Boolean)).size} />
        <Stat label="Aktif Makine" value={new Set(filtered.map((r) => r.makine)).size} />
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #2a3138", fontWeight: 700, fontSize: 14 }}>Sonuçlar ({filtered.length})</div>
        <div style={{ overflowX: "auto", maxHeight: 480, overflowY: "auto" }}>
          <table>
            <thead><tr><th>Tarih</th><th>Takım</th><th>Mağaza</th><th>Makine</th><th>Ürün</th><th>Adet</th><th></th></tr></thead>
            <tbody>
              {filtered.length === 0 && <tr><td colSpan={7} style={{ color: "#6b7178", textAlign: "center", padding: 24 }}>Sonuç bulunamadı.</td></tr>}
              {filtered.map((r) => (
                <tr key={r.id}>
                  <td style={{ fontFamily: "monospace" }}>{r.tarih}</td>
                  <td>{r.takim}</td>
                  <td>{r.magaza || "—"}</td>
                  <td>{r.makine}</td>
                  <td>{r.urun || "—"}</td>
                  <td style={{ fontFamily: "monospace", fontWeight: 700, color: "#e8a33d" }}>{r.adet}</td>
                  <td><button onClick={() => sil(r.id)} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", padding: 4 }}><Trash2 size={14} /></button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function Stat({ label, value, highlight }) {
  return (
    <div className="card" style={{ padding: "16px 18px" }}>
      <div style={{ fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em", color: "#8b929a", fontWeight: 600, marginBottom: 8 }}>{label}</div>
      <div style={{ fontSize: 26, fontWeight: 800, fontFamily: "monospace", color: highlight ? "#e8a33d" : "#e7e5e0" }}>{value}</div>
    </div>
  );
}

// ---------- Takım / Makine Listesi Yönetimi ----------
function ListeYonetimi({ title, baslikCogul, koleksiyon, placeholder, items, icon: Icon }) {
  const [val, setVal] = useState("");
  const [arama, setArama] = useState("");
  const [iceAktariliyor, setIceAktariliyor] = useState(false);
  const [mesaj, setMesaj] = useState("");
  const dosyaRef = useRef(null);

  const add = async () => {
    const name = val.trim();
    if (!name) return;
    if (items.some((i) => i.name.toLowerCase() === name.toLowerCase())) { setVal(""); return; }
    await addDoc(collection(db, koleksiyon), { name });
    setVal("");
  };
  const sil = async (id) => { await deleteDoc(doc(db, koleksiyon, id)); };

  const iceAktar = async (e) => {
    const dosya = e.target.files[0];
    if (!dosya) return;
    setIceAktariliyor(true);
    setMesaj("");
    try {
      const isimler = await excelDenIsimOku(dosya);
      const mevcut = new Set(items.map((i) => i.name.toLowerCase()));
      const yeniler = isimler.filter((n) => !mevcut.has(n.toLowerCase()));
      for (const name of yeniler) {
        await addDoc(collection(db, koleksiyon), { name });
      }
      setMesaj(`${yeniler.length} yeni ${title.toLowerCase()} eklendi${isimler.length - yeniler.length > 0 ? `, ${isimler.length - yeniler.length} zaten vardı` : ""}.`);
    } catch (err) {
      setMesaj("Dosya okunamadı. .xlsx veya .csv dosyası olduğundan emin olun.");
    }
    setIceAktariliyor(false);
    e.target.value = "";
    setTimeout(() => setMesaj(""), 4000);
  };

  const disaAktar = () => {
    excelIndir(items.map((i) => ({ [title]: i.name })), `${baslikCogul.toLowerCase()}.xlsx`, baslikCogul);
  };

  const filtrelenmis = useMemo(() => {
    if (!arama.trim()) return items;
    const q = arama.trim().toLowerCase();
    return items.filter((i) => i.name.toLowerCase().includes(q));
  }, [items, arama]);

  return (
    <div style={{ display: "grid", gap: 20 }}>
      <div className="card" style={{ padding: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontWeight: 700, fontSize: 15, display: "flex", alignItems: "center", gap: 8 }}>
            <Icon size={17} color="#e8a33d" /> {baslikCogul} Listesi
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <input ref={dosyaRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={iceAktar} />
            <button className="btn-ghost" onClick={() => dosyaRef.current?.click()} disabled={iceAktariliyor}>
              <Upload size={14} /> {iceAktariliyor ? "Aktarılıyor…" : "Excel'den İçe Aktar"}
            </button>
            <button className="btn-ghost" onClick={disaAktar}><Download size={14} /> Excel'e Aktar</button>
          </div>
        </div>

        <div style={{ display: "flex", gap: 10 }}>
          <input
            className="input"
            placeholder={placeholder}
            value={val}
            onChange={(e) => setVal(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && add()}
          />
          <button onClick={add} style={{ display: "flex", alignItems: "center", gap: 6, background: "#e8a33d", color: "#14181c", border: "none", borderRadius: 7, padding: "0 16px", fontWeight: 700, fontSize: 13.5, cursor: "pointer", whiteSpace: "nowrap" }}>
            <Plus size={15} /> Ekle
          </button>
        </div>
        <div style={{ fontSize: 12, color: mesaj ? "#e8a33d" : "#6b7178", marginTop: 10 }}>
          {mesaj || `Toplam ${items.length} kayıt. Excel dosyasında isimler tek sütunda alt alta olmalı (başlık satırı olabilir).`}
        </div>
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "12px 16px", borderBottom: "1px solid #2a3138", position: "relative" }}>
          <Search size={14} color="#6b7178" style={{ position: "absolute", left: 26, top: "50%", transform: "translateY(-50%)" }} />
          <input
            className="input"
            style={{ paddingLeft: 30 }}
            placeholder={`${baslikCogul} içinde ara…`}
            value={arama}
            onChange={(e) => setArama(e.target.value)}
          />
        </div>
        {filtrelenmis.length === 0 ? (
          <div style={{ color: "#6b7178", textAlign: "center", padding: 32, fontSize: 13.5 }}>
            {arama ? "Sonuç bulunamadı." : `Henüz ${baslikCogul.toLowerCase()} eklenmedi.`}
          </div>
        ) : (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))" }}>
            {filtrelenmis.map((i, idx) => (
              <div key={i.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 16px", borderBottom: "1px solid #21272d", borderRight: "1px solid #21272d" }}>
                <span style={{ fontSize: 13.5, display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ fontFamily: "monospace", color: "#8b929a", fontSize: 11 }}>{String(idx + 1).padStart(2, "0")}</span>
                  {i.name}
                </span>
                <button onClick={() => sil(i.id)} style={{ background: "none", border: "none", color: "#6b7178", cursor: "pointer", padding: 4 }}><X size={14} /></button>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
