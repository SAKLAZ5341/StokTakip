import React, { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { FileDown, FileUp, Plus, ArrowRightCircle } from "lucide-react";
import { SIPARIS_SABLON_BASLIKLAR, SIPARIS_SABLON_ORNEK, siparisNoUret, guvenliCagir } from "./yardimcilar";

function durumRozetStili(durum, s) {
  if (durum === "Planlamada") return { background: "rgba(34,197,94,0.15)", color: s.renk?.iyi || "#22c55e" };
  if (durum === "Projelendirildi") return { background: "rgba(45,212,191,0.15)", color: s.renk?.vurgu };
  return { background: "rgba(250,204,21,0.15)", color: s.renk?.uyari };
}

// Açık siparişler tablosu: Excel şablon indirme, Excel'den toplu içe aktarma,
// manuel sipariş girişi ve "Projelendir" (WBS'e aktarma) tetikleyicisi.
export default function SiparisPaneli({ api, ui, veri, yazabilir, siparisler, seciliSiparisId, onSec, onProjelendir }) {
  const [yeniAcik, setYeniAcik] = useState(false);
  const [yukleniyor, setYukleniyor] = useState(false);
  const dosyaRef = useRef(null);
  const s = ui.stil || {};

  const siraliListe = useMemo(
    () => [...(siparisler || [])].sort((a, b) => String(b.olusturma || "").localeCompare(String(a.olusturma || ""))),
    [siparisler]
  );

  function sablonIndir() {
    ui.sablonIndir(SIPARIS_SABLON_BASLIKLAR, SIPARIS_SABLON_ORNEK, "siparis_sablonu", "Siparişler");
  }

  function satirAl(r, anahtarlar) {
    for (const k of anahtarlar) {
      if (r[k] !== undefined && r[k] !== null && String(r[k]).trim() !== "") return r[k];
    }
    return "";
  }

  async function dosyaSecildi(e) {
    const dosya = e.target.files && e.target.files[0];
    if (!dosya) return;
    setYukleniyor(true);
    try {
      const buffer = await dosya.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array" });
      const sayfa = wb.Sheets[wb.SheetNames[0]];
      const satirlar = XLSX.utils.sheet_to_json(sayfa, { defval: "" });

      const islemler = satirlar
        .filter((r) => String(satirAl(r, ["Parca No", "Parça No"])).trim() !== "")
        .map((r) => ({
          tur: "ekle",
          koleksiyon: "planlama_siparisler",
          veri: {
            musteriAdi: String(satirAl(r, ["Musteri Adi", "Müşteri Adı"])).trim(),
            projeKodu: String(satirAl(r, ["Proje Kodu"])).trim(),
            kalemNo: String(satirAl(r, ["Kalem No"])).trim(),
            parcaNo: String(satirAl(r, ["Parca No", "Parça No"])).trim(),
            parcaAdi: String(satirAl(r, ["Parca Adi", "Parça Adı"])).trim(),
            revizyon: String(satirAl(r, ["Revizyon"]) || "00").trim(),
            miktar: Number(satirAl(r, ["Miktar"]) || 0),
            teslimTarihi: String(satirAl(r, ["Teslim Tarihi"])).trim(),
            sfc: String(satirAl(r, ["SFC"])).toUpperCase().startsWith("E"),
            tightestTolerance: String(satirAl(r, ["Dar Tolerans"])).trim(),
            requiredMachineClass: String(satirAl(r, ["Gerekli Makine Sinifi", "Gerekli Makine Sınıfı"])).trim(),
            durum: "Açık",
          },
        }));

      if (islemler.length === 0) {
        alert("Dosyada geçerli satır bulunamadı. Lütfen şablon başlıklarını kontrol edin.");
      } else {
        await api.topluYaz(islemler);
      }
    } catch (err) {
      alert("İçe aktarma sırasında hata oluştu: " + (err && err.message ? err.message : String(err)));
    } finally {
      setYukleniyor(false);
      if (dosyaRef.current) dosyaRef.current.value = "";
    }
  }

  return (
    <div className="card" style={{ padding: 16, marginBottom: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
        <h3 style={{ margin: 0, color: s.renk?.yazi }}>Açık Siparişler</h3>
        {yazabilir && (
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button className="btn-ghost" style={s.dugme} onClick={sablonIndir}>
              <FileDown size={16} style={{ marginRight: 6 }} /> Excel Şablon İndir
            </button>
            <button
              className="btn-ghost"
              style={s.dugme}
              onClick={() => dosyaRef.current && dosyaRef.current.click()}
              disabled={yukleniyor}
            >
              <FileUp size={16} style={{ marginRight: 6 }} /> {yukleniyor ? "Aktarılıyor..." : "Excel'den İçe Aktar"}
            </button>
            <input ref={dosyaRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={dosyaSecildi} />
            <button className="btn-ghost" style={s.anaDugme} onClick={() => setYeniAcik(true)}>
              <Plus size={16} style={{ marginRight: 6 }} /> Yeni Sipariş
            </button>
          </div>
        )}
      </div>

      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              {["Müşteri", "Sipariş No", "Parça No", "Parça Adı", "Rev.", "Miktar", "Teslim Tarihi", "Durum", ""].map((h) => (
                <th key={h} style={s.tabloBaslik}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {siraliListe.map((sip) => (
              <tr
                key={sip.id}
                onClick={() => onSec(sip.id)}
                style={{ cursor: "pointer", background: sip.id === seciliSiparisId ? "rgba(45,212,191,0.12)" : "transparent" }}
              >
                <td style={s.tabloHucre}>{sip.musteriAdi}</td>
                <td style={s.tabloHucre}>{siparisNoUret(sip)}</td>
                <td style={s.tabloHucre}>{sip.parcaNo}</td>
                <td style={s.tabloHucre}>{sip.parcaAdi}</td>
                <td style={s.tabloHucre}>{sip.revizyon || "00"}</td>
                <td style={s.tabloHucre}>{sip.miktar}</td>
                <td style={s.tabloHucre}>{sip.teslimTarihi ? ui.tarih.tr(sip.teslimTarihi) : "-"}</td>
                <td style={s.tabloHucre}>
                  <span className="pill" style={durumRozetStili(sip.durum, s)}>
                    {sip.durum || "Açık"}
                  </span>
                </td>
                <td style={s.tabloHucre}>
                  {yazabilir && (sip.durum || "Açık") === "Açık" && (
                    <button
                      className="btn-ghost"
                      style={s.dugme}
                      onClick={(e) => { e.stopPropagation(); onProjelendir(sip); }}
                      title="Bu siparişi projeye / WBS'e aktar"
                    >
                      <ArrowRightCircle size={16} style={{ marginRight: 4 }} /> Projelendir
                    </button>
                  )}
                </td>
              </tr>
            ))}
            {siraliListe.length === 0 && (
              <tr>
                <td colSpan={9} style={{ ...s.tabloHucre, textAlign: "center", color: s.renk?.soluk }}>
                  Kayıtlı sipariş yok. Excel şablonunu indirip toplu girebilir ya da "Yeni Sipariş" ile tek tek ekleyebilirsiniz.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {yeniAcik && <YeniSiparisFormu api={api} ui={ui} veri={veri} kapat={() => setYeniAcik(false)} />}
    </div>
  );
}

function YeniSiparisFormu({ api, ui, veri, kapat }) {
  const [form, setForm] = useState({
    musteriAdi: "", projeKodu: "", kalemNo: "", parcaNo: "", parcaAdi: "",
    revizyon: "00", miktar: 1, teslimTarihi: ui.tarih.bugun(), sfc: false,
    tightestTolerance: "", requiredMachineClass: "",
  });
  const [cariListesiAcik, setCariListesiAcik] = useState(false);
  const s = ui.stil || {};

  function alan(k, v) { setForm((f) => ({ ...f, [k]: v })); }

  async function kaydet() {
    if (!form.musteriAdi || !form.parcaNo || !form.miktar) {
      alert("Müşteri, Parça No ve Miktar zorunludur.");
      return;
    }
    const tamam = await guvenliCagir(
      () => api.ekle("planlama_siparisler", { ...form, miktar: Number(form.miktar), durum: "Açık" }),
      "Sipariş kaydedilemedi"
    );
    if (tamam) kapat();
  }

  return (
    <ui.Pencere
      acik
      baslik="Yeni Sipariş"
      genislik="560px"
      kapat={kapat}
      butonlar={[
        { etiket: "Vazgeç", onTikla: kapat },
        { etiket: "Kaydet", onTikla: kaydet, birincil: true },
      ]}
    >
      <div style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 10 }}>
        <div className="field-label" style={{ gridColumn: "1 / -1" }}>
          Müşteri
          <input
            className="input"
            style={s.giris}
            value={form.musteriAdi}
            placeholder="Müşteri adı yazın veya listeden seçin"
            onChange={(e) => alan("musteriAdi", e.target.value)}
          />
          {veri.cariler && veri.cariler.length > 0 && (
            <button type="button" className="btn-ghost" style={{ ...s.dugme, marginTop: 4 }} onClick={() => setCariListesiAcik((v) => !v)}>
              Cari Kartlarından Seç
            </button>
          )}
          {cariListesiAcik && (
            <div style={{ marginTop: 6, maxHeight: 160, overflowY: "auto", border: `1px solid ${s.renk?.kenar || "#2a4b52"}`, borderRadius: 8 }}>
              {(veri.cariler || []).map((c) => (
                <div
                  key={c.id}
                  onClick={() => { alan("musteriAdi", c.ad || c.unvan || ""); setCariListesiAcik(false); }}
                  style={{ padding: 8, cursor: "pointer", borderBottom: `1px solid ${s.renk?.kenar || "#2a4b52"}` }}
                >
                  {c.ad || c.unvan}
                </div>
              ))}
            </div>
          )}
        </div>
        <label className="field-label">Proje Kodu
          <input className="input" style={s.giris} value={form.projeKodu} onChange={(e) => alan("projeKodu", e.target.value)} />
        </label>
        <label className="field-label">Kalem No
          <input className="input" style={s.giris} value={form.kalemNo} onChange={(e) => alan("kalemNo", e.target.value)} />
        </label>
        <label className="field-label">Parça No
          <input className="input" style={s.giris} value={form.parcaNo} onChange={(e) => alan("parcaNo", e.target.value)} />
        </label>
        <label className="field-label">Parça Adı
          <input className="input" style={s.giris} value={form.parcaAdi} onChange={(e) => alan("parcaAdi", e.target.value)} />
        </label>
        <label className="field-label">Revizyon
          <input className="input" style={s.giris} value={form.revizyon} onChange={(e) => alan("revizyon", e.target.value)} />
        </label>
        <label className="field-label">Miktar
          <input className="input" type="number" min="1" style={s.giris} value={form.miktar} onChange={(e) => alan("miktar", e.target.value)} />
        </label>
        <label className="field-label">Teslim Tarihi
          <input className="input" type="date" style={s.giris} value={form.teslimTarihi} onChange={(e) => alan("teslimTarihi", e.target.value)} />
        </label>
        <label className="field-label">Dar Tolerans (mm)
          <input className="input" style={s.giris} placeholder="ör. 0.01" value={form.tightestTolerance} onChange={(e) => alan("tightestTolerance", e.target.value)} />
        </label>
        <label className="field-label">Gerekli Makine Sınıfı
          <input className="input" style={s.giris} placeholder="ör. CNC3 / CNCT / TASLAMA" value={form.requiredMachineClass} onChange={(e) => alan("requiredMachineClass", e.target.value)} />
        </label>
        <label className="field-label" style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <input type="checkbox" checked={form.sfc} onChange={(e) => alan("sfc", e.target.checked)} />
          Uçuş Emniyeti Kritik (SFC)
        </label>
      </div>
    </ui.Pencere>
  );
}
