import React, { useMemo, useState } from "react";
import { Plus, Trash2, CheckCircle2, AlertTriangle, Printer } from "lucide-react";
import { OPERASYON_TIPLERI, operasyonBul, makineOner, makineAdi, rotaAnahtari, guvenliCagir } from "./yardimcilar";

// Seçili WBS kalemi için rota/istasyon yönetimi.
// - Aynı parça+revizyon için daha önce onaylı bir rota varsa: adımlar YENİDEN
//   girilmez, sadece "Onayla ve Kullan" ile bu kaleme bağlanır.
// - Yoksa: adım adım rota kurulur; her makine gerektiren adımda sistem,
//   tightestTolerance'ı karşılayan makineleri önerir (uygun olmayanlar listeye
//   hiç girmez — AS9100 GD&T hard-block).
export default function RotaPaneli({ api, ui, veri, yazabilir, node, siparis, rotalar }) {
  const s = ui.stil || {};
  const anahtar = rotaAnahtari(node.parcaNo, node.revizyon);
  const mevcutRota = useMemo(
    () => (rotalar || []).find((r) => r.id === anahtar || r.anahtar === anahtar),
    [rotalar, anahtar]
  );
  const rotaMevcut = !!mevcutRota;

  const [taslakAdimlar, setTaslakAdimlar] = useState([]);
  const [operasyonKod, setOperasyonKod] = useState(OPERASYON_TIPLERI[0].kod);
  const [aciklama, setAciklama] = useState("");
  const [secilenMakineId, setSecilenMakineId] = useState("");
  const [zorunlu, setZorunlu] = useState(true);

  const secilenOp = operasyonBul(operasyonKod);
  const oneri = useMemo(() => {
    if (!secilenOp) return { uygunlar: [], geregi: false };
    return makineOner(secilenOp.kod, node.tightestTolerance, veri.makineler);
  }, [secilenOp, node.tightestTolerance, veri.makineler]);

  const gosterilenAdimlar = rotaMevcut ? mevcutRota.adimlar || [] : taslakAdimlar;

  function opDegisti(kod) {
    setOperasyonKod(kod);
    const op = operasyonBul(kod);
    if (op && op.makineGerekli) {
      const on = makineOner(op.kod, node.tightestTolerance, veri.makineler);
      setSecilenMakineId(on.uygunlar[0] ? on.uygunlar[0].id : "");
    } else {
      setSecilenMakineId("");
    }
  }

  function adimEkle() {
    if (!secilenOp) return;
    if (secilenOp.makineGerekli && oneri.uygunlar.length === 0) {
      alert(
        `"${secilenOp.ad}" için ${node.tightestTolerance || "belirtilen"} mm dar toleransı karşılayan uygun makine bulunamadı.\n` +
        `Bu adım eklenemez (AS9100 GD&T kilidi). Makine tanımlarını / kapasite bilgilerini kontrol edin.`
      );
      return;
    }
    if (secilenOp.makineGerekli && !secilenMakineId) {
      alert("Önerilen makinelerden birini seçin.");
      return;
    }
    setTaslakAdimlar((prev) => [
      ...prev,
      {
        opNo: (prev.length + 1) * 10,
        operasyonKod: secilenOp.kod,
        tanim: aciklama || secilenOp.ad,
        birim: secilenOp.birim,
        makineGerekli: secilenOp.makineGerekli,
        onerilenMakineIdler: oneri.uygunlar.map((m) => m.id),
        secilenMakineId: secilenOp.makineGerekli ? secilenMakineId : "",
        zorunlu,
      },
    ]);
    setAciklama("");
    setZorunlu(true);
  }

  function adimSil(idx) {
    setTaslakAdimlar((prev) => prev.filter((_, i) => i !== idx).map((a, i) => ({ ...a, opNo: (i + 1) * 10 })));
  }

  async function rotayiKaydet() {
    if (taslakAdimlar.length === 0) { alert("En az bir operasyon adımı ekleyin."); return; }
    const tamam = await guvenliCagir(async () => {
      await api.ekleNumarali("planlama_rotalar", anahtar, {
        anahtar, parcaNo: node.parcaNo, revizyon: node.revizyon || "00",
        adimlar: taslakAdimlar, onayDurumu: "onaylandi",
      });
      await api.guncelle("planlama_wbs", node.id, { rotaId: anahtar, durum: "Rotalandı" });
    }, "Rota kaydedilemedi");
    if (tamam) setTaslakAdimlar([]);
  }

  async function rotayiOnaylaVeKullan() {
    await guvenliCagir(
      () => api.guncelle("planlama_wbs", node.id, { rotaId: anahtar, durum: "Rotalandı (Mevcut Rota)" }),
      "Rota bağlanamadı"
    );
  }

  function isEmriYazdir() {
    ui.yazdir({
      belgeAdi: "İŞ EMRİ / WORK ORDER",
      dokumanKodu: "planlama_isemri",
      ustBilgiler: [
        ["Müşteri", siparis.musteriAdi],
        ["Proje Kodu", siparis.projeKodu],
        ["Parça No", node.parcaNo],
        ["Parça Adı", node.parcaAdi],
        ["Revizyon", node.revizyon || "00"],
        ["Miktar", String(node.miktar)],
        ["WBS Kodu", node.wbsKodu],
        ["SFC", siparis.sfc ? "EVET" : "HAYIR"],
        ["Tarih", ui.tarih.tr(ui.tarih.bugun())],
      ],
      kolonlar: [
        { baslik: "Op.No", gen: "16mm", hiza: "ort", al: (r) => r.opNo },
        { baslik: "Operasyon Tanımı", al: (r) => r.tanim },
        { baslik: "Birim", gen: "40mm", al: (r) => r.birim },
        { baslik: "Makine / İstasyon", gen: "35mm", al: (r) => (r.makineGerekli ? makineAdi(r.secilenMakineId, veri.makineler) : "-") },
      ],
      satirlar: gosterilenAdimlar,
      imzalar: ["Operatör", "Kontrol Eden"],
    });
  }

  return (
    <div>
      {rotaMevcut ? (
        <div style={{ marginBottom: 12, padding: 10, borderRadius: 8, background: (s.renk?.iyi || "#15803d") + "1a", border: `1px solid ${(s.renk?.iyi || "#15803d")}55` }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6, color: s.renk?.iyi || "#15803d", fontWeight: 600 }}>
            <CheckCircle2 size={16} /> Bu parça/revizyon için onaylı rota kütüphanede mevcut.
          </div>
          <div style={{ fontSize: 12, color: s.renk?.soluk, marginTop: 4 }}>
            Revizyon değişmediği sürece adımlar yeniden girilmez — sadece onaylayıp bu kalemde kullanın.
          </div>
          {yazabilir && node.rotaId !== anahtar && (
            <button className="btn-ghost" style={{ ...s.anaDugme, marginTop: 8 }} onClick={rotayiOnaylaVeKullan}>
              <CheckCircle2 size={14} style={{ marginRight: 6 }} /> Onayla ve Bu Kalemde Kullan
            </button>
          )}
        </div>
      ) : (
        <div style={{ marginBottom: 12, padding: 10, borderRadius: 8, background: (s.renk?.uyari || "#a16207") + "1a", border: `1px solid ${(s.renk?.uyari || "#a16207")}55`, color: s.renk?.uyari || "#a16207", fontSize: 13 }}>
          <AlertTriangle size={14} style={{ marginRight: 6 }} />
          Bu parça/revizyon için kayıtlı rota yok. Aşağıdan yeni rota tanımlayın — kaydedildikten sonra
          aynı parça+revizyon tekrar geldiğinde otomatik önerilecek ve sadece onay istenecektir.
        </div>
      )}

      <div style={{ overflowX: "auto", marginBottom: 12 }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              <th style={s.tabloBaslik}>Op.No</th>
              <th style={s.tabloBaslik}>Operasyon</th>
              <th style={s.tabloBaslik}>Birim</th>
              <th style={s.tabloBaslik}>Makine / İstasyon</th>
              <th style={s.tabloBaslik}>Durum</th>
              {yazabilir && !rotaMevcut && <th style={s.tabloBaslik}></th>}
            </tr>
          </thead>
          <tbody>
            {gosterilenAdimlar.map((a, idx) => (
              <tr key={idx}>
                <td style={s.tabloHucre}>{a.opNo}</td>
                <td style={s.tabloHucre}>{a.tanim}</td>
                <td style={s.tabloHucre}>{a.birim}</td>
                <td style={s.tabloHucre}>
                  {a.makineGerekli ? makineAdi(a.secilenMakineId, veri.makineler) : <span style={{ color: s.renk?.soluk }}>—</span>}
                </td>
                <td style={s.tabloHucre}>
                  {a.zorunlu === false ? (
                    <span style={{ fontSize: 11, color: s.renk?.uyari || "#a16207" }}>Opsiyonel</span>
                  ) : (
                    <span style={{ fontSize: 11, color: s.renk?.soluk }}>Zorunlu</span>
                  )}
                </td>
                {yazabilir && !rotaMevcut && (
                  <td style={s.tabloHucre}>
                    <button className="btn-ghost" style={{ ...s.dugme, padding: "2px 6px" }} onClick={() => adimSil(idx)}>
                      <Trash2 size={14} />
                    </button>
                  </td>
                )}
              </tr>
            ))}
            {gosterilenAdimlar.length === 0 && (
              <tr>
                <td colSpan={6} style={{ ...s.tabloHucre, textAlign: "center", color: s.renk?.soluk }}>Henüz adım yok.</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {yazabilir && !rotaMevcut && (
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "flex-end", padding: 10, border: `1px dashed ${s.renk?.kenar || "#d5dfec"}`, borderRadius: 8, marginBottom: 12 }}>
          <label className="field-label" style={{ minWidth: 200 }}>Operasyon
            <select className="input" style={s.giris} value={operasyonKod} onChange={(e) => opDegisti(e.target.value)}>
              {OPERASYON_TIPLERI.map((o) => <option key={o.kod} value={o.kod}>{o.ad}</option>)}
            </select>
          </label>

          {secilenOp && secilenOp.makineGerekli && (
            <label className="field-label" style={{ minWidth: 220 }}>
              Sistem Önerisi (Tolerans Uyumlu)
              {oneri.uygunlar.length > 0 ? (
                <select className="input" style={s.giris} value={secilenMakineId} onChange={(e) => setSecilenMakineId(e.target.value)}>
                  {oneri.uygunlar.map((m) => (
                    <option key={m.id} value={m.id}>{m.ad || m.makineAdi || m.id}</option>
                  ))}
                </select>
              ) : (
                <div style={{ fontSize: 12, color: s.renk?.hata || "#dc2626", padding: "6px 0" }}>
                  Uygun makine yok — adım eklenemez
                </div>
              )}
            </label>
          )}

          <label className="field-label" style={{ flex: 1, minWidth: 160 }}>Açıklama (opsiyonel)
            <input className="input" style={s.giris} value={aciklama} onChange={(e) => setAciklama(e.target.value)} />
          </label>

          <label className="field-label" style={{ display: "flex", alignItems: "center", gap: 6, minWidth: 140 }}>
            <input type="checkbox" checked={zorunlu} onChange={(e) => setZorunlu(e.target.checked)} />
            Zorunlu adım
          </label>

          <button className="btn-ghost" style={s.anaDugme} onClick={adimEkle}>
            <Plus size={14} style={{ marginRight: 6 }} /> Adım Ekle
          </button>
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "space-between", gap: 8, flexWrap: "wrap" }}>
        {yazabilir && !rotaMevcut && (
          <button className="btn-ghost" style={s.anaDugme} onClick={rotayiKaydet}>
            <CheckCircle2 size={14} style={{ marginRight: 6 }} /> Rotayı Kaydet ve Onayla
          </button>
        )}
        <button className="btn-ghost" style={s.dugme} onClick={isEmriYazdir} disabled={gosterilenAdimlar.length === 0}>
          <Printer size={14} style={{ marginRight: 6 }} /> İş Emri Yazdır
        </button>
      </div>
    </div>
  );
}
