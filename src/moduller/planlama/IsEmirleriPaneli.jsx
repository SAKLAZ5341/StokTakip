import React, { useMemo, useState } from "react";
import { Plus, AlertTriangle, Wrench, Settings } from "lucide-react";
import { hepsiniCizelgele, makineDolulukOzeti, IS_EMRI_TURLERI, isEmriTuruBul } from "./cizelgeMotoru";
import { makineOner, rotaAnahtari, guvenliCagir } from "./yardimcilar";

const AYARLAR_VARSAYILAN = { gunlukKapasiteSaat: 8 };

// Bu panel bir siparişin "Planlamada" aşamasını yönetir: WBS kalemlerinden
// (kısmi miktarlarla, birden çok kez) İş Emri oluşturma, her adıma istasyon
// atama (sistem önerisi + kapasite doluluk uyarısı) ve TÜM aktif iş
// emirlerinin CANLI hesaplanan çizelgesini (hepsiniCizelgele) gösterme.
// Makine kuyruğu atölye genelinde PAYLAŞILDIĞI için hesap tumIsEmirleri
// (sadece bu sipariş değil) üzerinden yapılır — ekranda bu siparişe filtrelenir.
export default function IsEmirleriPaneli({
  api, ui, veri, yazabilir,
  siparis, tumSiparisler, tumWbs, tumIsEmirleri, malzemeTalepleri, rotalar,
  ayarlarListesi,
}) {
  const s = ui.stil || {};
  const [yeniFormNode, setYeniFormNode] = useState(null);
  const [malzemeDuzenle, setMalzemeDuzenle] = useState(null);

  const ayarKaydi = (ayarlarListesi || [])[0];
  const ayarlar = ayarKaydi || AYARLAR_VARSAYILAN;

  async function gunlukSaatGuncelle(deger) {
    const sayi = Number(deger);
    if (!sayi || sayi <= 0) return;
    await guvenliCagir(async () => {
      if (ayarKaydi) await api.guncelle("planlama_ayarlar", ayarKaydi.id, { gunlukKapasiteSaat: sayi });
      else await api.ekleNumarali("planlama_ayarlar", "AYAR-GENEL", { gunlukKapasiteSaat: sayi });
    }, "Ayar güncellenemedi");
  }

  const cizelge = useMemo(
    () => hepsiniCizelgele(tumIsEmirleri, tumWbs, malzemeTalepleri, tumSiparisler, ayarlar),
    [tumIsEmirleri, tumWbs, malzemeTalepleri, tumSiparisler, ayarlar]
  );

  const siparisWbs = useMemo(() => (tumWbs || []).filter((w) => w.siparisId === siparis.id), [tumWbs, siparis.id]);
  const siparisIsEmirleri = useMemo(() => (tumIsEmirleri || []).filter((ie) => ie.siparisId === siparis.id), [tumIsEmirleri, siparis.id]);

  function planliMiktar(nodeId) {
    return siparisIsEmirleri
      .filter((ie) => ie.wbsId === nodeId && ie.durum !== "İptal")
      .reduce((top, ie) => top + (Number(ie.planlananMiktar) || 0), 0);
  }

  return (
    <div>
      <div className="card" style={{ padding: 16, marginBottom: 16 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 12 }}>
          <h3 style={{ margin: 0, color: s.renk?.yazi }}>WBS Kalemleri — İş Emri Oluştur</h3>
          {yazabilir && (
            <label className="field-label" style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12 }}>
              <Settings size={14} /> Günlük Kapasite (sa.)
              <input
                className="input" type="number" min="1" style={{ ...s.giris, width: 60 }}
                defaultValue={ayarlar.gunlukKapasiteSaat}
                onBlur={(e) => gunlukSaatGuncelle(e.target.value)}
              />
            </label>
          )}
        </div>
        <p style={{ fontSize: 12, color: s.renk?.soluk, marginTop: 0 }}>
          Bir kalem 100 adet olsa bile kısmi miktarla ("10 adet") iş emri açabilirsiniz;
          kalan miktar sıfırlanana kadar aynı kaleme yeni iş emirleri (SERİ, DELTA, FAI...)
          eklemeye devam edebilirsiniz.
        </p>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                {["WBS Kodu", "Parça", "Toplam", "Planlanan", "Kalan", ""].map((h) => (
                  <th key={h} style={s.tabloBaslik}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {siparisWbs.map((node) => {
                const planli = planliMiktar(node.id);
                const kalan = (node.miktar || 0) - planli;
                return (
                  <tr key={node.id}>
                    <td style={{ ...s.tabloHucre, fontFamily: "monospace", color: s.renk?.vurgu }}>{node.wbsKodu}</td>
                    <td style={s.tabloHucre}>
                      {node.parcaAdi} <span style={{ color: s.renk?.soluk, fontSize: 12 }}>({node.parcaNo})</span>
                    </td>
                    <td style={s.tabloHucre}>{node.miktar}</td>
                    <td style={s.tabloHucre}>{planli}</td>
                    <td style={{ ...s.tabloHucre, color: kalan > 0 ? (s.renk?.uyari) : (s.renk?.iyi), fontWeight: 600 }}>{kalan}</td>
                    <td style={s.tabloHucre}>
                      {yazabilir && kalan > 0 && (
                        <button className="btn-ghost" style={{ ...s.dugme, padding: "4px 8px" }} onClick={() => setYeniFormNode(node)}>
                          <Plus size={14} style={{ marginRight: 4 }} /> İş Emri
                        </button>
                      )}
                      {yazabilir && node.malzemeGerekli && node.malzemeKaynagi !== "musteri" && (
                        <button
                          className="btn-ghost" style={{ ...s.dugme, padding: "4px 8px", marginLeft: 4 }}
                          onClick={() => setMalzemeDuzenle(node)} title="Malzeme teslim durumunu güncelle"
                        >
                          <Wrench size={14} />
                        </button>
                      )}
                    </td>
                  </tr>
                );
              })}
              {siparisWbs.length === 0 && (
                <tr><td colSpan={6} style={{ ...s.tabloHucre, textAlign: "center", color: s.renk?.soluk }}>WBS kalemi yok.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      <div className="card" style={{ padding: 16 }}>
        <h3 style={{ margin: "0 0 12px", color: s.renk?.yazi }}>Çizelge (Canlı Hesap)</h3>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                {["İş Emri", "Tür", "Parça", "Planlanan", "Başlangıç", "Bitiş", "Termin Durumu", "Gecikme"].map((h) => (
                  <th key={h} style={s.tabloBaslik}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {siparisIsEmirleri.map((ie) => {
                const sonuc = cizelge.isEmriSonuclari[ie.id];
                const node = siparisWbs.find((n) => n.id === ie.wbsId);
                return (
                  <tr key={ie.id}>
                    <td style={s.tabloHucre}>{ie.isEmriNo}</td>
                    <td style={s.tabloHucre}>{isEmriTuruBul(ie.tur).ad}</td>
                    <td style={s.tabloHucre}>{node ? node.parcaAdi : ie.wbsId}</td>
                    <td style={s.tabloHucre}>{ie.planlananMiktar}</td>
                    <td style={s.tabloHucre}>{sonuc ? ui.tarih.tr(sonuc.genelBaslangic.slice(0, 10)) : "-"}</td>
                    <td style={s.tabloHucre}>{sonuc && !sonuc.belirsiz ? ui.tarih.tr(sonuc.genelBitis.slice(0, 10)) : "Belirsiz"}</td>
                    <td style={{ ...s.tabloHucre, color: terminRenk(sonuc, s) }}>{sonuc ? sonuc.terminDurumu : "-"}</td>
                    <td style={s.tabloHucre}>{sonuc && sonuc.gecikmeGun > 0 ? `${sonuc.gecikmeGun} gün` : "-"}</td>
                  </tr>
                );
              })}
              {siparisIsEmirleri.length === 0 && (
                <tr><td colSpan={8} style={{ ...s.tabloHucre, textAlign: "center", color: s.renk?.soluk }}>Henüz iş emri yok.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {yeniFormNode && (
        <YeniIsEmriFormu
          api={api} ui={ui} veri={veri}
          siparis={siparis} node={yeniFormNode}
          kalanMiktar={(yeniFormNode.miktar || 0) - planliMiktar(yeniFormNode.id)}
          mevcutIsEmirleri={tumIsEmirleri} rotalar={rotalar} cizelge={cizelge}
          kapat={() => setYeniFormNode(null)}
        />
      )}

      {malzemeDuzenle && (
        <MalzemeDurumDuzenle api={api} ui={ui} node={malzemeDuzenle} malzemeTalepleri={malzemeTalepleri} kapat={() => setMalzemeDuzenle(null)} />
      )}
    </div>
  );
}

function terminRenk(sonuc, s) {
  if (!sonuc) return s.renk?.soluk;
  if (sonuc.terminDurumu === "Zamanında") return s.renk?.iyi;
  if (sonuc.terminDurumu === "Gecikiyor") return s.renk?.hata;
  return s.renk?.uyari;
}

function YeniIsEmriFormu({ api, ui, veri, siparis, node, kalanMiktar, mevcutIsEmirleri, rotalar, cizelge, kapat }) {
  const s = ui.stil || {};
  const anahtar = rotaAnahtari(node.parcaNo, node.revizyon);
  const rota = (rotalar || []).find((r) => r.id === anahtar || r.anahtar === anahtar);

  const [tur, setTur] = useState("SERI");
  const [miktar, setMiktar] = useState(kalanMiktar);
  const [reworkKaynak, setReworkKaynak] = useState("");
  // Rota adımı makineyi "secilenMakineId" alanında saklar; iş emri adımı ise
  // "makineId" ile çalışır (çizelge motoru bu adı okur). Rotada önceden seçilmiş
  // istasyon burada ön-doldurulur, aksi halde kullanıcı her iş emrinde makineyi
  // yeniden seçmek zorunda kalırdı.
  const [adimlar, setAdimlar] = useState(() =>
    (rota && rota.adimlar ? rota.adimlar : []).map((a) => ({
      ...a,
      dahil: a.zorunlu !== false,
      makineId: a.makineId || a.secilenMakineId || "",
      sureSaat: "",
    }))
  );

  function adimGuncelle(idx, alan, deger) {
    setAdimlar((prev) => prev.map((a, i) => (i === idx ? { ...a, [alan]: deger } : a)));
  }

  const reworkAdaylari = (mevcutIsEmirleri || []).filter((ie) => ie.wbsId === node.id && ie.tur !== "FAI_REWORK");

  async function kaydet() {
    if (!rota) { alert("Bu parça/revizyon için onaylı rota yok. Önce Projelendirme'den rota tanımlayın."); return; }
    if (!miktar || Number(miktar) <= 0) { alert("Geçerli bir miktar girin."); return; }
    if (Number(miktar) > kalanMiktar) { alert(`En fazla ${kalanMiktar} adet planlayabilirsiniz.`); return; }
    if (tur === "FAI_REWORK" && !reworkKaynak) { alert("Rework için kaynak iş emri seçin."); return; }
    const eksikMakine = adimlar.some((a) => a.dahil && a.makineGerekli && !a.makineId);
    if (eksikMakine) { alert("Dahil edilen zorunlu/opsiyonel adımlarda makine seçilmemiş."); return; }
    const eksikSure = adimlar.some((a) => a.dahil && (!a.sureSaat || Number(a.sureSaat) <= 0));
    if (eksikSure) { alert("Dahil edilen tüm adımlar için tahmini süre (saat) girilmeli."); return; }

    const isEmriNo = api.sonrakiNo(mevcutIsEmirleri || [], "IE");
    const tamam = await guvenliCagir(
      () =>
        api.ekleNumarali("planlama_isEmirleri", isEmriNo, {
          isEmriNo, siparisId: siparis.id, wbsId: node.id,
          tur, planlananMiktar: Number(miktar),
          reworkKaynakIsEmriId: tur === "FAI_REWORK" ? reworkKaynak : null,
          adimlar: adimlar.map((a) => ({ ...a, sureSaat: Number(a.sureSaat) || 0 })),
          durum: "Planlandı",
        }),
      "İş emri oluşturulamadı"
    );
    if (tamam) kapat();
  }

  return (
    <ui.Pencere
      acik baslik={`Yeni İş Emri — ${node.wbsKodu}`} genislik="760px" kapat={kapat}
      butonlar={[{ etiket: "Vazgeç", onTikla: kapat }, { etiket: "İş Emrini Oluştur", onTikla: kaydet, birincil: true }]}
    >
      {!rota && (
        <div style={{ padding: 10, borderRadius: 8, background: (s.renk?.hata || "#dc2626") + "1a", color: s.renk?.hata || "#dc2626", marginBottom: 12, fontSize: 13 }}>
          <AlertTriangle size={14} style={{ marginRight: 6 }} /> Bu parça/revizyon için onaylı rota bulunamadı.
        </div>
      )}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(0,1fr))", gap: 10, marginBottom: 12 }}>
        <label className="field-label">İş Emri Türü
          <select className="input" style={s.giris} value={tur} onChange={(e) => setTur(e.target.value)}>
            {IS_EMRI_TURLERI.map((t) => <option key={t.kod} value={t.kod}>{t.ad}</option>)}
          </select>
        </label>
        <label className="field-label">Miktar (kalan: {kalanMiktar})
          <input className="input" type="number" min="1" max={kalanMiktar} style={s.giris} value={miktar} onChange={(e) => setMiktar(e.target.value)} />
        </label>
        {tur === "FAI_REWORK" && (
          <label className="field-label" style={{ gridColumn: "1 / -1" }}>Kaynak İş Emri
            <select className="input" style={s.giris} value={reworkKaynak} onChange={(e) => setReworkKaynak(e.target.value)}>
              <option value="">Seçin...</option>
              {reworkAdaylari.map((ie) => <option key={ie.id} value={ie.id}>{ie.isEmriNo}</option>)}
            </select>
          </label>
        )}
      </div>

      {rota && (
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                {["Dahil", "Operasyon", "Makine (Sistem Önerisi)", "Doluluk", "Süre (sa)"].map((h) => (
                  <th key={h} style={s.tabloBaslik}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {adimlar.map((a, idx) => {
                const oneri = a.makineGerekli ? makineOner(a.operasyonKod, node.tightestTolerance, veri.makineler) : { uygunlar: [] };
                return (
                  <tr key={idx} style={{ opacity: a.dahil ? 1 : 0.5 }}>
                    <td style={s.tabloHucre}>
                      <input
                        type="checkbox" checked={a.dahil}
                        disabled={a.zorunlu !== false}
                        onChange={(e) => adimGuncelle(idx, "dahil", e.target.checked)}
                        title={a.zorunlu !== false ? "Zorunlu adım — çıkarılamaz" : "Opsiyonel adım"}
                      />
                    </td>
                    <td style={s.tabloHucre}>{a.tanim}</td>
                    <td style={s.tabloHucre}>
                      {a.makineGerekli ? (
                        oneri.uygunlar.length > 0 ? (
                          <select
                            className="input" style={s.giris} value={a.makineId || ""} disabled={!a.dahil}
                            onChange={(e) => adimGuncelle(idx, "makineId", e.target.value)}
                          >
                            <option value="">Seçin...</option>
                            {oneri.uygunlar.map((m) => <option key={m.id} value={m.id}>{m.ad || m.makineAdi || m.id}</option>)}
                          </select>
                        ) : (
                          <span style={{ color: s.renk?.hata, fontSize: 12 }}>Uygun makine yok</span>
                        )
                      ) : (
                        <span style={{ color: s.renk?.soluk }}>—</span>
                      )}
                    </td>
                    <td style={s.tabloHucre}>
                      {a.makineGerekli && a.makineId ? <DolulukEtiketi makineId={a.makineId} cizelge={cizelge} ui={ui} /> : "-"}
                    </td>
                    <td style={s.tabloHucre}>
                      <input
                        className="input" type="number" step="0.5" min="0" style={{ ...s.giris, width: 70 }}
                        disabled={!a.dahil} value={a.sureSaat} onChange={(e) => adimGuncelle(idx, "sureSaat", e.target.value)}
                      />
                    </td>
                  </tr>
                );
              })}
              {adimlar.length === 0 && (
                <tr><td colSpan={5} style={{ ...s.tabloHucre, textAlign: "center", color: s.renk?.soluk }}>Rotada adım yok.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      )}
    </ui.Pencere>
  );
}

function DolulukEtiketi({ makineId, cizelge, ui }) {
  const ozet = makineDolulukOzeti(makineId, cizelge);
  if (!ozet.sonBitis) return <span style={{ fontSize: 12, color: "#15803d" }}>Boş</span>;
  return (
    <span style={{ fontSize: 12 }} title={`${ozet.isSayisi} iş, ${ozet.toplamSaat.toFixed(1)} sa. yük`}>
      Müsait: {ui.tarih.tr(ozet.sonBitis.slice(0, 10))}
    </span>
  );
}

function MalzemeDurumDuzenle({ api, ui, node, malzemeTalepleri, kapat }) {
  const s = ui.stil || {};
  const talepler = (malzemeTalepleri || []).filter((t) => t.wbsId === node.id);

  async function guncelle(talep, alanlar) {
    await guvenliCagir(() => api.guncelle("planlama_malzemeTalepleri", talep.id, alanlar), "Malzeme talebi güncellenemedi");
  }

  return (
    <ui.Pencere acik baslik={`Malzeme Durumu — ${node.wbsKodu}`} genislik="520px" kapat={kapat} butonlar={[{ etiket: "Kapat", onTikla: kapat }]}>
      {talepler.length === 0 && <p style={{ color: s.renk?.soluk }}>Bu kalem için gönderilmiş malzeme talebi yok.</p>}
      {talepler.map((t) => (
        <div key={t.id} style={{ padding: 10, border: `1px solid ${s.renk?.kenar || "#d5dfec"}`, borderRadius: 8, marginBottom: 8 }}>
          <div style={{ fontWeight: 600, marginBottom: 6 }}>{t.stokKodu} — {t.stokAdi}</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 8 }}>
            <label className="field-label">Durum
              <select className="input" style={s.giris} value={t.durum} onChange={(e) => guncelle(t, { durum: e.target.value })}>
                <option value="beklemede">Beklemede</option>
                <option value="siparişte">Siparişte</option>
                <option value="karsilandi">Karşılandı</option>
              </select>
            </label>
            <label className="field-label">Tahmini Teslim
              <input className="input" type="date" style={s.giris} value={t.tahminiTeslimTarihi || ""} onChange={(e) => guncelle(t, { tahminiTeslimTarihi: e.target.value })} />
            </label>
            <label className="field-label" style={{ gridColumn: "1 / -1" }}>Fiili Teslim (karşılandıysa)
              <input className="input" type="date" style={s.giris} value={t.fiiliTeslimTarihi || ""} onChange={(e) => guncelle(t, { fiiliTeslimTarihi: e.target.value })} />
            </label>
          </div>
        </div>
      ))}
    </ui.Pencere>
  );
}
