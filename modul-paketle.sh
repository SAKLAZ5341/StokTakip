#!/usr/bin/env bash
# =====================================================================
#  MODÜL PAKETLEYİCİ
#
#  Bir modül klasörünü (src/moduller/<ad>) tek bir ".sh" dosyasına çevirir.
#  Üretilen dosya, StokTakip deposunun bulunduğu terminale yapıştırılıp
#  çalıştırılır; klasörü yerine koyar, commit'ler ve gönderir.
#
#  Kullanım:
#     bash modul-paketle.sh planlama
#     -> planlama-modul-20260820.sh dosyasını üretir
#
#  Notlar:
#   - Paket YALNIZCA src/moduller/<ad>/ klasörünü taşır. Ana programın
#     hiçbir dosyasına dokunamaz — betik bunu hem paketlerken hem
#     açarken kontrol eder.
#   - Modülün menüye bağlanması ana programda bir kez yapılır; sonraki
#     sürümlerde sadece bu klasör değişir.
# =====================================================================
set -e

AD="$1"
if [ -z "$AD" ]; then
  echo "Kullanım: bash modul-paketle.sh <modul-adi>"
  echo "Örnek   : bash modul-paketle.sh planlama"
  exit 1
fi

KLASOR="src/moduller/$AD"
if [ ! -d "$KLASOR" ]; then
  echo "HATA: $KLASOR klasörü yok."
  exit 1
fi

# Giriş dosyası var mı?
if ! ls "$KLASOR"/*.jsx >/dev/null 2>&1; then
  echo "HATA: $KLASOR içinde .jsx dosyası bulunamadı."
  exit 1
fi

# Klasör dışına import var mı? (sözleşme gereği yasak)
if grep -rn "from *['\"]\.\./\.\./" "$KLASOR" 2>/dev/null; then
  echo
  echo "HATA: Modül klasörünün DIŞINA import var (yukarıdaki satırlar)."
  echo "Sözleşme gereği modül ana programın dosyalarını import edemez."
  echo "İhtiyacın olan şeyi prop olarak iste."
  exit 1
fi

TARIH=$(date +%Y%m%d)
CIKTI="$AD-modul-$TARIH.sh"
GECICI=$(mktemp -d)

tar czf "$GECICI/modul.tgz" "$KLASOR"
BOYUT=$(wc -c < "$GECICI/modul.tgz")
DOSYA_SAYISI=$(find "$KLASOR" -type f | wc -l | tr -d ' ')

{
cat <<HEAD_SON
#!/usr/bin/env bash
# =====================================================================
#  STOK TAKİP — "$AD" MODÜLÜ ($TARIH)
#  $DOSYA_SAYISI dosya · $BOYUT bayt (sıkıştırılmış)
#
#  Bu dosyanın TAMAMINI kopyalayıp Codespace terminaline yapıştır.
#  Sadece src/moduller/$AD/ klasörünü değiştirir; başka hiçbir dosyaya
#  dokunmaz.
# =====================================================================
set -e
cd /workspaces/StokTakip
git pull --ff-only

rm -f /tmp/modul-$AD.b64 /tmp/modul-$AD.tgz
cat > /tmp/modul-$AD.b64 <<'B64_SON'
HEAD_SON

base64 -w 76 < "$GECICI/modul.tgz"

cat <<TAIL_SON
B64_SON

base64 -d /tmp/modul-$AD.b64 > /tmp/modul-$AD.tgz

# Güvenlik: arşivin içinde sadece src/moduller/$AD/ altındaki dosyalar olmalı
DISARIDA=\$(tar tzf /tmp/modul-$AD.tgz | grep -v "^src/moduller/$AD/" || true)
if [ -n "\$DISARIDA" ]; then
  echo "HATA: Paket, modül klasörünün dışına dosya koymaya çalışıyor:"
  echo "\$DISARIDA"
  exit 1
fi

rm -rf "src/moduller/$AD"
tar xzf /tmp/modul-$AD.tgz
echo "src/moduller/$AD güncellendi (\$(find src/moduller/$AD -type f | wc -l | tr -d ' ') dosya)."

git add "src/moduller/$AD"
git commit -m "feat($AD): modul guncellendi ($TARIH)"
git push

echo
echo "==================================================================="
echo " TAMAM. GitHub > Actions yesil tik verince saklaz.net.tr guncel olur."
echo "==================================================================="
TAIL_SON
} > "$CIKTI"

rm -rf "$GECICI"
chmod +x "$CIKTI"

echo "Paket hazır: $CIKTI"
echo "  modül      : $AD ($DOSYA_SAYISI dosya)"
echo "  paket boyu : $(wc -c < "$CIKTI") bayt"
echo
echo "Bu dosyanın tamamını kopyalayıp StokTakip terminaline yapıştır."
