"""
Kaçak Ekipleri - Günlük İş Planlama Dashboard
"""

import streamlit as st
import pandas as pd
import numpy as np
import folium
from folium import plugins
from streamlit_folium import st_folium
from shapely.geometry import Point, shape
from io import BytesIO
from datetime import datetime
import math
import json
import base64
import os

# ============================================================
# SAYFA AYARLARI
# ============================================================
st.set_page_config(
    page_title="Ekip İş Planlama Dashboard",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# EKİP RENKLERİ (havuz - en fazla 20 ekip desteklenir)
# ============================================================
EKIP_RENK_HAVUZU = [
    ("Ekip1", "#e74c3c"),   ("Ekip2", "#3498db"),   ("Ekip3", "#2ecc71"),
    ("Ekip4", "#f39c12"),   ("Ekip5", "#9b59b6"),   ("Ekip6", "#1abc9c"),
    ("Ekip7", "#e67e22"),   ("Ekip8", "#e84393"),   ("Ekip9", "#00cec9"),
    ("Ekip10", "#6c5ce7"),  ("Ekip11", "#fd79a8"),  ("Ekip12", "#a29bfe"),
    ("Ekip13", "#ffeaa7"),  ("Ekip14", "#55efc4"),  ("Ekip15", "#74b9ff"),
    ("Ekip16", "#dfe6e9"),  ("Ekip17", "#b2bec3"),  ("Ekip18", "#636e72"),
    ("Ekip19", "#d63031"),  ("Ekip20", "#0984e3"),
]
ATANMAMIS_RENK = "#95a5a6"  # Gri

def aktif_ekipler(sayi):
    """Kullanıcının seçtiği ekip sayısına göre ekip adı-renk dict'i döner."""
    return dict(EKIP_RENK_HAVUZU[:sayi])


# ============================================================
# YARDIMCI FONKSİYONLAR
# ============================================================

def koordinat_temizle(val):
    """Koordinat değerindeki virgülleri noktaya çevirir ve float'a dönüştürür."""
    if pd.isna(val):
        return np.nan
    if isinstance(val, (int, float)):
        return float(val)
    try:
        return float(str(val).strip().replace(",", "."))
    except (ValueError, TypeError):
        return np.nan


def koordinatlari_dogrula(df):
    """Koordinatları doğrular ve düzeltir.
    - Enlem/Boylam karışmışsa takas eder
    - Türkiye sınırları dışındakileri filtreler
    - Temizleme istatistiklerini döner (Vektörize edilmiş yüksek performanslı versiyon)
    """
    istatistik = {"toplam_once": len(df), "takas": 0, "gecersiz": 0}

    # Türkiye koordinat sınırları
    ENLEM_MIN, ENLEM_MAX = 35.5, 42.5   # Kuzey-Güney
    BOYLAM_MIN, BOYLAM_MAX = 25.5, 45.0  # Doğu-Batı

    # --- ADIM 1: Enlem-Boylam karışıklığını tespit et ve düzelt ---
    takas_mask = (df["Enlem"] >= BOYLAM_MIN) & (df["Enlem"] <= BOYLAM_MAX) & \
                 (df["Boylam"] >= ENLEM_MIN) & (df["Boylam"] <= ENLEM_MAX)
                 
    istatistik["takas"] = takas_mask.sum()
    if istatistik["takas"] > 0:
        df.loc[takas_mask, ["Enlem", "Boylam"]] = df.loc[takas_mask, ["Boylam", "Enlem"]].values

    # --- ADIM 2: Türkiye sınırları dışındakileri filtrele ---
    gecerli_mask = (
        (df["Enlem"] >= ENLEM_MIN) & (df["Enlem"] <= ENLEM_MAX) &
        (df["Boylam"] >= BOYLAM_MIN) & (df["Boylam"] <= BOYLAM_MAX)
    )
    istatistik["gecersiz"] = (~gecerli_mask).sum()
    df = df[gecerli_mask].reset_index(drop=True)

    istatistik["toplam_sonra"] = len(df)
    return df, istatistik


def jitter_uygula(df):
    """Aynı koordinattaki noktaları dairesel olarak hafifçe kaydırır (~10m).
    Vektörize edilmiş yüksek performanslı versiyon.
    """
    JITTER_YARICAP = 0.00012  # ~13 metre (derece cinsinden)

    # Başlangıçta gösterim = orijinal
    df["Enlem_Gosterim"] = df["Enlem"].copy()
    df["Boylam_Gosterim"] = df["Boylam"].copy()

    # Aynı lokasyondaki nokta sayılarını bul
    grup_sayilari = df.groupby(["Enlem", "Boylam"])["Enlem"].transform("size")
    mask = grup_sayilari > 1

    if mask.any():
        sira = df[mask].groupby(["Enlem", "Boylam"]).cumcount()
        grup_buyuklugu = grup_sayilari[mask]
        acilar = (2 * np.pi * sira) / grup_buyuklugu
        
        df.loc[mask, "Enlem_Gosterim"] += JITTER_YARICAP * np.sin(acilar)
        df.loc[mask, "Boylam_Gosterim"] += JITTER_YARICAP * np.cos(acilar)

    return df


def haversine_mesafe(lat1, lon1, lat2, lon2):
    """İki koordinat arasındaki kuş uçuşu mesafeyi km cinsinden hesaplar (Numpy ile Vektörize)."""
    R = 6371.0
    lat1_r, lon1_r = np.radians(lat1), np.radians(lon1)
    lat2_r, lon2_r = np.radians(lat2), np.radians(lon2)
    dlat = lat2_r - lat1_r
    dlon = lon2_r - lon1_r
    a = np.sin(dlat / 2) ** 2 + np.cos(lat1_r) * np.cos(lat2_r) * np.sin(dlon / 2) ** 2
    c = 2 * np.arctan2(np.sqrt(a), np.sqrt(1 - a))
    return R * c


def excel_oku(uploaded_file):
    """Yüklenen Excel dosyasını okur, gerekli sütunları kontrol eder ve temizler."""
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        return None, None, f"❌ **Excel dosyası okunamadı!**\n\nDosya bozuk olabilir veya formatı (xlsx/xls) desteklenmiyor olabilir.\nHata detayı: `{str(e)}`"

    if df.empty:
        return None, None, None, "⚠️ **Yüklenen Excel dosyası boş!**\n\nLütfen içinde veri olan bir dosya yükleyin."

    # --- SÜTUN TESPİT MANTIĞI (Öncelikli) ---
    sutun_haritasi = {}
    bulunan_sutunlar = df.columns.tolist()

    for col in bulunan_sutunlar:
        col_str = str(col).strip()
        cl = col_str.lower().replace(" ", "_").replace("ı", "i").replace("ş", "s").replace("ç", "c").replace("ö", "o").replace("ü", "u")
        
        # 1. Hizmet Noktası No - Kesin eşleşmeler
        if cl in ["hizmet_noktasi_no", "hizmet_noktasi", "tesisat_no", "tesisat_numarasi", "hizmet_no"]:
            sutun_haritasi["Hizmet_Noktasi_No"] = col
        
        # 2. Enlem/Boylam (tek sütun - "36,25459068/30,00748026" formatında)
        elif cl in ["enlem/boylam", "enlem_boylam", "enlem/boylem"]:
            sutun_haritasi["Enlem_Boylam"] = col

    # --- FALLBACK: Eğer hala bulunamadıysa geniş arama (Benzerleri eleyerek) ---
    if "Hizmet_Noktasi_No" not in sutun_haritasi:
        for col in bulunan_sutunlar:
            cl = str(col).lower().replace(" ", "_").replace("ı", "i").replace("ş", "s")
            # "hizmet" ve "no" geçsin ama "anlasma" geçmesin
            if "hizmet" in cl and "no" in cl and "anlasma" not in cl:
                sutun_haritasi["Hizmet_Noktasi_No"] = col
                break

    # Enlem/Boylam fallback
    if "Enlem_Boylam" not in sutun_haritasi:
        for col in bulunan_sutunlar:
            cl = str(col).strip().lower()
            if "enlem" in cl and "boylam" in cl:
                sutun_haritasi["Enlem_Boylam"] = col
                break

    # Bölge, İl, İlçe için esnek arama
    for col in bulunan_sutunlar:
        cl_str = str(col).strip()
        cl = cl_str.lower().replace("ı", "i").replace("ş", "s").replace("ç", "c").replace("ö", "o").replace("ü", "u")
        
        # Bölge tespiti
        if "bolge" in cl or "region" in cl or ("lge" in cl and len(cl) < 10):
            if "Bolge" not in sutun_haritasi: sutun_haritasi["Bolge"] = col
        
        # İl tespiti (İlçe ile karışmaması için "ilce" içermemeli)
        if (cl == "il" or cl == "i̇l" or cl == "sehir" or cl == "city" or " il" in " " + cl) and "ilce" not in cl and "ilçe" not in cl:
            if "Il" not in sutun_haritasi: sutun_haritasi["Il"] = col
            
    # İlçe tespiti
    if "Ilce" not in sutun_haritasi:
        for col in bulunan_sutunlar:
            cl = str(col).strip().lower()
            if "ilce" in cl or "ilçe" in cl or "district" in cl or "lce" in cl:
                sutun_haritasi["Ilce"] = col
                break

    # --- FALLBACK: Eğer hala bulunamadıysa ilk 3 sütunu varsayılan olarak kullan ---
    if "Bolge" not in sutun_haritasi and len(bulunan_sutunlar) > 0:
        sutun_haritasi["Bolge"] = bulunan_sutunlar[0]
    if "Il" not in sutun_haritasi and len(bulunan_sutunlar) > 1:
        sutun_haritasi["Il"] = bulunan_sutunlar[1]
    if "Ilce" not in sutun_haritasi and len(bulunan_sutunlar) > 2:
        sutun_haritasi["Ilce"] = bulunan_sutunlar[2]

    # Zorunlu sütun kontrolü
    eksik = []
    if "Hizmet_Noktasi_No" not in sutun_haritasi: eksik.append("Hizmet Noktası No (veya Tesisat No)")
    if "Enlem_Boylam" not in sutun_haritasi: eksik.append("Enlem/Boylam")

    if eksik:
        hata_mesaji = "### ❌ Eksik Sütun Hatası\n\nExcel dosyanızda aşağıdaki zorunlu sütunlar bulunamadı:\n\n"
        for e in eksik:
            hata_mesaji += f"- **{e}**\n"
        hata_mesaji += f"\n\n**Sizin Excel'deki Sütunlar:** `{', '.join([str(c) for c in bulunan_sutunlar])}`"
        hata_mesaji += "\n\n💡 *Lütfen Excel'deki başlıkları yukarıdaki isimlerle uyumlu olacak şekilde düzeltip tekrar yükleyin.*"
        return None, None, None, hata_mesaji

    # Veriyi hazırla
    try:
        sonuc = pd.DataFrame()
        sonuc["Tesisat_No"] = df[sutun_haritasi["Hizmet_Noktasi_No"]].apply(
            lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (float, int)) else str(x)
        )

        # --- ENLEM/BOYLAM PARSE (tek sütundan ayır) Vektörize ---
        eb_str = df[sutun_haritasi["Enlem_Boylam"]].astype(str).str.strip()
        parcalar = eb_str.str.split("/", expand=True)
        if len(parcalar.columns) >= 2:
            sonuc["Enlem"] = pd.to_numeric(parcalar[0].str.strip().str.replace(",", "."), errors='coerce')
            sonuc["Boylam"] = pd.to_numeric(parcalar[1].str.strip().str.replace(",", "."), errors='coerce')
        else:
            sonuc["Enlem"] = np.nan
            sonuc["Boylam"] = np.nan

        sonuc["Bolge"] = df[sutun_haritasi.get("Bolge", "")].astype(str) if "Bolge" in sutun_haritasi else ""
        sonuc["Il"] = df[sutun_haritasi.get("Il", "")].astype(str) if "Il" in sutun_haritasi else ""
        sonuc["Ilce"] = df[sutun_haritasi.get("Ilce", "")].astype(str) if "Ilce" in sutun_haritasi else ""
        sonuc["Ekip"] = "Atanmamış"

        # --- OPSİYONEL SÜTUNLAR ---
        opsiyonel_eslestirme = {
            "Sayac_Seri_No": ["sayac_seri_no", "sayac seri no", "sayaç seri"],
            "Adres": ["adres", "address"],
            "Sayac_Tip_Adi": ["sayac_tipi", "sayac tipi", "sayaç tipi", "sayac_tip_adi"],
            "AG_OG": ["ag/og", "ag_og", "agog"],
            "Saha_Aktivitesi_Yonergeleri": ["saha_aktivitesi_yonergeleri", "saha aktivitesi yonergeleri"],
            "Rezerv_Kwh": ["rezerv_kwh", "rezerv kwh"],
            "Ihbar_Sekli_2": ["ihbar_sekli_2", "ihbar sekli 2"],
            "Tahakkuk_Carpani": ["tahakkuk_carpani", "tahakkuk carpani"],
            "Hizmet_Nok_Tip_Kodu": ["hizmet_nok_tip_kodu", "hizmet nok tip kodu"],
        }
        for hedef_ad, aranacaklar in opsiyonel_eslestirme.items():
            for col in bulunan_sutunlar:
                cl_norm = str(col).lower().replace(" ", "_").replace("ı", "i").replace("ş", "s").replace("ç", "c").replace("ü", "u").replace("ö", "o").replace("İ", "i").replace("\u0307", "")
                if cl_norm in [a.replace(" ", "_") for a in aranacaklar]:
                    if hedef_ad in ["Sayac_Seri_No", "Tahakkuk_Carpani"]:
                        sonuc[hedef_ad] = df[col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True).replace('nan', '')
                    else:
                        sonuc[hedef_ad] = df[col].astype(str).replace("nan", "")
                    break
            if hedef_ad not in sonuc.columns:
                sonuc[hedef_ad] = ""

        # Tarife haritası
        tarife_map = {
            "E-AYD": "Aydınlatma",
            "E-DTR": "Trafo",
            "E-MES": "Mesken",
            "E-SAN": "Sanayi",
            "E-TAR": "Tarımsal",
            "E-TIC": "Ticarethane"
        }
        if "Hizmet_Nok_Tip_Kodu" in sonuc.columns:
            sonuc["Tarife"] = sonuc["Hizmet_Nok_Tip_Kodu"].map(tarife_map).fillna(sonuc["Hizmet_Nok_Tip_Kodu"])
        else:
            sonuc["Tarife"] = ""

        # Boş veya NaN olan değerleri "Boş" olarak değiştir (Optimizasyon)
        for col in ["AG_OG", "Tarife", "Ihbar_Sekli_2", "Saha_Aktivitesi_Yonergeleri"]:
            if col in sonuc.columns:
                sonuc[col] = sonuc[col].replace(["", "nan", "NaN", "None", None, np.nan], "Boş")
                sonuc[col] = sonuc[col].fillna("Boş")

        # --- KOORDİNAT TEMİZLİĞİ ---
        koordinatsiz_mask = sonuc["Enlem"].isna() | sonuc["Boylam"].isna() | (sonuc["Enlem"] == 0) | (sonuc["Boylam"] == 0)
        koordinatsiz_df = sonuc[koordinatsiz_mask].copy()
        
        sonuc = sonuc[~koordinatsiz_mask].reset_index(drop=True)
        koordinatsiz_df = koordinatsiz_df.reset_index(drop=True)
        
        if len(sonuc) == 0:
            return None, None, None, "❌ **Geçerli Koordinat Yok!**\n\nExcel'deki tüm Enlem/Boylam değerleri boş veya geçersiz (0) görünüyor."

        # 2) Enlem-Boylam takası ve sınır kontrolü
        sonuc, istatistik = koordinatlari_dogrula(sonuc)

        # 3) Aynı konumdaki noktaları dairesel kaydır (jitter)
        sonuc = jitter_uygula(sonuc)

        return sonuc, istatistik, koordinatsiz_df, None

    except Exception as e:
        return None, None, None, f"❌ **Veri işleme hatası!**\n\nSütunlar bulundu ancak veriler işlenirken bir sorun çıktı.\nHata: `{str(e)}`"



def harita_olustur(df, ekip_renkleri, arama_noktasi=None, goster_rezerv_kwh=False):
    """Folium haritası - ekip bazlı katmanlarla toggle desteği + arama noktası."""
    merkez_lat = df["Enlem"].mean()
    merkez_lon = df["Boylam"].mean()
    zoom = 12

    # Arama noktası varsa oraya zoom
    if arama_noktasi is not None:
        merkez_lat, merkez_lon = arama_noktasi
        zoom = 17

    m = folium.Map(location=[merkez_lat, merkez_lon], zoom_start=zoom, tiles="OpenStreetMap")

    # Draw plugini
    draw = plugins.Draw(
        export=False, position="topleft",
        draw_options={
            "polyline": False, "circle": False, "circlemarker": False, "marker": False,
            "polygon": {"allowIntersection": False, "showArea": True},
            "rectangle": {"showArea": True},
        },
        edit_options={"edit": False}
    )
    draw.add_to(m)

    # Fullscreen eklentisi
    plugins.Fullscreen(
        position="topleft",
        title="Tam Ekran",
        title_cancel="Tam Ekrandan Çık",
        force_separate_button=True
    ).add_to(m)

    # --- Ekip bazlı FeatureGroup'lar (toggle için) ---
    cluster_ayar = {"maxClusterRadius": 40, "disableClusteringAtZoom": 17, "spiderfyOnMaxZoom": True}

    # Atanmamış grubu (FeatureGroup)
    fg_atanmamis = folium.FeatureGroup(name=f"⬤ Atanmamış", show=True)

    # Ekip grupları
    ekip_gruplari = {}
    for ekip_adi, renk in ekip_renkleri.items():
        fg = folium.FeatureGroup(name=f'<span style="color:{renk}">⬤</span> {ekip_adi}', show=True)
        ekip_gruplari[ekip_adi] = fg

    # Vektörize HTML Popup ve Tooltip Oluşturma
    df_temp = df.copy()
            
    for col in ["AG_OG", "Tarife", "Ihbar_Sekli_2", "Saha_Aktivitesi_Yonergeleri"]:
        if col in df_temp.columns:
            df_temp[col] = df_temp[col].replace(["", "nan", "NaN", "None", None, np.nan], "Boş")
            df_temp[col] = df_temp[col].fillna("Boş")
        else:
            df_temp[col] = "Boş"
            
    popup_series = (
        "<b>Ekip:</b> " + df_temp["Ekip"].astype(str) + "<br>"
        "<b>İlçe:</b> " + df_temp["Ilce"].astype(str) + "<br>"
        "<b>Tesisat:</b> " + df_temp["Tesisat_No"].astype(str) + "<br>"
        "<b>AG/OG:</b> " + df_temp["AG_OG"].astype(str) + "<br>"
        "<b>Tarife:</b> " + df_temp["Tarife"].astype(str) + "<br>"
        "<b>İhbar:</b> " + df_temp["Ihbar_Sekli_2"].astype(str) + "<br>"
        "<b>Yönerge:</b> " + df_temp["Saha_Aktivitesi_Yonergeleri"].astype(str)
    )
    if goster_rezerv_kwh and "Rezerv_Kwh" in df_temp.columns:
        popup_series += "<br><b>Rezerv Kwh:</b> " + df_temp["Rezerv_Kwh"].astype(str)
        
    df_temp["popup_html_val"] = popup_series
    df_temp["tooltip_val"] = df_temp["Tesisat_No"].astype(str) + " - " + df_temp["Ekip"].astype(str)

    # FastMarkerCluster ile Vektörize Hızlı Renderlama (Optimizasyon)
    callback = """
    function (row) {
        var marker = L.circleMarker(new L.LatLng(row[0], row[1]), {color: row[2], fill: true, fillColor: row[2], fillOpacity: 0.85, radius: 6});
        var popup = L.popup({maxWidth: 250}).setContent(row[3]);
        marker.bindPopup(popup);
        marker.bindTooltip(row[4]);
        return marker;
    };
    """

    for ekip, group in df_temp.groupby("Ekip"):
        if ekip == "Atanmamış":
            renk = ATANMAMIS_RENK
            hedef_fg = fg_atanmamis
        else:
            renk = ekip_renkleri.get(ekip, "#333333")
            hedef_fg = ekip_gruplari.get(ekip, fg_atanmamis)
            
        data_kume = group[["Enlem_Gosterim", "Boylam_Gosterim"]].copy()
        data_kume["renk"] = renk
        data_kume["popup"] = group["popup_html_val"]
        data_kume["tooltip"] = group["tooltip_val"]
        
        fmc = plugins.FastMarkerCluster(
            data_kume.values.tolist(), 
            callback=callback,
            **cluster_ayar
        )
        fmc.add_to(hedef_fg)

    # Grupları haritaya ekle
    fg_atanmamis.add_to(m)
    for fg in ekip_gruplari.values():
        fg.add_to(m)

    # Arama noktası varsa özel marker
    if arama_noktasi is not None:
        folium.Marker(
            location=arama_noktasi,
            icon=folium.Icon(color="red", icon="star", prefix="fa"),
            popup="📍 Aranan Hizmet Noktası",
            tooltip="📍 Aranan Nokta"
        ).add_to(m)

    # LayerControl — ekipleri açıp kapatma
    folium.LayerControl(collapsed=False).add_to(m)

    return m


def koordinat_sirasi_tespit(geometry):
    """Leaflet Draw'dan gelen GeoJSON koordinatlarının sırasını tespit eder.
    GeoJSON standardı [lng, lat] dir ama Leaflet bazen [lat, lng] gönderir.
    Türkiye koordinatları: Enlem ~36-42, Boylam ~26-45
    Eğer ilk değer >35 ise muhtemelen [lat, lng] sırasındadır."""
    try:
        coords = geometry.get("coordinates", [])
        # Polygon -> [[[x,y], ...]]
        if geometry.get("type") == "Polygon" and coords:
            first_point = coords[0][0]  # İlk nokta
            # Türkiye'de enlem ~36-42, boylam ~26-45
            # Eğer ilk değer (x) > 35 ise, muhtemelen lat,lng sırasında (Leaflet formatı)
            if first_point[0] > 35:
                return "lat_lng"  # Leaflet formatı: [lat, lng]
            else:
                return "lng_lat"  # GeoJSON standardı: [lng, lat]
    except (IndexError, TypeError, KeyError):
        pass
    return "lng_lat"  # Varsayılan GeoJSON standardı


def polygon_koordinatlarini_duzelt(geometry):
    """Eğer koordinatlar [lat, lng] sırasındaysa [lng, lat]'e çevirir."""
    sira = koordinat_sirasi_tespit(geometry)
    if sira == "lat_lng":
        # Koordinatları ters çevir: [lat, lng] -> [lng, lat]
        yeni_coords = []
        for ring in geometry["coordinates"]:
            yeni_ring = [[p[1], p[0]] for p in ring]
            yeni_coords.append(yeni_ring)
        geometry = dict(geometry)
        geometry["coordinates"] = yeni_coords
    return geometry


def alan_icindeki_isleri_bul(df, geojson_data):
    """Çizilen alanın içindeki atanmamış işleri bulur. Vektörize bounding box ile hızlandırılmış versiyon."""
    if not geojson_data:
        return []

    secilen_indexler = []
    hata_olustu = False
    
    atanmamis_mask = df["Ekip"] == "Atanmamış"
    if not atanmamis_mask.any():
        return []
        
    df_atanmamis = df[atanmamis_mask]
    
    for feature in geojson_data.get("features", []):
        try:
            duzeltilmis_geom = polygon_koordinatlarini_duzelt(feature["geometry"])
            cizim = shape(duzeltilmis_geom)
            
            # Bounding box ile ön filtreleme (Hız optimizasyonu)
            minx, miny, maxx, maxy = cizim.bounds
            bbox_mask = (df_atanmamis["Boylam_Gosterim"] >= minx) & \
                        (df_atanmamis["Boylam_Gosterim"] <= maxx) & \
                        (df_atanmamis["Enlem_Gosterim"] >= miny) & \
                        (df_atanmamis["Enlem_Gosterim"] <= maxy)
                        
            adaylar = df_atanmamis[bbox_mask]
            
            if not adaylar.empty:
                import shapely
                # Vektörize Shapely point oluşturma ve alan kontrolü (Optimizasyon)
                pts = shapely.points(adaylar["Boylam_Gosterim"].values, adaylar["Enlem_Gosterim"].values)
                contains_mask = shapely.contains(cizim, pts)
                secilen_indexler.extend(adaylar[contains_mask].index.tolist())
        except Exception as e:
            hata_olustu = True
            continue

    if hata_olustu and not secilen_indexler:
        st.warning("⚠️ Çizilen alanın geometrisi okunamadı. Lütfen alanı tekrar çizmeyi deneyin.")

    return list(set(secilen_indexler))


def en_yakin_isleri_bul(df, hedef_lat, hedef_lon, adet):
    """Hedef noktaya en yakın atanmamış işleri bulur. Vektörize mesafe hesabı kullanır."""
    atanmamis = df[df["Ekip"] == "Atanmamış"].copy()
    if atanmamis.empty:
        return []

    # Numpy tabanlı vektörize haversine hesabı anında sonuç verir
    atanmamis["mesafe"] = haversine_mesafe(hedef_lat, hedef_lon, atanmamis["Enlem"], atanmamis["Boylam"])
    atanmamis = atanmamis.sort_values("mesafe")
    return atanmamis.head(adet).index.tolist()


def tablo_sutunlari_hazirla(df_kaynak):
    """Dashboard tablosu için sütunları hazırlar."""
    sonuc = pd.DataFrame()
    sonuc["Tesisat No"] = df_kaynak["Tesisat_No"].values
    sonuc["Bölge"] = df_kaynak["Bolge"].values
    sonuc["İl"] = df_kaynak["Il"].values
    sonuc["İlçe"] = df_kaynak["Ilce"].values
    sonuc["Atandığı Ekip"] = df_kaynak["Ekip"].values
    for col, baslik in [("Sayac_Seri_No", "Sayaç Seri No"), ("Adres", "Adres"),
                         ("Enlem", "Enlem"), ("Boylam", "Boylam"),
                         ("Sayac_Tip_Adi", "Sayaç Tipi"),
                         ("Tarife", "Tarife"),
                         ("AG_OG", "AG/OG"),
                         ("Saha_Aktivitesi_Yonergeleri", "Yönerge"),
                         ("Rezerv_Kwh", "Rezerv Kwh"),
                         ("Ihbar_Sekli_2", "İhbar"),
                         ("Tahakkuk_Carpani", "Tahakkuk Çarpanı")]:
        if col in df_kaynak.columns:
            sonuc[baslik] = df_kaynak[col].values
    return sonuc


def export_sutunlari_hazirla(df_kaynak):
    """Excel export (Master) için sütunları hazırlar (Bölge, İl, Sayaç Tipi hariç + Kaçak/Kwh)."""
    sonuc = pd.DataFrame()
    sonuc["Tesisat No"] = df_kaynak["Tesisat_No"].values
    sonuc["İlçe"] = df_kaynak["Ilce"].values
    sonuc["Atandığı Ekip"] = df_kaynak["Ekip"].values
    for col, baslik in [("Sayac_Seri_No", "Sayaç Seri No"), ("Adres", "Adres"),
                         ("Enlem", "Enlem"), ("Boylam", "Boylam"),
                         ("Tarife", "Tarife"),
                         ("AG_OG", "AG/OG"),
                         ("Saha_Aktivitesi_Yonergeleri", "Yönerge"),
                         ("Rezerv_Kwh", "Rezerv Kwh"),
                         ("Ihbar_Sekli_2", "İhbar"),
                         ("Tahakkuk_Carpani", "Tahakkuk Çarpanı")]:
        if col in df_kaynak.columns:
            sonuc[baslik] = df_kaynak[col].values
    sonuc["Kaçak tutuldu mu?"] = ""
    sonuc["Kwh"] = ""
    return sonuc


def ekip_export_sutunlari_hazirla(df_kaynak):
    """Ekiplerin kendi sayfaları için sütunları hazırlar (Bölge, İl, Sayaç Tipi hariç + Kaçak)."""
    sonuc = pd.DataFrame()
    sonuc["Tesisat No"] = df_kaynak["Tesisat_No"].values
    sonuc["İlçe"] = df_kaynak["Ilce"].values
    sonuc["Atandığı Ekip"] = df_kaynak["Ekip"].values
    for col, baslik in [("Sayac_Seri_No", "Sayaç Seri No"), ("Adres", "Adres"),
                         ("Enlem", "Enlem"), ("Boylam", "Boylam"),
                         ("Tarife", "Tarife"),
                         ("AG_OG", "AG/OG"),
                         ("Saha_Aktivitesi_Yonergeleri", "Yönerge"),
                         ("Rezerv_Kwh", "Rezerv Kwh"),
                         ("Ihbar_Sekli_2", "İhbar"),
                         ("Tahakkuk_Carpani", "Tahakkuk Çarpanı")]:
        if col in df_kaynak.columns:
            sonuc[baslik] = df_kaynak[col].values
    sonuc["Kaçak tutuldu mu?"] = ""
    return sonuc


def excel_export(df):
    """Atama sonuçlarını Excel dosyasına yazar ve BytesIO döner."""
    try:
        output = BytesIO()
        # Atanmış ve atanmamış verileri ayır
        atanmis_df = df[df["Ekip"] != "Atanmamış"]
        atanmamis_df = df[df["Ekip"] == "Atanmamış"]

        if atanmis_df.empty and atanmamis_df.empty:
            return None

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # 1. Master sayfası — Sadece atanmış işler (Tüm sütunlar dahil)
            if not atanmis_df.empty:
                master = export_sutunlari_hazirla(atanmis_df)
                master.to_excel(writer, sheet_name="Master", index=False)

            # 2. Atanmayan İşler sayfası (Bölge, İl, Sayaç Tipi hariç)
            if not atanmamis_df.empty:
                atanmayan = export_sutunlari_hazirla(atanmamis_df)
                # Atanmayan işlerde Kaçak/Kwh sütunları gereksiz - kaldır
                atanmayan = atanmayan.drop(columns=["Kaçak tutuldu mu?", "Kwh"], errors="ignore")
                atanmayan.to_excel(writer, sheet_name="Atanmayan İşler", index=False)

            # 3. Her ekip için ayrı sayfa (Kaçak var, Kwh yok)
            if not atanmis_df.empty:
                atanmis_ekipler = sorted(atanmis_df["Ekip"].unique())
                for ekip in atanmis_ekipler:
                    ekip_verisi = atanmis_df[atanmis_df["Ekip"] == ekip]
                    ekip_df = ekip_export_sutunlari_hazirla(ekip_verisi)
                    ekip_df.to_excel(writer, sheet_name=ekip, index=False)

            # 4. Koordinatsız İşler
            if st.session_state.df_koordinatsiz is not None and not st.session_state.df_koordinatsiz.empty:
                koor_df = export_sutunlari_hazirla(st.session_state.df_koordinatsiz)
                koor_df = koor_df.drop(columns=["Kaçak tutuldu mu?", "Kwh", "Enlem", "Boylam", "Atandığı Ekip"], errors="ignore")
                koor_df.to_excel(writer, sheet_name="Koordinatsız İşler", index=False)

        output.seek(0)
        return output
    except Exception as e:
        st.error(f"❌ Excel dosyası oluşturulurken hata oluştu: {str(e)}")
        return None


# ============================================================
# SESSION STATE BAŞLATMA
# ============================================================
if "df" not in st.session_state:
    st.session_state.df = None
if "df_koordinatsiz" not in st.session_state:
    st.session_state.df_koordinatsiz = None
if "son_cizim" not in st.session_state:
    st.session_state.son_cizim = None
if "ekip_sayisi" not in st.session_state:
    st.session_state.ekip_sayisi = 5
if "arama_noktasi" not in st.session_state:
    st.session_state.arama_noktasi = None


# ============================================================
# CUSTOM CSS
# ============================================================
st.markdown("""
<style>
    .stApp { background-color: #f8f9fa; }
    div[data-testid="stMetric"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 15px 20px;
        border-radius: 12px;
        color: white !important;
        box-shadow: 0 4px 15px rgba(102,126,234,0.3);
    }
    div[data-testid="stMetric"] label { color: rgba(255,255,255,0.85) !important; }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] { color: white !important; font-weight: 700; }
    .baslik-kutu {
        background: linear-gradient(135deg, #0f0c29, #302b63, #24243e);
        padding: 20px 30px;
        border-radius: 14px;
        margin-bottom: 20px;
        color: white;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .baslik-kutu .baslik-metin {
        text-align: left;
    }
    .baslik-kutu .baslik-metin h1 {
        margin: 0;
    }
    .baslik-kutu .baslik-metin p {
        margin: 5px 0 0 0;
        opacity: 0.85;
    }
    .baslik-kutu .baslik-logo img {
        height: 60px;
        border-radius: 8px;
        background: white;
        padding: 4px 8px;
    }
    section[data-testid="stSidebar"] > div { background: linear-gradient(180deg, #1a1a2e 0%, #16213e 100%); }
    section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span, section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 { color: #e0e0e0 !important; }
    /* Soru işareti (Tooltip) ikonlarını belirginleştir (Beyaz ve Parlak) */
    .stTooltipIcon svg,
    [data-testid="stTooltipIcon"] svg,
    [data-testid="stTooltipHoverTarget"] svg {
        color: #ffffff !important;
        fill: #ffffff !important;
        stroke: #ffffff !important;
        opacity: 1 !important;
        width: 1.5rem !important;
        height: 1.5rem !important;
        filter: drop-shadow(0px 0px 4px rgba(255,255,255,0.8));
        margin-left: 5px;
    }
    
    .stTooltipIcon:hover svg,
    [data-testid="stTooltipIcon"]:hover svg,
    [data-testid="stTooltipHoverTarget"]:hover svg {
        transform: scale(1.2);
        transition: 0.2s;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================
# LOGO YÜKLEME
# ============================================================
def logo_base64_oku():
    """Logo.png dosyasını base64 string olarak döner."""
    logo_yolu = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Logo.png")
    if os.path.exists(logo_yolu):
        with open(logo_yolu, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None

_logo_b64 = logo_base64_oku()


# ============================================================
# ANA ARAYÜZ
# ============================================================
_logo_html = f'<div class="baslik-logo"><img src="data:image/png;base64,{_logo_b64}" alt="Logo"></div>' if _logo_b64 else ""
st.markdown(
    f'<div class="baslik-kutu">'
    f'<div class="baslik-metin"><h1>🗺️ Ekip İş Planlama Dashboard</h1><p>Kaçak Ekipleri — Saha Planlama</p></div>'
    f'{_logo_html}'
    f'</div>',
    unsafe_allow_html=True
)

# --- SIDEBAR ---
with st.sidebar:
    st.header("📊 Kontrol Paneli")

    # İstatistikler ve Ekip Dağılımı
    if st.session_state.df is not None:
        df = st.session_state.df
        toplam = len(df)
        atanmis = len(df[df["Ekip"] != "Atanmamış"])
        kalan = toplam - atanmis

        st.subheader("📈 İş Durumu")
        st.metric("Toplam İş", toplam)
        st.metric("Atanmış", atanmis)
        st.metric("Kalan", kalan)

        if toplam > 0:
            oran = atanmis / toplam
            st.progress(oran, text=f"%{oran*100:.1f} tamamlandı")

        st.divider()

        # Ekip bazlı dağılım
        ekip_renkleri = aktif_ekipler(st.session_state.ekip_sayisi)
        st.subheader("👥 Ekip Dağılımı")
        ekip_sayilari = df[df["Ekip"] != "Atanmamış"]["Ekip"].value_counts()
        if not ekip_sayilari.empty:
            for ekip_adi, sayi in ekip_sayilari.items():
                renk = ekip_renkleri.get(ekip_adi, "#333")
                st.markdown(f'<span style="color:{renk};font-weight:bold;">● {ekip_adi}:</span> {sayi} iş', unsafe_allow_html=True)
        else:
            st.info("Henüz atama yapılmadı.")

    st.divider()

    # Ekip sayısı seçimi
    st.subheader("👥 Ekip Sayısı", help="Bölgenizdeki ekip sayısını giriniz")
    st.session_state.ekip_sayisi = st.number_input(
        "Kaç ekip olacak?", min_value=1, max_value=20,
        value=st.session_state.ekip_sayisi, step=1, key="ekip_sayisi_input"
    )

    st.divider()

    # Excel yükleme
    st.subheader("📁 Bekleyen İşleri Ekle", help="Kaçak Kontrol Detayda Bölge seçilip ardından Saha aktivite durum kodu: P ve H seçilerek bekleyen işleri excel formatında indirip buraya yükleyin")
    uploaded_files = st.file_uploader("Excel dosyası seçin", type=["xlsx", "xls"], key="excel_upload", accept_multiple_files=True)

    current_file_hashes = [f"{f.name}_{f.size}" for f in uploaded_files] if uploaded_files else []
    
    if current_file_hashes != st.session_state.get("last_file_hashes", []):
        st.session_state["last_file_hashes"] = current_file_hashes
        if uploaded_files:
            with st.spinner("Excel dosyaları okunuyor ve koordinatlar temizleniyor..."):
                tum_dfler = []
                tum_koordinatsiz = []
                tum_istatistikler = {"takas": 0, "gecersiz": 0}
                hatalar = []

                for dosya in uploaded_files:
                    df_yeni, istatistik, k_df, hata = excel_oku(dosya)
                    if hata:
                        hatalar.append(f"**{dosya.name}**: {hata}")
                    elif df_yeni is not None:
                        tum_dfler.append(df_yeni)
                        if k_df is not None and not k_df.empty:
                            tum_koordinatsiz.append(k_df)
                        tum_istatistikler["takas"] += istatistik.get("takas", 0)
                        tum_istatistikler["gecersiz"] += istatistik.get("gecersiz", 0)

                if hatalar:
                    st.session_state.upload_errors = hatalar
                else:
                    st.session_state.upload_errors = []
                
                if tum_dfler:
                    birlestirilmis_df = pd.concat(tum_dfler, ignore_index=True)
                    st.session_state.df = birlestirilmis_df
                    
                    if tum_koordinatsiz:
                        st.session_state.df_koordinatsiz = pd.concat(tum_koordinatsiz, ignore_index=True)
                    else:
                        st.session_state.df_koordinatsiz = None

                    st.session_state.son_cizim = None
                    st.session_state.upload_success_msg = f"✅ **{len(birlestirilmis_df)}** adet iş başarıyla yüklendi!"
                    st.session_state.upload_stats = tum_istatistikler
                    st.rerun()
        else:
            st.session_state.df = None
            st.session_state.df_koordinatsiz = None
            st.session_state.upload_success_msg = None
            st.session_state.upload_errors = []
            st.session_state.upload_stats = {}
            st.rerun()

    # Eğer sayfa yenilendiyse ve mesajlar varsa göster
    if st.session_state.get("upload_errors"):
        for h in st.session_state.upload_errors:
            st.error(h)
    if st.session_state.get("upload_success_msg"):
        st.success(st.session_state.upload_success_msg)
        stats = st.session_state.get("upload_stats", {})
        if stats.get("takas", 0) > 0:
            st.info(f"🔄 {stats['takas']} satırda Enlem↔Boylam takası yapıldı")
        if stats.get("gecersiz", 0) > 0:
            st.warning(f"⚠️ {stats['gecersiz']} satır geçersiz koordinat nedeniyle çıkarıldı")




# ============================================================
# ANA İÇERİK
# ============================================================
if st.session_state.df is None:
    st.info("👈 Lütfen sol panelden bir **Kaçak Kontrol KST** Excel dosyası yükleyin.")
    st.stop()

df = st.session_state.df
ekip_renkleri = aktif_ekipler(st.session_state.ekip_sayisi)
ekip_secenekleri = list(ekip_renkleri.keys())

# --- ARAMA FONKSİYONU (Callback) ---
def hizmet_ara_callback():
    """Arama butonuna tıklandığında veya Enter'a basıldığında çalışır."""
    if "arama_no" in st.session_state and st.session_state.arama_no.strip():
        arama_temiz = st.session_state.arama_no.strip().lstrip('0') or '0'
        df_temp = st.session_state.df
        eslesen = df_temp[df_temp["Tesisat_No"].str.lstrip('0') == arama_temiz]
        
        if not eslesen.empty:
            st.session_state.arama_noktasi = (eslesen.iloc[0]["Enlem"], eslesen.iloc[0]["Boylam"])
        else:
            st.session_state.arama_noktasi = None
    else:
        st.session_state.arama_noktasi = None

# --- HARİTA ---
st.subheader("🗺️ Saha Haritası")
st.caption("Haritada dikdörtgen/çokgen çizerek iş seçimi yapın. Sağ üstten ekip katmanlarını açıp kapatabilirsiniz.")

ag_og_secim = st.radio("AG/OG Filtresi", ["Tümü", "AG", "OG"], horizontal=True, key="ag_og_filtre")
tarife_listesi = ["Tümü"] + sorted([str(t) for t in df["Tarife"].unique() if str(t).strip() != ""])
tarife_secim = st.selectbox("Tarife Filtresi", tarife_listesi, key="tarife_filtre")

# Rezerv Kwh filter conditional
def parse_kwh(val):
    try:
        return float(str(val).replace(",", "."))
    except (ValueError, TypeError):
        return 0.0

rezerv_secim = "Tümü"
rezerv_deger = 0.0

if "Rezerv_Kwh" in df.columns:
    has_rezerv_data = True
else:
    has_rezerv_data = False

if has_rezerv_data:
    col_rez1, col_rez2 = st.columns([1, 1])
    with col_rez1:
        rezerv_secim = st.radio("Rezerv Kwh Filtresi", ["Tümü", "Değer Üzeri"], horizontal=True, key="rezerv_filtre")
    with col_rez2:
        if rezerv_secim == "Değer Üzeri":
            rezerv_deger = st.number_input("Filtrelenecek Minimum Kwh Değeri", min_value=0.0, value=300.0, step=10.0, key="rezerv_deger_input")

df_harita = df.copy()

# AG/OG Filtresi
if ag_og_secim == "AG":
    df_harita = df_harita[df_harita["AG_OG"].str.upper() == "AG"]
elif ag_og_secim == "OG":
    df_harita = df_harita[df_harita["AG_OG"].str.upper() == "OG"]

# Tarife Filtresi
if tarife_secim != "Tümü":
    df_harita = df_harita[df_harita["Tarife"] == tarife_secim]

# Rezerv Kwh Filtresi
if rezerv_secim == "Değer Üzeri":
    df_harita["Rezerv_Kwh_Num"] = df_harita["Rezerv_Kwh"].apply(parse_kwh)
    df_harita = df_harita[df_harita["Rezerv_Kwh_Num"] >= rezerv_deger]
    df_harita = df_harita.drop(columns=["Rezerv_Kwh_Num"])

st.markdown(f"**Haritada Gözüken İş Sayısı:** {len(df_harita)}")

harita = harita_olustur(df_harita, ekip_renkleri, st.session_state.arama_noktasi, goster_rezerv_kwh=(rezerv_secim == "Değer Üzeri"))
harita_sonuc = st_folium(harita, width=None, height=550, key="harita_ana",
                          returned_objects=["all_drawings"])

# Çizim verilerini yakala
if harita_sonuc is not None:
    drawings = harita_sonuc.get("all_drawings")
    if drawings and len(drawings) > 0:
        son_drawing = drawings[-1]
        if "geometry" in son_drawing:
            geojson = {"type": "FeatureCollection", "features": [son_drawing]}
            st.session_state.son_cizim = geojson
        else:
            st.session_state.son_cizim = None
    else:
        st.session_state.son_cizim = None

# --- KOORDİNATSIZ İŞLER ---
if st.session_state.df_koordinatsiz is not None and not st.session_state.df_koordinatsiz.empty:
    with st.expander(f"⚠️ Koordinatsız İşler ({len(st.session_state.df_koordinatsiz)} adet)", expanded=False):
        st.caption("Aşağıdaki işlerin koordinat bilgisi eksik olduğu için haritada gösterilemiyor.")
        koor_gosterim = st.session_state.df_koordinatsiz.copy()
        if "Tesisat_No" in koor_gosterim.columns:
            st.dataframe(koor_gosterim, hide_index=True, use_container_width=True)

# --- MANUEL ATAMA ---
st.subheader("✏️ Manuel İş Atama")

col_m1, col_m2, col_m3 = st.columns([2, 1, 1])

with col_m1:
    if st.session_state.son_cizim:
        secilen_idx = alan_icindeki_isleri_bul(df_harita, st.session_state.son_cizim)
        st.success(f"📌 Çizilen alanda **{len(secilen_idx)}** atanmamış iş bulundu.")
    else:
        secilen_idx = []
        st.info("Haritada bir alan çizin, ardından ekip seçip atayın.")

with col_m2:
    hedef_ekip_manuel = st.selectbox("Hedef Ekip", ekip_secenekleri, key="manuel_ekip")

with col_m3:
    st.write("")
    st.write("")
    if st.button("🎯 Seçili İşleri Ekip'e Ata", key="btn_manuel", type="primary",
                 disabled=(len(secilen_idx) == 0)):
        df.loc[secilen_idx, "Ekip"] = hedef_ekip_manuel
        st.session_state.df = df
        st.session_state.son_cizim = None
        st.success(f"✅ {len(secilen_idx)} iş **{hedef_ekip_manuel}**'e atandı!")
        st.rerun()

st.divider()

# --- ARAMA ARAYÜZÜ ---
st.subheader("🔍 Hizmet Noktası Ara")
col_s1, col_s2 = st.columns([3, 1])
with col_s1:
    st.text_input("Hizmet Noktası No yazın ve haritada gösterin",
                  placeholder="Örn: 3350000", key="arama_no", 
                  label_visibility="collapsed", on_change=hizmet_ara_callback)
with col_s2:
    st.button("📍 Haritada Göster", key="btn_ara", on_click=hizmet_ara_callback, type="secondary", use_container_width=True)

st.write("")

# --- OTOMATİK ATAMA ---
st.subheader("🤖 Otomatik Atama")
st.caption("Üstteki arama çubuğuna yazdığınız hizmet noktasına en yakın işleri seçilen ekibe atar.")

col_a1, col_a2 = st.columns([1, 1])

with col_a1:
    is_sayisi = st.number_input("Atanacak İş Sayısı", min_value=1, max_value=500,
                                 value=25, step=5, key="is_sayisi")

with col_a2:
    hedef_ekip_oto = st.selectbox("Hedef Ekip", ekip_secenekleri, key="oto_ekip")

if st.button("🚀 Otomatik Ata", key="btn_oto", type="primary"):
    val = st.session_state.get("arama_no", "").strip()
    if not val:
        st.error("❌ Önce üstteki arama çubuğuna bir hizmet noktası numarası girin.")
    else:
        arama_temiz = val.lstrip('0')
        if not arama_temiz: arama_temiz = "0"
        
        eslesen = df[df["Tesisat_No"].str.lstrip('0') == arama_temiz]
        
        if eslesen.empty:
            st.error(f"❌ '{val}' numaralı hizmet noktası bulunamadı!")
        else:
            hedef_lat = eslesen.iloc[0]["Enlem"]
            hedef_lon = eslesen.iloc[0]["Boylam"]
            yakin_idx = en_yakin_isleri_bul(df_harita, hedef_lat, hedef_lon, is_sayisi)
            if not yakin_idx:
                st.warning("⚠️ Atanacak atanmamış iş bulunamadı.")
            else:
                df.loc[yakin_idx, "Ekip"] = hedef_ekip_oto
                st.session_state.df = df
                st.success(f"✅ {len(yakin_idx)} iş **{hedef_ekip_oto}**'e otomatik atandı!")
                st.rerun()

st.divider()

# --- ATAMA SIFIRLA ---
col_r1, col_r2 = st.columns([1, 3])
with col_r1:
    if st.button("🗑️ Tüm Atamaları Sıfırla", key="btn_sifirla", type="secondary"):
        df["Ekip"] = "Atanmamış"
        st.session_state.df = df
        st.session_state.son_cizim = None
        st.rerun()

with col_r2:
    sifirla_cols = st.columns([1, 1])
    with sifirla_cols[0]:
        sifirla_ekip = st.selectbox("Ekip seçin", ekip_secenekleri, key="sifirla_ekip")
    with sifirla_cols[1]:
        st.write("")
        st.write("")
        if st.button(f"🔄 {sifirla_ekip} Atamasını Sıfırla", key="btn_ekip_sifirla"):
            df.loc[df["Ekip"] == sifirla_ekip, "Ekip"] = "Atanmamış"
            st.session_state.df = df
            st.rerun()

st.divider()

# --- ATANMIŞ İŞLER TABLOSU ---
st.subheader("📋 Atanmış İşler Tablosu")
atanmis_df = df[df["Ekip"] != "Atanmamış"]

if not atanmis_df.empty:
    tablo = tablo_sutunlari_hazirla(atanmis_df)
    st.dataframe(tablo, width='stretch', hide_index=True)
else:
    st.info("Henüz ekiplere atanmış iş yok.")

st.divider()

# --- EXCEL EXPORT ---
st.subheader("📥 Excel'e Aktar")

if not atanmis_df.empty:
    zaman = datetime.now().strftime("%Y%m%d_%H%M%S")
    dosya_adi = f"İşler_{zaman}.xlsx"
    excel_data = excel_export(df)
    if excel_data:
        st.download_button(
            label="📥 Excel'e Aktar (İndir)", data=excel_data,
            file_name=dosya_adi,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", key="btn_export"
        )
    else:
        st.error("❌ Excel dosyası hazırlanamadı.")
else:
    st.warning("⚠️ Henüz ekiplere atanmış iş yok. Atama yapıldıktan sonra Excel indirilebilir.")



# --- ALT BİLGİ ---
st.markdown("---")
st.caption("🔧 Ekip İş Planlama Dashboard | Kaçak Saha Operasyonları Müdürlüğü")
