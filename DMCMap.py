import pandas as pd
import geopandas as gpd
import folium
import streamlit as st
from streamlit_folium import st_folium
import os

st.set_page_config(page_title="Interatieve Wereldkaart met DMC Data", layout="wide")

# --- Pad naar bestanden ---
EXCEL_PATH = r"DMCMap.xlsx"
SHAPE_PATH = r"ne_110m_admin_0_countries.geojson"

# --- Data laden (alleen Land + kolommen C–H, en alleen rijen t/m 41) ---
df = pd.read_excel(EXCEL_PATH, usecols="B:H", nrows=41)

# --- Kolomnamen (categorieën) exact zoals ze in je Excel staan ---
CATEGORIES = ["Leisure", "Premium Leisure", "Business", "MICE", "FIT", "Groups"]

# Zorg dat ontbrekende categorie-kolommen niet crashen
for c in CATEGORIES:
    if c not in df.columns:
        df[c] = False

# --- NL -> EN mapping (vul aan wanneer nodig) ---
land_mapping = {
    "Duitsland": "Germany",
    "Nederland": "Netherlands",
    "Frankrijk": "France",
    "Spanje": "Spain",
    "Italië": "Italy",
    "België": "Belgium",
    "Zuid-Afrika": "South Africa",
    "Verenigd Koninkrijk": "United Kingdom",
    "USA": "United States of America",
    "Marokko": "Morocco",
    "Portugal": "Portugal",
    "Griekenland": "Greece",
    "Turkije": "Turkey",
    "Oostenrijk": "Austria",
    "Zwitserland": "Switzerland",
    "Polen": "Poland",
    "Tsjechië": "Czech Republic",
    "Hongarije": "Hungary",
    "Roemenië": "Romania",
    "Brazilië": "Brazil",
    "Argentinië": "Argentina",
    "Chili": "Chile",
    "Peru": "Peru",
    "Mexico": "Mexico",
    "Canada": "Canada",
    "Australië": "Australia",
    "Nieuw-Zeeland": "New Zealand",
    "Japan": "Japan",
    "Zuid-Korea": "South Korea",
    "China": "China",
    "India": "India",
    "Thailand": "Thailand",
    "Vietnam": "Vietnam",
    "Maleisië": "Malaysia",
    "Singapore": "Singapore",
    "Filipijnen": "Philippines",
    "Indonesië": "Indonesia",
    "UAE": "United Arab Emirates",
    "Saudi-Arabië": "Saudi Arabia",
    "Egypte": "Egypt",
}

# --- Landen splitsen op komma, trimmen en vertalen NL -> EN ---
df = df.assign(Land=df["Land"].astype(str).str.split(","))  # kan ook "Namibië, Zuid-Afrika" etc.
df = df.explode("Land")
df["Land"] = df["Land"].str.strip().replace(land_mapping)

# --- Categorieën booleans netjes maken (True/NaN -> bool) ---
for c in CATEGORIES:
    # Als Excel True/NaN had, dit maakt het betrouwbaar boolean
    df[c] = df[c].fillna(False)
    # Als ergens strings "True"/"False" zitten:
    if df[c].dtype == object:
        df[c] = df[c].astype(str).str.lower().map({"true": True, "false": False}).fillna(False)

# --- Aggregeren per land: ANY over rijen zodat elk land één rij wordt ---
agg = df.groupby("Land", dropna=False)[CATEGORIES].any().reset_index()

# --- Wereldkaart inladen ---
if not os.path.exists(SHAPE_PATH):
    st.error(f"Shapefile niet gevonden: {SHAPE_PATH}")
    st.stop()

world = gpd.read_file(SHAPE_PATH)

# Natural Earth gebruikt meestal 'ADMIN' als landennaam
country_field = "ADMIN" if "ADMIN" in world.columns else ("NAME" if "NAME" in world.columns else None)
if country_field is None:
    st.error("Kon geen landennaam-kolom vinden in shapefile (verwacht 'ADMIN' of 'NAME').")
    st.stop()

# --- Merge: één rij per wereld-land + (optioneel) matching uit Excel ---
world = world.merge(agg, how="left", left_on=country_field, right_on="Land")

# --- Sidebar checkboxes (standaard UIT) ---
st.sidebar.header("Categorieën")
selected = [c for c in CATEGORIES if st.sidebar.checkbox(c, value=False)]

st.title("Interatieve Wereldkaart met DMC Data")

# --- Kleuren per categorie ---
category_colors = {
    "Leisure": "#3B82F6",           # blauw
    "Premium Leisure": "#A855F7",   # paars
    "Business": "#EF4444",          # rood
    "MICE": "#F59E0B",              # oranje
    "FIT": "#EAB308",               # geel
    "Groups": "#06B6D4",            # cyaan
}
COLOR_NOT_IN_EXCEL = "lightgrey"
COLOR_IN_EXCEL_DEFAULT = "green"

def determine_color(row):
    # Land komt niet uit Excel (merge gaf NaN in categorie-kolommen) -> grijs
    if row[CATEGORIES].isna().all():
        return COLOR_NOT_IN_EXCEL

    # Land staat wel in Excel
    if selected:
        # Geef kleur van de eerste geselecteerde categorie die True is
        for c in selected:
            val = row.get(c, False)
            if pd.notna(val) and bool(val):
                return category_colors.get(c, COLOR_IN_EXCEL_DEFAULT)
        # Geen geselecteerde categorie True -> groen
        return COLOR_IN_EXCEL_DEFAULT
    else:
        # Geen selectie -> groen
        return COLOR_IN_EXCEL_DEFAULT

# Bepaal kleur per land
world["__color__"] = world.apply(determine_color, axis=1)

# --- Folium map maken ---
m = folium.Map(location=[20, 0], zoom_start=3, tiles="cartodbpositron")

for _, r in world.iterrows():
    # geometry tekenen
    folium.GeoJson(
        r["geometry"],
        style_function=lambda x, color=r["__color__"]: {
            "fillColor": color,
            "color": "black",
            "weight": 0.5,
            "fillOpacity": 0.7,
        },
        tooltip=f"{r[country_field]}",
    ).add_to(m)

# --- Legenda toevoegen ---
legend_items = "".join(
    f'<div style="margin:2px 0;"><span style="display:inline-block;width:12px;height:12px;background:{col};margin-right:6px;border:1px solid #444"></span>{cat}</div>'
    for cat, col in category_colors.items()
)
legend_html = f"""
<div style="
     position: fixed;
     bottom: 5px; left: 5px;
     z-index: 9999;
     background: white;
     padding: 10px 12px;
     border: 1px solid #ccc;
     border-radius: 6px;
     box-shadow: 0 1px 4px rgba(0,0,0,.2);
     font-size: 14px;
     color: black;">
  <div style="font-weight:600;margin-bottom:6px;">Legenda</div>
  <div style="margin:2px 0;">
    <span style="display:inline-block;width:12px;height:12px;background:{COLOR_NOT_IN_EXCEL};margin-right:6px;border:1px solid #444"></span>
    Geen DMC
  </div>
  <div style="margin:2px 0;">
    <span style="display:inline-block;width:12px;height:12px;background:{COLOR_IN_EXCEL_DEFAULT};margin-right:6px;border:1px solid #444"></span>
    DMC (geen selectie)
  </div>
  {legend_items}
</div>
"""

m.get_root().html.add_child(folium.Element(legend_html))

# --- Tonen in Streamlit ---
st_data = st_folium(m, width=1920, height=1080)


