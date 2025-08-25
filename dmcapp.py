import os
import pandas as pd
import geopandas as gpd
import folium
from flask import Flask, render_template_string, request

# --- Bestanden ---
EXCEL_PATH = "DMCMap.xlsx"
GEOJSON_PATH = "ne_110m_admin_0_countries.geojson"

CATEGORIES = ["Leisure", "Premium Leisure", "Business", "MICE", "FIT", "Groups"]
category_colors = {"Leisure":"#4F46E5","Premium Leisure":"#7C3AED","Business":"#DC2626",
                   "MICE":"#EA580C","FIT":"#D97706","Groups":"#0891B2"}
COLOR_NOT_IN_EXCEL="#64748B"
COLOR_IN_EXCEL_DEFAULT="#10B981"
COLOR_BOTH = "#FBBF24"

# --- Data processing ---
def process_dataframe(df, is_uk_table=False):
    df.columns = [c.split(".")[0] for c in df.columns]
    for c in range(2,2+len(CATEGORIES)):
        if c >= len(df.columns): df[df.columns[c-1]]=False
        df.iloc[:,c] = df.iloc[:,c].fillna(False).astype(bool)
    df = df.assign(Land=df.iloc[:,1].astype(str).str.split(",")).explode("Land")
    df["Land"] = df["Land"].str.normalize("NFKD").str.encode("ascii","ignore").str.decode("utf-8").str.replace(r"\s+"," ",regex=True).str.strip()
    if is_uk_table:
        uk_mapping = {"United States":"United States of America","Alaska":"United States of America","Hawaii":"United States of America",
                      "UK":"United Kingdom","United Kingdom":"United Kingdom","South-Africa":"South Africa","Cape Town":"South Africa",
                      "Hong Kong S.A.R.":"Hong Kong","Macao S.A.R":"Macau","Myanmar (Burma)":"Myanmar","Ivory Coast":"Côte d'Ivoire",
                      "DR Congo":"Democratic Republic of the Congo","Congo":"Republic of the Congo","Bosnia":"Bosnia and Herzegovina",
                      "Herzegovina":"Bosnia and Herzegovina","Galapagos":"Ecuador","Zanzibar":"Tanzania","Rodrigues":"Mauritius",
                      "Borneo":"Malaysia","Caribbean":"Jamaica","Tibet":"China","Greenland":"Denmark","Korea":"South Korea",
                      "Australie":"Australia","Argentina":"Argentina","Australië":"Australia","Brazil":"Brazil","Canada":"Canada",
                      "China":"China","Egypt":"Egypt","France":"France","Germany":"Germany","India":"India","Italy":"Italy",
                      "Japan":"Japan","Spain":"Spain"}
        df["Land"]=df["Land"].replace(uk_mapping)
    else:
        nlbe_mapping = { 
            "Duitsland":"Germany","Nederland":"Netherlands","Frankrijk":"France","Spanje":"Spain","Italië":"Italy","België":"Belgium",
            "Zuid-Afrika":"South Africa","Zuid Afrika":"South Africa","Verenigd Koninkrijk":"United Kingdom","UK":"United Kingdom",
            "USA":"United States of America","Marokko":"Morocco","Portugal":"Portugal","Griekenland":"Greece","Turkije":"Turkey",
            "Oostenrijk":"Austria","Zwitserland":"Switzerland","Polen":"Poland","Tsjechië":"Czech Republic","Tsjechie":"Czech Republic",
            "Slowakije":"Slovakia","Hongarije":"Hungary","Roemenië":"Romania","Brazilië":"Brazil","Argentinië":"Argentina",
            "Chili":"Chile","Peru":"Peru","Mexico":"Mexico","Canada":"Canada","Australië":"Australia","Australie":"Australia",
            "Nieuw-Zeeland":"New Zealand","Nieuw Zeeland":"New Zealand","Japan":"Japan","Zuid-Korea":"South Korea","China":"China",
            "India":"India","Thailand":"Thailand","Vietnam":"Vietnam","Maleisië":"Malaysia","Singapore":"Singapore",
            "Filipijnen":"Philippines","Indonesië":"Indonesia","UAE":"United Arab Emirates","Dubai":"United Arab Emirates",
            "Saudi-Arabië":"Saudi Arabia","Egypte":"Egypt","Namibië":"Namibia","Botswana":"Botswana","Zimbabwe":"Zimbabwe",
            "Zambia":"Zambia","Mozambique":"Mozambique","Kenya":"Kenya","Tanzania":"Tanzania","Rwanda":"Rwanda","Uganda":"Uganda",
            "Ethiopië":"Ethiopia","Madagascar":"Madagascar","Mauritius":"Mauritius","Seychellen":"Seychelles","Cuba":"Cuba",
            "Dominicaanse Republiek":"Dominican Republic","Jamaica":"Jamaica","Costa Rica":"Costa Rica","Panama":"Panama",
            "Ecuador":"Ecuador","Colombia":"Colombia","Venezuela":"Venezuela","Bolivia":"Bolivia","Paraguay":"Paraguay",
            "Uruguay":"Uruguay","Guatemala":"Guatemala","Honduras":"Honduras","El Salvador":"El Salvador","Nicaragua":"Nicaragua",
            "Albanië":"Albania","Kroatië":"Croatia","Slovenië":"Slovenia","Montenegro":"Montenegro","Servië":"Serbia",
            "Bosnië":"Bosnia and Herzegovina","Bosnia":"Bosnia and Herzegovina","Herzegovina":"Bosnia and Herzegovina",
            "Noord-Macedonië":"North Macedonia","Bulgarije":"Bulgaria","Cyprus":"Cyprus","Malta":"Malta","Ierland":"Ireland",
            "IJsland":"Iceland","Noorwegen":"Norway","Zweden":"Sweden","Denemarken":"Denmark","Finland":"Finland",
            "Estland":"Estonia","Letland":"Latvia","Litouwen":"Lithuania","Oekraïne":"Ukraine","Wit-Rusland":"Belarus",
            "Rusland":"Russia","Kazachstan":"Kazakhstan","Oezbekistan":"Uzbekistan","Georgië":"Georgia","Armenië":"Armenia",
            "Azerbeidzjan":"Azerbaijan","Israël":"Israel","Libanon":"Lebanon","Jordanië":"Jordan","Syrië":"Syria",
            "Irak":"Iraq","Iran":"Iran","Qatar":"Qatar","Koeweit":"Kuwait","Bahrein":"Bahrain","Jemen":"Yemen","Oman":"Oman",
            "Bangladesh":"Bangladesh","Pakistan":"Pakistan","Sri Lanka":"Sri Lanka","Nepal":"Nepal","Bhutan":"Bhutan",
            "Malediven":"Maldives","Myanmar":"Myanmar","Laos":"Laos","Cambodja":"Cambodia","Mongolië":"Mongolia",
            "Taiwan":"Taiwan","Hong Kong":"Hong Kong","Hongkong":"Hong Kong","Macau":"Macau","Papoea-Nieuw-Guinea":"Papua New Guinea",
            "Fiji":"Fiji","Samoa":"Samoa","Tonga":"Tonga","Vanuatu":"Vanuatu","Nieuw-Caledonië":"New Caledonia",
            "Salomonseilanden":"Solomon Islands","Alaska":"United States of America","Hawaii":"United States of America",
            "Galapagos":"Ecuador","Zanzibar":"Tanzania","Rodrigues":"Mauritius","Réunion":"Réunion","Bonaire":"Bonaire",
            "Curaçao":"Curaçao","Aruba":"Aruba","Suriname":"Suriname","Guyana":"Guyana","Frans-Guyana":"French Guiana",
            "Belize":"Belize","Bahama's":"Bahamas","Barbados":"Barbados","Saint Lucia":"Saint Lucia","Dominica":"Dominica",
            "Grenada":"Grenada","Saint Vincent":"Saint Vincent and the Grenadines","Antigua":"Antigua and Barbuda",
            "Trinidad":"Trinidad and Tobago","Kaapverdië":"Cape Verde","Senegal":"Senegal","Gambia":"Gambia",
            "Guinee":"Guinea","Liberia":"Liberia","Ivoorkust":"Ivory Coast","Ghana":"Ghana","Togo":"Togo","Benin":"Benin",
            "Nigeria":"Nigeria","Kameroen":"Cameroon","Gabon":"Gabon","Congo":"Republic of the Congo",
            "DR Congo":"Democratic Republic of the Congo","Angola":"Angola","Comoren":"Comoros","Mayotte":"Mayotte",
            "Baltische Staten":"Baltic States","Benelux":"Benelux","Rwanda & Zanziba":"Rwanda","and Belize":"Belize",
            "marokko":"Morocco","nan":None}
        df["Land"]=df["Land"].replace(nlbe_mapping)
    return df

def aggregate_data(df):
    dmc_per_land={}
    for land, group in df.groupby(df.iloc[:,1]):
        key=str(land).strip().lower()
        dmc_per_land[key]={CATEGORIES[i]:group.loc[group.iloc[:,i+2]==True, group.columns[0]].tolist() for i in range(len(CATEGORIES))}
    return dmc_per_land

df_table1 = process_dataframe(pd.read_excel(EXCEL_PATH, usecols="A:H"), is_uk_table=False)
df_table2 = process_dataframe(pd.read_excel(EXCEL_PATH, usecols="O:V"), is_uk_table=True)
dmc_table1 = aggregate_data(df_table1)
dmc_table2 = aggregate_data(df_table2)
world = gpd.read_file(GEOJSON_PATH)
country_field = "ADMIN" if "ADMIN" in world.columns else "NAME"

def find_matching_key(country,dmc_dict):
    key=country.strip().lower()
    if key in dmc_dict: return key
    return None

def make_tooltip(row, dmc_dict, selected_categories, table_name, combined_dict=None, combined_table_name=None):
    land = row[country_field]
    key = find_matching_key(land, dmc_dict)
    key_combined = find_matching_key(land, combined_dict) if combined_dict else None
    lines = [f"<b>{land}</b>"]
    if key:
        for c in selected_categories:
            dmcs = dmc_dict.get(key, {}).get(c, [])
            if dmcs: lines.append(f"<span style='color:{category_colors.get(c)}'>{c} ({table_name})</span>: {', '.join(dmcs)}")
    if combined_dict and key_combined:
        for c in selected_categories:
            dmcs = combined_dict.get(key_combined, {}).get(c, [])
            if dmcs: lines.append(f"<span style='color:{category_colors.get(c)}'>{c} ({combined_table_name})</span>: {', '.join(dmcs)}")
    if len(lines) == 1: lines.append("<i>Geen DMC's</i>")
    return "<br>".join(lines)

def determine_color(row, dmc_dict, selected_categories, combined_dict=None):
    key = find_matching_key(row[country_field], dmc_dict)
    key_combined = find_matching_key(row[country_field], combined_dict) if combined_dict else None
    has_main = any(dmc_dict.get(key, {}).get(c) for c in selected_categories) if key else False
    has_combined = any(combined_dict.get(key_combined, {}).get(c) for c in selected_categories) if key_combined else False
    if has_main and has_combined: return COLOR_BOTH
    elif has_main or has_combined: return COLOR_IN_EXCEL_DEFAULT
    else: return COLOR_NOT_IN_EXCEL

def build_map(selected_categories, show_nlbe=True, show_uk=True):
    combined = show_nlbe and show_uk
    m = folium.Map(location=[20,0], zoom_start=3, tiles="cartodbdark_matter")
    for _, r in world.iterrows():
        color = determine_color(r, dmc_table1 if show_nlbe else {}, selected_categories, combined_dict=dmc_table2 if combined else None)
        style = {"fillColor": color, "color":"#334155", "weight":1, "fillOpacity":0.7}
        tooltip = folium.Tooltip(
            make_tooltip(r, dmc_table1 if show_nlbe else {}, selected_categories, "NL&BE" if show_nlbe else "UK",
                         combined_dict=dmc_table2 if combined else None,
                         combined_table_name="UK" if combined else None), sticky=True)
        folium.GeoJson(r["geometry"], style_function=lambda x, style=style: style, tooltip=tooltip).add_to(m)
    # legenda
    legend_html='<div style="position: fixed; bottom: 50px; left: 50px; width: 220px; height: 220px; background-color: white; border:2px solid grey; z-index:9999; font-size:14px; padding:10px;"><b>Legenda</b><br>'
    for cat,color in category_colors.items():
        legend_html+=f'<i style="background:{color};width:12px;height:12px;display:inline-block"></i> {cat}<br>'
    legend_html+=f'<i style="background:{COLOR_BOTH};width:12px;height:12px;display:inline-block"></i> NL&BE + UK<br>'
    legend_html+=f'<i style="background:{COLOR_NOT_IN_EXCEL};width:12px;height:12px;display:inline-block"></i> Geen DMC</div>'
    m.get_root().html.add_child(folium.Element(legend_html))
    return m._repr_html_()

# --- Flask webapp ---
app = Flask(__name__)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
<title>DMC Kaart</title>
<style>
body { margin:0; font-family:Arial,sans-serif; }
.sidebar { width:300px; background:#f8fafc; border-right:1px solid #cbd5e1; position:fixed; height:100%; padding:10px; overflow:auto; }
.main { margin-left:310px; padding:10px; }
input[type=checkbox]{ margin:5px 0; }
button { margin-top:10px; padding:6px 12px; font-weight:600; border-radius:8px; border:1px solid #cbd5e1; background:#fff; cursor:pointer; }
button:hover { background:#f1f5f9; }
</style>
</head>
<body>
<div class="sidebar">
  <form method="get">
    <h3>Tabellen</h3>
    <label><input type="checkbox" name="nlbe" value="1" {% if show_nlbe %}checked{% endif %}>NL&BE</label><br>
    <label><input type="checkbox" name="uk" value="1" {% if show_uk %}checked{% endif %}>UK</label>
    <h3>Categorieën</h3>
    {% for cat in categories %}
      <label><input type="checkbox" name="cat" value="{{cat}}" {% if cat in selected_categories %}checked{% endif %}>{{cat}}</label><br>
    {% endfor %}
    <button type="submit">Update Kaart</button>
  </form>
</div>
<div class="main">
  {{ map_html|safe }}
</div>
</body>
</html>
"""

@app.route("/", methods=["GET"])
def index():
    selected_categories = request.args.getlist("cat") or CATEGORIES
    show_nlbe = "nlbe" in request.args
    show_uk = "uk" in request.args
    map_html = build_map(selected_categories, show_nlbe, show_uk)
    return render_template_string(HTML_TEMPLATE, map_html=map_html, categories=CATEGORIES,
                                  selected_categories=selected_categories, show_nlbe=show_nlbe, show_uk=show_uk)

if __name__=="__main__":
    app.run(debug=True)
