import json


def code_cell(source: str):
    return {
        "cell_type": "code",
        "execution_count": None,
        "metadata": {},
        "outputs": [],
        "source": source
    }

cells = []

cell0 = '''# Cellule 0 – installations et imports essentiels
!pip -q install "folium>=0.14.0" "openpyxl>=3.1.0"

import re
import unicodedata
from pathlib import Path

import pandas as pd
import folium
from folium.plugins import MarkerCluster, FeatureGroupSubGroup, Fullscreen
import ipywidgets as widgets
from IPython.display import display, clear_output
'''

cell1 = '''# Cellule 1 – configuration générale, constantes et helpers
pd.options.display.max_rows = 10

BASE_DIR = Path("/mnt/data")
EXCEL_FILE = BASE_DIR / "BASE RETV2.xlsx"
SHEET_NAME = 0
OUTPUT_HTML = Path("map_pdv_retv2.html")

COL_LAT = "LATITUDE"
COL_LON = "LONGITUDE"
COL_SOLDE = "SOLDE"
COL_NOM_RETAILER = "NOM RETAILER"
COL_NUM_RETAILER = "N° RETAILER"
COL_NOMPRENOM = "NOM & PRENOM"
COL_CONTACT = "CONTACT"
COL_CLASSIFICATION = "CLASSIFICATION"
COL_NOM_DISTRIBUTEUR = "NOM DISTRIBUTEUR"

BALANCE_BINS = [
    (float("-inf"), 5000,        "#d73027"),
    (5000,        20000,         "#fc8d59"),
    (20000,       50000,         "#fee08b"),
    (50000,       100000,        "#d9ef8b"),
    (100000,      200000,        "#91cf60"),
    (200000,      float("inf"),  "#1a9850"),
]

lat_min, lat_max = -6.0, 12.0
lon_min, lon_max = -10.0, 0.0

CLASSIFICATION_ORDER = ["650", "AGENCE MOOV", "PREMIUM", "TOP 8500 PDV", "ESPACE SERVICE", "N/A"]
CLASSIFICATION_CANONICAL = {opt.upper(): opt for opt in CLASSIFICATION_ORDER if opt != "N/A"}
DEFAULT_CENTER = (5.345317, -4.024429)
FALLBACK_LABEL = "Autres"


def normalize_label(label: str) -> str:
    if label is None:
        return ""
    s = unicodedata.normalize("NFKD", str(label))
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"[^A-Za-z0-9]+", "", s)
    return s.upper()


def find_col(df_cols, wanted: str):
    target = normalize_label(wanted)
    for c in df_cols:
        if normalize_label(c) == target:
            return c
    raise KeyError(f"Colonne introuvable: {wanted}")


def to_float_coord(x):
    if pd.isna(x):
        return None
    s = str(x).strip().replace("\u00A0", " ")
    s = s.replace(",", ".")
    s = re.sub(r"[^\d\.\-]+", "", s)
    try:
        return float(s)
    except ValueError:
        return None


def to_number_amount(x):
    if pd.isna(x):
        return 0.0
    s = str(x).strip().replace("\u00A0", " ")
    s = re.sub(r"[^\d\.\-]", "", s)
    if s in ("", "-", "."):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def color_for_balance(amount: float) -> str:
    for lower, upper, color in BALANCE_BINS:
        if lower < amount <= upper:
            return color
    return "#999999"


def bin_label(lower, upper):
    def fmt(x):
        if x == float("-inf"):
            return "−∞"
        if x == float("inf"):
            return "+∞"
        return f"{int(x):,}".replace(",", " ")
    if lower == float("-inf"):
        return f"< {fmt(upper)} FCFA"
    if upper == float("inf"):
        return f"> {fmt(lower)} FCFA"
    return f"{fmt(lower)} — {fmt(upper)} FCFA"


def label_for_amount(amount):
    for lower, upper, _ in BALANCE_BINS:
        if lower < amount <= upper:
            return bin_label(lower, upper)
    return bin_label(BALANCE_BINS[-1][0], BALANCE_BINS[-1][1])


def build_legend_html(added_count: int) -> str:
    def fmt(x):
        if x == float("-inf"):
            return "−∞"
        if x == float("inf"):
            return "+∞"
        return f"{int(x):,}".replace(",", " ")
    items = []
    for lower, upper, color in BALANCE_BINS:
        label = f"{fmt(lower)} — {fmt(upper)} FCFA"
        items.append(f"""
        <div style="display:flex; align-items:center; margin-bottom:4px;">
            <span style="display:inline-block; width:14px; height:14px; background:{color}; border:1px solid #333; margin-right:6px;"></span>
            <span style="font-size:12px;">{label}</span>
        </div>
        """)
    return f"""
    <div style="
        position: fixed; bottom: 20px; left: 20px; z-index: 9999;
        background: rgba(255,255,255,0.95);
        padding: 10px 12px; border: 1px solid #ccc; border-radius: 6px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.15);">
      <div style="font-weight:600; margin-bottom:6px;">Légende – Solde</div>
      {''.join(items)}
      <div style="font-size:11px; color:#555; margin-top:6px;">
        Points: {added_count} | Zoom pour plus de détails
      </div>
    </div>
    """


def prepare_join_key(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    text = re.sub(r"\s+", "", text)
    return text.upper()


def normalize_classification_value(value):
    if pd.isna(value):
        return "N/A"
    text = str(value).strip()
    if not text:
        return "N/A"
    upper = text.upper()
    if upper in CLASSIFICATION_CANONICAL:
        return CLASSIFICATION_CANONICAL[upper]
    if upper in {"N/A", "NA"} or "NON DISPON" in upper:
        return "N/A"
    return "N/A"
'''

cell2 = '''# Cellule 2 – chargement de la base Excel et du dataset PDV
load_messages = []
df_base_raw = pd.DataFrame()
df_pdv_raw = pd.DataFrame()

if not EXCEL_FILE.exists():
    load_messages.append(f"⚠️ Fichier introuvable : {EXCEL_FILE}")
else:
    try:
        df_base_raw = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, engine="openpyxl")
        df_pdv_raw = df_base_raw.copy()
        print(f"Base chargée : {df_base_raw.shape[0]} lignes, {df_base_raw.shape[1]} colonnes.")
    except Exception as exc:
        load_messages.append(f"⚠️ Erreur lors de la lecture de '{EXCEL_FILE.name}' : {exc}")
        df_base_raw = pd.DataFrame()
        df_pdv_raw = pd.DataFrame()
'''

cell3 = '''# Cellule 3 – jointure, nettoyage et normalisation de la classification
LAT_COL = LON_COL = SOLDE_COL = NOMR_COL = NUMR_COL = NOMPRENOM_COL = CONTACT_COL = NOM_DISTR_COL = None
CLASSIFICATION_FILTER_AVAILABLE = True

df_clean = pd.DataFrame()

if df_pdv_raw.empty:
    CLASSIFICATION_FILTER_AVAILABLE = False
    df_clean = pd.DataFrame(columns=df_pdv_raw.columns)
    print("⚠️ Aucun PDV exploitable : la carte restera vide.")
else:
    column_lookup = {}
    missing_critical = False
    for key, label in [
        ("LAT", COL_LAT),
        ("LON", COL_LON),
        ("SOLDE", COL_SOLDE),
        ("NOMR", COL_NOM_RETAILER),
        ("NUMR", COL_NUM_RETAILER),
        ("NOMPRENOM", COL_NOMPRENOM),
        ("CONTACT", COL_CONTACT),
        ("NOM_DISTR", COL_NOM_DISTRIBUTEUR),
    ]:
        try:
            column_lookup[key] = find_col(df_pdv_raw.columns, label)
        except KeyError as exc:
            load_messages.append(f"⚠️ {exc}")
            missing_critical = True

    classification_col_in_pdv = None
    try:
        classification_col_in_pdv = find_col(df_pdv_raw.columns, COL_CLASSIFICATION)
        column_lookup["CLASSIFICATION"] = classification_col_in_pdv
    except KeyError:
        column_lookup["CLASSIFICATION"] = None

    if missing_critical:
        CLASSIFICATION_FILTER_AVAILABLE = False
        df_clean = pd.DataFrame(columns=df_pdv_raw.columns)
        print("⚠️ Colonnes critiques manquantes : aucune donnée cartographique.")
    else:
        df_clean = df_pdv_raw.copy()

        lat_col = column_lookup["LAT"]
        lon_col = column_lookup["LON"]
        solde_col = column_lookup["SOLDE"]
        nomr_col = column_lookup["NOMR"]
        numr_col = column_lookup["NUMR"]
        nompren_col = column_lookup["NOMPRENOM"]
        contact_col = column_lookup["CONTACT"]
        distr_col = column_lookup["NOM_DISTR"]

        df_clean[lat_col] = df_clean[lat_col].map(to_float_coord)
        df_clean[lon_col] = df_clean[lon_col].map(to_float_coord)
        df_clean[solde_col] = df_clean[solde_col].map(to_number_amount)

        df_clean = df_clean.dropna(subset=[lat_col, lon_col]).copy()
        df_clean = df_clean[(df_clean[lat_col].between(lat_min, lat_max)) & (df_clean[lon_col].between(lon_min, lon_max))].copy()

        df_clean[nomr_col] = df_clean[nomr_col].fillna("Nom retailer inconnu")
        df_clean[numr_col] = df_clean[numr_col].fillna("N/A")
        df_clean[solde_col] = df_clean[solde_col].fillna(0)
        df_clean[nompren_col] = df_clean[nompren_col].fillna("N/A")
        df_clean[contact_col] = df_clean[contact_col].fillna("Non Disponible")
        df_clean[distr_col] = df_clean[distr_col].fillna("Non Disponible")

        classification_series = None
        if not df_base_raw.empty:
            try:
                base_join_col = find_col(df_base_raw.columns, COL_NUM_RETAILER)
                base_class_col = find_col(df_base_raw.columns, COL_CLASSIFICATION)
                base_lookup = df_base_raw[[base_join_col, base_class_col]].copy()
                base_lookup["__join_key__"] = base_lookup[base_join_col].map(prepare_join_key)
                base_lookup = base_lookup.drop_duplicates("__join_key__")
                base_lookup = base_lookup.rename(columns={base_class_col: "__classification_base__"})
                df_clean["__join_key__"] = df_clean[numr_col].map(prepare_join_key)
                df_clean = df_clean.merge(base_lookup[["__join_key__", "__classification_base__"]], on="__join_key__", how="left")
                classification_series = df_clean["__classification_base__"]
            except KeyError as exc:
                load_messages.append(f"⚠️ {exc}")

        if classification_series is None and classification_col_in_pdv:
            classification_series = df_clean[classification_col_in_pdv]

        if classification_series is None:
            CLASSIFICATION_FILTER_AVAILABLE = False
            df_clean["classification_display"] = "Non Disponible"
            df_clean["classification_filter"] = "N/A"
            load_messages.append("⚠️ Classification indisponible : filtre désactivé.")
        else:
            if not isinstance(classification_series, pd.Series):
                classification_series = pd.Series(classification_series, index=df_clean.index)
            display_values = classification_series.fillna("Non Disponible").replace("", "Non Disponible")
            filter_values = classification_series.map(normalize_classification_value).fillna("N/A")
            df_clean["classification_display"] = display_values
            df_clean["classification_filter"] = filter_values
            CLASSIFICATION_FILTER_AVAILABLE = df_clean["classification_filter"].notna().any()

        if "__classification_base__" in df_clean.columns:
            df_clean = df_clean.drop(columns=["__classification_base__"])
        if "__join_key__" in df_clean.columns:
            df_clean = df_clean.drop(columns=["__join_key__"])

        df_clean["balance_label"] = df_clean[solde_col].map(label_for_amount)

        LAT_COL = lat_col
        LON_COL = lon_col
        SOLDE_COL = solde_col
        NOMR_COL = nomr_col
        NUMR_COL = numr_col
        NOMPRENOM_COL = nompren_col
        CONTACT_COL = contact_col
        NOM_DISTR_COL = distr_col

        print(f"{len(df_clean)} PDV prêts pour la carte.")

if "classification_filter" not in df_clean.columns:
    df_clean["classification_filter"] = pd.Series(dtype=object)
if "classification_display" not in df_clean.columns:
    df_clean["classification_display"] = pd.Series(dtype=object)
if df_clean.empty:
    CLASSIFICATION_FILTER_AVAILABLE = False
'''

cell4 = '''# Cellule 4 – widgets (KPI + panneau filtre par classification)
kpi_html = widgets.HTML(
    value="<div style='font-weight:600;'>Aperçu dynamique</div><div>Initialisation...</div>",
    layout=widgets.Layout(width='280px', padding='10px', border='1px solid #d1d5db', border_radius='6px', background_color='#ffffff')
)

multi_toggle = widgets.ToggleButton(
    value=False,
    description="Multi: OFF",
    tooltip="Activer/Désactiver le mode multi-sélection",
    layout=widgets.Layout(width='140px')
)

multi_caption = widgets.HTML(
    "<span style='font-size:12px; color:#6b7280;'>Multi OFF = sélection exclusive</span>",
    layout=widgets.Layout(margin='0 0 0 8px')
)

all_button = widgets.Button(
    description="ALL",
    tooltip="Réinitialiser le filtre",
    button_style='info',
    layout=widgets.Layout(width='70px')
)

class_buttons = {
    label: widgets.ToggleButton(
        description=label,
        value=False,
        layout=widgets.Layout(width='150px')
    )
    for label in CLASSIFICATION_ORDER
}

buttons_box = widgets.Box(
    [all_button] + [class_buttons[label] for label in CLASSIFICATION_ORDER],
    layout=widgets.Layout(display='flex', flex_flow='row wrap', gap='8px', margin='6px 0 0 0')
)

selection_status = widgets.HTML(
    "<span style='font-size:12px; color:#2563eb;'>Filtre actif : ALL</span>",
    layout=widgets.Layout(margin='8px 0 0 0')
)

warning_widget = widgets.HTML(layout=widgets.Layout(margin='4px 0 4px 0'))
if load_messages:
    warning_widget.value = "<div style='font-size:12px; color:#92400e; background:#fef3c7; padding:8px; border-radius:6px;'>" + "<br/>".join(load_messages) + "</div>"

filter_children = [widgets.HTML("<div style='font-weight:600; font-size:14px;'>Filtre par classification</div>")]
if load_messages:
    filter_children.append(warning_widget)
filter_children.append(widgets.HBox([multi_toggle, multi_caption], layout=widgets.Layout(align_items='center', gap='8px')))
filter_children.append(buttons_box)
filter_children.append(selection_status)

filter_panel = widgets.VBox(
    filter_children,
    layout=widgets.Layout(width='280px', padding='10px', border='1px solid #d1d5db', border_radius='6px', background_color='#ffffff', margin='12px 0 0 0')
)

left_panel = widgets.VBox(
    [kpi_html, filter_panel],
    layout=widgets.Layout(width='300px', gap='12px')
)

map_output = widgets.Output(layout=widgets.Layout(width='100%', height='760px', border='1px solid #d1d5db', border_radius='6px'))

if df_clean.empty or not CLASSIFICATION_FILTER_AVAILABLE:
    multi_toggle.disabled = True
    all_button.disabled = True
    for btn in class_buttons.values():
        btn.disabled = True
    if df_clean.empty:
        selection_status.value = "<span style='color:#6b7280;'>Aucun PDV à filtrer.</span>"
    else:
        selection_status.value = "<span style='color:#92400e;'>Classification indisponible : filtre désactivé.</span>"
'''

cell5 = '''# Cellule 5 – état du filtre et gestionnaires d'événements
CLASSIFICATION_FILTER_AVAILABLE = bool(CLASSIFICATION_FILTER_AVAILABLE and not df_clean.empty)

def build_filter_mask(df: pd.DataFrame, selections, multi_on: bool) -> pd.Series:
    if df.empty:
        return pd.Series([], dtype=bool)
    if (not CLASSIFICATION_FILTER_AVAILABLE) or selections is None:
        return pd.Series(True, index=df.index)
    if isinstance(selections, set):
        active = [cls for cls in selections if cls]
    else:
        active = [cls for cls in selections if cls]
    if len(active) == 0:
        return pd.Series(False, index=df.index)
    return df["classification_filter"].isin(active)

active_selections = None  # None => ALL
programmatic_update = False


def ordered_selection():
    if active_selections is None:
        return []
    return [label for label in CLASSIFICATION_ORDER if label in active_selections]


def update_selection_status():
    if df_clean.empty:
        selection_status.value = "<span style='color:#6b7280;'>Aucun PDV à filtrer.</span>"
        return
    if not CLASSIFICATION_FILTER_AVAILABLE:
        selection_status.value = "<span style='color:#92400e;'>Classification indisponible : filtre désactivé.</span>"
        return
    if active_selections is None:
        selection_status.value = "<span style='font-size:12px; color:#2563eb;'>Filtre actif : ALL</span>"
    elif len(active_selections) == 0:
        selection_status.value = "<span style='font-size:12px; color:#b91c1c;'>Filtre actif : Aucun PDV</span>"
    else:
        items = ", ".join(ordered_selection())
        selection_status.value = f"<span style='font-size:12px;'>Filtre actif : {items}</span>"


def update_controls():
    global programmatic_update
    programmatic_update = True
    multi_toggle.description = f"Multi: {'ON' if multi_toggle.value else 'OFF'}"
    all_button.button_style = 'info' if active_selections is None else ''
    for label, btn in class_buttons.items():
        selected = False
        if active_selections is not None:
            selected = label in active_selections
        btn.value = selected
        btn.button_style = 'success' if selected else ''
    programmatic_update = False
    update_selection_status()


def refresh_outputs():
    if df_clean.empty:
        with map_output:
            clear_output(wait=True)
            display(widgets.HTML("<b>Aucun PDV à afficher.</b>"))
        update_kpis(df_clean)
        return
    mask = build_filter_mask(df_clean, active_selections, multi_toggle.value)
    df_filtered = df_clean[mask].copy()
    update_kpis(df_filtered)
    m = rebuild_map(df_filtered)
    with map_output:
        clear_output(wait=True)
        display(m)
    try:
        m.save(str(OUTPUT_HTML))
    except Exception:
        pass


def handle_multi_toggle(change):
    global active_selections
    if programmatic_update or change.get("name") != "value":
        return
    if not change["new"]:
        if isinstance(active_selections, set) and len(active_selections) > 1:
            for label in CLASSIFICATION_ORDER:
                if label in active_selections:
                    active_selections = {label}
                    break
    update_controls()
    refresh_outputs()


def handle_all_button(_):
    global active_selections
    if programmatic_update:
        return
    active_selections = None
    update_controls()
    refresh_outputs()


def handle_class_button(change, label):
    global active_selections
    if programmatic_update or change.get("name") != "value":
        return
    if multi_toggle.value:
        if active_selections is None or not isinstance(active_selections, set):
            active_selections = set()
        if change["new"]:
            active_selections.add(label)
        else:
            active_selections.discard(label)
            if len(active_selections) == 0:
                active_selections = set()
    else:
        if change["new"]:
            active_selections = {label}
        else:
            active_selections = None
    update_controls()
    refresh_outputs()


multi_toggle.observe(handle_multi_toggle, names="value")
all_button.on_click(handle_all_button)
for label, btn in class_buttons.items():
    btn.observe(lambda change, label=label: handle_class_button(change, label), names="value")
'''

cell6 = '''# Cellule 6 – calcul KPI dynamique

def format_int(value):
    try:
        return f"{int(round(float(value))):,}".replace(",", " ")
    except (TypeError, ValueError):
        return "0"


def format_currency(value):
    try:
        return f"{format_int(value)} FCFA"
    except (TypeError, ValueError):
        return "0 FCFA"


def update_kpis(df_filtered: pd.DataFrame):
    total = len(df_clean)
    filtered = len(df_filtered)

    solde_col = SOLDE_COL if SOLDE_COL and SOLDE_COL in df_filtered.columns else None
    if solde_col:
        sum_solde = float(df_filtered[solde_col].sum()) if filtered else 0.0
        mean_solde = float(df_filtered[solde_col].mean()) if filtered else 0.0
        median_solde = float(df_filtered[solde_col].median()) if filtered else 0.0
    else:
        sum_solde = mean_solde = median_solde = 0.0

    if filtered and CLASSIFICATION_FILTER_AVAILABLE and "classification_filter" in df_filtered.columns:
        counts = df_filtered["classification_filter"].value_counts()
    else:
        counts = pd.Series(dtype=int)

    rows = []
    if CLASSIFICATION_FILTER_AVAILABLE:
        for label in CLASSIFICATION_ORDER:
            count = int(counts.get(label, 0))
            pct = (count / filtered * 100) if filtered else 0.0
            rows.append(f"<tr><td>{label}</td><td style='text-align:right;'>{format_int(count)}</td><td style='text-align:right;'>{pct:.1f}%</td></tr>")
    else:
        rows.append("<tr><td colspan='3' style='text-align:center; color:#6b7280;'>Classification indisponible</td></tr>")

    table_html = "".join(rows)

    kpi_html.value = f"""
    <div style="font-family: 'Source Sans Pro', Arial, sans-serif; font-size: 13px; color: #1f2937;">
      <div style="font-weight: 600; font-size: 15px; margin-bottom: 8px;">Aperçu dynamique</div>
      <div style="display:flex; justify-content:space-between; margin-bottom:4px;">
        <span>Total PDV</span><span>{format_int(total)}</span>
      </div>
      <div style="display:flex; justify-content:space-between; margin-bottom:4px;">
        <span>PDV filtrés</span><span>{format_int(filtered)}</span>
      </div>
      <div style="display:flex; justify-content:space-between; margin-bottom:4px;">
        <span>Somme des soldes filtrés</span><span>{format_currency(sum_solde)}</span>
      </div>
      <div style="display:flex; justify-content:space-between; margin-bottom:4px;">
        <span>Moyenne filtrée</span><span>{format_currency(mean_solde)}</span>
      </div>
      <div style="display:flex; justify-content:space-between; margin-bottom:8px;">
        <span>Médiane filtrée</span><span>{format_currency(median_solde)}</span>
      </div>
      <div style="font-weight: 600; margin-bottom: 4px;">Répartition par classification</div>
      <table style="width:100%; border-collapse: collapse;">
        <thead>
          <tr style="text-align:left; border-bottom:1px solid #e5e7eb; font-size:12px; color:#4b5563;">
            <th>Classe</th><th style="text-align:right;">PDV</th><th style="text-align:right;">Part</th>
          </tr>
        </thead>
        <tbody>
          {table_html}
        </tbody>
      </table>
    </div>
    """
'''

cell7 = '''# Cellule 7 – reconstruction de la carte Folium filtrée

def rebuild_map(df_filtered: pd.DataFrame) -> folium.Map:
    if LAT_COL is None or LON_COL is None:
        return folium.Map(location=list(DEFAULT_CENTER), zoom_start=6, tiles='OpenStreetMap')

    if df_filtered.empty:
        center_lat, center_lon = DEFAULT_CENTER
    else:
        center_lat = float(df_filtered[LAT_COL].median())
        center_lon = float(df_filtered[LON_COL].median())

    m = folium.Map(
        location=[center_lat, center_lon],
        zoom_start=7,
        tiles="OpenStreetMap",
        control_scale=True
    )

    folium.TileLayer('CartoDB positron', name='Fond clair', show=False).add_to(m)
    folium.TileLayer('CartoDB dark_matter', name='Fond sombre', show=False).add_to(m)
    folium.TileLayer(
        tiles="https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
        attr="Esri", name="Satellite", show=False
    ).add_to(m)

    Fullscreen(position='topleft', title='Mode présentation', title_cancel='Quitter la présentation').add_to(m)

    icon_create_fn = """
    function(cluster){
      var counts = {};
      var domColor = '#3c78d8';
      var children = cluster.getAllChildMarkers();
      for (var i=0;i<children.length;i++){
        var c = (children[i].options && (children[i].options.fillColor || children[i].options.color)) || '#3c78d8';
        counts[c] = (counts[c]||0)+1;
      }
      var max = 0;
      for (var k in counts){
        if (counts[k] > max){ max = counts[k]; domColor = k; }
      }
      var count = cluster.getChildCount();
      return new L.DivIcon({
        html: '<div class="mc-dyn" style="background:'+domColor+'"><span>'+count+'</span></div>',
        className: 'marker-cluster',
        iconSize: new L.Point(40,40)
      });
    }
    """

    parent_cluster = MarkerCluster(
        name="PDV (cluster global)",
        control=False,
        icon_create_function=icon_create_fn,
        showCoverageOnHover=False,
        spiderfyOnMaxZoom=True,
        disableClusteringAtZoom=16
    ).add_to(m)

    groups = {}
    labels_in_order = []
    for lower, upper, color in BALANCE_BINS:
        label = bin_label(lower, upper)
        labels_in_order.append(label)
        fg = FeatureGroupSubGroup(parent_cluster, name=label, overlay=True, show=True)
        fg.add_to(m)
        groups[label] = {"fg": fg, "color": color}

    fallback_fg = FeatureGroupSubGroup(parent_cluster, name=FALLBACK_LABEL, overlay=True, show=True)
    fallback_fg.add_to(m)
    groups[FALLBACK_LABEL] = {"fg": fallback_fg, "color": "#999999"}

    added = 0
    solde_col = SOLDE_COL if SOLDE_COL and SOLDE_COL in df_filtered.columns else None

    for _, row in df_filtered.iterrows():
        lat = row.get(LAT_COL)
        lon = row.get(LON_COL)
        if pd.isna(lat) or pd.isna(lon):
            continue
        solde = float(row.get(solde_col, 0)) if solde_col else 0.0
        color = color_for_balance(solde)
        label = row.get("balance_label") or label_for_amount(solde)
        resolved_label = label if label in groups else FALLBACK_LABEL

        nomr = str(row.get(NOMR_COL, "N/A"))
        numr = str(row.get(NUMR_COL, "N/A"))
        nompren = str(row.get(NOMPRENOM_COL, "N/A"))
        contact = str(row.get(CONTACT_COL, "Non Disponible"))
        nom_distr = str(row.get(NOM_DISTR_COL, "Non Disponible"))
        classification_display = str(row.get("classification_display", "Non Disponible"))

        popup_html = (
            f"<b>NOM RETAILER:</b> {nomr}<br>"
            f"<b>NOM &amp; PRENOM :</b> {nompren}<br>"
            f"<b>N° RETAILER:</b> {numr}<br>"
            f"<b>Solde:</b> {format_currency(solde)}<br>"
            f"<b>CONTACT:</b> {contact}<br>"
            f"<b>CLASSIFICATION:</b> {classification_display}<br>"
            f"<b>NOM DISTRIBUTEUR:</b> {nom_distr}"
        )

        folium.CircleMarker(
            location=[lat, lon],
            radius=5,
            color=color,
            fill=True,
            fill_color=color,
            fill_opacity=0.8,
            weight=1,
            popup=folium.Popup(popup_html, max_width=350),
            tooltip=nomr
        ).add_to(groups[resolved_label]["fg"])
        added += 1

    legend_html = build_legend_html(added)
    m.get_root().html.add_child(folium.Element(legend_html))

    cluster_css = """
    <style>
      .mc-dyn {
        width: 40px; height: 40px; line-height: 40px;
        border-radius: 50%;
        border: 2px solid rgba(0,0,0,0.35);
        color: #fff; font-weight: 700; text-align: center;
        box-shadow: 0 0 0 2px rgba(255,255,255,0.4) inset;
      }
      .mc-dyn span { color: #fff; }
    </style>
    """
    m.get_root().html.add_child(folium.Element(cluster_css))

    folium.LayerControl().add_to(m)
    return m
'''

cell8 = '''# Cellule 8 – affichage final (panneau gauche + carte Folium)
app_layout = widgets.HBox(
    [left_panel, map_output],
    layout=widgets.Layout(width='100%', align_items='flex-start', gap='16px')
)

update_controls()
refresh_outputs()
display(app_layout)
'''

cells.extend([
    code_cell(cell0),
    code_cell(cell1),
    code_cell(cell2),
    code_cell(cell3),
    code_cell(cell4),
    code_cell(cell5),
    code_cell(cell6),
    code_cell(cell7),
    code_cell(cell8),
])

notebook = {
    "cells": cells,
    "metadata": {
        "kernelspec": {
            "display_name": "Python 3",
            "language": "python",
            "name": "python3"
        },
        "language_info": {
            "name": "python",
            "version": "3.10"
        }
    },
    "nbformat": 4,
    "nbformat_minor": 5
}

with open('PDV_Classification_Filter.ipynb', 'w', encoding='utf-8') as f:
    json.dump(notebook, f, indent=2)
