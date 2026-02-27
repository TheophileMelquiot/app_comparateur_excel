import unidecode
import gradio as gr
import pandas as pd
import tempfile
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import numpy as np
from openpyxl.utils import get_column_letter

def auto_adjust_column_width(ws):
    for column_cells in ws.columns:
        max_length = 0
        column = column_cells[0].column

        for cell in column_cells:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass

        adjusted_width = max_length + 2
        ws.column_dimensions[get_column_letter(column)].width = adjusted_width

def compute_unique_ids(file, sheet_name, header_row):
    if file is None or sheet_name is None or header_row is None:
        return pd.DataFrame()

    try:
        # Lecture brute SANS header
        df_raw = pd.read_excel(
            file.name,
            sheet_name=sheet_name,
            header=None
        )

        header_row_index = int(header_row) - 1

        # On récupère la vraie ligne d'en-tête
        new_header = df_raw.iloc[header_row_index]

        # On coupe le dataframe à partir de la ligne suivante
        df = df_raw.iloc[header_row_index + 1:].copy()

        # On applique les bons noms de colonnes
        df.columns = new_header

        # Supprimer colonnes complètement vides
        df = df.dropna(axis=1, how='all')

        # Supprimer lignes complètement vides
        df = df.dropna(how='all')

        # Nettoyage noms de colonnes
        df.columns = df.columns.astype(str).str.strip()

        total_rows = len(df)
        unique_counts = df.nunique(dropna=True)

        result = pd.DataFrame({
            "Colonne": unique_counts.index.astype(str),
            "Nb_valeurs_uniques": unique_counts.values
        }).sort_values(by="Nb_valeurs_uniques", ascending=False)

        top5 = result.head(5).copy()
        top5["Total_lignes"] = total_rows
        top5["% unicité"] = (top5["Nb_valeurs_uniques"] / total_rows * 100).round(2)

        return top5

    except Exception as e:
        return pd.DataFrame({"Erreur": [str(e)]})

def normalize_dataframe(df):

    def clean_value(v):

        # NaN → chaîne vide
        if pd.isna(v):
            return ""

        # Numérique → float propre
        if isinstance(v, (int, float, np.number)):
            try:
                return float(v)
            except:
                return ""

        # Texte
        v = str(v)

        v = v.replace("\xa0", " ")
        v = v.replace("\n", " ")
        v = v.replace("\r", " ")

        v = " ".join(v.split())  # supprime espaces multiples
        v = v.strip().lower()

        return v

    for col in df.columns:
        if col != "merge_key":
            df[col] = df[col].apply(clean_value)

    return df

def safe_compare(v1, v2, decimals =4 ):

    # Deux NaN
    if pd.isna(v1) and pd.isna(v2):
        return True

    # Un seul NaN
    if pd.isna(v1) or pd.isna(v2):
        return False


    # Si numérique → arrondi contrôlé
    try:
        f1 = float(v1)
        f2 = float(v2)

        return round(f1, decimals) == round(f2, decimals)

    except:
        pass

    # Sinon texte
    return str(v1).strip() == str(v2).strip()

    # Comparaison texte normalisée
    return str(v1).strip() == str(v2).strip()
# 🔁 Normalisation des noms de colonnes
def normalize_colname(col):
    return unidecode.unidecode(str(col)).strip().lower()

# ⚙️ Fonction de comparaison et export
def compare_excels(file1, sheet1, header1, col1, file2, sheet2, header2, col2):
    try:
        df1_raw = pd.read_excel(file1.name, sheet_name=sheet1, header=int(header1)-1, decimal=',')
        df2_raw = pd.read_excel(file2.name, sheet_name=sheet2, header=int(header2)-1, decimal=',')


        df1_raw.columns = [normalize_colname(c) for c in df1_raw.columns]
        df2_raw.columns = [normalize_colname(c) for c in df2_raw.columns]

        col1_norm = normalize_colname(col1)
        col2_norm = normalize_colname(col2)

        df1_raw = df1_raw.rename(columns={col1_norm: "merge_key"})
        df2_raw = df2_raw.rename(columns={col2_norm: "merge_key"})
        df1 = df1_raw.copy()
        df2 = df2_raw.copy()

        # 🔥 TRIER POUR GARANTIR LE MÊME ORDRE
        df1 = df1.sort_values(by=["merge_key"] + [c for c in df1.columns if c != "merge_key"])
        df2 = df2.sort_values(by=["merge_key"] + [c for c in df2.columns if c != "merge_key"])

        # 🔥 AJOUTER IDENTIFIANT TECHNIQUE D'OCCURRENCE
        df1["occurrence"] = df1.groupby("merge_key").cumcount()
        df2["occurrence"] = df2.groupby("merge_key").cumcount()

        # 🔥 NORMALISER
        df1 = normalize_dataframe(df1)
        df2 = normalize_dataframe(df2) 
        merged = pd.merge(
    df1,
    df2,
    on=["merge_key", "occurrence"],
    how="outer",
    suffixes=('_1', '_2')
)

        common_cols = [
            c for c in df1.columns
            if c not in ["merge_key", "occurrence"] and c in df2.columns
        ]        

        result = merged[['merge_key']].copy()
        for col in common_cols:
            result[f"{col}_1"] = merged[f"{col}_1"]
            result[f"{col}_2"] = merged[f"{col}_2"]

        wb = Workbook()
        ws = wb.active

        red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
        error =0

        for r in dataframe_to_rows(result, index=False, header=True):
            ws.append(r)

        # itérer avec un index qui avance de 2 colonnes par paire
        for idx, col in enumerate(common_cols):
            # position Excel : col A = 1 ; on a merge_key en 1, donc la 1ère paire commence en 2
            col1_idx = 2 + idx * 2
            col2_idx = col1_idx + 1

            # itérer sur les lignes du DataFrame merged (par index numérique)
            for row_idx in range(len(merged)):
                excel_row = row_idx + 2  # +2 = 1ère ligne header + 1-based excel rows

                val1 = merged.at[row_idx, f"{col}_1"]
                val2 = merged.at[row_idx, f"{col}_2"]

                # comparer en string (sécurisé) — ou utiliser safe_compare si tu préfères tolérance numérique
                if not safe_compare(val1, val2):
                    ws.cell(row=excel_row, column=col1_idx).fill = red_fill
                    ws.cell(row=excel_row, column=col2_idx).fill = red_fill
                    error += 1

                if not safe_compare(val1, val2):
                    print("DIFF DEBUG:", val1, type(val1), val2, type(val2))

        print(error)

        auto_adjust_column_width(ws)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            wb.save(tmp.name)
            temp_path = tmp.name

        return temp_path, ""

    except Exception as e:
        return None, f"❌ Erreur : {e}"

# 🎛️ Interface Gradio
with gr.Blocks() as app:
    gr.Markdown("## 🔍 Comparaison de fichiers Excel sur colonne clé avec export Excel coloré")

    with gr.Row():
        file1 = gr.File(label="📁 Fichier Excel 1")
        file2 = gr.File(label="📁 Fichier Excel 2")

    with gr.Row():
        sheet1 = gr.Dropdown(label="Onglet Fichier 1", choices=[], value=None)
        sheet2 = gr.Dropdown(label="Onglet Fichier 2", choices=[], value=None)

    with gr.Row():
        header1 = gr.Number(label="Ligne des en-têtes Fichier 1", value=1)
        header2 = gr.Number(label="Ligne des en-têtes Fichier 2", value=1)

    with gr.Row():
        preview1 = gr.Dataframe(label="Aperçu Fichier 1")
        preview2 = gr.Dataframe(label="Aperçu Fichier 2")

    with gr.Row():
        btn_unique1 = gr.Button("🔎 Analyser Unique ID Fichier 1")
        btn_unique2 = gr.Button("🔎 Analyser Unique ID Fichier 2")

    with gr.Row():
        unique_result1 = gr.Dataframe(label="Top 5 colonnes les plus uniques - Fichier 1")
        unique_result2 = gr.Dataframe(label="Top 5 colonnes les plus uniques - Fichier 2") 

    with gr.Row():
        btn_cols1 = gr.Button("🔃 Charger colonnes Fichier 1")
        dropdown_cols1 = gr.Dropdown(label="Colonne clé Fichier 1", choices=[], interactive=True)

    with gr.Row():
        btn_cols2 = gr.Button("🔃 Charger colonnes Fichier 2")
        dropdown_cols2 = gr.Dropdown(label="Colonne clé Fichier 2", choices=[], interactive=True)

    with gr.Row():
        btn_compare = gr.Button("📤 Comparer et exporter Excel")
        output_file = gr.File(label="📄 Fichier Excel comparé")
        error_msg = gr.Textbox(label="Message d'erreur", interactive=False)

    # 📑 Fonctions auxiliaires
    def get_sheet_names(file):
        if file is None:
            return gr.update(choices=[], value=None)
        try:
            xls = pd.ExcelFile(file.name)
            sheets = xls.sheet_names
            default = sheets[0] if sheets else None
            return gr.update(choices=sheets, value=default)
        except:
            return gr.update(choices=[], value=None)

    def read_excel(file, sheet_name, header_row):
        if file is None or sheet_name is None or header_row is None:
            return pd.DataFrame()
        try:
            df = pd.read_excel(file.name, sheet_name=sheet_name, header=int(header_row)-1)
            return df.head(10)
        except:
            return pd.DataFrame()

    def get_columns(file, sheet_name, header_row):
        if file is None or sheet_name is None or header_row is None:
            return gr.update(choices=[], value=None)
        try:
            df = pd.read_excel(file.name, sheet_name=sheet_name, header=int(header_row)-1)
            cols = list(df.columns)
            return gr.update(choices=cols, value=cols[0] if cols else None)
        except:
            return gr.update(choices=[], value=None)

    # ⚙️ Connexion des événements Gradio
    file1.change(fn=get_sheet_names, inputs=file1, outputs=sheet1)
    file2.change(fn=get_sheet_names, inputs=file2, outputs=sheet2)

    sheet1.change(fn=read_excel, inputs=[file1, sheet1, header1], outputs=preview1)
    sheet2.change(fn=read_excel, inputs=[file2, sheet2, header2], outputs=preview2)

    btn_cols1.click(fn=get_columns, inputs=[file1, sheet1, header1], outputs=dropdown_cols1)
    btn_cols2.click(fn=get_columns, inputs=[file2, sheet2, header2], outputs=dropdown_cols2)

    btn_unique1.click(
        fn=compute_unique_ids,
        inputs=[file1, sheet1, header1],
        outputs=unique_result1
    )

    btn_unique2.click(
        fn=compute_unique_ids,
        inputs=[file2, sheet2, header2],
        outputs=unique_result2
    )

    btn_compare.click(
        fn=compare_excels,
        inputs=[file1, sheet1, header1, dropdown_cols1, file2, sheet2, header2, dropdown_cols2],
        outputs=[output_file, error_msg]
    )
#
# 🚀 Lancement de l'app
app.launch(inbrowser=True)




