import pandas as pd
import argparse
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

def find_non_empty_ranges(df):
    ranges = []
    current_range = None
    for i, row in df.iterrows():
        for j, value in enumerate(row):
            if pd.notna(value):
                if current_range is None:
                    current_range = (i, i, j, j)
                else:
                    current_range = (current_range[0], i, current_range[2], j)
        if i == df.index[-1] and current_range is not None:
            ranges.append(current_range)
    return ranges

def export_to_pdf(df, output_filename):
    pdf = canvas.Canvas(output_filename, pagesize=letter)

    cell_width = 100
    cell_height = 20
    x_offset = 50
    y_offset = 750

    current_x = x_offset
    current_y = y_offset

    for i, row in enumerate(df.iterrows()):
        for j, value in enumerate(row[1]):
            if pd.notna(value):
                x = current_x
                y = current_y
                cell_text = str(value)
                if len(cell_text) > cell_width // 8:
                    cell_text = cell_text[:cell_width // 8 - 3] + '...'
                pdf.drawString(x + 2, y + 2, cell_text)
                current_x += cell_width

                if current_x + cell_width > pdf._pagesize[0]:
                    current_x = x_offset
                    current_y -= cell_height

                    if current_y - cell_height < 50:
                        pdf.showPage()
                        current_x = x_offset
                        current_y = pdf._pagesize[1] - cell_height

        current_x = x_offset
        current_y -= cell_height

        if current_y - cell_height < 50:
            pdf.showPage()
            current_x = x_offset
            current_y = pdf._pagesize[1] - cell_height

    pdf.save()

def main():
    parser = argparse.ArgumentParser(description="Exporte les plages non vides d'un onglet Excel vers des fichiers PDF.")
    parser.add_argument("-f", "--file", required=True, help="Adresse du fichier Excel (.xls*)")
    parser.add_argument("-w", "--worksheet", required=True, help="Nom de l'onglet")
    args = parser.parse_args()

    try:
        df = pd.read_excel(args.file, sheet_name=args.worksheet, header=None)
        non_empty_ranges = find_non_empty_ranges(df)
        if not non_empty_ranges:
            raise ValueError("Aucune plage non vide trouvée dans l'onglet spécifié.")

        for idx, range_coords in enumerate(non_empty_ranges):
            start_row, end_row, start_col, end_col = range_coords
            non_empty_range = df.iloc[start_row:end_row + 1, start_col:end_col + 1]
            export_to_pdf(non_empty_range, f"output_{idx + 1}.pdf")

        print("Exportation terminée. Voir les fichiers 'output_1.pdf', 'output_2.pdf', etc.")
    except FileNotFoundError:
        print(f"Fichier introuvable : {args.file}")
    except ValueError as e:
        print(e)

if __name__ == "__main__":
    main()
