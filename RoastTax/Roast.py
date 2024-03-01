import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import colors
from openpyxl.styles import Font, Color
from openpyxl.styles import Border, Side, Alignment
from datetime import datetime
import os
# Initialize the Excel workbook
wb = Workbook()
wb.remove(wb.active)  # Remove the default sheet created

# Define your file paths (assuming they are named as 'January 2024.xlsx', 'February 2024.xlsx', etc.)
file_paths = ['2024/Januar 2024.xlsx', '2024/Februar 2024.xlsx', '2024/März 2024.xlsx', '2024/April 2024.xlsx', 
              '2024/Mai 2024.xlsx', '2024/Juni 2024.xlsx', '2024/Juli 2024.xlsx', '2024/August 2024.xlsx', 
              '2024/September 2024.xlsx', '2024/Oktober 2024.xlsx', '2024/November 2024.xlsx', '2024/Dezember 2024.xlsx']

roast_names = ['LR MH','LR KEB', 'LR Fortezza', 'LR DKB', 'LR BM', 'LR JT', 'LR Kafrika', 'LR CS', 'BIO']

id = 1

monthly_totals = {}
for file_path in file_paths:
    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        df_reversed = df.iloc[::-1]
        new_rows = []
        weight_before_sum = 0
        weight_after_sum = 0
        last_date = None
        monthly_before_sum = 0
        monthly_after_sum = 0
        monthly_before_sum_check = 0
        monthly_after_sum_check = 0

        for index, row in df_reversed.iterrows():
            
            date = pd.to_datetime(row.iloc[1]).date()
            roast_title = row.iloc[3]
            weight_before = row.iloc[4]
            weight_after = row.iloc[5]
            monthly_before_sum += weight_before
            monthly_after_sum += weight_after
            name_found = ''
            text = ''

            for name in roast_names:
                if name in roast_title:
                    name_found = name
                    text = roast_title.split(name)[1].split(':')[0].strip() if ':' in roast_title else ''
                    break
            if not name_found:
                text = roast_title.split(':')[0]

            if date != last_date and last_date is not None:
                new_rows[-1].update({'Rohk.': f"{weight_before_sum:.1f}", 'Röstk.': f"{weight_after_sum:.1f}"})
                weight_before_sum = 0
                weight_after_sum = 0
            
            if date != last_date:
                last_date = date
            else:
                date = ''
            
            weight_before_sum += weight_before
            weight_after_sum += weight_after
            if date != '':
                date_obj = datetime.strptime(str(date), "%Y-%m-%d")
                date = date_obj.strftime("%d.%m.%Y")
            new_rows.append({
                'ID': id,
                'Datum': str(date),
                'Name': name_found,
                'Sorte': text,
                'Rohk..': f"{weight_before:.1f}",
                'Rohk.':'',
                'Röstk..': f"{weight_after:.1f}",
                'Röstk.':''
            
            })
            id += 1
            monthly_before_sum_check += weight_before
            monthly_after_sum_check += weight_after
        if new_rows:
            new_rows[-1].update({'Rohk.': f"{weight_before_sum:.1f}", 'Röstk.': f"{weight_after_sum:.1f}"})


        new_df = pd.DataFrame(new_rows)
        title = file_path.replace('.xlsx', '').replace('2024/', '')
        ws = wb.create_sheet(title=title)  # Create a sheet named after the month
        ws.freeze_panes = 'A3' 
    # Insert the total row at the top
        ws.merge_cells('A1:C1')
        cell = ws['A1']
        cell.value = f'Röstbuch Schneid {title}'
        cell.alignment = Alignment(horizontal='center')
        font=Font(size=12, bold=True)
        cell.font=font
        ws['D1'] = 'Gesamt:'
        ws['E1'] = f"{monthly_before_sum:.1f}"
        ws['G1'] = f"{monthly_after_sum:.1f}"
        ws['F1'] = f"{monthly_before_sum_check:.1f}"
        ws['H1'] = f"{monthly_after_sum_check:.1f}"

        monthly_totals[title] = (monthly_before_sum, monthly_after_sum)

        thin_border = Border(bottom=Side(style='thin'))

        for col, header in enumerate(new_df.columns, start=1):
            ws.cell(row=2, column=col, value=header)

        max_length = {}
        for rowIndex, row in enumerate(new_df.itertuples(), start=3):  # Adjust start for new row index
            for colIndex, cell_value in enumerate(row[1:], start=1):  # row is a tuple including index at first position
                cell = ws.cell(row=rowIndex, column=colIndex, value=cell_value)
                if colIndex == 2 and getattr(row, 'Datum') != '':
                    for i in range(1, len(new_df.columns) + 1):
                        ws.cell(row=rowIndex - 1, column=i).border = thin_border
                if cell_value and isinstance(cell_value, str):
                    max_length[colIndex] = max(max_length.get(colIndex, 0), len(cell_value))

        for i, width in max_length.items():
            ws.column_dimensions[get_column_letter(i)].width = width + 2
        print(f'Blatt für {title} populiert\n')
    else:
    # File does not exist
        title = file_path.replace('.xlsx', '').replace('2024/', '')
        print(f"Datei für {title} kann nicht gefunden werden")
summary_ws = wb.create_sheet(title="Röstk.Ges.2024", index=0)
summary_ws.append(["Monat", "Rohk", "Röstk"])

max_length = [0, 0, 0]
for month, totals in monthly_totals.items():
    summary_ws.append([month, f"{totals[0]:.1f}", f"{totals[1]:.1f}"])
    max_length[0] = max(max_length[0], len(month))
    max_length[1] = max(max_length[1], len(str(totals[0])))
    max_length[2] = max(max_length[2], len(str(totals[1])))

for i, width in enumerate(max_length, start=1):
    summary_ws.column_dimensions[get_column_letter(i)].width = width + 2
print('Zusammenfassungsblatt erstellt\n')
wb.save('Röstbuch 2024 automatisiert.xlsx')
print('Erfolg!')