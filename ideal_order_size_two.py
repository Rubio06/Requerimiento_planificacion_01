import pandas as pd
import math
import json
import os
from decimal import Decimal, ROUND_UP, ROUND_DOWN, ROUND_HALF_UP
from xlsxwriter import Workbook

# Funciones de cálculo y evaluación (se mantienen iguales)
def calculate_ideal_size(days, last_sales_day, tvu, remove_product):
    if any(val == 'PRODUCTO SIN CARGA' for val in [days, last_sales_day, tvu, remove_product]):
        return 0
    return int(math.trunc((last_sales_day / days) * (tvu - remove_product)))

def evaluate_masterpack(ddi_master_pack, tvu, stock, last_sales_day):
    if ddi_master_pack in ['PRODUCTO SIN VENTA', 'PRODUCTO SIN CARGA']:
        return ddi_master_pack.upper()
    
    try:
        ddi_master_pack = float(ddi_master_pack)
    except ValueError:
        return 'PRODUCTO SIN CARGA'

    if ddi_master_pack <= tvu:
        return 'OK'
    elif stock == 0 and last_sales_day == 0:
        return 'PRODUCTO SIN CARGA'
    elif stock > 0 and last_sales_day == 0:
        return 'PRODUCTO SIN VENTA'
    else:
        return 'BAJAR MASTERPACK'

def evaluate_packqty(ddi_packqty, tvu, stock, last_sales_day):
    if ddi_packqty in ['PRODUCTO SIN VENTA', 'PRODUCTO SIN CARGA']:
        return ddi_packqty.upper()
    
    try:
        ddi_packqty = float(ddi_packqty)
    except ValueError:
        return 'PRODUCTO SIN CARGA'

    if ddi_packqty <= tvu:
        return 'OK'
    elif stock == 0 and last_sales_day == 0:
        return 'PRODUCTO SIN CARGA'
    elif stock > 0 and last_sales_day == 0:
        return 'PRODUCTO SIN VENTA'
    else:
        return 'BAJAR PACKQTY'

def evaluate_innerpack(ddi_innerpack, tvu, stock, last_sales_day):
    if ddi_innerpack in ['PRODUCTO SIN VENTA', 'PRODUCTO SIN CARGA']:
        return ddi_innerpack.upper()
    
    try:
        ddi_innerpack = float(ddi_innerpack)
    except ValueError:
        return 'PRODUCTO SIN CARGA'

    if ddi_innerpack <= tvu:
        return 'OK'
    elif stock == 0 and last_sales_day == 0:
        return 'PRODUCTO SIN CARGA'
    elif stock > 0 and last_sales_day == 0:
        return 'PRODUCTO SIN VENTA'
    else:
        return 'BAJAR INNERPACK'

def safe_round(value):
    try:
        value_data = Decimal(str(value))
        round_up_value = value_data.quantize(Decimal('0.00'), rounding=ROUND_UP)
        round_down_value = value_data.quantize(Decimal('0.00'), rounding=ROUND_DOWN)
        
        if round_up_value == value_data:
            return round_up_value
        elif round_down_value == value_data:
            return round_down_value
        else:
            return value_data.quantize(Decimal('0.00'), rounding=ROUND_HALF_UP)
    except:
        return value.upper()

def clean_value(value):
    if pd.isna(value) or value in [float('inf'), float('-inf')]:
        return ''
    return value

resultArray = []

print('============== INGRESO DE USUARIO ==============')
days = 28
last_sales_day = int(input('Ingrese la venta de los últimos días: '))
tvu = 60
remove_product = 7
ddi_master_pack = input('Ingrese el DDI del masterpack: ')
ddi_packqty = input('Ingrese el DDI del packqty: ')
ddi_innerpack = input('Ingrese el DDI del innerpack: ')
stock = int(input('Ingrese la cantidad de stock: '))

print('============== INGRESO DE ARCHIVO ==============')
archive = input("Ingrese el nombre del archivo: ") + '.xlsx'
sheet_name = input("Ingrese el nombre de la hoja: ")
df = pd.read_excel('archivos/' + archive, sheet_name=sheet_name, header=1, index_col=None)

df['DDI_Masterpack'] = df['DDI_Masterpack'].apply(safe_round)
df['DDI_Packqty'] = df['DDI_Packqty'].apply(safe_round)
df['DDI_Innerpack'] = df['DDI_Innerpack'].apply(safe_round)

print('============== ARCHIVO EXCEL ==============')
found = False
for index, row in df.iterrows():
    row_ddi_masterpack = row['DDI_Masterpack']
    row_vta_ultdiascant = int(row['VtaUltDiasCant']) if pd.notna(row['VtaUltDiasCant']) else None
    row_ddi_packqty = row['DDI_Packqty']
    row_ddi_innerpack = row['DDI_Innerpack']
    row_stock = int(row['Stock']) if pd.notna(row['Stock']) else None

    ddi_masterpack_match = (row_ddi_masterpack == safe_round(ddi_master_pack))
    vta_ultdiascant_match = (row_vta_ultdiascant == last_sales_day)
    ddi_packqty_match = (row_ddi_packqty == safe_round(ddi_packqty))
    ddi_innerpack_match = (row_ddi_innerpack == safe_round(ddi_innerpack))
    stock_match = (row_stock == stock)

    if ddi_masterpack_match and vta_ultdiascant_match and ddi_packqty_match and ddi_innerpack_match and stock_match:
        found = True
        break

if found:
    ideal_size = calculate_ideal_size(days, last_sales_day, tvu, remove_product)
    validate_masterpack = evaluate_masterpack(ddi_master_pack, tvu, stock, last_sales_day)
    validate_packqty = evaluate_packqty(ddi_packqty, tvu, stock, last_sales_day)
    validate_innerpack = evaluate_innerpack(ddi_innerpack, tvu, stock, last_sales_day)
    
    results = {
        'days': days,
        'stock': stock,
        'VtaUltDiasCant': last_sales_day,
        'tvu': tvu,
        'remove_product': remove_product,
        'ddi_master_pack': ddi_master_pack,
        'ddi_packqty': ddi_packqty,
        'ddi_innerpack': ddi_innerpack,
        'validate_packqty': validate_packqty,
        'validate_innerpack': validate_innerpack,
        'validate_masterpack': validate_masterpack,
        'ideal_size': ideal_size
    }
    

    output_dir = 'json_results'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    output_file_json = os.path.join(output_dir, "result.json")
    resultArray.append(results)
    with open(output_file_json, 'w') as f:
        json.dump(resultArray, f, indent=4)
    
    print(f'Archivo JSON guardado correctamente en la ruta {output_dir} con el nombre result.json')

    ######### ARCHIVO EXCEL ############
    output_file_excel = os.path.join(output_dir, "result.xlsx")
    
    if os.path.exists(output_file_excel):
        with pd.ExcelFile(output_file_excel) as xls:
            existing_df = pd.read_excel(xls, sheet_name='Results', header=2, index_col=None)
            existing_data = existing_df.copy()
    else:
        existing_data = pd.DataFrame(columns=[
            'days', 'stock', 'VtaUltDiasCant', 'tvu', 'remove_product', 
            'ddi_master_pack', 'ddi_packqty', 'ddi_innerpack', 
            'validate_packqty', 'validate_innerpack', 'validate_masterpack', 
            'ideal_size'
        ])

    new_record_df = pd.DataFrame([results])
    
    # Usar map en lugar de applymap
    existing_data = existing_data.applymap(clean_value)
    new_record_df = new_record_df.applymap(clean_value)
    
    # Convertir tipos de datos para el DataFrame de nuevos registros
    new_record_df = new_record_df.astype({
        'days': 'int',
        'stock': 'int',
        'VtaUltDiasCant': 'int',
        'tvu': 'int',
        'remove_product': 'int',
        'ideal_size': 'int',
        'ddi_master_pack': 'str',
        'ddi_packqty': 'str',
        'ddi_innerpack': 'str',
        'validate_packqty': 'str',
        'validate_innerpack': 'str',
        'validate_masterpack': 'str'
    })

    with pd.ExcelWriter(output_file_excel, engine='xlsxwriter') as writer:
        worksheet = writer.book.add_worksheet('Results')

        title_format = writer.book.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Arial',
        })
        worksheet.merge_range('A1:L1', 'RESULTADOS VALIDACIÓN', title_format)
        
        header_format = writer.book.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'border': 1,
            'align': 'center',
            'bg_color': '#fcf3cf'

        })

        field_titles = {
            'days': 'DÍAS',
            'stock': 'STOCK',
            'VtaUltDiasCant': 'VTA_ULT_DÍAS',
            'tvu': 'TVU',
            'remove_product': 'REM_PROD',
            'ddi_master_pack': 'DDI_MASTERPACK',
            'ddi_packqty': 'DDI_PACKQTY',
            'ddi_innerpack': 'DDI_INNERPACK',
            'validate_packqty': 'VALIDAR_PACKQTY',
            'validate_innerpack': 'VALIDAR_INNERPACK',
            'validate_masterpack': 'VALIDAR_MASTERPACK',
            'ideal_size': 'TAMAÑO_IDEAL'
        }

        for col_num, (col_name, col_title) in enumerate(field_titles.items()):
            worksheet.write(1, col_num, col_title, header_format)

        start_row = 2
        for index, row in new_record_df.iterrows():
            for col_num, value in enumerate(row):
                worksheet.write(start_row, col_num, value)
            start_row += 1

    print(f'Archivo Excel guardado correctamente en la ruta {output_dir} con el nombre result.xlsx')
else:
    print('No se encontraron coincidencias.')