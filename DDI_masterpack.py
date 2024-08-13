import pandas as pd
import math
import json
import os
from decimal import Decimal, ROUND_UP, ROUND_DOWN, ROUND_HALF_UP
from xlsxwriter import Workbook

def calculate_ideal_size(days, last_sales_day, tvu, remove_product):
    if any(val == 'PRODUCTO SIN CARGA' for val in [days, last_sales_day, tvu, remove_product]):
        return '0'
    return str(math.trunc((last_sales_day / days) * (tvu - remove_product)))

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

resultArray = []

def get_valid_int(prompt):
    while True:
        try:
            return int(input(prompt))
        except ValueError:
            print("Entrada no válida. Por favor, ingrese un número entero.")

def get_valid_float(prompt):
    while True:
        try:
            return float(input(prompt))
        except ValueError:
            print("Entrada no válida. Por favor, ingrese un número válido.")

print('============== INGRESO DE USUARIO ==============')
days = 28
last_sales_day = get_valid_int('Ingrese la venta de los últimos días: ')
tvu = 60
remove_product = 7
ddi_master_pack = get_valid_float('Ingrese el DDI del masterpack: ')
ddi_packqty = get_valid_float('Ingrese el DDI del packqty: ')
ddi_innerpack = get_valid_float('Ingrese el DDI del innerpack: ')
stock = get_valid_int('Ingrese la cantidad de stock: ')

print('============== INGRESO DE ARCHIVO ==============')
archive = input("Ingrese el nombre del archivo: ") + '.xlsx'
sheet_name = input("Ingrese el nombre de la hoja: ")
df = pd.read_excel('archivos/' + archive, sheet_name=sheet_name, header=1, index_col=None)

# Aplicar safe_round a las columnas relevantes
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
    
    # Guardar en formato JSON
    output_file_json = os.path.join(output_dir, "result.json")
    with open(output_file_json, 'w') as f:
        resultArray.append(results)
        json.dump(resultArray, f, indent=4)  # Cambiado para guardar una lista de resultados
    
    print(f'Archivo JSON guardado correctamente en la ruta {output_dir} con el nombre result.json')

    ######### ARCHIVO EXCEL ############
    # Guardar en formato Excel
    output_file_excel = os.path.join(output_dir, "result.xlsx")
    
    # Leer el archivo existente si existe
    if os.path.exists(output_file_excel):
        with pd.ExcelFile(output_file_excel) as xls:
            existing_df = pd.read_excel(xls, sheet_name='Results', header=1)
            existing_data = existing_df.iloc[2:].reset_index(drop=True)  # Leer datos existentes excluyendo título y encabezados
    else:
        existing_data = pd.DataFrame()  # Crear un DataFrame vacío si no hay datos existentes

    # Crear un DataFrame para el nuevo registro
    new_record_df = pd.DataFrame([results])

    # Concatenar los datos existentes con el nuevo registro
    all_data_df = pd.concat([existing_data, new_record_df], ignore_index=True)

    # Guardar el DataFrame actualizado en el archivo Excel
    with pd.ExcelWriter(output_file_excel, engine='xlsxwriter') as writer:
        worksheet = writer.book.add_worksheet('Results')
        
        title_format = writer.book.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Arial',
        })
        worksheet.merge_range('A1:M1', 'RESULTADOS VALIDACIÓN', title_format)
        
        header_format = writer.book.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#DCE6F1',
        })
        
        # Escribir encabezados
        headers = list(results.keys())
        worksheet.write_row('A2', headers, header_format)
        
        # Escribir datos
        for row_num, record in enumerate(all_data_df.values.tolist(), start=3):
            worksheet.write_row(f'A{row_num}', record)
        
        # Ajustar el ancho de las columnas
        for i, header in enumerate(headers):
            worksheet.set_column(i, i, len(header) + 2)
    
    print(f'Archivo Excel guardado correctamente en la ruta {output_dir} con el nombre result.xlsx')