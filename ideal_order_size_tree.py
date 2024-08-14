import pandas as pd
import math
import json
import os
from decimal import Decimal, ROUND_HALF_UP

# Funciones de cálculo y evaluación
def calculate_ideal_size(days, last_sales_day, tvu, remove_product):
    if days is None or last_sales_day is None:
        return '0'
    
    if any(val in ['PRODUCTO SIN CARGA', 'PRODUCTO SIN VENTA'] for val in [days, last_sales_day, tvu, remove_product]):
        return '0'
    
    if days == 0:
        return '0'

    return str(math.trunc((last_sales_day / days) * (tvu - remove_product)))

def evaluate_masterpack(ddi_master_pack, tvu, stock, last_sales_day):
    if ddi_master_pack in ['PRODUCTO SIN VENTA', 'PRODUCTO SIN CARGA']:
        return ddi_master_pack.upper()
    
    if pd.isna(ddi_master_pack):
        return 'PRODUCTO SIN CARGA'
    
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
    
    if pd.isna(ddi_packqty):
        return 'PRODUCTO SIN CARGA'
    
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
    
    if pd.isna(ddi_innerpack):
        return 'PRODUCTO SIN CARGA'
    
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
        return value_data.quantize(Decimal('0.00'), rounding=ROUND_HALF_UP)
    except:
        return value

def clean_value(value):
    if pd.isna(value) or value in [float('inf'), float('-inf')]:
        return ''
    return value

print('============== INGRESO DE ARCHIVO ==============')
archive = input("Ingrese el nombre del archivo (sin extensión): ") + '.xlsx'
sheet_name = input("Ingrese el nombre de la hoja: ")

# Leer el archivo Excel existente
output_file_excel = f'archivos/{archive}'
df = pd.read_excel(output_file_excel, sheet_name=sheet_name, header=1, index_col=None)

# Asignar valores predeterminados
days = 28
tvu = 60
remove_product = 7

resultArray = []

print(df.columns)

print('============== ARCHIVO EXCEL ==============')
for index, row in df.iterrows():
    last_sales_day = int(row['VtaUltDiasCant']) if pd.notna(row['VtaUltDiasCant']) else None
    ddi_master_pack = row['DDI_Masterpack']
    ddi_packqty = row['DDI_Packqty']
    ddi_innerpack = row['DDI_Innerpack']
    stock = int(row['Stock']) if pd.notna(row['Stock']) else None

    ideal_size = calculate_ideal_size(days, last_sales_day, tvu, remove_product)
    validate_masterpack = evaluate_masterpack(ddi_master_pack, tvu, stock, last_sales_day)
    validate_packqty = evaluate_packqty(ddi_packqty, tvu, stock, last_sales_day)
    validate_innerpack = evaluate_innerpack(ddi_innerpack, tvu, stock, last_sales_day)

    results = {
        'Dias': days,
        'Stock': stock,
        'VtaUltDiasCant': last_sales_day,
        'Tvu': tvu,
        'Remover producto': remove_product,
        'DDI Masterpack': safe_round(ddi_master_pack),
        'DDI Packqty': safe_round(ddi_packqty),
        'DDI Innerpack': safe_round(ddi_innerpack),
        'Validar Packqty': validate_packqty,
        'Validar Innerpack': validate_innerpack,
        'Validar Masterpack': validate_masterpack,
        'Tamaño ideal': ideal_size
    }
    
    resultArray.append(results)

df = pd.DataFrame(resultArray)

# Definir el nombre del archivo Excel
output_file_excel = 'resultados.xlsx'

# Guardar en archivo Excel
with pd.ExcelWriter(output_file_excel, engine='xlsxwriter') as writer:
    # Crear hoja de resultados
    df.to_excel(writer, sheet_name='Resultados', index=False, startrow=2)

    # Obtener el objeto de la hoja
    worksheet = writer.sheets['Resultados']
    
    # Definir formatos
    title_format = writer.book.add_format({
        'bold': True,
        'align': 'center',
    })

    header_format = writer.book.add_format({
        'bold': True,
        'align': 'center',
        'bg_color': '#d9ead3',
        'border': 1
    })

    data_format = writer.book.add_format({
        'align': 'center',
        'border': 1,
    })

    green_format = writer.book.add_format({
        'align': 'center',
        'border': 1,
        'bg_color': 'green'
    })

    # Escribir título y encabezados
    worksheet.merge_range('A1:L1', 'RESULTADOS VALIDACIÓN', title_format)
    worksheet.merge_range('A2:L2', '', title_format)

    for col_num, col_name in enumerate(df.columns):
        worksheet.write(2, col_num, col_name, header_format)

    # Aplicar formato a las celdas
    for row_num, row_data in df.iterrows():
        for col_num, cell_value in enumerate(row_data):
            # Determinar el formato a utilizar
            if col_num >= 8 and col_num <= 11:  # Columnas I (8) a L (11)
                if 10 <= row_num + 3 <= 20:  # Filas 10 a 20 en el Excel (ajustado por el offset)
                    format_to_use = green_format
                else:
                    format_to_use = data_format
            else:
                format_to_use = data_format

            # Verificar el tipo de valor y escribir de acuerdo al formato
            if isinstance(cell_value, (int, float, Decimal)):
                if not pd.isna(cell_value) and not (math.isinf(cell_value) or math.isnan(cell_value)):
                    worksheet.write_number(row_num + 3, col_num, cell_value, format_to_use)
                else:
                    worksheet.write_string(row_num + 3, col_num, '', format_to_use)
            else:
                worksheet.write_string(row_num + 3, col_num, str(cell_value), format_to_use)

print(f'Archivo Excel guardado correctamente en la ruta {output_file_excel}')
