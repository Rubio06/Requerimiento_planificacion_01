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

# Aplicar safe_round a las columnas relevantes
df['DDI_Masterpack'] = df['DDI_Masterpack'].apply(safe_round)
df['DDI_Packqty'] = df['DDI_Packqty'].apply(safe_round)
df['DDI_Innerpack'] = df['DDI_Innerpack'].apply(safe_round)
print('============== ARCHIVO EXCEL ==============')


found = False
for index, row in df.iterrows():
    # Obtener valores de la fila
    row_ddi_masterpack = row['DDI_Masterpack']
    row_vta_ultdiascant = int(row['VtaUltDiasCant']) if pd.notna(row['VtaUltDiasCant']) else None
    row_ddi_packqty = row['DDI_Packqty']
    row_ddi_innerpack = row['DDI_Innerpack']
    row_stock = int(row['Stock']) if pd.notna(row['Stock']) else None

    # Comparaciones
    ddi_masterpack_match = (row_ddi_masterpack == safe_round(ddi_master_pack))
    vta_ultdiascant_match = (row_vta_ultdiascant == last_sales_day)
    ddi_packqty_match = (row_ddi_packqty == safe_round(ddi_packqty))
    ddi_innerpack_match = (row_ddi_innerpack == safe_round(ddi_innerpack))
    stock_match = (row_stock == stock)
    
    # Validar todos los campos
    if ddi_masterpack_match and vta_ultdiascant_match and ddi_packqty_match and ddi_innerpack_match and stock_match:
        found = True
        break
    
if found:
    ideal_size = calculate_ideal_size(days, last_sales_day, tvu, remove_product)
    validate_masterpack = evaluate_masterpack(ddi_master_pack, tvu, stock, last_sales_day)
    validate_packqty = evaluate_packqty(ddi_packqty, tvu, stock, last_sales_day)
    validate_innerpack = evaluate_innerpack(ddi_innerpack, tvu, stock, last_sales_day)
    results = {
        'validate_packqty': validate_packqty,
        'validate_innerpack': validate_innerpack,
        'validate_masterpack': validate_masterpack,
        'ideal_size': ideal_size
    }

    output_dir = 'json_results'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    output_file_json = os.path.join(output_dir, "result.json")
    with open(output_file_json, 'w') as f:
        resultArray.append(results)
        json.dump(resultArray, f, indent=4)
        
    print('============== GUARDANDO ARCHIVOS JSON Y EXCEL ==============')
    print(f'Archivo JSON guardado correctamente en la ruta {output_dir} con el nombre result.json')
    
    output_file_excel = os.path.join(output_dir, "result.xlsx")
    with pd.ExcelWriter(output_file_excel, engine='xlsxwriter') as writer:
        # Crear una hoja con el nombre 'Results'
        writer.book.add_worksheet('Results')
        
        # Acceder a la hoja creada
        worksheet = writer.sheets['Results']
        
        # Definir un formato para el título
        title_format = writer.book.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Arial',           
        })
        
        format_data = writer.book.add_format({
            'bold': True,
            'font_size': 11,
            'valign': 'vcenter',
            'font_name': 'Arial',
            'align': 'left'
           
        })
        
        # Escribir el título en la primera fila (empezando desde la columna B)
        worksheet.merge_range('B1:E1', 'RESULTADOS VALIDACIÓN', title_format)
    
        
        bold_format = writer.book.add_format({
            'bold': True,
            'font_size': 12,
            'align': 'left'
        })
        
        # Definir formato para el valor de days
        normal_format = writer.book.add_format({
            'font_size': 12,
            'align': 'left'

        })
    
        header_format = writer.book.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 2,
            'align': 'center',
            'font_name': 'Arial'

        })

        worksheet.write_rich_string('B3:C3', bold_format, 'Dias: ', normal_format, str(days), format_data)
        worksheet.write_rich_string('B4:C4', bold_format, 'Vta. de los ultimos dias: ', normal_format, str(last_sales_day), format_data)
        worksheet.write_rich_string('B5:C5', bold_format, 'T. de vida util: ', normal_format, str(tvu), format_data)
        worksheet.write_rich_string('B6:C6', bold_format, 'Producto removido: ', normal_format, str(remove_product), format_data)
        worksheet.write_rich_string('B7:C7', bold_format, 'DDI Masterpack: ', normal_format, str(ddi_master_pack), format_data)
        worksheet.write_rich_string('B8:C8', bold_format, 'DDI packqty: ',  normal_format, str(ddi_packqty), format_data)
        worksheet.write_rich_string('B9:C9', bold_format, 'DDI innerpack: ', normal_format, str(ddi_innerpack), format_data)
        worksheet.write_rich_string('B10:C10', bold_format, 'Stock: ', normal_format, str(stock), format_data)
                
        # Escribir los encabezados en la fila 3 (empezando desde la columna B)
        custom_headers = [
            'PACKQTY',
            'INNERPACK',
            'MASTERPACK',
            'TAMAÑO IDEAL',
        ]
        
        #results.keys()
        for col_num, value in enumerate(custom_headers):
            worksheet.write(12, col_num + 1, value, header_format)  # Ajustar el índice de columna para la fila 3
        
        # Definir un formato para las celdas de la tabla
        cell_format = writer.book.add_format({
            'border': 2,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Arial'           

        })
        
        # Escribir los datos en la fila 4 (empezando desde la columna B)
        for col_num, (key, value) in enumerate(results.items()):
            worksheet.write(13, col_num + 1, value, cell_format)  # Ajustar el índice de columna para la fila 4
            
        # Ajustar el ancho de las columnas a un tamaño fijo
        fixed_width = 25  # Ajusta el ancho fijo que deseas
        for col_num in range(len(results.keys())):
            worksheet.set_column(col_num + 1, col_num + 1, fixed_width)  # Ajustar el índice de columna
            
        worksheet.set_row(0, 17)  # Altura de la fila del título
        worksheet.set_row(2, 17)  # Altura de la fila del dato del masterpack
        worksheet.set_row(3, 17)  # Altura de la fila de los encabezados
        worksheet.set_row(4, 17)  # Altura de la fila de los datos
        worksheet.set_row(5, 17)  # Altura de la fila de los datos
        worksheet.set_row(6, 17)  # Altura de la fila de los datos
        worksheet.set_row(7, 17)  # Altura de la fila de los datos
        worksheet.set_row(8, 17)  # Altura de la fila de los datos
        worksheet.set_row(9, 17)  # Altura de la fila de los datos
        worksheet.set_row(12, 17)  # Altura de la fila de los datos
        worksheet.set_row(13, 17)  # Altura de la fila de los datos
        
        print(f'Archivo Excel guardado correctamente en la ruta {output_dir} con el nombre result.xlsx')
else:
    print("No se encontraron coincidencias en el archivo Excel.")
