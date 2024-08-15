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
    # Campos adicionales
    fecha = row['Fecha'] 
    id_tienda = row['idtienda']
    tda_nombre = row['TdaNombre']
    dist_nombre = row['DistNombre']
    tda_clasif_oper = row['TdaClasifOper']
    id_producto = row['IdProducto']
    prod_nombre = row['ProdNombre']
    dpto_prod = row['DptoProd']
    clase_prod = row['ClaseProd']
    bloqueo_general = row['BloqueoGeneral']
    bloqueo_tienda_compra = row['BloqueoTiendaCompra']
    id_proveedor = row['Idproveedor']
    provee_nombre = row['ProveeNombre']
    stock = int(row['Stock']) if pd.notna(row['Stock']) else None
    
    packqty = row['Packqty'] if pd.notna(row['Packqty']) else ''
    inner_display = row['InnerDisplay'] if pd.notna(row['InnerDisplay']) else ''
    master_pack = row['MasterPack'] if pd.notna(row['MasterPack']) else ''
    costo_promedio = row['CostoPromedio'] if pd.notna(row['CostoPromedio']) else ''
    vta_ult_dias_soles_costo = row['VtaUltDiasSolesCosto'] if pd.notna(row['VtaUltDiasSolesCosto']) else ''
    last_sales_day = int(row['VtaUltDiasCant']) if pd.notna(row['VtaUltDiasCant']) else None
    int_stock1 = row['IntStock1'] if pd.notna(row['IntStock1']) else ''
    on_order = row['On_Order'] if pd.notna(row['On_Order']) else ''
    cantidad_confirmada_slip = row['Cantidad_Confirmada_Slip'] if pd.notna(row['Cantidad_Confirmada_Slip']) else ''
    cantidad_pendiente_directiva = row['Cantidad_Pendiente_Directiva'] if pd.notna(row['Cantidad_Pendiente_Directiva']) else ''
    ddi_packqty = row['DDI_Packqty']
    ddi_innerpack = row['DDI_Innerpack']
    ddi_master_pack = row['DDI_Masterpack']
    validate_packqty = evaluate_packqty(ddi_packqty, tvu, stock, last_sales_day)
    validate_innerpack = evaluate_innerpack(ddi_innerpack, tvu, stock, last_sales_day)
    validate_masterpack = evaluate_masterpack(ddi_master_pack, tvu, stock, last_sales_day)
    ideal_size = calculate_ideal_size(days, last_sales_day, tvu, remove_product)
    
    # Campos adicionales para agregar
    results = {
        'Fecha': fecha.strftime('%d/%m/%Y') if pd.notna(fecha) else '',
        'idtienda': id_tienda,
        'TdaNombre': tda_nombre,
        'DistNombre': dist_nombre,
        'TdaClasifOper': tda_clasif_oper,
        'IdProducto': id_producto,
        'ProdNombre': prod_nombre,
        'DptoProd': dpto_prod,
        'ClaseProd': clase_prod,
        'BloqueoGeneral': bloqueo_general,
        'BloqueoTiendaCompra': bloqueo_tienda_compra,
        'Idproveedor': id_proveedor,
        'ProveeNombre': provee_nombre,
        'Dias': days,
        'Stock': stock,
        'Packqty': packqty,
        'InnerDisplay': inner_display,
        'MasterPack': master_pack,
        'CostoPromedio': costo_promedio,
        'VtaUltDiasSolesCosto': vta_ult_dias_soles_costo,
        'VtaUltDiasCant': last_sales_day,
        'IntStock1': int_stock1,
        'On_Order': on_order,
        'Cantidad_Confirmada_Slip': cantidad_confirmada_slip,
        'Cantidad_Pendiente_Directiva': cantidad_pendiente_directiva,
        'DDI_Packqty': safe_round(ddi_packqty),
        'DDI_Innerpack': safe_round(ddi_innerpack),
        'DDI_Masterpack': safe_round(ddi_master_pack),
        'TVU': tvu,
        'Valida_Packqty': validate_packqty,
        'Valida_IP': validate_innerpack,
        'Valida_MP': validate_masterpack,
        'Tamaño ideal (unidades)': ideal_size
    }
    
    resultArray.append(results)

df = pd.DataFrame(resultArray)

# Definir el nombre del archivo Excel
output_file_excel = 'UMP_PROV487_20240604_RESULTS.xlsx'

# Guardar en archivo Excel
with pd.ExcelWriter(output_file_excel, engine='xlsxwriter') as writer:
    # Crear hoja de resultados
    df.to_excel(writer, sheet_name='Resultados', index=False, startrow=1)

    # Obtener el objeto de la hoja
    worksheet = writer.sheets['Resultados']
    
    format_data = writer.book.add_format({
        'border': 1,
        'align': 'center'
    })
    
    format_data = writer.book.add_format({
        'align': 'center'
    })
    
    border_format = writer.book.add_format({'border': 1})


    for col_num, col in enumerate(df.columns):
        max_length = max(df[col].astype(str).apply(len).max(), len(col))
        worksheet.set_column(col_num, col_num, max_length, format_data)
    
    green_format = writer.book.add_format({
        'bg_color': '#d0ece7',
        'border': 1,
        'align': 'center',
    })
    border_format = writer.book.add_format({
        'border': 1,
        'align': 'center'
    })
    
    cell_format = writer.book.add_format({
        'align': 'center',
        'bg_color': '#fcf3cf'
    })
    
    row = 0  # Fila específica (indexada desde 0)
    col = 28   # Columna específica (indexada desde 0)

    # Agregar el dato en la celda específica
    worksheet.write(row, col, remove_product, cell_format)
    
    row = 0  # Fila específica (indexada desde 0)
    col = 19   # Columna específica (indexada desde 0)

    # Agregar el dato en la celda específica
    worksheet.write(row, col, days, cell_format)
    
    col_DptoProd = df.columns.get_loc('DptoProd')
    worksheet.set_column(col_DptoProd, col_DptoProd, 25) 
    
    col_ProveeNombre = df.columns.get_loc('ProveeNombre')
    worksheet.set_column(col_ProveeNombre, col_ProveeNombre, 35) 
    
    col_TVU = df.columns.get_loc('TVU')
    worksheet.set_column(col_TVU, col_TVU, 10) 
    
    # Aplicar formato a toda la tabla de datos
    num_rows = len(df.index) + 2
    worksheet.conditional_format(2, 0, num_rows, len(df.columns) - 1, {'type': 'no_blanks', 'format': border_format})


    start_col_name = 'Valida_Packqty'
    end_col_name = 'Tamaño ideal (unidades)'
    
    start_col_index = df.columns.get_loc(start_col_name)
    end_col_index = df.columns.get_loc(end_col_name)
    
    start_row = 2 
    end_row = start_row + len(df) - 1
    
    for col_num in range(start_col_index, end_col_index + 1):
        for row_num in range(start_row, end_row + 1):
            worksheet.write(row_num, col_num, df.iloc[row_num - start_row, col_num], green_format)

