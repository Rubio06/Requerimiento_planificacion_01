import pandas as pd
import math
import json
import os

def calculate_ideal_size(days, last_sales_day, tvu, remove_product):
    if pd.isnull(last_sales_day) or isinstance(last_sales_day, pd._libs.tslibs.nattype.NaTType):
        return '0'
        
    return math.trunc((last_sales_day / days) * (tvu - remove_product))

def evaluate_masterpack(ddi_master_pack, tvu, stock, last_sales_day):
    if pd.isnull(ddi_master_pack) or ddi_master_pack == 'PRODUCTO SIN VENTA':
        return 'PRODUCTO SIN VENTA'
    
    try:
        ddi_master_pack = float(ddi_master_pack)
    except ValueError:
        return 'PRODUCTO SIN VENTA'
    
    if ddi_master_pack <= tvu:
        return 'OK'
    elif stock == 0 and last_sales_day == 0:
        return 'PRODUCTO SIN CARGA'
    elif stock > 0 and last_sales_day == 0:
        return 'PRODUCTO SIN VENTA'
    else:
        return 'BAJAR MASTERPACK'

def evaluate_packqty(ddi_packqty, tvu, stock, last_sales_day):
    if pd.isnull(ddi_packqty) or ddi_packqty == 'PRODUCTO SIN VENTA':
        return 'PRODUCTO SIN VENTA'
    
    try:
        ddi_packqty = float(ddi_packqty)
    except ValueError:
        return 'PRODUCTO SIN VENTA'
    
    if ddi_packqty <= tvu:
        return 'OK'
    elif stock == 0 and last_sales_day == 0:
        return 'PRODUCTO SIN CARGA'
    elif stock > 0 and last_sales_day == 0:
        return 'PRODUCTO SIN VENTA'
    else:
        return 'BAJAR PACKQTY'

def evaluate_innerpack(ddi_innerpack, tvu, stock, last_sales_day):
    if pd.isnull(ddi_innerpack) or ddi_innerpack == 'PRODUCTO SIN VENTA':
        return 'PRODUCTO SIN VENTA'
    
    try:
        ddi_innerpack = float(ddi_innerpack)
    except ValueError:
        return 'PRODUCTO SIN VENTA'
    
    if ddi_innerpack <= tvu:
        return 'OK'
    elif stock == 0 and last_sales_day == 0:
        return 'PRODUCTO SIN CARGA'
    elif stock > 0 and last_sales_day == 0:
        return 'PRODUCTO SIN VENTA'
    else:
        return 'BAJAR INNERPACK'

def calculate_and_validate(row, days, tvu, remove_product):
    last_sales_day = row['VtaUltDiasCant']
    ddi_master_pack = row['DDI_Masterpack']
    ddi_packqty = row['DDI_Packqty']
    ddi_innerpack = row['DDI_Innerpack']
    stock = row['Stock']
    
    ideal_size = calculate_ideal_size(days, last_sales_day, tvu, remove_product)
    validate_masterpack = evaluate_masterpack(ddi_master_pack, tvu, stock, last_sales_day)
    validate_packqty = evaluate_packqty(ddi_packqty, tvu, stock, last_sales_day)
    validate_innerpack = evaluate_innerpack(ddi_innerpack, tvu, stock, last_sales_day)
    
    results = {
        'validate_masterpack': validate_masterpack,
        'validate_packqty': validate_packqty,
        'validate_innerpack': validate_innerpack,
        'ideal_size': ideal_size
    }
    
    return results

####################### ENTRADA DE DATOS ###############################

# Configuración inicial
try:
    path: str = 'archivos'
    input_file: str = input('Ingrese el nombre del archivo excel: ') + '.xlsx'
    sheet_name: str = input('Ingrese el nombre de la hoja: ')        
    days: int = 28
    tvu: int = 60
    remove_product: int = 7
    fullpath = os.path.join(path, input_file)

    # Leer el archivo de Excel
    df = pd.read_excel(fullpath, sheet_name = sheet_name, header = 1, index_col = None)

    # Solicitar el índice de la fila que desea procesar
    row_index = int(input("Ingrese el índice de la fila que desea procesar: "))

    # Validar que el índice proporcionado esté dentro del rango del DataFrame
    if row_index < 0 or row_index >= len(df):
        print(f"Índice fuera de rango. El DataFrame tiene {len(df)} filas.")
    else:
        # Procesar la fila específica
        row = df.iloc[row_index]
        results = calculate_and_validate(row, days, tvu, remove_product)

        # Mostrar el resultado en la consola
        print(json.dumps(results, indent=4))

        # Guardar el resultado en un archivo JSON
        output_dir = 'json_results'
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        output_file = os.path.join(output_dir, f"result_{row_index}.json")

        with open(output_file, 'w') as f:
            json.dump(results, f, indent=4)

        print(f"Resultado para fila {row_index} guardado en {output_file}")

    print("Proceso completado.")
    
except Exception as e:
    print(f"el error es: {e}")



