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
        return 'BAJAR INERPACK'

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

# Solicitar variables de entrada

path = 'archivos'
input_file = input("Ingrese la ruta del archivo Excel: ") + '.xlsx'
sheet_name = input("Ingrese el nombre de la hoja en el archivo Excel: ")
days = int(input("Ingrese el número de días: "))
tvu = int(input("Ingrese el valor de TVU: "))
remove_product = int(input("Ingrese el valor de Remove Product: "))

fullpath = os.path.join(path, input_file)


df = pd.read_excel(fullpath, sheet_name=sheet_name, header=1, index_col=None)

output_dir = 'json_results'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

for index, row in df.iterrows():
    results = calculate_and_validate(row, days, tvu, remove_product)
    
    output_file = os.path.join(output_dir, f"result_{index}.json")
    with open(output_file, 'w') as f:
        json.dump(results, f, indent=4)
    
    print(f"Resultado para fila {index} guardado en {output_file}")

print("Todos los resultados se han guardado en archivos JSON.")

