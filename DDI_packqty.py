stock = int(input("Ingrese la cantidad de stock: "))
VtaUltDiasCant = int(input("Ingrese el VtaUltDiasCant: "))
packqty = int(input("Ingrese el packqty: "))
dias = 28 

def evaluateDdiPackqty(stock, VtaUltDiasCant, packqty):
    if stock == 0 and VtaUltDiasCant == 0:
        return "PRODUCTO SIN CARGA"
    elif stock > 0 and VtaUltDiasCant == 0:
        return "PRODUCTO SIN VENTA"
    else:
        return round(packqty / (VtaUltDiasCant / dias), 2)
    
        
    
validateDdiPackqty= evaluateDdiPackqty(stock, VtaUltDiasCant, packqty)

print(validateDdiPackqty)

