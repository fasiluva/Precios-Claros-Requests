import openpyxl 
from openpyxl.styles import PatternFill 
from pprint import pprint

"""
. El precio por unidad del producto se calcula automaticamente en el excel. No hay que programarlo.
. Hay recetas que usan otras recetas para elaborar productos (por ejemplo, el pan). Estos no son encontrados,
  ya que no estan en el excel de productos. Corregir?

"""


def extraerPrecios(sheet):

    """
    extraerPrecios :: Worksheet -> dict(str : float)

    . Extrae todos los productos de la lista del excel. Guarda el nombre del mismo y
      el precio estandarizado dividido la cantidad de unidades en el.

    Por ejemplo, en 1 kg de anchoas hay aproximadamente 65 anchoas, y el precio del kilo de anchoas 
    sale 20.000, entonces el precio de cada anchoa ronda los 300 pesos. En este caso, el precio de cada 
    anchoa es lo que nos interesa.
    """

    productos = {}

    for row in sheet.iter_rows(min_row = 7, min_col = 2, max_col = 9):
        if row[0].value != None:
            # Si la celda del nombre del producto no es None, entonces no termino la lista.
            try: 
                pprint(row[0].value)
                productos[row[0].value.casefold()] = float(row[6].value) / float(row[7].value) 

            except TypeError:
                # Si al intentar hacer la division hubo un TypeError, entonces no se pudo calcular el precio
                # estandar anteriormente, o la casilla de 'agrupados en' esta vacia. Continua la iteracion.
                continue
        
        else:
            # Si la celda es None, termino la lista.
            break

    return productos


def main(): 

    # Ruta del archivo
    archivo_xlsx = 'Costodecarta(2).xlsx'

    try: 
        workbook = openpyxl.load_workbook(filename=archivo_xlsx)
        sheet = workbook.get_sheet_by_name("Maestro de provedores")

    except:
        print("ERROR: Archivo de materias primas no encontrado.\n")
        return

    productos = extraerPrecios(sheet)
    
    sheet = workbook.get_sheet_by_name("CARTA")

    columnaNombres = sheet['D']

    #! CAMBIAR NOMBRE DE ESTA VARIABLE A desfaseArch.
    i = 6

    while True: 
        try: 
            if columnaNombres[i].fill.patternType == 'solid':
                # Si la celda tiene un subrayado, entonces esta celda es el nombre del platillo.

                i += 1 # El siguiente item de la lista es el nombre de un producto de la receta.

                while columnaNombres[i].value != None:
                    # Mientras que la celda en la que estoy no sea None, entonces esta es el nombre de un producto.

                    try: 
                        # Intenta buscar el producto en la lista del excel.
                        productoActual = productos[columnaNombres[i].value.casefold()]

                    except KeyError:
                        # Si no encontro el producto, lo notifica, cambia el color de la celda y continua a la siguiente iteracion.
                        print("\nNO ENCONTRADO: ", columnaNombres[i].value.casefold())
                        sheet[f'E{i + 1}'].fill = PatternFill(fill_type='solid', fgColor='E6B8B7')
                        i += 1
                        continue 
                        
                    else: 
                        # Si encontro el producto en la lista del excel:
                        pprint(productoActual) 
                    
                        try: 
                            # Intenta escribir en la casilla el precio del producto estandarizado.
                            sheet[f'E{i + 1}'] = float(productoActual)
                            sheet[f'E{i + 1}'].fill = PatternFill(fill_type= 'solid', fgColor= 'B8CCE4') # Color azul

                        except: 
                            # Si no pudo escribir el precio estandarizado, marca la casilla como incompleta y continua la siguiente iteracion.
                            sheet[f'E{i + 1}'].fill = PatternFill(fill_type='solid', fgColor='E6B8B7') # Color rojo
                            i += 1
                            continue
                            # print("Producto ", columnaNombres[i].value, " no tiene precio estandar.")
                    
                    i += 1

        except IndexError:
            # Si hubo un IndexError, se llego al final del excel y no hay mas recetas. Sale del while.
            break
        
        else: 
            i += 1

    workbook.save(filename=archivo_xlsx)


main()

