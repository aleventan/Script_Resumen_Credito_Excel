from PyPDF2 import PdfReader
import re

def abrirPDF(nombre_archivo_xls):
    # Lee el pdf
    reader = PdfReader(nombre_archivo_xls)
    # Armo las listas donde voy a almacenar los consumos, pagos e impuestos
    consumos = []
    pagos = []
    impuestos = []
    paginas = reader.pages
    # Recorre cada pagina para tomar los datos
    for z, pagina in enumerate(paginas):
        text = pagina.extract_text()
        # Separa las lineas
        lines = text.splitlines()
        for i, line in enumerate(lines):
            # Tomo la fecha de vencimiento para poder manejar las tablas a partir de aca
            if z == 0 and i == 10:
                fecha_vto = line[:9].replace(' ', '-')
                # print(fecha_vto)
            fecha_match = re.search(r'\d{2}\.\d{2}\.\d{2}', line)
            if fecha_match:

                # Si la linea es más grande, genera otra y la agrega al final de la lista de la primera iteración para revisarla luego
                if len(line) > 110:
                    nueva_liena = line[105:]
                    lines.append(nueva_liena)
                # Elimina los espacios en blanco al final
                line = line.lstrip()
                # Toma la fecha, consumo, cuota e importe, todo segun los resumenes del banco Nacion.
                fecha = line[:8]
                id_consumo = line[9:23].lstrip().rstrip()
                cuota = line[60:68].rstrip()
                importe = line[85:97]
                # Si es cuota muestra el consumo sin la pabara couta
                palabra_cuota = line[53:58]
                if palabra_cuota == 'Cuota':
                    consumo = line[24:53].rstrip()
                else:
                    consumo = line[24:60].rstrip()
                    # Ya que yt factura cada vez con codigo diferente, lo agrrego a mano
                    if consumo[:15] == 'GOOGLE *YouTube':
                        consumo= consumo[:15]

                # si hay 4 digitos al final del consumo, los borra
                if consumo[-4:].isdigit():
                    consumo = consumo[:-4].rstrip()
                else:
                    pass

                #Formatear importe
                importe = importe.replace('.', '').replace(',', '.')

                # Verificar si el número es negativo y multiplicar por -1 si es necesario
                if importe.endswith('-'):
                    importe = float(importe[:-1]) * -1
                else:
                    importe = float(importe)

                # Crear un diccionario con los campos extraídos
                item = {
                    'id_consumo': id_consumo,
                    'fecha': fecha,
                    'consumo': consumo,
                    'cuota': cuota,
                    'importe': importe
                }

                # Agregar el diccionario a la lista de pagos
                if consumo == 'SU PAGO EN PESOS':
                    pagos.append(item)
                # Agrega a los impuestos
                elif consumo[:2] == 'DB' or consumo[:8] == 'IMPUESTO' or consumo[:9] == 'INTERESES' or consumo[:3] == 'IVA' :
                    if len(consumo) >= 20 and consumo[26] == '$':
                        # print(consumo[26])
                        consumo = line[24:50].rstrip()
                    else:
                        consumo = line[24:60].rstrip()

                    if consumo[:6] == 'DB IVA':
                        importe = float(line[75:].strip().replace(',', '.'))
                        cuota = ''
                    item['consumo'] = consumo
                    item['importe'] = importe
                    item['cuota'] = cuota
                    impuestos.append(item)
                # Sino, agrega a los consumos
                else:
                    consumos.append(item)

    return consumos, impuestos, pagos, fecha_vto

if __name__ == '__main__':
    pass