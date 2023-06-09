from clases import tablaExcel
import traceback

# 'resumen_nacion_junio.pdf'
try:
    nombre_excel = "Detalle Visa Nacion.xlsx"
    input_pdf = input('Ingrese el nombre del archivo, sin \'.pdf\' (Tiene que estar en el mismo directorio que el script): ')
    nombre_resumen_pdf = input_pdf + '.pdf'
    hojas = {0: 'Cuotas', 1: 'Un_pago', 2: 'Impuestos', 3: 'Pagos_Realizados', 4: 'Mes_actual'}
    for key, value in hojas.items():
        # Crea el nombre de la tabla en la que va a trabajar
        tabla = 'Table_' + value
        nueva_tabla = tablaExcel()
        # Toma los datos del pdf
        nueva_tabla.tomarDatosPDF(nombre_resumen_pdf, tipo_dato=key)

        # Abre el archivo excel
        nueva_tabla.abrirExcel(nombre_excel, value)

        # Para cualquier hoja que no sea Mes actual
        if value != 'Mes_actual':
            # SE BORRA EL FORMATO DE LA TABLA SELECCIONADA #
            nueva_tabla.formatoInicial(tabla)

            # Si es cuotas, ejecuta lo siguiente
            if value == 'Cuotas':
                # AGREGAR CONSUMOS EN CUOTAS #
                nueva_tabla.agregarConsumosCuotas()

                # TITULOS #
                nueva_tabla.agregarTitulos()

                # SUMA TOTALES #
                nueva_tabla.sumaTotal()

                # SE LE DA EL FORMATO A LA TABLA #
                nueva_tabla.formatoFinal(tabla)

                # OCULTO COLUMNAS Y FILAS NO RELEVANTES AL MES #
                nueva_tabla.ocultarColFilas()
            else:
                # Varía la funcion segun la tabla, el resto es igual
                if value == 'Un_pago':
                    # AGREGAR CONSUMOS EN UN PAGO #
                    # Ejecuto 2 veces agregarUnPago, ya que A VECES no inserta todos los datos, pero corriendolo dos veces si
                    nueva_tabla.agregarUnPago()
                    nueva_tabla.agregarUnPago()
                elif value == 'Impuestos':
                    # AGREGAR IMPUESTOS #
                    nueva_tabla.agregarImpuestos()
                elif value == 'Pagos_Realizados':
                    # AGREGAR PAGOS #
                    nueva_tabla.agregarPagos()
                # TITULOS #
                nueva_tabla.agregarTitulos(2)

                # SUMA #
                nueva_tabla.sumaTotal(col_inicio=2)

                # SE LE DA EL FORMATO A LA TABLA #
                nueva_tabla.formatoFinal(tabla, col_inicio=2, fila_inicio=2)

                # OCULTO COLUMNAS Y FILAS NO RELEVANTES AL MES #
                nueva_tabla.ocultarColFilas(fila_inicio=2)

            # SE GUARDA EL ARCHIVO #
            nueva_tabla.guardarTabla()

            # SE CIERRA EL ARCHIVO #
            nueva_tabla.cerrarExcel()

        else:
            # Para el mes actual, es un tipo de hoja diferente #
            # Tengo que recorrer todos los datos para que queden almacenados, ya que los necesito de nuevo #
            for i in range(1, len(hojas)):
                nueva_tabla.tomarDatosPDF(nombre_resumen_pdf, tipo_dato=i)

            # CREA LA HOJA DEL MES ACTUAL #
            nueva_tabla.agregarMesActual()

            # SE GUARDA EL ARCHIVO #
            nueva_tabla.guardarTabla()

            # SE CIERRA EL ARCHIVO #
            nueva_tabla.cerrarExcel()
        print('Se creo correctamente la hoja ' + value)

except Exception as e:
    # Obtener información detallada de la excepción y la pila de llamadas
    traceback_info = traceback.format_exc()

    # Mostrar el nombre de la función y el número de línea donde se produjo el error
    print("Mensaje de error:", str(e))

    # Obtener la pila de llamadas
    stack_trace = traceback.extract_tb(e.__traceback__)
    # Obtener el último elemento de la pila de llamadas
    last_call = stack_trace[-1]

    # Mostrar el nombre del archivo, nombre de la función y número de línea
    print("Se produjo un error en el archivo:", last_call.filename)
    print("En la función:", last_call.name)
    print("Número de línea:", last_call.lineno)




