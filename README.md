# Script Resumen Tarjeta de Credito a Excel
## Script que permite organizar los consumos realizados con la tarjeta de crédito en un archivo Excel actualizable resúmen a resúmen.

### Contenido
- [Descripción](#descripcion)
- [Instalación y uso](#instalacion)
- [Licencia](#licencia)

### Descripción <a name="descripcion"></a>
Script que recorre el resúmen de la tarjeta de crédito (por ahora solo funciona con la tarjeta Visa del banco Nación) y lo agrega a un archivo Excel. En el mismo se podrán ver 5 hojas: 
* Mes Actual: En ella se veran reflejados en 3 diferentes tablas los movimientos actuales del mes: consumos, impuestos y pagos realizados. Además figura la fecha de vencimiento.
* Cuotas: Tabla que muestra todos los consumos realizados en cuotas, la cantidad de cuotas que son y por cuál va. Además se puede ver el total a pagar en dicho mes y en los venideros solo con consumos en cuotas.
* Un Pago: Tabla que agrupa todos los consumos en 'un pago'. Allí se meustran las sumas de los importes del mismo tipo de consumo detallado mes a mes.
* Impuestos: Tabla con el total de todos los impuestos a pagar en cada mes.
* Pagos Realizados: Lista de totales abonados a la tarjeta en el mes previo al del resúmen.

### Instalación y uso<a name="instalacion"></a>
Para poder ejecutar el Script, es necesario primero instalar las librerías necesarias. Las mismas se encuentran en el archivo requirements.txt.
Luego es necesario ejecutar el archivo main.py. __Es necesario tener los resúmenes descargados desde la página web de visa y guardarlos donde se encuentra el script.__ 
Una vez ejecutado el archivo hay que ingresar el nombre del __archivo PDF (Sin '.pdf')__ y presionar Enter. Al finalizar, se habrán cargado los datos dentro del archivo 'Detalle Visa Nacion.xlsx'.

**IMPORTANTE:**
 Siempre se va a escribir sobre el archivo existente y deja almacenados los datos previos, para tener un regístro de todo lo que se viene abonando mes a mes.
Hay que tener en cuenta que, para mantener un orden, se ejecute primero los meses anteriores a los posteriores. Esto se debe a que el script va agregando las columnas en las tablas a medida que se ejecuta, sin distinguir por nombres de los meses. 

### Licencia <a name="licencia"></a>
Este proyecto está licenciado bajo la licencia Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0).

#### ¿Qué significa esto?
Esta licencia permite a otros distribuir, remezclar y construir sobre tu trabajo, siempre y cuando se dé crédito al autor original y no se utilice con fines comerciales. Esto significa que retienes los derechos de autor de tu trabajo y puedes decidir cómo se utiliza. Al utilizar esta licencia, puedes asegurarte de que tu trabajo se comparta y se utilice de manera que esté alineada con tus valores.

#### Cómo utilizar este trabajo
Si deseas utilizar este trabajo para fines no comerciales, puedes hacerlo siempre y cuando des crédito al autor original. Si deseas utilizar este trabajo con fines comerciales, debes contactar al autor para obtener permiso.
