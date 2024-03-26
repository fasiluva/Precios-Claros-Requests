Programa hecho por Valentin Sosa, vea mas de mis trabajos en:
GitHub: https://github.com/fasiluva 
Linkedin: https://www.linkedin.com/in/valentin-sosa-aa55a9294/.

# Descripción del proyecto.

Este proyecto consta de dos etapas, cada una con un programa escrito en Python:

* **Extracción de datos**

El archivo extractPrices.py extrae información relevante de una serie de productos especificados en un archivo Excel 2007. El programa lee el Excel que contiene una lista de productos, los busca en la página de [Precios Claros](https://www.preciosclaros.gob.ar/#!/buscar-productos) y obtiene la marca, la cantidad y el precio de cada producto. Luego, actualiza las columnas del Excel con estos datos para ser utilizados por el siguiente programa.

* **Actualización de costos**

El archivo mealsCost.py actualiza los costos de los platos de un restaurante, que están en una hoja diferente dentro del mismo archivo Excel utilizado en la etapa anterior. Este programa complementa los datos obtenidos en la etapa de extracción.

## Propósitos

* Funcionar como un programa auxiliar para restaurantes en Argentina, facilitando la actualización de precios, o para otros proyectos que requieran una automatización casi completa en la extracción de precios.
* Demostrar mi habilidad para desarrollar programas que optimicen procesos y actividades en diversas áreas y sectores.

## Alcance 

* **Limitación geográfica**: Precios Claros es una página web argentina, por lo que sería inútil utilizarla en comercios o proyectos de otros países.
* **Limitación del lado del servidor**: La etapa de extracción hace un uso directo del servidor de la página, por lo que se ve afectada por sus horarios de mantenimiento, lentitud en las solicitudes, etc.
* **Limitación del programador y del usuario**: Se enumeran los posibles problemas para las diferentes etapas:
    
    ###### Etapa de extracción:  

    * El producto puede no ser encontrado en la página como se espera, debido a:
        * Mala escritura del nombre. Esto incluye faltas ortográficas, redacción confusa, o simplemente que en la página se encuentre con otro nombre (por ejemplo, "queso muzarella" != "queso mozzarella").
        * Se encuentra un producto diferente al esperado (por ejemplo, buscar "berenjenas" y encontrar "escabeche de berenjenas").
        * Si el producto no tiene marca específica (por ejemplo, "pan" o "huevos"), es posible que se encuentre su precio promediado.
        * El producto no está en la página.

    ###### Etapa de actualización:

    * Si el producto no fue encontrado en el Excel, o no se encontraron los datos relevantes de dicho producto en la etapa de extracción, dejará un espacio vacío marcado en rojo en las casillas donde debería aparecer en la hoja de recetas.
    * Muy raramente podría ocurrir un fallo en el programa, debido a que la medida que figura en la lista de productos no pueda ser transformada a la utilizada en la receta (por ejemplo, unidades extrañas como 'oz', 'mg', 'pizca', etc).

---

## Formato del Excel

#### Hoja de maestro de proveedores (lista de productos)

![Formato de la tabla Excel](https://github.com/fasiluva/Precios-Claros-WebClient/blob/main/docs/TablaDeProductos.png?raw=true)
Imágen 1: Resumen de los datos de la tabla.

* **¿Por qué la marca puede ser calculada por el programa o modificable por el usuario?** Cada casilla de la columna "Marca" puede tener 5 valores posibles:
    * **Nada**: significa que el producto no se buscó nunca antes, o la casilla fue borrada intencionalmente.
    * **Una marca del producto**: el producto fue buscado anteriormente y encontrado en la página, por lo que se escribe la marca que se encontró.
    * **NA**: el producto fue buscado y no se encontró. Cuando se vuelva a ejecutar el programa, las casillas con el atributo NA no serán buscadas nuevamente, con el fin de optimizar la velocidad del programa al realizar menos consultas. Para volver a buscarlo, simplemente borre la casilla NA.
    * **ANCLADO**: el producto no será buscado en la página, ya que sus datos fueron ingresados manualmente. Esto es especialmente útil para los productos que no se encuentran y para la etapa de actualización, ya que los productos con este valor en la marca son leídos correctamente.
    * **PROMEDIO**: indica que se encontró un precio generico para el producto, ya sea por conveniencia de la página o porque para el producto en si, es indiferente tener una marca (por ejemplo, pan, huevos, frutas y verduras, etc). Consejo: si se desea buscar un producto con una marca en específica, escribirla en "Detalle".

* **¿Qué es la columna "Agrupados en"?** Se usa para especificar cuántas "porciones" conforman el producto. Por ejemplo, un paquete de masitas "Surtido Bagley" tiene (aproximadamente) 80 masitas, entonces decimos que dicho producto "está agrupado" en 80 masitas. Esta columna pierde sentido si se trata de productos que no están divididos en porciones, por ejemplo, aceite de girasol o crema de leche. En esos casos, simplemente se pone 1 en la casilla. Esta columna es útil para la etapa de actualización, para las recetas que especifican las proporciones por porción y no por paquete. Por ejemplo, en el ejemplo de carta proporcionado, la receta "Pincho de dátiles" mide los dátiles por unidad, pero al extraer dicho producto de la página, se encuentra por kg. Entonces se calcula una aproximación de cuántos dátiles hay en un kg y se escribe en la casilla. Esta casilla es opcional si no se realizará la etapa de actualización, o no está interesado en tener tanta precisión con los numeros.

![Ejemplo de la tabla Excel](https://github.com/fasiluva/Precios-Claros-WebClient/blob/main/docs/excel-example1.PNG?raw=true)
Imágen 2: Ejemplo de tabla y sus resultados.

#### Hoja de carta 

![Formato de la columna de recetas](https://github.com/fasiluva/Precios-Claros-WebClient/blob/main/docs/FormatoDeLaCarta.png?raw=true)
Imágen 3: Formato de columna 'D' de la hoja de recetas (formato de recetas).

![Formato de receta en Excel](https://github.com/fasiluva/Precios-Claros-WebClient/blob/main/docs/receta-example.PNG?raw=true)
Imágen 4: Formato de receta en Excel.

* Las casillas que tienen la letra en **negrita** están calculados con una macro del propio Excel.

---

#### Requisitos tecnicos

1. Python 3.0 o superior.
2. Librerias/módulos de Python: `openpyxl` y `requests`.
3. Archivo .xlsx (Excel 2007 o superior), llamado `products`.
4. Un navegador web de los siguientes: Google Chrome, Opera, Opera GX, Firefox (desconozco el funcionamiento del programa en otros navegadores).

#### Requisitos del excel

* Estar todo en un archivo .xlsx (Excel de 2007 o superior)

##### Hoja de lista de productos:

1. No tener espacios entre productos
2. Tener la cantidad exacta de columnas como se muestra en el archivo de ejemplo
5. La lista de productos comienza en la linea 7 (se puede modificar en los archivo .py).

##### Hoja de recetas

1. Las recetas comienzan en la linea 7 (se puede modificar en los archivos .py)
1. El archivo de recetas debe tener: los nombres de la receta y su resaltado, las cantidades de los productos y los espacios entre productos y nombre de receta **EXACTAMENTE COMO ESTA EN LA IMÁGEN 4**. Sino puede sobreescribir en lugares del excel que no corresponden o leer mal el archivo.   

## Ejecución

1. Clone el repositorio, o descargue el Zip y descomprímalo.
2. Vaya a la carpeta `dist`.
3. Busque los archivos `extractPrices.exe` y `mealsCost.exe`.  El primero se encarga de extraer los precios y el segundo de actualizar la carta. 
4. Sáquelos de esa carpeta y póngalo en la misma carpeta que el Excel. Puede borrar la carpeta `dist` posteriormente. Luego ejecútelos.
5. Puede crear un acceso directo a dichos archivos y acceder a ellos donde quiera, pero no los mueva de esa carpeta. Aquí terminan los pasos si todo funciona bien.

* Si surgue algun problema para ejecutarlos, puede recurrir a los siguientes pasos:

6. Vuelva a la carpeta principal del programa.
7. Abra una terminal en dicha carpeta, y ejecute los .py de manera manual: `python extractPrices.py` y `python mealsCost.py`.

    * Si los programas no se ejecutan de manera esperada:
        8. Revise tener instalados todos los módulos pedidos anteriormente.
        9. Vea que la pagina web este actualmente activa y funcionando. Si no esta activa, lo más probable es que se encuentre en mantenimiento. Reintente mas tarde.
        10. Contacte con el desarrollador.
    
    * Si los programas funcionan, descargue el módulo `pyinstaller` para Python, vamos a guardar los archivos en un `.exe`:
        11. Ejecute en la terminal: `pyinstaller -F extractPrices.py && pyinstaller -F mealsCost.py`. Ésto creara la carpeta `dist` en el mismo directorio.
        12. Vuelva al paso 2.
