# Script de Procesamiento de Tarjetas de Descuento

Este script de Python procesa un archivo de Excel que contiene información sobre las tarjetas de descuento. El script realiza las siguientes tareas:

1. Carga un archivo de Excel desde una ruta especificada.
2. Filtra las tarjetas de descuento que están duplicadas.
3. Aplica un filtro de color a las celdas que contienen tarjetas de descuento duplicadas.
4. Crea una nueva hoja en el libro de trabajo y copia ciertos datos de la hoja original a la nueva hoja si la tarjeta de descuento está duplicada.
5. Crea otra hoja en el libro de trabajo y copia ciertos datos de la hoja original a la nueva hoja si la tarjeta de descuento está duplicada.
6. Aplica fórmulas a ciertas celdas en la nueva hoja y formatea las celdas para tener un máximo de dos decimales.
7. Ajusta el ancho de las columnas en todas las hojas para adaptarse al contenido.
8. Guarda el libro de trabajo procesado en una nueva ubicación.

## Requisitos

Este script requiere Python y las siguientes bibliotecas de Python:

- `openpyxl`
- `decouple`
- `datetime`

Puedes instalar estas bibliotecas con pip:

```bash
pip install openpyxl python-decouple

```

### Uso
Para usar este script, debes tener un archivo de configuración .env en el mismo directorio que el script. Este archivo debe contener las siguientes variables:

- `FILE_PATH`: La ruta al archivo de Excel que deseas procesar.
- `SHEET1`: El nombre de la primera hoja en el archivo de Excel.
- `SHEET2`: El nombre que deseas para la segunda hoja creada en el archivo de Excel.
- `SHEET3`: El nombre que deseas para la tercera hoja creada en el archivo de Excel.
- `FILENAME_SAVE`: El nombre que deseas para el archivo de Excel procesado.
- `FILE_PATH_PROCESSED`: La ruta donde deseas guardar el archivo de Excel procesado.

Una vez que hayas configurado tu archivo .env, puedes ejecutar el script con Python:

```bash
python tarjetaDescuentoScript.py
```

El script procesará el archivo de Excel y guardará el archivo procesado en la ubicación especificada.
