# Uso del software como aplicación de línea de comandos

Este es el modo alternativo de uso del software. Es útil cuando se intenta utilizar como parte de un script. Para usarlo, debe ubicar el ejecutable, que debe estar en `C:\Program Files\InsertImageDocument`, y tiene el nombre `editdocuments.exe`. Si lo desea, puede agregar esta carpeta a la variable de entorno PATH o simplemente ejecutar el archivo desde esta ubicación.

## Uso

Debe usar el nombre del ejecutable seguido de cualquiera de las siguientes opciones. Si no se ingresa ninguna opción, el software iniciará la interfaz gráfica.
```
editdocuments.exe --opción_1 valor_1 --opción_2 valor_2 ... --opción_n valor_n
```

Donde `opción` podría ser lo siguiente:

* `-t, --type` (Predeterminado: `Word`) Tipo de documento. Valores válidos: `Word`, `PDF`.

* `-A, --is-absolute` (Predeterminado: `false`) La referencia es absoluta a la página. Cuando es cierto, utiliza una página específica y una esquina de la página como referencia (solo para tipo PDF).

* `-N, --page` (Predeterminado: 1) Número de página (usado cuando `--is-absolute` es `true`).

* `-r, --page-reference` (Predeterminado: `bottom_left`) Referencia de página (se usa cuando `--is-absolute` es `true`). Valores válidos: `bottom_left`, `top_left`, `top_right`, `bottom_right`.

* `-f, --inputfile` Archivos de entrada a procesar.

* `-d, --inputfolder` Carpeta de entrada a procesar.

* `-p, --picturepath` Archivo de imagen a insertar.

* `-H, --placeholder` Marcador de posición donde se inserta la imagen.

* `--verbose` (predeterminado: `false`) Imprime todos los mensajes en la salida estándar.

* `-V, --visible` (Predeterminado: `false`) Establecer la aplicación MSWord como visible (solo para el tipo Word).

* `-S, --foldersave` Establecer carpeta para guardar PDF.

* `-u, --subfoldersave` Guardar en subcarpeta relativa a la ruta del documento

* `-s, --save` (Predeterminado: `false`) Guardar en archivo.

* `--leftoffset` (Predeterminado: 0) Desplazamiento izquierdo de la imagen (en puntos).

* `--bottomoffset` (Predeterminado: 0) Desplazamiento inferior para la imagen (en puntos).

* `--width` (Predeterminado: 200) Ancho de la imagen (en puntos).

* `--height` (Predeterminado: 150) Altura de la imagen (en puntos).

* `--help` Muestra esta pantalla de ayuda.

* `--version` Muestra la información de la versión.

Ten en cuenta que la equivalencia de puntos es 72 puntos = 1 pulgada = 2.54 cm.

## Ejemplos

Hay varios casos de uso del software. En los siguientes ejemplos mostramos los más comunes:

### Para un solo archivo de Word

Este comando recibe el archivo de Word como entrada e inserta la imagen en una posición alrededor del texto del marcador de posición. Su contraparte en el modo GUI es [Insertar imagen en archivos de MS Word (para archivos individuales)](gui.md#para-archivos-individuales)

```bash
.\editdocuments.exe --inputfile "D:\Luighi\Documents\test-iid\document_en.docx" --picturepath "D:\Luighi\Documents\test-iid\signature_hw.png" --placeholder "signature" --subfoldersave "PDF" --bottomoffset "-7" --leftoffset "-40" --width "140" --height "56" --verbose true
```

### Para una carpeta con archivos de Word

Este comando inserta la imagen en todos los archivos de Word en la carpeta de entrada especificada. Su equivalente en el modo GUI es [Insertar imagen en archivos de MS Word (para una carpeta que contiene archivos de Word)](gui.md#para-una-carpeta-que-contiene-archivos-de-word)

```bash
.\editdocuments.exe --inputfolder "D:\Luighi\Documents\test-iid\test_folder" --picturepath "D:\Luighi\Documents\test-iid\signature_hw.png" --placeholder "signature" --subfoldersave "PDF" --bottomoffset "-7" --leftoffset "-40" --width "140" --height "56" --verbose true
```


### Para una carpeta con archivos PDF (relativo al texto)

En este caso, la imagen se inserta en todos los archivos PDF en la carpeta de entrada especificada. Su equivalente en el modo GUI es [Insertar imagen en archivos PDF (en relación con el texto - para una carpeta que contiene archivos PDF)](gui.md#para-una-carpeta-que-contiene-archivos-pdf)

```bash
.\editdocuments.exe --type PDF --inputfolder "D:\Luighi\Documents\test-iid\test_folder_pdf" --picturepath "D:\Luighi\Documents\test-iid\signature_hw.png" --placeholder "signature" --subfoldersave "PDF" --bottomoffset "-7" --leftoffset "-40" --width "140" --height "56" --verbose true
```

### Para una carpeta con archivos PDF (relativo a la página)

Este comando inserta la imagen con una posición relativa a la página en lugar del texto. Su contraparte en el modo GUI es [Insertar imagen en archivos de MS Word (relativa a la página - para una carpeta que contiene archivos PDF)](gui.md#para-una-carpeta-que-contiene-archivos-pdf-1)

```bash
.\editdocuments.exe --type PDF --inputfolder "D:\Luighi\Documents\test-iid\test_folder_pdf" --picturepath "D:\Luighi\Documents\test-iid\signature_hw.png" --is-absolute true --page 1 --page-reference top_left --subfoldersave "PDF" --bottomoffset "-80" --leftoffset "57" --width "85" --height "71" --verbose true
```