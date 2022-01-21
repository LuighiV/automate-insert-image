# Uso del software como Aplicación de Interfaz Gráfica

Este es el uso predeterminado del software, que se puede iniciar desde el menú Inicio y abre automáticamente la interfaz gráfica para ingresar a las opciones y procesar los archivos de Word / PDF. Este artículo se divide en dos secciones orientadas a los dos tipos de documentos de entrada para insertar la imagen:

1. [Insertar imagen en archivos de MS Word](#insertar-imagen-en-archivos-de-ms-word)

2. [Insertar imagen en archivos PDF](#insertar-imagen-en-archivos-pdf)

## Insertar imagen en archivos de MS Word

Tanto para archivos Word como PDF, hay dos opciones para usar el software: con archivos individuales o con una carpeta que contiene los archivos a procesar.

### Para archivos individuales:

1. Seleccione *Documento de Word* como *Formato de entrada* de documento.

2. Seleccione *Archivos* como *Tipo*.

3. Haga clic en *Explorar* en la fila *Dirección de carpeta/archivos* para seleccionar un archivo de MS Word válido o una lista de archivos.

    ![Opciones para abrir archivo de Word](~/images/es/word_file_1_2_3.png "Opciones para abrir archivo de Word")

     Abre el explorador de archivos donde debe localizar el documento y abrirlo (también puede seleccionar varios documentos e insertar la imagen en todos ellos).

    ![Abrir archivo(s) de Word](~/images/es/word_file_3.png "Abrir archivo(s) de Word")

4. Haga clic en *Explorar* en la fila *Dirección de imagen* para seleccionar un archivo de imagen válido a insertar en los documentos seleccionados anteriormente.

    ![Explorar archivo de imagen](~/images/es/word_file_4.png "Explorar archivo de imagen")

    También abre el explorador de archivos para seleccionar la imagen deseada:

    ![Abrir archivo de imagen](~/images/es/word_file_4_1.png "Abrir archivo de imagen")

    Cuando se selecciona el archivo a través del explorador de archivos, se muestra una vista previa de la imagen en el cuadro de *Vista previa*. También carga las dimensiones medidas en la unidad seleccionada, que se muestra en la fila *Unidades*.

5. Puede modificar el tamaño de la imagen a insertar y si se marca *Mantener relación de aspecto*, la otra dimensión se modifica en consecuencia.
    También puede modificar el *Desplazamiento desde la izquierda* y *Desplazamiento desde abajo*. También puede ingresar números negativos para obtener un desplazamiento hacia la izquierda y hacia abajo. 

    ![Establecer dimensiones](~/images/es/word_file_5.png "Establecer dimensiones")

6. Una vez que se establecen todas las dimensiones, puede ingresar el marcador de posición, que es una parte del texto en el documento que funciona como referencia.

    ![Establecer marcador de posición](~/images/es/word_file_6.png "Establecer marcador de posición")

    Tenga en cuenta que *Desplazamiento desde la izquierda* y *Desplazamiento desde abajo* son relativos a la posición de la esquina superior izquierda del texto en el documento. Para obtener información más detallada sobre el cálculo de la posición de la imagen, consulte [Posición de la imagen relativa al texto](position.md#referencia-relativa-al-texto).

7. En las opciones de guardado, puede seleccionar si desea guardar el PDF en la misma carpeta donde está el documento de MS Word o en una subcarpeta e ingresar el nombre. Si la carpeta no existe, el software la creará por usted.

    ![Elegir ubicación para el PDF guardado](~/images/es/word_file_7.png "Elegir ubicación para el PDF guardado")

8. Finalmente, puede hacer clic en *Generar* para procesar los documentos e insertar la imagen con las dimensiones ingresadas en la referencia ingresada.

    ![Todos los pasos para archivos de Word](~/images/es/word_file_8.png "Todos los pasos para archivos de Word")

    Cuando se inicia el proceso se abre un nuevo diálogo que muestra en una barra de progreso el avance del proceso e imprime mensajes informativos sobre el estado actual del flujo de procesamiento.
    
    <img src="~/images/es/word_file_progress.png" alt="Progreso de procesamiento" width="70%"/>

Luego, puede verificar el resultado abriendo el PDF. No olvide cerrar el PDF si desea volver a generarlo; de lo contrario, se producirá un error ya que el software no puede escribir el archivo si está abierto.

![Archivos de documentos de entrada y salida](~/images/es/word_file_diff.png "Archivos de documentos de entrada y salida")

### Para una carpeta que contiene archivos de Word:

Es bastante similar al anterior flujo de trabajo cambiando en el segundo paso:

1. Seleccione *Documento de Word* como tipo de entrada de documento.

2. Seleccione *Carpeta* como *Tipo*.

3. Haga clic en *Explorar* en la fila *Archivos/Ruta de carpeta* para seleccionar la carpeta que contiene los archivos de Word.

    ![Opciones para abrir la carpeta](~/images/es/word_folder_1_2_3.png "Opciones para abrir la carpeta")

     Abre el explorador de carpetas donde debe localizar la carpeta que contiene los archivos de Word.

    <img src="~/images/es/word_folder_3.png" alt="Abrir carpeta" width="50%"/>

A partir de aquí, continúa con los pasos 4 a 8 de [ficheros individuales](#para-archivos-individuales). Al hacer clic en el botón *Generar*, se abre el cuadro de diálogo de progreso que muestra el estado de procesamiento de todos los documentos de Word dentro de la carpeta seleccionada.

<img src="~/images/es/word_folder_progress.png" alt="Progreso de procesamiento" width="70%"/>

Al abrir la carpeta, puede verificar los archivos generados (cada uno por documento de Word que se encuentra en la carpeta seleccionada).

![Carpetas de documentos de entrada y salida](~/images/es/word_folder_diff.png "Carpetas de documentos de entrada y salida")


## Insertar imagen en archivos PDF

A diferencia del caso de los archivos de Word, para PDF hay dos opciones para la referencia: relativa al texto y relativa a la página, mientras que Word solo tiene relativa al texto.

### Con referencia relativa al texto

En este caso, es bastante similar al caso de Word, solo que cambia tanto para los archivos individuales como para los casos de carpetas, el *Formato de entrada* del documento.

#### Para archivos individuales:

1. Seleccione *Documento PDF* como tipo de entrada de documento.

2. Seleccione *Archivos* como *Tipo*.

3. Haga clic en *Explorar* en la fila *Dirección de carpeta/archivos* para seleccionar un archivo PDF válido o una lista de archivos.
    
    ![Opciones para abrir archivo PDF](~/images/es/pdf_file_1_2_3.png "Opciones para abrir archivo PDF")

     Abre el explorador de archivos donde debes localizar el documento y abrirlo (también puedes seleccionar varios documentos e insertar la imagen en todos ellos).

    ![Abrir archivo(s) PDF](~/images/es/pdf_file_3.png "Abrir archivo(s) PDF")

Continúe del 4 al 8 como el caso de [Word para archivos individuales](#insertar-imagen-en-archivos-de-ms-word). Asegúrese de que la opción *Relativo al texto* esté seleccionada en la fila *Tipo de referencia*.

Después de que el programa haya completado esa tarea con éxito, puede verificar el PDF resultante en la subcarpeta creada (con el nombre PDF, si no se cambió en las opciones) en el directorio del archivo original.

![Archivos de documentos de entrada y salida](~/images/es/pdf_file_diff.png "Archivos de documentos de entrada y salida")

#### Para una carpeta que contiene archivos PDF:

1. Seleccione *Documento PDF* como tipo de entrada de documento.

2. Seleccione *Carpeta* como *Tipo*.

3. Haga clic en *Explorar* en la fila *Dirección de archivo/carpeta* para seleccionar la carpeta que contiene los archivos PDF.

    ![Opciones para abrir la carpeta](~/images/es/pdf_folder_1_2_3.png "Opciones para abrir la carpeta")

     Abre el explorador de carpetas donde debe localizar la carpeta que contiene los archivos PDF.

    <img src="~/images/es/pdf_folder_3.png" alt="Abrir carpeta" width="50%"/>

A partir de aquí se continúa con los 4 a los 8 pasos de [Word para archivos individuales](#insertar-imagen-en-archivos-de-ms-word). Al hacer clic en el botón *Generar*, se abre el cuadro de diálogo de progreso que muestra el estado de procesamiento de todos los documentos PDF dentro de la carpeta seleccionada.

<img src="~/images/es/pdf_folder_progress.png" alt="Progreso de procesamiento" width="70%"/>

Al abrir la carpeta, puede verificar los archivos generados (cada uno por documento pdf de entrada que se encuentra en la carpeta seleccionada).

![Carpetas de documentos de entrada y salida](~/images/es/pdf_folder_diff.png "Carpetas de documentos de entrada y salida")


### Con referencia relativa a la página

En este caso varía, ya que incorpora sus propias opciones para referir la posición de la imagen con respecto a la página.

#### Para archivos individuales:

1. Seleccione *Documento PDF* como tipo de entrada de documento.

2. Seleccione *Archivos* como *Tipo*.

3. Haga clic en *Explorar* en la fila *Dirección de carpeta/archivos* para seleccionar un archivo PDF válido o una lista de archivos.

4. Haga clic en *Explorar* en la fila *Dirección de imagen* para seleccionar un archivo de imagen válido para insertar en los documentos seleccionados anteriormente.
    Cuando se selecciona el archivo a través del explorador de archivos, se muestra una vista previa de la imagen en el cuadro de vista previa. También carga las dimensiones medidas en la unidad seleccionada, que se muestra en la fila *Unidades*.

5. Puede modificar el tamaño de la imagen a insertar y si se marca *Mantener relación de aspecto*, la otra dimensión se modifica en consecuencia.
    También puede modificar el *Desplazamiento desde la izquierda* y *Desplazamiento desde abajo*. También puede ingresar números negativos para obtener un desplazamiento hacia la derecha y hacia abajo.

    ![Establecer dimensiones](~/images/es/pdf_file_page_1_to_5.png "Establecer dimensiones")

6. Seleccione la opción *Relativo a la página* en la fila *Tipo de referencia*.

7. Ingrese el número de página y seleccione qué esquina de la página será la referencia de página para la inserción de la imagen. Para obtener información más detallada sobre el cálculo de la posición de la imagen con respecto a la página, consulte [Posición de la imagen con respecto a la página] (position.md#referencia-relativa-a-la-página).

    ![Establecer referencia de página](~/images/es/pdf_file_page_6_7.png "Establecer referencia de página")

8. Puede modificar el nombre de la subcarpeta donde se exportarán los documentos, si lo desea.

9. Haga clic en *Generar* para crear un nuevo documento PDF con la imagen insertada.

    ![Todos los pasos](~/images/es/pdf_file_page_9.png "Todos los pasos")

    Cuando se inicia el proceso se abre un nuevo diálogo que muestra en una barra de progreso el avance del proceso e imprime mensajes informativos sobre el estado actual del flujo de procesamiento.
    
    <img src="~/images/es/pdf_file_page_progress.png" alt="Progreso de procesamiento" width="70%"/>

Luego, puede verificar el resultado abriendo el PDF. No olvide cerrar el PDF si desea volver a generarlo; de lo contrario, se producirá un error ya que el software no puede escribir el archivo si está abierto.

![Archivos de documentos de entrada y salida](~/images/es/pdf_file_page_diff.png "Archivos de documentos de entrada y salida")

#### Para una carpeta que contiene archivos PDF:

1. Seleccione *Documento PDF* como tipo de entrada de documento.

2. Seleccione *Carpeta* como *Tipo*.

3. Haga clic en *Explorar* en la fila *Dirección de archivo/carpeta* para seleccionar la carpeta que contiene los archivos PDF.

    ![Opciones para abrir la carpeta](~/images/es/pdf_folder_1_2_3.png "Opciones para abrir la carpeta")

     Abre el explorador de carpetas donde debe localizar la carpeta que contiene los archivos PDF.

<img src="~/images/es/pdf_folder_3.png" alt="Abrir carpeta" width="50%"/>

A partir de aquí se continúa con los pasos 4 al 8 de [archivos individuales](#para-archivos-individuales-2). Asegúrese de que la referencia esté establecida en *Relativo a la página* como se especifica en el paso 6. Al hacer clic en el botón *Generar*, se abre el cuadro de diálogo de progreso que muestra el estado de procesamiento de todos los documentos PDF dentro de la carpeta seleccionada.

<img src="~/images/es/pdf_folder_page_progress.png" alt="Progreso de procesamiento" width="70%"/>

Al abrir la carpeta, puede verificar los archivos generados (cada uno por documento pdf de entrada que se encuentra en la carpeta seleccionada).

![Carpetas de documentos de entrada y salida](~/images/es/pdf_folder_diff.png "Carpetas de documentos de entrada y salida")