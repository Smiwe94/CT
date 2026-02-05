# CronosTask (Add-in de Excel)

Este repositorio contiene `Actividades.html` convertido en un complemento de Excel (Office Add-in) para sincronizar tareas con un libro almacenado en OneDrive.

## Archivos clave

- `Actividades.html`: interfaz del panel (task pane) con integración a Office.js.
- `manifest.xml`: manifiesto del complemento para cargarlo como complemento privado.

## Cómo usarlo en Excel (OneDrive)

1. Hospeda `Actividades.html` en un servidor HTTPS accesible (por ejemplo, una carpeta en OneDrive/SharePoint o un servidor web seguro).
2. Actualiza en `manifest.xml` la URL de `SourceLocation` y `SupportUrl` con tu URL HTTPS pública.
3. Sideload del complemento:
   - Excel en la web: **Insertar** → **Complementos** → **Mis complementos** → **Agregar un complemento personalizado** → **Desde archivo** y selecciona `manifest.xml`.
4. Abre un libro en OneDrive y carga el complemento.

## Usarlo desde SharePoint (.aspx) como plataforma colaborativa

Puedes abrir `Actividades.html` desde una **Site Page** de SharePoint (se guarda como `.aspx`) para que el equipo trabaje en la nube.

1. Sube `Actividades.html` a una biblioteca del sitio (por ejemplo, **Documentos compartidos**).
2. Crea una página del sitio en SharePoint (esto genera un `.aspx`).
3. Inserta el web part **Integrar (Embed)** y pega la URL HTTPS del `Actividades.html`.
4. Usa un archivo de Excel compartido en OneDrive/SharePoint y abre el complemento en Excel Web para sincronizar.

> Nota: Office.js solo puede escribir/leer el Excel cuando está abierto en el complemento de Excel. En la página `.aspx`, la interfaz trabaja en modo independiente y luego sincroniza al abrir el add-in.

## Sincronización con Excel

- **Guardar Excel**: guarda las tareas en una hoja llamada `Cronograma` dentro del libro abierto.
- **Cargar Excel**: lee la hoja `Cronograma` y actualiza la interfaz.
- **Modo independiente**: puedes abrir `Actividades.html` en una ventana del navegador para trabajar sin Excel y luego, al abrir el complemento, pulsar **Guardar Excel** para enviar los cambios a la hoja.
- **Auto-guardado**: la interfaz guarda cambios en el navegador y, si el complemento está conectado, intenta sincronizar automáticamente con Excel.
- **Auto-cargar cambios**: cuando el complemento está abierto en Excel, recarga cada 30s para traer cambios de otros usuarios.

> Nota: si la hoja no existe, el complemento la crea automáticamente.
