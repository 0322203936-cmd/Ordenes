# Ordenes Wholesale

Aplicación web local para capturar órdenes wholesale, consultar productos y transporte, generar DayOrders y preparar Passing Sheets.

## Ejecutar localmente

```bash
npm start
```

Después abrir `http://localhost:3050`.

También se puede iniciar en Windows con `INICIAR_SISTEMA.bat`.

## Estructura principal

- `index.html`, `styles.css`: interfaz tipo Excel.
- `app.js`: interacción, edición, macros y exportaciones.
- `formulaEngine.js`: cálculo de fórmulas del libro.
- `server.js`: servidor HTTP y API de guardado.
- `posco-export.js`: crea el Excel compatible con POSCO desde `PO Nuevo` y detecta productos pendientes de BOM/BOOM.
- `order-pdf.js`: crea el PDF de la orden para adjuntarlo al correo.
- `passing-pdf.js`: crea el PDF del Passing Sheet para adjuntarlo al correo.
- `email-sender.js`: envía PDF, POSCO y el detalle de BOM mediante Microsoft Graph.
- `data/initial_workbook.json`: datos iniciales.
- `data/workbook_state.json`: estado guardado actualmente.

## Generar Orden y Passing Sheet

El botón `Generar Orden` abre la orden en una vista lista para usar `Imprimir` y prepara el PDF y el Excel `POSCO Wholesale` con las columnas requeridas para cargarlo en POSCO. El botón `Enviar` manda ambos archivos por correo. Si alguna línea tiene estado `PENDIENTE DE BOM` o `PENDIENTE DE BOOM`, el sistema muestra la propuesta de BOM y la incluye en el mensaje.

El botón `Enviar` aparece después de generar una orden y manda por Microsoft Graph el PDF, el Excel POSCO y los pendientes de BOM/BOOM en el cuerpo del mensaje. Configura en Render `SHAREPOINT_TENANT_ID`, `SHAREPOINT_CLIENT_ID`, `SHAREPOINT_CLIENT_SECRET`, `MAIL_SENDER` y `WHOLESALE_PDF_RECIPIENTS`. Los tres primeros nombres se conservan para reutilizar las credenciales del bot de referencia. La aplicación de Microsoft Entra debe tener permiso de aplicación `Mail.Send` con consentimiento de administrador.

El botón `Generar Passing Sheet` prepara la vista imprimible del Passing Sheet y habilita `Enviar` para mandar ese PDF por correo a los mismos destinatarios.

## Neon PostgreSQL

Cuando `DATABASE_URL` está configurada, el servidor guarda y lee el libro desde Neon PostgreSQL. En el primer arranque crea la tabla `workbook_state` y carga `data/initial_workbook.json` como plantilla inicial si la tabla está vacía.

En PowerShell, configurar la conexión solo para la sesión actual:

```powershell
$env:DATABASE_URL = "postgresql://..."
npm start
```

También se puede copiar `.env.example` como `.env` y configurar la variable en el entorno local.

La cadena real nunca debe escribirse en el repositorio. En Render se configurará como variable de entorno secreta.

## Nota de despliegue

Sin `DATABASE_URL`, la aplicación conserva el modo local y escribe en `data/workbook_state.json`. Con `DATABASE_URL`, Neon es la fuente principal y el JSON solo se utiliza como semilla inicial.
