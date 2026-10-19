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
- `data/initial_workbook.json`: datos iniciales.
- `data/workbook_state.json`: estado guardado actualmente.

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
