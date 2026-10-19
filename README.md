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

## Nota de despliegue

El guardado actual escribe en un archivo JSON local. Para producción en Render se deberá migrar posteriormente a almacenamiento persistente o una base de datos.
