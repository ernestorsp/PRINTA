# PRINTA

Aplicación interna sencilla para organizar órdenes y finanzas de Printa Crea.

## Funciones

- Órdenes con dos estados: Pendiente y Listo.
- Origen de la orden: TikTok, Printa Crea u Otro.
- Total pagado y saldo pendiente.
- Fotos de referencia, producto terminado y etiqueta de envío.
- Registro de gastos con categoría, nota y recibo.
- Resumen mensual de ingresos, gastos, ganancia y ventas por origen.

## Base de datos

Google Sheet `PRINTA`:

- `ORDENES`
- `GASTOS`
- `CONFIGURACION`

## Publicación en Apps Script

1. Crear o abrir un proyecto de Google Apps Script.
2. Copiar `Code.gs`, `Index.html` y `appsscript.json`.
3. Implementar como aplicación web.
4. Ejecutar como el usuario que implementa.
5. Autorizar acceso a Google Sheets y Google Drive.

El ID del Sheet ya está configurado dentro de `Code.gs`.
