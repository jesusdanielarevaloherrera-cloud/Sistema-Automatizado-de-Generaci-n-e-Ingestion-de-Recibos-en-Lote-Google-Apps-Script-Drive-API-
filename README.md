# 🧾 Sistema Automatizado de Generación de Recibos PDF (Google Apps Script)

Un sistema de automatización empresarial desarrollado en **Google Apps Script** para Google Sheets y Google Drive. Permite agrupar transacciones masivas por cliente, aplicar filtros avanzados por fecha/producto y generar comprobantes de pago en formato PDF renderizados dinámicamente mediante HTML/CSS.

![Demostración del proyecto](docs/demo.gif)

---

## 🚀 Características Principales

* **Filtros Avanzados e Interfaz Modal:** Selección interactiva por mes, año, rango de días (semanas), segmento de cliente y categoría/producto.
* **Procesamiento Eficiente en Lote (Bulk Data Processing):** Lectura y actualización en memoria (`getDisplayValues` / `setValues`) para maximizar el rendimiento y minimizar llamadas a la API de Sheets.
* **Generación de PDF In-Memory:** Generación directa de blobs PDF desde plantillas HTML/CSS sin depender de Google Docs, reduciendo el tiempo de ejecución en más de un 70%.
* **Gestión de Drive API:** Creación y organización automática de carpetas mensuales (`Recibos - [Mes] [Año]`) en Google Drive.
* **Manejo de Tiempos de Ejecución (Execution Timeout Control):** Lógica interna de control de tiempo que detiene de manera segura la corrida al alcanzar 4.5 minutos para evitar el error de tiempo de espera nativo de Apps Script.
* **Prevención de Duplicados:** Marcado dinámico con sello de fecha y hora transaccional (`Recibo_Generado`) y opción de forzar reprocesado.

---

## 🛠️ Tecnologías Utilizadas

* **Lenguaje:** JavaScript (Google Apps Script - ES6)
* **Entorno Workspace:** Google Sheets API, Google Drive API, HTMLService, Utilities
* **Frontend Interno:** HTML5, CSS3 (Estilos en línea y tablas sintácticas para renderizado estricto en PDF)

---

## 📋 Estructura de la Tabla de Datos (`ventas`)

Para que el script funcione correctamente, la hoja de cálculo de Google Sheets debe contar con una cabecera que incluya las siguientes columnas (las coincidencias de nombres son flexibles):

| Columna | Descripción |
| :--- | :--- |
| `Fecha` | Fecha de la transacción (ISO `YYYY-MM-DD` o `DD/MM/YYYY`) |
| `ID Venta` | Identificador único de la transacción |
| `Sucursal` | Nombre o código de la sucursal |
| `Vendedor` | Nombre del vendedor asignado |
| `cliente` / `Cliente` | Nombre completo del cliente |
| `Categoría` | Categoría del producto |
| `Producto` | Descripción del ítem o producto |
| `Cantidad` | Unidades vendidas |
| `Total Venta ($)` | Monto total procesado |
| `Método Pago` | Forma de pago empleada |
| `Recibo_Generado` | *(Creada automáticamente)* Fecha y hora del procesamiento |

---

## 🔧 Instalación y Configuración

1. Abre tu hoja de cálculo en **Google Sheets**.
2. Ve al menú superior: **Extensiones > Apps Script**.
3. Copia el contenido del archivo [`src/Codigo.gs`](src/Codigo.gs) y pégalo en el editor del proyecto.
4. Guarda el proyecto (**Ctrl + S** o icono de disco).
5. Recarga tu hoja de cálculo. Aparecerá un nuevo menú superior llamado **Recibos**.
6. Haz clic en **Recibos > Generar recibos** y concede los permisos de Google Workspace solicitados la primera vez.

---

## 📊 Vista Previa del Recibo PDF

El archivo HTML integrado compila un diseño limpio con paleta azul corporativa, encabezado comercial (`TODO MAYOR`), desglose detallado en tabla y pie de página no fiscal.

---

## 📄 Licencia

Este proyecto está bajo la Licencia MIT. Consulta el archivo `LICENSE` para más información.
