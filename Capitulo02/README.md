# Preparación del dataset de ventas

## Metadatos

| Elemento | Valor |
|---|---|
| Duración | 105 minutos |
| Complejidad | Difícil |
| Nivel de Bloom | Aplicar |

## Descripción General

En esta práctica prepararás una versión curada, trazable y reutilizable del dataset artificial masivo **Ventas Retail LATAM 2026.1**, compuesto por aproximadamente 750,000 transacciones. Usarás el brief analítico de la práctica anterior para priorizar los campos que afectan directamente el análisis comercial, perfilarás la fuente en Excel y Snowflake, y aplicarás reglas justificadas de estandarización y calidad.

El resultado será una base curada común para las prácticas posteriores: un libro de controles en Excel, una vista o tabla curada en Snowflake y una consulta parametrizada de Power Query en Power BI Desktop. Los registros no eliminados conservarán una bandera de calidad y una razón de revisión para evitar ocultar incertidumbre analítica.

## Objetivos de Aprendizaje

- [ ] Perfilar un dataset masivo e identificar su estructura, granularidad, completitud, validez, consistencia y unicidad.
- [ ] Aplicar reglas de normalización para región, canal, categoría, fechas, cantidades y descuentos.
- [ ] Detectar transacciones repetidas por `TRANSACCION_ID`, duplicados exactos y posibles duplicados de negocio.
- [ ] Crear una versión curada documentada en Snowflake y controles operativos equivalentes en Excel y Power BI.
- [ ] Conservar evidencia de registros sospechosos mediante `BANDERA_CALIDAD` y `RAZON_REVISION`.

## Prerrequisitos

### Conocimientos requeridos

- Haber completado el archivo `01_brief_analitico_ventas.xlsx` durante la Práctica 1.
- Comprender que la granularidad debe validarse antes de contar transacciones, clientes o ventas.
- Manejar filtros, tablas, fórmulas y Power Query básico en Excel.
- Conocer las sentencias SQL `SELECT`, `WHERE`, `GROUP BY`, `CASE`, `COUNT`, `MIN`, `MAX` y funciones básicas de fecha.
- Reconocer que una anomalía detectada es un hallazgo que requiere validación, no una razón automática para eliminar datos.

### Accesos requeridos

- Acceso de lectura a `DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1`.
- Rol `DAF_ANALYST_ROLE` y warehouse `DAF_LAB_WH`.
- Conexión configurada de Snowflake CLI llamada `da_foundations_lab`.
- Permisos para crear objetos en el esquema `CURATED` o, si no existen, permisos para crear una vista con sufijo individual.
- Permisos locales para usar Microsoft Excel para Microsoft 365 y Power BI Desktop.

> **Advertencia de seguridad y gobierno:** no ejecutes `UPDATE`, `DELETE`, `TRUNCATE`, `DROP` ni `ALTER` sobre objetos del esquema `RAW`. La tabla `RAW.VENTAS_TRANSACCIONES_2026_1` es una fuente inmutable.

## Entorno de Laboratorio

### Directorios obligatorios

Abre PowerShell y ejecuta los siguientes comandos antes de comenzar:

```powershell
New-Item -ItemType Directory -Force -Path `
  "C:\DAF\Batch_01\00_source", `
  "C:\DAF\Batch_01\01_brief", `
  "C:\DAF\Batch_01\02_quality", `
  "C:\DAF\Batch_01\03_descriptive", `
  "C:\DAF\Batch_01\04_exploration", `
  "C:\DAF\Batch_01\05_dashboard", `
  "C:\DAF\Batch_01\sql"
```

Copia o confirma la disponibilidad de los siguientes archivos:

| Archivo | Ubicación esperada | Uso |
|---|---|---|
| `01_brief_analitico_ventas.xlsx` | `C:\DAF\Batch_01\01_brief\` | Priorización de campos y métricas |
| `DATA_ANALYTICS_FOUNDATIONS_2026_1.xlsx` | `C:\DAF\Batch_01\00_source\` | Inspección y muestra operativa en Excel |
| `02_control_calidad_ventas.xlsx` | `C:\DAF\Batch_01\02_quality\` | Entregable de controles |
| `02_preparacion_ventas.pbix` | `C:\DAF\Batch_01\02_quality\` | Consulta parametrizada de Power BI |
| `02_00_01_preparacion_ventas.sql` | `C:\DAF\Batch_01\sql\` | Script SQL del laboratorio |

### Recursos técnicos

| Recurso | Configuración de referencia |
|---|---|
| Sistema operativo | Windows 11 Pro, 64 bits |
| Memoria | 16 GB mínimo; 32 GB recomendado |
| Espacio libre | 20–25 GB en SSD |
| Excel | Microsoft 365, 64 bits |
| Power BI Desktop | Versión de 64 bits |
| Snowflake CLI | 3.6.0 o compatible |
| Cuenta Snowflake | `<ORG_ACCOUNT>` |
| Base de datos | `DATA_ANALYTICS_FOUNDATIONS` |
| Esquemas | `RAW`, `CURATED`, `ANALYTICS` |
| Warehouse | `DAF_LAB_WH` |
| Rol | `DAF_ANALYST_ROLE` |

### Convención de campos

El laboratorio utiliza los nombres de referencia siguientes. Antes de ejecutar transformaciones, verifica los nombres reales de las columnas mediante `DESCRIBE TABLE`. Si tu fuente usa variantes como `FECHA_VENTA` en lugar de `FECHA_TRANSACCION`, adapta exclusivamente el mapeo de columnas y documenta el cambio.

| Campo de referencia | Uso analítico |
|---|---|
| `TRANSACCION_ID` | Identificador principal de la transacción |
| `CLIENTE_ID` | Segmentación y detección de duplicados de negocio |
| `PRODUCTO_ID` | Producto vendido |
| `FECHA_TRANSACCION` | Análisis temporal |
| `REGION` | Segmentación geográfica |
| `CANAL` | Segmentación comercial |
| `CATEGORIA` | Segmentación de producto |
| `CANTIDAD` | Unidades transaccionadas |
| `IMPORTE` | Importe de venta o devolución |
| `DESCUENTO` | Proporción de descuento esperada entre 0 y 1 |

## Instrucciones Paso a Paso

### Paso 1: Revisar el brief y definir criterios de calidad prioritarios

**Objective:** Convertir el brief analítico de la Práctica 1 en criterios verificables de calidad para la preparación del dataset.

1. Abre el archivo:

   ```text
   C:\DAF\Batch_01\01_brief\01_brief_analitico_ventas.xlsx
   ```

2. Identifica las preguntas analíticas definidas en la práctica anterior. Localiza especialmente las relacionadas con:
   - ventas o importe por período;
   - ventas por región;
   - ventas por canal;
   - ventas por categoría;
   - clientes, productos o transacciones;
   - descuentos, devoluciones o cantidades.

3. Crea el archivo `02_control_calidad_ventas.xlsx` en:

   ```text
   C:\DAF\Batch_01\02_quality\
   ```

4. Crea una hoja llamada `Criterios_brief` y registra una tabla con estas columnas:

   | Pregunta o decisión | Campo crítico | Riesgo de calidad | Regla propuesta | Tratamiento |
   |---|---|---|---|---|

5. Incluye, como mínimo, los siguientes criterios:

   | Pregunta o decisión | Campo crítico | Riesgo de calidad | Regla propuesta | Tratamiento |
   |---|---|---|---|---|
   | Tendencia mensual de ventas | `FECHA_TRANSACCION` | Texto, nulo o fecha fuera de cobertura | Convertir a `DATE`; revisar fechas inválidas | Marcar para revisión |
   | Ventas por región | `REGION` | Variantes de escritura | Normalizar espacios, mayúsculas y equivalencias | Estandarizar |
   | Ventas por canal | `CANAL` | Categorías inconsistentes | Aplicar catálogo de canales permitidos | Estandarizar o revisar |
   | Ventas por categoría | `CATEGORIA` | Valores nulos o variantes | Normalizar y marcar desconocidos | Estandarizar o revisar |
   | Unidades vendidas | `CANTIDAD` | Cero o negativa en venta no devuelta | Validar cantidad mayor que cero | Marcar para revisión |
   | Impacto de descuentos | `DESCUENTO` | Valores fuera de 0 a 1 | Validar rango permitido | Marcar para revisión |
   | Conteo de transacciones | `TRANSACCION_ID` | Repetición del identificador | Identificar repetidos | Marcar para revisión |

6. Agrega una nota debajo de la tabla:

   > Los registros sospechosos no se eliminan automáticamente. La versión curada conserva la fila, una bandera de calidad y una razón de revisión, salvo que exista evidencia de un duplicado exacto técnicamente repetido.

**Expected output:** El libro `02_control_calidad_ventas.xlsx` contiene una hoja `Criterios_brief` con reglas vinculadas a decisiones analíticas concretas.

**Verification:** Confirma que cada métrica o segmentación del brief tiene al menos un campo crítico y una regla de calidad asociada.

---

### Paso 2: Inspeccionar la estructura y la granularidad de las fuentes

**Objective:** Confirmar qué representa cada fila, revisar columnas y comparar las versiones Excel y Snowflake sin modificar la fuente RAW.

1. Abre `DATA_ANALYTICS_FOUNDATIONS_2026_1.xlsx` desde:

   ```text
   C:\DAF\Batch_01\00_source\DATA_ANALYTICS_FOUNDATIONS_2026_1.xlsx
   ```

2. Identifica la hoja o tabla que contiene las transacciones. No conviertas manualmente todo el archivo ni copies las 750,000 filas a una hoja tradicional si el libro excede la capacidad de respuesta de Excel.

3. Registra en una nueva hoja llamada `Perfil_inicial`:
   - nombre de la fuente;
   - fecha de consulta;
   - número de columnas;
   - campos disponibles;
   - tipo detectado;
   - hipótesis de granularidad.

4. Escribe una hipótesis inicial como esta, adaptándola a la evidencia:

   > Hipótesis de granularidad: cada fila representa una transacción comercial identificada por `TRANSACCION_ID`. Esta hipótesis debe validarse comprobando si el identificador es único y si un mismo identificador aparece asociado a múltiples productos o líneas.

5. Abre PowerShell y valida la conexión a Snowflake:

   ```powershell
   snow sql -c da_foundations_lab -q "SELECT CURRENT_ROLE(), CURRENT_WAREHOUSE(), CURRENT_DATABASE(), CURRENT_SCHEMA();"
   ```

6. Inspecciona la estructura de la tabla fuente:

   ```powershell
   snow sql -c da_foundations_lab -q "USE ROLE DAF_ANALYST_ROLE; USE WAREHOUSE DAF_LAB_WH; DESCRIBE TABLE DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1;"
   ```

7. Guarda el resultado de la estructura en un archivo de evidencia:

   ```powershell
   snow sql -c da_foundations_lab -q "DESCRIBE TABLE DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1;" | Out-File -Encoding utf8 "C:\DAF\Batch_01\02_quality\02_describe_raw.txt"
   ```

8. Ejecuta una muestra limitada. Ajusta el nombre de las columnas si difiere del convenio del laboratorio:

   ```sql
   SELECT
       TRANSACCION_ID,
       CLIENTE_ID,
       PRODUCTO_ID,
       FECHA_TRANSACCION,
       REGION,
       CANAL,
       CATEGORIA,
       CANTIDAD,
       IMPORTE,
       DESCUENTO
   FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1
   LIMIT 20;
   ```

**Expected output:** Se dispone de una hipótesis explícita de granularidad, una lista de columnas y el archivo `02_describe_raw.txt`.

**Verification:** Comprueba que los campos requeridos por el brief pueden localizarse en la tabla RAW o que documentaste sus nombres equivalentes.

---

### Paso 3: Ejecutar el perfilado inicial en Snowflake

**Objective:** Medir filas, completitud, valores distintos, cobertura temporal, rangos y categorías antes de realizar transformaciones.

1. Crea el archivo:

   ```text
   C:\DAF\Batch_01\sql\02_00_01_preparacion_ventas.sql
   ```

2. Agrega el siguiente bloque inicial al archivo. Ejecuta cada consulta desde Snowsight o Snowflake CLI.

   ```sql
   USE ROLE DAF_ANALYST_ROLE;
   USE WAREHOUSE DAF_LAB_WH;
   USE DATABASE DATA_ANALYTICS_FOUNDATIONS;
   USE SCHEMA RAW;

   SELECT
       COUNT(*) AS FILAS_TOTALES,
       COUNT(DISTINCT TRANSACCION_ID) AS TRANSACCIONES_DISTINTAS,
       COUNT(DISTINCT CLIENTE_ID) AS CLIENTES_DISTINTOS,
       COUNT(DISTINCT PRODUCTO_ID) AS PRODUCTOS_DISTINTOS
   FROM VENTAS_TRANSACCIONES_2026_1;
   ```

3. Registra los resultados en la hoja `Perfil_inicial`. Compara el volumen con el rango esperado de 500,000 a 1,000,000 de transacciones y con la referencia aproximada de 750,000 filas.

4. Ejecuta el perfil de completitud. Esta consulta considera `NULL`, valores vacíos y algunos marcadores textuales de ausencia:

   ```sql
   SELECT
       COUNT(*) AS FILAS_TOTALES,

       COUNT_IF(TRANSACCION_ID IS NULL OR TRIM(TRANSACCION_ID) = '') AS NULOS_TRANSACCION_ID,
       COUNT_IF(CLIENTE_ID IS NULL OR TRIM(CLIENTE_ID) = '') AS NULOS_CLIENTE_ID,
       COUNT_IF(PRODUCTO_ID IS NULL OR TRIM(PRODUCTO_ID) = '') AS NULOS_PRODUCTO_ID,
       COUNT_IF(FECHA_TRANSACCION IS NULL) AS NULOS_FECHA,
       COUNT_IF(REGION IS NULL OR TRIM(REGION) IN ('', '-', 'N/A', 'NA')) AS NULOS_REGION,
       COUNT_IF(CANAL IS NULL OR TRIM(CANAL) IN ('', '-', 'N/A', 'NA')) AS NULOS_CANAL,
       COUNT_IF(CATEGORIA IS NULL OR TRIM(CATEGORIA) IN ('', '-', 'N/A', 'NA')) AS NULOS_CATEGORIA,
       COUNT_IF(CANTIDAD IS NULL) AS NULOS_CANTIDAD,
       COUNT_IF(IMPORTE IS NULL) AS NULOS_IMPORTE,
       COUNT_IF(DESCUENTO IS NULL) AS NULOS_DESCUENTO
   FROM VENTAS_TRANSACCIONES_2026_1;
   ```

5. Ejecuta el perfil de fechas y métricas:

   ```sql
   SELECT
       MIN(TRY_TO_DATE(FECHA_TRANSACCION)) AS FECHA_MINIMA,
       MAX(TRY_TO_DATE(FECHA_TRANSACCION)) AS FECHA_MAXIMA,
       COUNT_IF(TRY_TO_DATE(FECHA_TRANSACCION) IS NULL) AS FECHAS_NO_CONVERTIBLES,
       MIN(CANTIDAD) AS CANTIDAD_MINIMA,
       MAX(CANTIDAD) AS CANTIDAD_MAXIMA,
       MIN(IMPORTE) AS IMPORTE_MINIMO,
       MAX(IMPORTE) AS IMPORTE_MAXIMO,
       MIN(DESCUENTO) AS DESCUENTO_MINIMO,
       MAX(DESCUENTO) AS DESCUENTO_MAXIMO
   FROM VENTAS_TRANSACCIONES_2026_1;
   ```

6. Ejecuta un perfil de categorías. Si el tipo de dato ya es fecha, puedes omitir conversiones adicionales.

   ```sql
   SELECT 'REGION' AS CAMPO, UPPER(TRIM(REGION)) AS VALOR_NORMALIZADO, COUNT(*) AS FILAS
   FROM VENTAS_TRANSACCIONES_2026_1
   GROUP BY 1, 2

   UNION ALL

   SELECT 'CANAL' AS CAMPO, UPPER(TRIM(CANAL)) AS VALOR_NORMALIZADO, COUNT(*) AS FILAS
   FROM VENTAS_TRANSACCIONES_2026_1
   GROUP BY 1, 2

   UNION ALL

   SELECT 'CATEGORIA' AS CAMPO, UPPER(TRIM(CATEGORIA)) AS VALOR_NORMALIZADO, COUNT(*) AS FILAS
   FROM VENTAS_TRANSACCIONES_2026_1
   GROUP BY 1, 2
   ORDER BY CAMPO, FILAS DESC;
   ```

7. En Excel, crea una hoja `Hallazgos_iniciales` con estas columnas:

   | Prioridad | Hecho observado | Interpretación provisional | Impacto analítico | Acción de validación |
   |---|---|---|---|---|

8. Registra al menos cinco hallazgos. No afirmes que una anomalía es un error sin evidencia. Usa formulaciones como:
   - “Se observaron variantes de escritura en `REGION`; podrían fragmentar la segmentación.”
   - “Existen importes negativos; podrían representar devoluciones.”
   - “Existen descuentos fuera del rango esperado; requieren revisión.”
   - “`TRANSACCION_ID` presenta repeticiones; se debe validar si son duplicados o detalle legítimo.”

**Expected output:** Un perfil inicial cuantitativo documentado en Excel y consultas reproducibles guardadas en el archivo SQL.

**Verification:** Verifica que registraste filas totales, valores distintos, nulos, fechas mínima y máxima, rangos de métricas y categorías frecuentes.

---

### Paso 4: Detectar duplicados exactos, repetición de identificadores y duplicados de negocio

**Objective:** Diferenciar entre registros técnicamente repetidos, identificadores repetidos y coincidencias que requieren revisión de negocio.

1. Ejecuta la detección de `TRANSACCION_ID` repetidos:

   ```sql
   SELECT
       TRANSACCION_ID,
       COUNT(*) AS REPETICIONES
   FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1
   WHERE TRANSACCION_ID IS NOT NULL
   GROUP BY TRANSACCION_ID
   HAVING COUNT(*) > 1
   ORDER BY REPETICIONES DESC, TRANSACCION_ID;
   ```

2. No asumas que todas las repeticiones son duplicados exactos. Revisa una muestra de identificadores repetidos:

   ```sql
   SELECT
       TRANSACCION_ID,
       CLIENTE_ID,
       PRODUCTO_ID,
       FECHA_TRANSACCION,
       REGION,
       CANAL,
       CATEGORIA,
       CANTIDAD,
       IMPORTE,
       DESCUENTO
   FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1
   WHERE TRANSACCION_ID IN (
       SELECT TRANSACCION_ID
       FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1
       WHERE TRANSACCION_ID IS NOT NULL
       GROUP BY TRANSACCION_ID
       HAVING COUNT(*) > 1
   )
   ORDER BY TRANSACCION_ID
   LIMIT 100;
   ```

3. Detecta posibles duplicados de negocio. La coincidencia se define, para fines de revisión, por cliente, producto, fecha, importe y canal:

   ```sql
   SELECT
       CLIENTE_ID,
       PRODUCTO_ID,
       TRY_TO_DATE(FECHA_TRANSACCION) AS FECHA_TRANSACCION_DATE,
       IMPORTE,
       UPPER(TRIM(CANAL)) AS CANAL_NORMALIZADO,
       COUNT(*) AS REPETICIONES
   FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1
   GROUP BY
       CLIENTE_ID,
       PRODUCTO_ID,
       TRY_TO_DATE(FECHA_TRANSACCION),
       IMPORTE,
       UPPER(TRIM(CANAL))
   HAVING COUNT(*) > 1
   ORDER BY REPETICIONES DESC;
   ```

4. Registra la siguiente interpretación en la hoja `Hallazgos_iniciales`:

   > Una coincidencia por cliente, producto, fecha, importe y canal es un posible duplicado de negocio, no una prueba concluyente. Un cliente puede adquirir el mismo producto más de una vez el mismo día por el mismo importe y canal.

5. Agrega una hoja llamada `Muestra_revision` en `02_control_calidad_ventas.xlsx`.

6. Exporta una muestra pequeña de registros sospechosos desde Snowflake para revisión operativa. Si tu versión de Snowflake CLI permite redirección de salida, usa:

   ```powershell
   snow sql -c da_foundations_lab -q "
   SELECT
       TRANSACCION_ID,
       CLIENTE_ID,
       PRODUCTO_ID,
       FECHA_TRANSACCION,
       REGION,
       CANAL,
       CATEGORIA,
       CANTIDAD,
       IMPORTE,
       DESCUENTO
   FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1
   WHERE TRANSACCION_ID IN (
       SELECT TRANSACCION_ID
       FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1
       GROUP BY TRANSACCION_ID
       HAVING COUNT(*) > 1
   )
   LIMIT 200;" | Out-File -Encoding utf8 "C:\DAF\Batch_01\02_quality\02_muestra_ids_repetidos.txt"
   ```

7. Documenta en la hoja `Muestra_revision` que el archivo es una evidencia de inspección y no una extracción oficial completa.

**Expected output:** Existen consultas separadas para IDs repetidos y posibles duplicados de negocio; sus diferencias están documentadas.

**Verification:** Confirma que ninguna consulta de posibles duplicados se utilizó para eliminar filas automáticamente.

---

### Paso 5: Crear la vista o tabla curada en Snowflake

**Objective:** Aplicar reglas de estandarización y generar una fuente curada con trazabilidad de la calidad de cada registro.

1. En el archivo SQL, agrega el siguiente script. Si no tienes permiso para crear una tabla, crea una vista. Si no tienes permiso para usar el nombre común, agrega tu sufijo, por ejemplo `_STUDENT01`.

2. Ajusta los valores del catálogo según las categorías reales observadas en el Paso 3. El ejemplo trata variantes frecuentes sin inventar categorías de negocio.

   ```sql
   USE ROLE DAF_ANALYST_ROLE;
   USE WAREHOUSE DAF_LAB_WH;
   USE DATABASE DATA_ANALYTICS_FOUNDATIONS;
   USE SCHEMA CURATED;

   CREATE OR REPLACE VIEW VENTAS_TRANSACCIONES_CURADAS_2026_1 AS
   WITH fuente AS (
       SELECT
           TRANSACCION_ID,
           CLIENTE_ID,
           PRODUCTO_ID,
           TRY_TO_DATE(FECHA_TRANSACCION) AS FECHA_TRANSACCION,
           REGION,
           CANAL,
           CATEGORIA,
           CANTIDAD,
           IMPORTE,
           DESCUENTO
       FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1
   ),
   estandarizada AS (
       SELECT
           TRANSACCION_ID,
           CLIENTE_ID,
           PRODUCTO_ID,
           FECHA_TRANSACCION,

           CASE
               WHEN REGION IS NULL OR TRIM(REGION) IN ('', '-', 'N/A', 'NA') THEN 'SIN_REGION'
               WHEN UPPER(TRIM(REGION)) IN ('NORTE', 'ZONA NORTE') THEN 'NORTE'
               WHEN UPPER(TRIM(REGION)) IN ('CENTRO', 'ZONA CENTRO') THEN 'CENTRO'
               WHEN UPPER(TRIM(REGION)) IN ('SUR', 'ZONA SUR') THEN 'SUR'
               WHEN UPPER(TRIM(REGION)) IN ('ORIENTE', 'ESTE', 'ZONA ORIENTE') THEN 'ORIENTE'
               WHEN UPPER(TRIM(REGION)) IN ('OCCIDENTE', 'OESTE', 'ZONA OCCIDENTE') THEN 'OCCIDENTE'
               ELSE UPPER(TRIM(REGION))
           END AS REGION_NORMALIZADA,

           CASE
               WHEN CANAL IS NULL OR TRIM(CANAL) IN ('', '-', 'N/A', 'NA') THEN 'SIN_CANAL'
               WHEN UPPER(TRIM(CANAL)) IN ('TIENDA', 'TIENDA FISICA', 'FISICO', 'FÍSICO') THEN 'TIENDA'
               WHEN UPPER(TRIM(CANAL)) IN ('ONLINE', 'E-COMMERCE', 'ECOMMERCE', 'WEB') THEN 'ONLINE'
               WHEN UPPER(TRIM(CANAL)) IN ('MARKETPLACE', 'MARKET PLACE') THEN 'MARKETPLACE'
               ELSE UPPER(TRIM(CANAL))
           END AS CANAL_NORMALIZADO,

           CASE
               WHEN CATEGORIA IS NULL OR TRIM(CATEGORIA) IN ('', '-', 'N/A', 'NA') THEN 'SIN_CATEGORIA'
               ELSE UPPER(TRIM(CATEGORIA))
           END AS CATEGORIA_NORMALIZADA,

           CANTIDAD,
           IMPORTE,
           DESCUENTO,

           COUNT(*) OVER (PARTITION BY TRANSACCION_ID) AS REPETICIONES_TRANSACCION_ID,

           COUNT(*) OVER (
               PARTITION BY
                   CLIENTE_ID,
                   PRODUCTO_ID,
                   FECHA_TRANSACCION,
                   IMPORTE,
                   UPPER(TRIM(CANAL))
           ) AS REPETICIONES_POTENCIAL_DUPLICADO_NEGOCIO
       FROM fuente
   )
   SELECT
       TRANSACCION_ID,
       CLIENTE_ID,
       PRODUCTO_ID,
       FECHA_TRANSACCION,
       REGION_NORMALIZADA AS REGION,
       CANAL_NORMALIZADO AS CANAL,
       CATEGORIA_NORMALIZADA AS CATEGORIA,
       CANTIDAD,
       IMPORTE,
       DESCUENTO,

       CASE
           WHEN TRANSACCION_ID IS NULL OR TRIM(TRANSACCION_ID) = '' THEN 'REVISAR'
           WHEN FECHA_TRANSACCION IS NULL THEN 'REVISAR'
           WHEN FECHA_TRANSACCION < '2024-01-01'::DATE
                OR FECHA_TRANSACCION > '2025-12-31'::DATE THEN 'REVISAR'
           WHEN REGION_NORMALIZADA = 'SIN_REGION' THEN 'REVISAR'
           WHEN CANAL_NORMALIZADO = 'SIN_CANAL' THEN 'REVISAR'
           WHEN CATEGORIA_NORMALIZADA = 'SIN_CATEGORIA' THEN 'REVISAR'
           WHEN CANTIDAD IS NULL THEN 'REVISAR'
           WHEN CANTIDAD <= 0 AND COALESCE(IMPORTE, 0) >= 0 THEN 'REVISAR'
           WHEN DESCUENTO IS NOT NULL AND (DESCUENTO < 0 OR DESCUENTO > 1) THEN 'REVISAR'
           WHEN REPETICIONES_TRANSACCION_ID > 1 THEN 'REVISAR'
           WHEN REPETICIONES_POTENCIAL_DUPLICADO_NEGOCIO > 1 THEN 'REVISAR'
           ELSE 'VALIDO'
       END AS BANDERA_CALIDAD,

       CONCAT_WS(
           '; ',
           IFF(TRANSACCION_ID IS NULL OR TRIM(TRANSACCION_ID) = '', 'TRANSACCION_ID_FALTANTE', NULL),
           IFF(FECHA_TRANSACCION IS NULL, 'FECHA_NO_CONVERTIBLE_O_FALTANTE', NULL),
           IFF(
               FECHA_TRANSACCION < '2024-01-01'::DATE
               OR FECHA_TRANSACCION > '2025-12-31'::DATE,
               'FECHA_FUERA_DE_COBERTURA',
               NULL
           ),
           IFF(REGION_NORMALIZADA = 'SIN_REGION', 'REGION_FALTANTE', NULL),
           IFF(CANAL_NORMALIZADO = 'SIN_CANAL', 'CANAL_FALTANTE', NULL),
           IFF(CATEGORIA_NORMALIZADA = 'SIN_CATEGORIA', 'CATEGORIA_FALTANTE', NULL),
           IFF(CANTIDAD IS NULL, 'CANTIDAD_FALTANTE', NULL),
           IFF(CANTIDAD <= 0 AND COALESCE(IMPORTE, 0) >= 0, 'CANTIDAD_NO_VALIDA_PARA_VENTA', NULL),
           IFF(DESCUENTO IS NOT NULL AND (DESCUENTO < 0 OR DESCUENTO > 1), 'DESCUENTO_FUERA_DE_RANGO', NULL),
           IFF(REPETICIONES_TRANSACCION_ID > 1, 'TRANSACCION_ID_REPETIDO', NULL),
           IFF(
               REPETICIONES_POTENCIAL_DUPLICADO_NEGOCIO > 1,
               'POSIBLE_DUPLICADO_NEGOCIO',
               NULL
           )
       ) AS RAZON_REVISION,

       REPETICIONES_TRANSACCION_ID,
       REPETICIONES_POTENCIAL_DUPLICADO_NEGOCIO
   FROM estandarizada;
   ```

3. Si la creación falla por privilegios, intenta crear una vista con un nombre individual:

   ```sql
   CREATE OR REPLACE VIEW VENTAS_TRANSACCIONES_CURADAS_2026_1_STUDENT01 AS
   -- pegar aquí el mismo SELECT final de la vista anterior
   SELECT * FROM VENTAS_TRANSACCIONES_CURADAS_2026_1;
   ```

   Si tampoco puedes crear objetos, guarda el `SELECT` completo como evidencia en el archivo SQL y consulta al instructor antes de continuar.

4. Consulta el resumen de calidad:

   ```sql
   SELECT
       BANDERA_CALIDAD,
       COUNT(*) AS FILAS,
       ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) AS PORCENTAJE
   FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
   GROUP BY BANDERA_CALIDAD
   ORDER BY BANDERA_CALIDAD;
   ```

5. Consulta las razones de revisión más frecuentes:

   ```sql
   SELECT
       RAZON_REVISION,
       COUNT(*) AS FILAS
   FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
   WHERE BANDERA_CALIDAD = 'REVISAR'
   GROUP BY RAZON_REVISION
   ORDER BY FILAS DESC
   LIMIT 20;
   ```

6. Registra en Excel la decisión de tratamiento:

   | Tipo de hallazgo | Tratamiento aplicado | Justificación |
   |---|---|---|
   | Variantes de región, canal y categoría | Estandarización | Evita fragmentación de segmentos |
   | Fechas inválidas o fuera de cobertura | Conservar y marcar | Puede afectar comparabilidad temporal |
   | Cantidad no válida para venta | Conservar y marcar | Puede representar devolución o error operativo |
   | Descuento fuera de 0 a 1 | Conservar y marcar | Requiere confirmar escala o captura |
   | ID repetido | Conservar y marcar | Debe confirmarse granularidad |
   | Posible duplicado de negocio | Conservar y marcar | La coincidencia no prueba duplicidad |

**Expected output:** Una vista curada en `CURATED` con campos normalizados, `BANDERA_CALIDAD`, `RAZON_REVISION` y conteos de repetición.

**Verification:** Ejecuta:

```sql
SELECT *
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
LIMIT 20;
```

Confirma que existen las columnas `BANDERA_CALIDAD` y `RAZON_REVISION`, y que las categorías aparecen normalizadas.

---

### Paso 6: Documentar controles y una muestra operativa en Excel

**Objective:** Crear evidencia legible de los controles aplicados sin intentar usar Excel como motor principal para procesar todo el dataset masivo.

1. En `02_control_calidad_ventas.xlsx`, crea una hoja llamada `Resumen_controles`.

2. Agrega la siguiente estructura:

   | Control | Dimensión de calidad | Consulta o regla | Resultado observado | Decisión |
   |---|---|---|---|---|

3. Documenta al menos estos controles:
   - conteo de filas;
   - nulos en campos críticos;
   - fechas mínima y máxima;
   - fechas no convertibles o fuera de cobertura;
   - cantidad no válida;
   - descuento fuera de rango;
   - categorías de región, canal y categoría;
   - `TRANSACCION_ID` repetido;
   - posibles duplicados de negocio;
   - distribución de `BANDERA_CALIDAD`.

4. Crea una hoja llamada `Muestra_curada`.

5. Desde Snowflake, exporta una muestra balanceada de registros válidos y en revisión. Puedes usar esta consulta:

   ```sql
   SELECT *
   FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
   QUALIFY ROW_NUMBER() OVER (
       PARTITION BY BANDERA_CALIDAD
       ORDER BY TRANSACCION_ID
   ) <= 100;
   ```

6. Copia el resultado a `Muestra_curada` como valores. Agrega filtros y convierte el rango en una tabla de Excel con el nombre `tbl_muestra_curada`.

7. Crea una hoja `Catalogo_normalizacion` con tres secciones:
   - Región original y región normalizada.
   - Canal original y canal normalizado.
   - Categoría original y categoría normalizada.

8. Añade esta nota metodológica:

   > Excel contiene una muestra operativa y los controles documentados. La fuente de análisis masiva y reproducible es la vista o tabla curada en Snowflake. No se considera equivalente copiar manualmente subconjuntos no controlados.

**Expected output:** Un libro Excel con criterios, perfil inicial, hallazgos, resumen de controles, catálogo de normalización y muestra curada.

**Verification:** Confirma que una persona distinta puede identificar, a partir del libro, qué reglas se aplicaron y cuál es la fuente curada oficial.

---

### Paso 7: Crear una consulta parametrizada en Power BI Desktop

**Objective:** Conectar Power BI a la fuente curada de Snowflake y conservar una consulta de Power Query reutilizable y trazable.

1. Abre Power BI Desktop.

2. Selecciona **Obtener datos** > **Snowflake**.

3. Configura la conexión:
   - **Servidor:** `<ORG_ACCOUNT>.snowflakecomputing.com`
   - **Warehouse:** `DAF_LAB_WH`
   - **Base de datos:** `DATA_ANALYTICS_FOUNDATIONS`

4. Autentícate con tu usuario institucional. No guardes contraseñas en archivos de texto, scripts, PBIX o XLSX.

5. En el navegador de objetos, selecciona:

   ```text
   CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
   ```

   Si usaste un sufijo individual, selecciona tu objeto curado correspondiente.

6. Selecciona **Transformar datos** en lugar de cargar directamente.

7. En Power Query, cambia el nombre de la consulta a:

   ```text
   Ventas_Curadas
   ```

8. Crea un parámetro:
   - Nombre: `pIncluirRegistrosRevision`
   - Tipo: Lógico
   - Valor actual: `false`

9. En el Editor avanzado, conserva el origen generado por Power BI y agrega un filtro condicionado. El patrón esperado es similar al siguiente; Power BI puede generar nombres de pasos diferentes:

   ```powerquery
   let
       Source = Snowflake.Databases(
           "<ORG_ACCOUNT>.snowflakecomputing.com",
           "DAF_LAB_WH"
       ),
       Database = Source{[Name="DATA_ANALYTICS_FOUNDATIONS"]}[Data],
       CuratedSchema = Database{[Name="CURATED"]}[Data],
       VentasCuradas = CuratedSchema{[Name="VENTAS_TRANSACCIONES_CURADAS_2026_1"]}[Data],
       Tipos = Table.TransformColumnTypes(
           VentasCuradas,
           {
               {"FECHA_TRANSACCION", type date},
               {"CANTIDAD", Int64.Type},
               {"IMPORTE", type number},
               {"DESCUENTO", type number},
               {"BANDERA_CALIDAD", type text},
               {"RAZON_REVISION", type text}
           }
       ),
       FilasAnaliticas = if pIncluirRegistrosRevision
           then Tipos
           else Table.SelectRows(Tipos, each [BANDERA_CALIDAD] = "VALIDO")
   in
       FilasAnaliticas
   ```

10. Si el conector genera navegación diferente, conserva esa navegación y aplica únicamente la lógica de tipos y filtro sobre `BANDERA_CALIDAD`.

11. Selecciona **Cerrar y aplicar**.

12. Crea una página llamada `Control de calidad` y agrega:
   - una tarjeta con el conteo de filas;
   - un gráfico de barras con `BANDERA_CALIDAD`;
   - una tabla con `RAZON_REVISION` y conteo de filas;
   - un segmentador de `REGION`;
   - un segmentador de `CANAL`.

13. Guarda el archivo como:

   ```text
   C:\DAF\Batch_01\02_quality\02_preparacion_ventas.pbix
   ```

**Expected output:** Un archivo PBIX con una consulta llamada `Ventas_Curadas`, un parámetro `pIncluirRegistrosRevision` y una página de control de calidad.

**Verification:** Cambia temporalmente `pIncluirRegistrosRevision` a `true`, actualiza la vista previa y confirma que aparecen registros con `BANDERA_CALIDAD = "REVISAR"`. Restablece el valor a `false` antes de guardar la versión analítica predeterminada.

## Validación y Pruebas

Completa las siguientes validaciones antes de entregar el laboratorio.

| Prueba | Consulta o acción | Resultado esperado |
|---|---|---|
| Existencia de objeto curado | `SHOW VIEWS LIKE 'VENTAS_TRANSACCIONES_CURADAS_2026_1' IN SCHEMA DATA_ANALYTICS_FOUNDATIONS.CURATED;` | La vista existe o está documentada con sufijo individual |
| Conservación de filas | Comparar `COUNT(*)` de RAW y CURATED | El conteo debe mantenerse si no se eliminaron duplicados exactos |
| Bandera de calidad | Agrupar por `BANDERA_CALIDAD` | Existen al menos las categorías `VALIDO` y/o `REVISAR` según los hallazgos |
| Razones de revisión | Filtrar `BANDERA_CALIDAD = 'REVISAR'` | Cada registro en revisión tiene una razón no vacía |
| Fechas | Revisar `MIN` y `MAX` de `FECHA_TRANSACCION` | Cobertura esperada entre 2024-01-01 y 2025-12-31; excepciones marcadas |
| Cantidad | Filtrar `CANTIDAD <= 0 AND IMPORTE >= 0` | Los casos sospechosos están marcados para revisión |
| Descuento | Filtrar `DESCUENTO < 0 OR DESCUENTO > 1` | Los casos fuera de rango están marcados para revisión |
| Normalización | Agrupar por `REGION`, `CANAL`, `CATEGORIA` | Las variantes conocidas fueron consolidadas |
| Trazabilidad Excel | Revisar `02_control_calidad_ventas.xlsx` | Contiene criterios, perfil, hallazgos, controles y muestra |
| Trazabilidad Power BI | Abrir `02_preparacion_ventas.pbix` | La consulta usa la fuente `CURATED` y el parámetro funciona |

Ejecuta además esta consulta de control final:

```sql
SELECT
    COUNT(*) AS FILAS_TOTALES,
    COUNT_IF(BANDERA_CALIDAD = 'VALIDO') AS FILAS_VALIDAS,
    COUNT_IF(BANDERA_CALIDAD = 'REVISAR') AS FILAS_EN_REVISION,
    COUNT_IF(BANDERA_CALIDAD = 'REVISAR' AND RAZON_REVISION IS NULL) AS REVISION_SIN_RAZON
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1;
```

El valor de `REVISION_SIN_RAZON` debe ser `0`.

## Solución de Problemas

### Problema 1: Power BI tarda demasiado, se queda sin memoria o no responde al abrir el archivo Excel masivo

**Síntomas:** Excel o Power BI Desktop permanece sin responder, la actualización consume gran cantidad de memoria o el archivo tarda excesivamente en cargar.

**Causa probable:** Intento de cargar las aproximadamente 750,000 transacciones desde Excel en una hoja tradicional o de importar datos sin filtrar desde una fuente local innecesaria.

**Solución:**
1. Usa Snowflake como fuente principal para el procesamiento masivo.
2. En Excel, conserva controles y muestras operativas, no una copia manual completa.
3. En Power BI, conecta a la vista `CURATED` y aplica el filtro por `BANDERA_CALIDAD` en Power Query.
4. Cierra aplicaciones no necesarias y verifica que Power BI y Excel sean versiones de 64 bits.
5. Si es necesario, trabaja inicialmente con una muestra limitada para validar la lógica y actualiza el conjunto completo solo al finalizar.

### Problema 2: La creación de la vista en CURATED falla con un error de privilegios

**Síntomas:** Snowflake devuelve mensajes como `SQL access control error`, `insufficient privileges` o no permite crear la vista con el nombre común.

**Causa probable:** El rol asignado tiene permisos de lectura sobre RAW pero no permisos de creación en `CURATED`, o el nombre común ya está reservado para otro estudiante.

**Solución:**
1. Confirma el rol y warehouse activos:

   ```sql
   SELECT CURRENT_ROLE(), CURRENT_WAREHOUSE(), CURRENT_DATABASE(), CURRENT_SCHEMA();
   ```

2. Ejecuta:

   ```sql
   USE ROLE DAF_ANALYST_ROLE;
   USE WAREHOUSE DAF_LAB_WH;
   USE DATABASE DATA_ANALYTICS_FOUNDATIONS;
   USE SCHEMA CURATED;
   ```

3. Intenta crear la vista con un sufijo individual, por ejemplo:

   ```text
   VENTAS_TRANSACCIONES_CURADAS_2026_1_STUDENT01
   ```

4. Si el problema persiste, conserva el script completo en `C:\DAF\Batch_01\sql\02_00_01_preparacion_ventas.sql`, guarda la evidencia del error y solicita al instructor permisos de creación o un esquema personal autorizado.
5. No intentes resolver el problema modificando objetos del esquema `RAW`.

## Limpieza

1. Guarda todos los archivos abiertos:
   - `02_control_calidad_ventas.xlsx`
   - `02_preparacion_ventas.pbix`
   - `02_00_01_preparacion_ventas.sql`

2. Verifica que los entregables estén ubicados en:

   ```text
   C:\DAF\Batch_01\02_quality\
   C:\DAF\Batch_01\sql\
   ```

3. Cierra Excel y Power BI Desktop para liberar memoria y evitar bloqueos de archivos.

4. No elimines ni modifiques:
   - `DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1`
   - archivos fuente ubicados en `00_source`;
   - archivos del brief ubicados en `01_brief`.

5. Si creaste un objeto curado de práctica, consérvalo para las Prácticas 3, 4 y 5. No ejecutes `DROP` sobre la vista o tabla curada salvo instrucción explícita del instructor.

## Resumen

En esta práctica perfilaste un dataset artificial masivo de ventas y documentaste su condición inicial antes de usarlo para análisis descriptivos o visualizaciones. Confirmaste que la granularidad, los tipos de dato, la completitud, los rangos, las categorías y la unicidad afectan directamente la validez de cualquier conclusión comercial.

También creaste una fuente curada que normaliza campos de segmentación, convierte fechas, identifica condiciones sospechosas y conserva trazabilidad mediante `BANDERA_CALIDAD` y `RAZON_REVISION`. Esta base será la fuente común para los análisis de desempeño de las prácticas siguientes y para el dashboard de ventas.

**Entregables mínimos:**

- `C:\DAF\Batch_01\02_quality\02_control_calidad_ventas.xlsx`
- `C:\DAF\Batch_01\02_quality\02_preparacion_ventas.pbix`
- `C:\DAF\Batch_01\sql\02_00_01_preparacion_ventas.sql`
- Vista o tabla `CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1`, o versión con sufijo individual documentado.

**Principio clave:** un registro marcado para revisión no es automáticamente un registro inválido. La calidad de datos exige distinguir entre evidencia observada, interpretación provisional y decisión de tratamiento respaldada por el contexto de negocio.
