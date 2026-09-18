# Análisis descriptivo de ventas y clientes

## Metadatos

| Duration | Complexity | Bloom level |
|---|---|---|
| 116 minutos | Difícil | Aplicar |

## Descripción General

En esta práctica se realiza un análisis descriptivo sobre la fuente curada de ventas creada en la Práctica 2. Se definirá formalmente la población, la observación, la granularidad y las variables analíticas; posteriormente se calcularán medidas descriptivas, frecuencias, proporciones, tasas y probabilidades comerciales en Snowflake y Excel. Finalmente, se contrastarán resultados entre ambas herramientas y se documentarán límites de interpretación, sesgos potenciales y la diferencia entre correlación, asociación y causalidad.

## Objetivos de Aprendizaje

- [ ] Clasificar variables del dataset como categóricas, numéricas, identificadoras, temporales o indicadores.
- [ ] Definir población, observación, unidad de análisis, granularidad y posible muestra para las preguntas comerciales planteadas.
- [ ] Calcular e interpretar medidas de tendencia central, dispersión, cuartiles, percentiles, frecuencias, proporciones y tasas.
- [ ] Estimar probabilidades simples y condicionales relacionadas con descuentos, recurrencia y devoluciones.
- [ ] Contrastar resultados obtenidos en Excel y Snowflake, documentando diferencias metodológicas, riesgos de sesgo y límites de causalidad.

## Prerrequisitos

**Conocimientos requeridos**

- Haber completado la Práctica 2 y conocer las reglas de calidad aplicadas a la fuente curada.
- Comprender las métricas, segmentos y preguntas definidos en `01_brief_analitico_ventas.xlsx`.
- Manejar funciones de Excel tales como `PROMEDIO`, `MEDIANA`, `MIN`, `MAX`, `DESVEST.M`, `PERCENTIL.INC`, `CONTAR.SI.CONJUNTO` y tablas dinámicas.
- Manejar agregaciones SQL en Snowflake: `COUNT`, `COUNT_IF`, `AVG`, `MEDIAN`, `STDDEV`, `MIN`, `MAX`, `CORR` y `APPROX_PERCENTILE`.
- Comprender que una correlación estadística mide asociación lineal, pero no demuestra que una variable cause cambios en otra.

**Acceso requerido**

- Acceso de lectura al objeto curado:

```text
DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
```

- Si el objeto común no existe o no está disponible, usar la tabla o vista asignada por el instructor, por ejemplo:

```text
DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1_STUDENT01
```

- Rol `DAF_ANALYST_ROLE`.
- Warehouse `DAF_LAB_WH`.
- Acceso a Snowflake mediante Snowsight, Snowflake CLI o SnowSQL.
- Microsoft Excel para Microsoft 365, 64 bits.
- Directorio de trabajo `C:\DAF\Batch_01\`.

> **Restricción importante:** no ejecute `UPDATE`, `DELETE`, `TRUNCATE`, `DROP` ni `ALTER` sobre objetos del esquema `RAW`. Esta práctica es de lectura y análisis; no modifica la fuente curada ni la fuente RAW.

## Entorno de Laboratorio

| Componente | Configuración de referencia |
|---|---|
| Sistema operativo | Windows 11 Pro, 64 bits |
| Memoria RAM | 16 GB mínimo; 32 GB recomendados |
| Almacenamiento | Al menos 20 GB libres en SSD |
| Excel | Microsoft Excel para Microsoft 365, 64 bits |
| Snowflake | Snowsight o Snowflake CLI/SnowSQL |
| Warehouse | `DAF_LAB_WH` |
| Rol | `DAF_ANALYST_ROLE` |
| Base de datos | `DATA_ANALYTICS_FOUNDATIONS` |
| Esquemas | `RAW`, `CURATED`, `ANALYTICS` |
| Fuente de análisis | Tabla o vista curada de la Práctica 2 |
| Dataset | Ventas Retail LATAM 2026.1, artificial, entre 500,000 y 1,000,000 de transacciones |

1. Abra PowerShell y cree las carpetas requeridas si aún no existen:

```powershell
New-Item -ItemType Directory -Force `
  -Path C:\DAF\Batch_01\00_source,
        C:\DAF\Batch_01\01_brief,
        C:\DAF\Batch_01\02_quality,
        C:\DAF\Batch_01\03_descriptive,
        C:\DAF\Batch_01\04_exploration,
        C:\DAF\Batch_01\05_dashboard,
        C:\DAF\Batch_01\sql
```

2. Cree los siguientes archivos de trabajo:

```text
C:\DAF\Batch_01\sql\03_analisis_descriptivo_ventas.sql
C:\DAF\Batch_01\03_descriptive\03_analisis_descriptivo_ventas_clientes.xlsx
C:\DAF\Batch_01\03_descriptive\03_diccionario_y_alcance.md
C:\DAF\Batch_01\03_descriptive\03_validacion_excel_snowflake.md
```

3. Inicie una sesión de Snowflake. Si utiliza SnowSQL o Snowflake CLI, use la conexión definida para el curso sin incluir credenciales en archivos:

```bash
snow sql -c da_foundations_lab
```

4. Establezca el contexto de trabajo en Snowflake:

```sql
USE ROLE DAF_ANALYST_ROLE;
USE WAREHOUSE DAF_LAB_WH;
USE DATABASE DATA_ANALYTICS_FOUNDATIONS;
USE SCHEMA CURATED;
```

> **Nota sobre nombres de columnas:** esta guía utiliza nombres de referencia tales como `ID_TRANSACCION`, `ID_CLIENTE`, `FECHA_TRANSACCION`, `REGION`, `CANAL`, `CATEGORIA`, `SEGMENTO_CLIENTE`, `VENTA_NETA`, `CANTIDAD`, `DESCUENTO_PCT` y `ES_DEVOLUCION`. Verifique los nombres reales de su objeto curado durante el Paso 2 y sustitúyalos de forma consistente si difieren.

## Instrucciones Paso a Paso

### Paso 1: Preparar el expediente de análisis y revisar el brief

**Objetivo:** organizar los entregables y alinear el análisis descriptivo con la pregunta analítica, las métricas y los segmentos acordados.

1. Abra el archivo del brief creado en la práctica anterior:

```text
C:\DAF\Batch_01\01_brief\01_brief_analitico_ventas.xlsx
```

2. Identifique y registre en una hoja nueva del libro de esta práctica, llamada `Alcance`, los siguientes elementos:

| Elemento | Definición para esta práctica |
|---|---|
| Periodo de análisis | Del `2024-01-01` al `2025-12-31` |
| Fuente | Tabla o vista curada de la Práctica 2 |
| Población | Todas las transacciones aceptadas incluidas en la fuente curada durante el periodo 2024-2025 |
| Unidad de análisis | Una transacción comercial aceptada |
| Observación | Una fila que representa una transacción comercial, tras validar la granularidad |
| Clientes | Clientes únicos identificados mediante `ID_CLIENTE` |
| Segmentos principales | Región, canal, categoría y segmento de cliente |
| Métricas principales | Venta neta, cantidad, descuento, devolución y frecuencia de compra por cliente |

3. Documente una definición operativa para cada métrica. Use el siguiente formato como referencia:

| Métrica | Definición operativa | Numerador | Denominador | Unidad |
|---|---|---|---|---|
| Venta neta | Importe neto de la transacción después de descuentos y ajustes definidos en la fuente curada | No aplica | No aplica | Moneda del dataset |
| Cantidad | Unidades incluidas en la transacción | No aplica | No aplica | Unidades |
| Transacción con descuento | Transacción con descuento mayor que cero | Conteo de transacciones con descuento | Conteo de transacciones aceptadas | Proporción |
| Tasa de devoluciones | Participación de transacciones clasificadas como devolución | Conteo de devoluciones | Conteo de transacciones aceptadas | Porcentaje |
| Cliente recurrente | Cliente con dos o más transacciones aceptadas en el periodo | Clientes con frecuencia >= 2 | Clientes únicos | Proporción |

4. Registre explícitamente que el dataset es artificial y que no contiene datos personales reales.

5. Guarde el libro como:

```text
C:\DAF\Batch_01\03_descriptive\03_analisis_descriptivo_ventas_clientes.xlsx
```

**Resultado esperado**

Un libro de Excel creado con la hoja `Alcance`, incluyendo población, unidad de análisis, observación, periodo, segmentos y definiciones operativas de métricas.

**Verificación**

- La población no se define como “todos los clientes de LATAM”, sino como las transacciones aceptadas presentes en la fuente curada y dentro del periodo analizado.
- La observación se define como una transacción, no como un cliente.
- La definición de “cliente recurrente” incluye un umbral explícito: dos o más transacciones.

---

### Paso 2: Inspeccionar la estructura, las variables y la granularidad de la fuente curada

**Objetivo:** confirmar que la tabla curada tiene la estructura esperada y clasificar las variables antes de calcular estadísticas.

1. Ejecute la siguiente consulta para identificar las columnas reales de la fuente curada:

```sql
DESCRIBE TABLE DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1;
```

2. Si trabaja con una vista, use:

```sql
DESCRIBE VIEW DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1;
```

3. Consulte una muestra de registros:

```sql
SELECT *
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
LIMIT 20;
```

4. Cree una hoja llamada `Diccionario` en Excel y documente las variables disponibles. Use esta clasificación como modelo y adapte los nombres de campo reales:

| Variable de referencia | Tipo | Clasificación analítica | Uso descriptivo |
|---|---|---|---|
| `ID_TRANSACCION` | Texto o entero | Identificadora | Conteo de transacciones y validación de unicidad |
| `ID_CLIENTE` | Texto o entero | Identificadora | Conteo de clientes únicos y frecuencia de compra |
| `FECHA_TRANSACCION` | Fecha | Temporal | Cobertura del periodo y análisis temporal posterior |
| `REGION` | Texto | Categórica nominal | Frecuencias, proporciones y comparación de segmentos |
| `CANAL` | Texto | Categórica nominal | Frecuencias, proporciones y probabilidad condicional |
| `CATEGORIA` | Texto | Categórica nominal | Frecuencias y distribución de ventas |
| `SEGMENTO_CLIENTE` | Texto | Categórica nominal u ordinal | Frecuencias y comparación de grupos |
| `VENTA_NETA` | Decimal | Numérica continua | Tendencia central, dispersión y percentiles |
| `CANTIDAD` | Entero | Numérica discreta | Tendencia central, dispersión y correlación |
| `DESCUENTO_PCT` | Decimal | Numérica continua | Tendencia central, percentiles y tasa de descuento |
| `ES_DEVOLUCION` | Booleano, texto o entero | Indicador binario | Tasa de devoluciones |

5. Verifique la cobertura temporal, el volumen y la unicidad de la clave transaccional:

```sql
SELECT
    COUNT(*) AS filas_curadas,
    COUNT(DISTINCT ID_TRANSACCION) AS transacciones_distintas,
    COUNT(DISTINCT ID_CLIENTE) AS clientes_distintos,
    MIN(FECHA_TRANSACCION) AS fecha_minima,
    MAX(FECHA_TRANSACCION) AS fecha_maxima
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1;
```

6. Evalúe si la tabla está realmente a nivel de transacción. Compare filas y transacciones distintas:

```sql
SELECT
    COUNT(*) AS filas,
    COUNT(DISTINCT ID_TRANSACCION) AS id_transaccion_distintos,
    COUNT(*) - COUNT(DISTINCT ID_TRANSACCION) AS diferencia
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1;
```

7. Si `filas` es igual a `id_transaccion_distintos`, registre que la granularidad es una fila por transacción.

8. Si existen varias filas por `ID_TRANSACCION`, no trate cada fila como una transacción completa. Documente que la granularidad puede ser línea de detalle, producto por pedido u otra unidad. Solicite confirmación al instructor o agregue primero a nivel de transacción antes de continuar.

**Resultado esperado**

Un diccionario de datos clasificado y una validación documentada de volumen, clientes, periodo y granularidad.

**Verificación**

- La fecha mínima debe ser cercana a `2024-01-01` y la máxima cercana a `2025-12-31`.
- El número de filas debe encontrarse dentro del orden esperado para el dataset artificial masivo, entre 500,000 y 1,000,000 de transacciones.
- La diferencia entre filas y transacciones distintas debe ser cero si el objeto se encuentra a nivel transacción.
- Si hay diferencia, el análisis debe declarar explícitamente la granularidad real antes de usar conteos.

---

### Paso 3: Definir población, observación, muestra y límites de cobertura

**Objetivo:** formular el alcance estadístico del análisis sin extender conclusiones más allá de la fuente disponible.

1. Cree el archivo:

```text
C:\DAF\Batch_01\03_descriptive\03_diccionario_y_alcance.md
```

2. Incluya una sección con el siguiente contenido, adaptando el nombre del objeto utilizado:

```markdown
### Definición de población y observación

- Población analizada: todas las transacciones aceptadas de la fuente curada
  DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
  durante el periodo 2024-01-01 a 2025-12-31.
- Unidad de análisis: transacción comercial aceptada.
- Observación: una fila de la fuente curada, siempre que la validación de granularidad
  confirme una fila por ID_TRANSACCION.
- Unidad secundaria de análisis: cliente único identificado por ID_CLIENTE.
- Muestra: no se utiliza una muestra para los cálculos principales; se analiza la población
  disponible en la fuente curada. Cualquier extracción parcial para Excel debe identificarse
  como extracto operativo y no como una muestra aleatoria.
```

3. Documente las exclusiones de calidad heredadas de la Práctica 2. Incluya, como mínimo:

- Registros excluidos por identificador de transacción nulo o inválido.
- Registros excluidos por fecha fuera del periodo esperado.
- Registros excluidos por valores no válidos de venta neta, cantidad, canal, región u otros campos críticos.
- Duplicados eliminados o tratados.
- Reglas aplicadas a devoluciones, cancelaciones o transacciones de importe negativo.

4. Añada una advertencia de cobertura:

```markdown
Los resultados describen exclusivamente las transacciones aceptadas disponibles en la fuente
curada. No permiten inferir comportamiento de ventas fuera del periodo, de canales no incluidos,
de clientes no registrados o de poblaciones reales, debido a que el dataset es sintético.
```

5. Guarde también un resumen de cobertura calculado en Snowflake:

```sql
SELECT
    DATE_TRUNC('MONTH', FECHA_TRANSACCION) AS mes,
    COUNT(*) AS transacciones,
    COUNT(DISTINCT ID_CLIENTE) AS clientes_unicos,
    SUM(VENTA_NETA) AS venta_neta_total
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
GROUP BY 1
ORDER BY 1;
```

6. Revise si todos los meses entre enero de 2024 y diciembre de 2025 aparecen en el resultado. Si falta uno o más meses, marque el periodo como incompleto.

**Resultado esperado**

Un documento de alcance que diferencia correctamente población, muestra, unidad de análisis, observación y cobertura temporal.

**Verificación**

- La práctica no presenta una extracción de Excel como si fuera automáticamente representativa.
- Las conclusiones se limitan a la población curada disponible.
- Se identifica si existen periodos incompletos o meses sin registros.

---

### Paso 4: Calcular estadísticas descriptivas de venta neta, cantidad y descuento en Snowflake

**Objetivo:** calcular medidas de tendencia central, dispersión, posición y extremos sobre las variables numéricas de transacción.

1. Abra o cree el archivo:

```text
C:\DAF\Batch_01\sql\03_analisis_descriptivo_ventas.sql
```

2. Pegue y ejecute la siguiente consulta. Sustituya los nombres de columnas si son distintos en su objeto:

```sql
WITH base AS (
    SELECT
        VENTA_NETA,
        CANTIDAD,
        DESCUENTO_PCT
    FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
    WHERE FECHA_TRANSACCION >= '2024-01-01'
      AND FECHA_TRANSACCION < '2026-01-01'
)
SELECT
    'VENTA_NETA' AS variable,
    COUNT(VENTA_NETA) AS n_no_nulo,
    AVG(VENTA_NETA) AS media,
    MEDIAN(VENTA_NETA) AS mediana_p50,
    MIN(VENTA_NETA) AS minimo,
    MAX(VENTA_NETA) AS maximo,
    MAX(VENTA_NETA) - MIN(VENTA_NETA) AS rango,
    STDDEV(VENTA_NETA) AS desviacion_estandar_muestral,
    APPROX_PERCENTILE(VENTA_NETA, 0.25) AS p25,
    APPROX_PERCENTILE(VENTA_NETA, 0.50) AS p50_aproximado,
    APPROX_PERCENTILE(VENTA_NETA, 0.75) AS p75,
    APPROX_PERCENTILE(VENTA_NETA, 0.95) AS p95
FROM base

UNION ALL

SELECT
    'CANTIDAD' AS variable,
    COUNT(CANTIDAD) AS n_no_nulo,
    AVG(CANTIDAD) AS media,
    MEDIAN(CANTIDAD) AS mediana_p50,
    MIN(CANTIDAD) AS minimo,
    MAX(CANTIDAD) AS maximo,
    MAX(CANTIDAD) - MIN(CANTIDAD) AS rango,
    STDDEV(CANTIDAD) AS desviacion_estandar_muestral,
    APPROX_PERCENTILE(CANTIDAD, 0.25) AS p25,
    APPROX_PERCENTILE(CANTIDAD, 0.50) AS p50_aproximado,
    APPROX_PERCENTILE(CANTIDAD, 0.75) AS p75,
    APPROX_PERCENTILE(CANTIDAD, 0.95) AS p95
FROM base

UNION ALL

SELECT
    'DESCUENTO_PCT' AS variable,
    COUNT(DESCUENTO_PCT) AS n_no_nulo,
    AVG(DESCUENTO_PCT) AS media,
    MEDIAN(DESCUENTO_PCT) AS mediana_p50,
    MIN(DESCUENTO_PCT) AS minimo,
    MAX(DESCUENTO_PCT) AS maximo,
    MAX(DESCUENTO_PCT) - MIN(DESCUENTO_PCT) AS rango,
    STDDEV(DESCUENTO_PCT) AS desviacion_estandar_muestral,
    APPROX_PERCENTILE(DESCUENTO_PCT, 0.25) AS p25,
    APPROX_PERCENTILE(DESCUENTO_PCT, 0.50) AS p50_aproximado,
    APPROX_PERCENTILE(DESCUENTO_PCT, 0.75) AS p75,
    APPROX_PERCENTILE(DESCUENTO_PCT, 0.95) AS p95
FROM base
ORDER BY variable;
```

3. Interprete los resultados según estas reglas:

| Medida | Interpretación |
|---|---|
| Media | Promedio afectado por valores extremos |
| Mediana o P50 | Valor central; menos sensible a valores extremos |
| Mínimo y máximo | Límites observados, no necesariamente límites de negocio |
| Rango | Amplitud entre máximo y mínimo |
| Desviación estándar | Dispersión respecto a la media |
| P25 | Valor por debajo del cual se encuentra aproximadamente 25 % de los registros |
| P75 | Valor por debajo del cual se encuentra aproximadamente 75 % de los registros |
| P95 | Umbral alto; aproximadamente 95 % de los registros está en o por debajo de este valor |

4. Determine si la media y mediana de `VENTA_NETA` son muy diferentes. Si la media es notablemente superior a la mediana, registre una posible asimetría positiva: unas pocas ventas altas pueden elevar el promedio.

5. Evalúe la pertinencia de calcular la moda. Para `VENTA_NETA`, normalmente no es útil si el campo tiene muchos valores decimales diferentes. Para `CANTIDAD`, sí puede ser útil porque es una variable discreta.

6. Calcule una moda exploratoria para cantidad:

```sql
SELECT
    CANTIDAD,
    COUNT(*) AS frecuencia
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
WHERE CANTIDAD IS NOT NULL
GROUP BY CANTIDAD
ORDER BY frecuencia DESC, CANTIDAD
LIMIT 10;
```

7. Copie los resultados principales a una hoja de Excel llamada `Estadisticos_Snowflake`.

**Resultado esperado**

Una tabla con media, mediana, mínimo, máximo, rango, desviación estándar y percentiles para venta neta, cantidad y descuento.

**Verificación**

- Cada variable tiene un conteo no nulo documentado.
- Los percentiles están ordenados: `P25 <= P50 <= P75 <= P95`.
- La desviación estándar no se interpreta como una venta típica; representa dispersión.
- La moda de venta neta se marca como “no aplicable” si los importes decimales no se repiten de manera significativa.

---

### Paso 5: Calcular frecuencia de compra y recurrencia de clientes en Snowflake

**Objetivo:** cambiar correctamente la unidad de análisis desde transacción hacia cliente para describir la frecuencia de compra.

1. Ejecute la siguiente consulta para obtener una fila por cliente:

```sql
WITH frecuencia_cliente AS (
    SELECT
        ID_CLIENTE,
        COUNT(DISTINCT ID_TRANSACCION) AS frecuencia_compra,
        SUM(VENTA_NETA) AS venta_neta_cliente,
        MIN(FECHA_TRANSACCION) AS primera_compra,
        MAX(FECHA_TRANSACCION) AS ultima_compra
    FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
    WHERE FECHA_TRANSACCION >= '2024-01-01'
      AND FECHA_TRANSACCION < '2026-01-01'
      AND ID_CLIENTE IS NOT NULL
    GROUP BY ID_CLIENTE
)
SELECT
    COUNT(*) AS clientes_con_compra,
    AVG(frecuencia_compra) AS media_frecuencia,
    MEDIAN(frecuencia_compra) AS mediana_frecuencia,
    MIN(frecuencia_compra) AS minimo_frecuencia,
    MAX(frecuencia_compra) AS maximo_frecuencia,
    MAX(frecuencia_compra) - MIN(frecuencia_compra) AS rango_frecuencia,
    STDDEV(frecuencia_compra) AS desviacion_estandar_frecuencia,
    APPROX_PERCENTILE(frecuencia_compra, 0.25) AS p25_frecuencia,
    APPROX_PERCENTILE(frecuencia_compra, 0.50) AS p50_frecuencia,
    APPROX_PERCENTILE(frecuencia_compra, 0.75) AS p75_frecuencia,
    APPROX_PERCENTILE(frecuencia_compra, 0.95) AS p95_frecuencia,
    COUNT_IF(frecuencia_compra >= 2) AS clientes_recurrentes,
    COUNT_IF(frecuencia_compra >= 2) / NULLIF(COUNT(*), 0) AS probabilidad_cliente_recurrente
FROM frecuencia_cliente;
```

2. Ejecute una distribución de frecuencias de compra para identificar concentraciones:

```sql
WITH frecuencia_cliente AS (
    SELECT
        ID_CLIENTE,
        COUNT(DISTINCT ID_TRANSACCION) AS frecuencia_compra
    FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
    WHERE FECHA_TRANSACCION >= '2024-01-01'
      AND FECHA_TRANSACCION < '2026-01-01'
      AND ID_CLIENTE IS NOT NULL
    GROUP BY ID_CLIENTE
)
SELECT
    CASE
        WHEN frecuencia_compra = 1 THEN '1 compra'
        WHEN frecuencia_compra = 2 THEN '2 compras'
        WHEN frecuencia_compra BETWEEN 3 AND 5 THEN '3 a 5 compras'
        WHEN frecuencia_compra BETWEEN 6 AND 10 THEN '6 a 10 compras'
        ELSE '11 o más compras'
    END AS tramo_frecuencia,
    COUNT(*) AS clientes,
    COUNT(*) / SUM(COUNT(*)) OVER () AS proporcion_clientes
FROM frecuencia_cliente
GROUP BY 1
ORDER BY
    CASE tramo_frecuencia
        WHEN '1 compra' THEN 1
        WHEN '2 compras' THEN 2
        WHEN '3 a 5 compras' THEN 3
        WHEN '6 a 10 compras' THEN 4
        ELSE 5
    END;
```

3. Registre el cambio de granularidad en la hoja `Alcance`:

```text
Para el análisis de frecuencia de compra, la unidad de análisis es el cliente.
Cada observación corresponde a un cliente único y su frecuencia representa el número
de transacciones aceptadas durante el periodo 2024-2025.
```

4. Interprete cuidadosamente la probabilidad de recurrencia:

```text
P(cliente recurrente) = clientes con dos o más transacciones / clientes únicos con al menos una transacción.
```

5. No interprete esta métrica como probabilidad futura de recompra. Es una proporción observada dentro del periodo histórico disponible.

**Resultado esperado**

Un resumen descriptivo de frecuencia de compra por cliente y una estimación de la proporción de clientes recurrentes.

**Verificación**

- El denominador de recurrencia es el número de clientes únicos, no el número de transacciones.
- La frecuencia mínima debe ser al menos 1.
- El conteo de clientes recurrentes no puede superar el total de clientes con compra.

---

### Paso 6: Calcular frecuencias, proporciones y tasas comerciales en Snowflake

**Objetivo:** describir la distribución de transacciones por segmentos y calcular tasas operativas comparables.

1. Calcule frecuencias y proporciones por región:

```sql
SELECT
    REGION,
    COUNT(*) AS transacciones,
    COUNT(*) / SUM(COUNT(*)) OVER () AS proporcion_transacciones,
    COUNT(DISTINCT ID_CLIENTE) AS clientes_unicos,
    SUM(VENTA_NETA) AS venta_neta_total,
    AVG(VENTA_NETA) AS venta_neta_promedio
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
GROUP BY REGION
ORDER BY transacciones DESC;
```

2. Repita la consulta para `CANAL`, `CATEGORIA` y `SEGMENTO_CLIENTE`. Guarde cada resultado en una hoja o pestaña distinta en Excel:

```sql
SELECT
    CANAL,
    COUNT(*) AS transacciones,
    COUNT(*) / SUM(COUNT(*)) OVER () AS proporcion_transacciones,
    COUNT(DISTINCT ID_CLIENTE) AS clientes_unicos,
    SUM(VENTA_NETA) AS venta_neta_total,
    AVG(VENTA_NETA) AS venta_neta_promedio
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
GROUP BY CANAL
ORDER BY transacciones DESC;
```

3. Calcule las tasas globales de descuento y devolución. Ajuste la condición de devolución según el tipo real de campo. Si `ES_DEVOLUCION` es booleano, use la condición mostrada:

```sql
SELECT
    COUNT(*) AS transacciones_totales,
    COUNT_IF(DESCUENTO_PCT > 0) AS transacciones_con_descuento,
    COUNT_IF(DESCUENTO_PCT > 0) / NULLIF(COUNT(*), 0) AS tasa_descuento,
    COUNT_IF(ES_DEVOLUCION = TRUE) AS transacciones_devolucion,
    COUNT_IF(ES_DEVOLUCION = TRUE) / NULLIF(COUNT(*), 0) AS tasa_devoluciones
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1;
```

4. Si `ES_DEVOLUCION` usa valores como `SI`/`NO`, `1`/`0` o un estado de transacción, sustituya la condición. Ejemplos:

```sql
COUNT_IF(ES_DEVOLUCION = 'SI')
```

```sql
COUNT_IF(ES_DEVOLUCION = 1)
```

```sql
COUNT_IF(ESTADO_TRANSACCION = 'DEVOLUCION')
```

5. Calcule la tasa de devoluciones por canal:

```sql
SELECT
    CANAL,
    COUNT(*) AS transacciones,
    COUNT_IF(ES_DEVOLUCION = TRUE) AS devoluciones,
    COUNT_IF(ES_DEVOLUCION = TRUE) / NULLIF(COUNT(*), 0) AS tasa_devoluciones
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
GROUP BY CANAL
ORDER BY tasa_devoluciones DESC;
```

6. Registre las tasas como porcentajes con dos decimales en Excel. No compare tasas sin revisar también el número de transacciones del segmento.

**Resultado esperado**

Tablas de frecuencia y proporción por región, canal, categoría y segmento de cliente, junto con tasas globales y por canal de descuentos y devoluciones.

**Verificación**

- Las proporciones por cada dimensión deben sumar aproximadamente 1.00 o 100 %.
- La tasa de descuento debe estar entre 0 y 1.
- La tasa de devoluciones debe estar entre 0 y 1.
- Todo segmento con una tasa extrema debe revisarse junto con su tamaño de base.

---

### Paso 7: Calcular probabilidades simples y condicionales

**Objetivo:** traducir eventos comerciales definidos en el brief a probabilidades observadas con numerador y denominador explícitos.

1. Calcule la probabilidad simple de observar una transacción con descuento:

```sql
SELECT
    COUNT_IF(DESCUENTO_PCT > 0) AS eventos_con_descuento,
    COUNT(*) AS transacciones_totales,
    COUNT_IF(DESCUENTO_PCT > 0) / NULLIF(COUNT(*), 0) AS p_venta_con_descuento
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1;
```

2. Calcule la probabilidad condicional de devolución dado que el canal es digital. Sustituya `'Digital'` por el valor real de su dataset, por ejemplo `'Online'`, `'E-commerce'` o `'Web'`:

```sql
SELECT
    COUNT_IF(CANAL = 'Digital') AS transacciones_canal_digital,
    COUNT_IF(CANAL = 'Digital' AND ES_DEVOLUCION = TRUE) AS devoluciones_canal_digital,
    COUNT_IF(CANAL = 'Digital' AND ES_DEVOLUCION = TRUE)
        / NULLIF(COUNT_IF(CANAL = 'Digital'), 0) AS p_devolucion_dado_canal_digital
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1;
```

3. Si no conoce los valores exactos de `CANAL`, inspecciónelos primero:

```sql
SELECT
    CANAL,
    COUNT(*) AS transacciones
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
GROUP BY CANAL
ORDER BY transacciones DESC;
```

4. Registre las fórmulas de probabilidad en la hoja `Probabilidades`:

| Evento | Fórmula | Interpretación correcta |
|---|---|---|
| P(venta con descuento) | Transacciones con descuento / transacciones totales | Proporción observada de transacciones con descuento |
| P(cliente recurrente) | Clientes con 2 o más compras / clientes únicos | Proporción observada de clientes recurrentes |
| P(devolución \| canal digital) | Devoluciones digitales / transacciones digitales | Tasa de devolución dentro del canal digital |

5. Añada una nota metodológica:

```text
Estas probabilidades son frecuencias relativas observadas en la población curada del periodo.
No son predicciones causales ni estimaciones universales fuera del alcance del dataset.
```

**Resultado esperado**

Tres probabilidades documentadas con eventos, numeradores, denominadores, periodo y fuente.

**Verificación**

- El denominador de la probabilidad condicional incluye solo transacciones del canal digital.
- La probabilidad de cliente recurrente usa una tabla agregada a nivel cliente.
- Las probabilidades se comunican como proporciones observadas, no como certezas futuras.

---

### Paso 8: Importar la fuente curada en Excel y construir estadísticas equivalentes

**Objetivo:** obtener resultados comparables en Excel utilizando la misma versión curada del dataset.

1. Abra el libro:

```text
C:\DAF\Batch_01\03_descriptive\03_analisis_descriptivo_ventas_clientes.xlsx
```

2. En Excel, seleccione:

```text
Datos > Obtener datos > Desde base de datos > Desde Snowflake
```

3. Configure la conexión con los valores institucionales:

| Campo | Valor |
|---|---|
| Servidor | Cuenta Snowflake institucional, por ejemplo `<ORG_ACCOUNT>.snowflakecomputing.com` |
| Warehouse | `DAF_LAB_WH` |
| Base de datos | `DATA_ANALYTICS_FOUNDATIONS` |
| Esquema | `CURATED` |
| Tabla o vista | `VENTAS_TRANSACCIONES_CURADAS_2026_1` |
| Rol, si se solicita | `DAF_ANALYST_ROLE` |

4. Autentíquese con su usuario institucional. No guarde contraseñas en el libro Excel.

5. En Power Query, mantenga únicamente las columnas requeridas para esta práctica:

```text
ID_TRANSACCION
ID_CLIENTE
FECHA_TRANSACCION
REGION
CANAL
CATEGORIA
SEGMENTO_CLIENTE
VENTA_NETA
CANTIDAD
DESCUENTO_PCT
ES_DEVOLUCION
```

6. Cargue la consulta en una tabla de Excel llamada `Ventas`. Si el equipo tiene recursos limitados, seleccione cargar al Modelo de datos y utilice tablas dinámicas. Si carga los 750,000 registros a una hoja, confirme que no supera el límite de 1,048,576 filas de Excel.

7. Cree dos columnas auxiliares en la tabla `Ventas`:

| Columna | Fórmula de ejemplo |
|---|---|
| `ConDescuento` | `=--([@DESCUENTO_PCT]>0)` |
| `DevolucionFlag` | `=--([@ES_DEVOLUCION]=VERDADERO)` |

8. Si el campo de devolución utiliza texto, adapte la segunda fórmula:

```excel
=--([@ES_DEVOLUCION]="SI")
```

9. Cree una hoja llamada `Estadisticos_Excel` y construya una tabla de medidas para `VENTA_NETA`, `CANTIDAD` y `DESCUENTO_PCT`.

| Medida | Fórmula para `VENTA_NETA` |
|---|---|
| Conteo no nulo | `=CONTAR(Ventas[VENTA_NETA])` |
| Media | `=PROMEDIO(Ventas[VENTA_NETA])` |
| Mediana | `=MEDIANA(Ventas[VENTA_NETA])` |
| Mínimo | `=MIN(Ventas[VENTA_NETA])` |
| Máximo | `=MAX(Ventas[VENTA_NETA])` |
| Rango | `=MAX(Ventas[VENTA_NETA])-MIN(Ventas[VENTA_NETA])` |
| Desviación estándar muestral | `=DESVEST.M(Ventas[VENTA_NETA])` |
| P25 | `=PERCENTIL.INC(Ventas[VENTA_NETA];0,25)` |
| P50 | `=PERCENTIL.INC(Ventas[VENTA_NETA];0,5)` |
| P75 | `=PERCENTIL.INC(Ventas[VENTA_NETA];0,75)` |
| P95 | `=PERCENTIL.INC(Ventas[VENTA_NETA];0,95)` |

10. Para `CANTIDAD`, calcule una moda solo si existe repetición relevante:

```excel
=MODA.UNO(Ventas[CANTIDAD])
```

11. Si Excel devuelve `#N/A` al calcular una moda de `VENTA_NETA`, registre “No aplicable: valores decimales sin repetición significativa”. No fuerce una interpretación.

12. Calcule tasas globales en Excel:

| Métrica | Fórmula |
|---|---|
| Total de transacciones | `=FILAS(Ventas[ID_TRANSACCION])` |
| Transacciones con descuento | `=SUMA(Ventas[ConDescuento])` |
| Tasa de descuento | `=SUMA(Ventas[ConDescuento])/FILAS(Ventas[ID_TRANSACCION])` |
| Devoluciones | `=SUMA(Ventas[DevolucionFlag])` |
| Tasa de devoluciones | `=SUMA(Ventas[DevolucionFlag])/FILAS(Ventas[ID_TRANSACCION])` |

13. Aplique formato de porcentaje a las tasas y formato numérico apropiado a ventas, cantidades y percentiles.

**Resultado esperado**

Un libro Excel conectado a la fuente curada, con medidas descriptivas y tasas calculadas sobre la misma población utilizada en Snowflake.

**Verificación**

- El conteo de transacciones en Excel coincide con `COUNT(*)` de Snowflake, salvo que exista un filtro documentado.
- La media y mediana de `VENTA_NETA` son cercanas a los resultados de Snowflake.
- Las diferencias menores en P25, P75 o P95 son aceptables cuando Snowflake usa `APPROX_PERCENTILE` y Excel usa `PERCENTIL.INC`.

---

### Paso 9: Construir tablas dinámicas para segmentos y frecuencia de compra

**Objetivo:** calcular en Excel frecuencias, proporciones y frecuencia de compra por cliente sin confundir transacciones con clientes.

1. Cree una hoja llamada `Frecuencias_Excel`.

2. Inserte una tabla dinámica basada en la tabla `Ventas`.

3. Para analizar región:

- Filas: `REGION`.
- Valores: conteo de `ID_TRANSACCION`.
- Valores adicionales: suma de `VENTA_NETA`.
- Valores adicionales: promedio de `VENTA_NETA`.

4. Para obtener proporciones de transacciones:

- Agregue nuevamente `ID_TRANSACCION` al área Valores.
- Seleccione **Mostrar valores como**.
- Elija **% del total general**.
- Renombre el campo como `ProporcionTransacciones`.

5. Repita el procedimiento para:

```text
CANAL
CATEGORIA
SEGMENTO_CLIENTE
```

6. Cree una tabla dinámica para tasa de descuentos por canal:

- Filas: `CANAL`.
- Valores: promedio de `ConDescuento`.
- Formatee el resultado como porcentaje.
- Agregue también conteo de `ID_TRANSACCION`.

7. Cree una tabla dinámica para tasa de devoluciones por canal:

- Filas: `CANAL`.
- Valores: promedio de `DevolucionFlag`.
- Formatee como porcentaje.
- Agregue conteo de `ID_TRANSACCION`.

8. Cree una hoja llamada `Frecuencia_cliente_Excel`.

9. Inserte una tabla dinámica con la siguiente configuración:

- Filas: `ID_CLIENTE`.
- Valores: conteo de `ID_TRANSACCION`.

10. Copie los valores de la tabla dinámica y péguelos como valores en una tabla llamada `FrecuenciaCliente`, con dos columnas:

```text
ID_CLIENTE
FrecuenciaCompra
```

11. En una celda de resumen, calcule la proporción de clientes recurrentes:

```excel
=CONTAR.SI(FrecuenciaCliente[FrecuenciaCompra];">=2")/CONTAR(FrecuenciaCliente[FrecuenciaCompra])
```

12. Calcule estadísticas descriptivas de frecuencia de compra:

```excel
=PROMEDIO(FrecuenciaCliente[FrecuenciaCompra])
```

```excel
=MEDIANA(FrecuenciaCliente[FrecuenciaCompra])
```

```excel
=PERCENTIL.INC(FrecuenciaCliente[FrecuenciaCompra];0,95)
```

13. Registre que la tabla `FrecuenciaCliente` tiene granularidad de cliente y que la tabla `Ventas` tiene granularidad de transacción.

**Resultado esperado**

Tablas dinámicas que muestran frecuencias, proporciones, tasas por segmento y una tabla independiente con frecuencia de compra a nivel cliente.

**Verificación**

- Las proporciones por región, canal, categoría y segmento suman aproximadamente 100 %.
- La tasa de devolución por canal se calcula como promedio de un indicador binario, no como suma de importes.
- La recurrencia usa clientes únicos, no filas de la tabla transaccional.
- Los resultados de tasa de descuento y recurrencia son comparables con Snowflake.

---

### Paso 10: Contrastar resultados, evaluar correlación y documentar sesgos

**Objetivo:** validar métricas entre herramientas, evaluar la relación descuento-cantidad y comunicar límites de interpretación.

1. Cree el archivo:

```text
C:\DAF\Batch_01\03_descriptive\03_validacion_excel_snowflake.md
```

2. Incluya una tabla de contraste entre Snowflake y Excel. Complete con sus resultados:

| Métrica | Snowflake | Excel | Diferencia | Estado |
|---|---:|---:|---:|---|
| Número de transacciones |  |  |  | Coincide / Investigar |
| Clientes únicos |  |  |  | Coincide / Investigar |
| Media de venta neta |  |  |  | Coincide / Investigar |
| Mediana de venta neta |  |  |  | Coincide / Investigar |
| Tasa de descuento |  |  |  | Coincide / Investigar |
| Tasa de devoluciones |  |  |  | Coincide / Investigar |
| Probabilidad de cliente recurrente |  |  |  | Coincide / Investigar |

3. Considere aceptable una diferencia mínima en percentiles aproximados si Snowflake utilizó `APPROX_PERCENTILE`. Investigue diferencias en conteos, medias, medianas y tasas, porque normalmente deberían coincidir al usar la misma fuente y filtros.

4. Calcule la correlación entre descuento y cantidad en Snowflake:

```sql
SELECT
    COUNT(*) AS observaciones_validas,
    CORR(DESCUENTO_PCT, CANTIDAD) AS correlacion_descuento_cantidad
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
WHERE DESCUENTO_PCT IS NOT NULL
  AND CANTIDAD IS NOT NULL;
```

5. Calcule la misma correlación en Excel:

```excel
=COEF.DE.CORREL(Ventas[DESCUENTO_PCT];Ventas[CANTIDAD])
```

6. Interprete el coeficiente mediante una escala orientativa:

| Valor absoluto de r | Interpretación orientativa |
|---:|---|
| Cercano a 0 | Asociación lineal débil o inexistente |
| Entre 0.10 y 0.29 | Asociación lineal débil |
| Entre 0.30 y 0.49 | Asociación lineal moderada |
| 0.50 o superior | Asociación lineal relativamente fuerte |

7. Redacte una conclusión responsable. Use una estructura como la siguiente:

```markdown
Se observó una correlación de [valor] entre descuento y cantidad. Este resultado indica una
asociación lineal [débil/moderada/fuerte] en la población curada analizada. Sin embargo, no
demuestra que el descuento haya causado el cambio en cantidad, porque pueden intervenir
factores como categoría, campaña, estacionalidad, canal, perfil de cliente, disponibilidad
de inventario o reglas comerciales no observadas.
```

8. Documente al menos cuatro riesgos de sesgo o interpretación:

| Riesgo | Posible efecto | Mitigación o declaración |
|---|---|---|
| Exclusiones por calidad | La población curada puede diferir de la fuente RAW | Documentar reglas y volumen excluido |
| Periodo incompleto | Comparaciones temporales pueden ser engañosas | Confirmar cobertura mensual |
| Segmentos de tamaños distintos | Tasas extremas pueden provenir de bases pequeñas | Mostrar conteo y proporción junto a la tasa |
| Duplicidad o granularidad incorrecta | Conteos de clientes o transacciones pueden estar inflados | Validar unicidad de `ID_TRANSACCION` |
| Correlación interpretada como causalidad | Decisiones comerciales injustificadas | Declarar variables de confusión y requerir análisis adicional |
| Dataset sintético | Resultados no representan personas ni operaciones reales | Limitar conclusiones al entorno de práctica |

9. Guarde los archivos y prepare una lista breve de tres hallazgos candidatos para la Práctica 4. Cada hallazgo debe incluir:

- Métrica.
- Segmento o población.
- Periodo.
- Resultado numérico.
- Decisión potencial.
- Limitación o nivel de certeza.

10. Ejemplo de formato:

```markdown
Hallazgo candidato:
- Métrica: tasa de devoluciones.
- Segmento: canal digital.
- Periodo: 2024-01-01 a 2025-12-31.
- Resultado: [valor] %.
- Decisión potencial: revisar condiciones de devolución, información de producto o experiencia de compra digital.
- Limitación: la diferencia observada no identifica por sí sola la causa de las devoluciones.
```

**Resultado esperado**

Una validación documentada entre Excel y Snowflake, una correlación calculada y una lista de sesgos, límites y hallazgos candidatos.

**Verificación**

- Se contrastan al menos dos métricas entre Excel y Snowflake.
- La correlación calculada en ambas herramientas es consistente, salvo diferencias menores de redondeo.
- Ninguna conclusión afirma que el descuento causó un cambio en la cantidad solo por la correlación.
- Cada hallazgo incluye contexto temporal, segmento, métrica y limitación.

## Validación y Pruebas

Complete esta lista antes de cerrar la práctica.

| Prueba | Criterio de aceptación |
|---|---|
| Acceso a fuente curada | La consulta a la tabla o vista curada se ejecuta sin modificar datos |
| Cobertura temporal | Se documenta fecha mínima, fecha máxima y presencia o ausencia de meses incompletos |
| Granularidad | Se valida si una fila equivale a una transacción mediante `ID_TRANSACCION` |
| Población | Se define como transacciones aceptadas de 2024-2025 presentes en la fuente curada |
| Variables | Se clasifican variables numéricas, categóricas, temporales, identificadoras e indicadores |
| Estadísticas transaccionales | Se calculan media, mediana, mínimo, máximo, rango, desviación estándar, P25, P50, P75 y P95 |
| Moda | Se aplica solo cuando es interpretable, especialmente para cantidad |
| Frecuencia por cliente | Se calcula sobre una tabla agregada a nivel `ID_CLIENTE` |
| Frecuencias segmentadas | Se obtienen por región, canal, categoría y segmento de cliente |
| Tasas comerciales | Se calcula tasa de descuento y tasa de devoluciones con numerador y denominador explícitos |
| Probabilidades | Se calcula P(venta con descuento), P(cliente recurrente) y P(devolución \| canal digital) |
| Contraste de herramientas | Al menos dos resultados coinciden entre Excel y Snowflake |
| Correlación | Se calcula relación entre descuento y cantidad y se declara que no prueba causalidad |
| Sesgos | Se documentan exclusiones de calidad, periodos incompletos y segmentos con bases diferentes |
| Entregables | Los archivos SQL, Excel, Markdown de alcance y validación están guardados en `03_descriptive` y `sql` |

## Solución de Problemas

### Problema 1: Excel se vuelve lento, no responde o no puede cargar la tabla completa

**Síntomas**

- Excel tarda demasiado al cargar la consulta de Snowflake.
- El libro consume mucha memoria.
- La carga falla, muestra errores de recursos o supera la capacidad disponible.
- Las fórmulas sobre cientos de miles de filas tardan varios minutos.

**Causa probable**

El dataset contiene aproximadamente 750,000 transacciones y Excel debe procesar datos, caché de Power Query, tablas, fórmulas y tablas dinámicas. Cargar todas las columnas o usar fórmulas fila por fila puede exceder la memoria disponible.

**Solución**

1. Mantenga solo las columnas requeridas para esta práctica en Power Query.
2. Cargue la consulta al **Modelo de datos** en lugar de cargarla directamente a una hoja cuando sea necesario.
3. Use tablas dinámicas para agregaciones por región, canal, categoría y cliente.
4. Evite fórmulas de búsqueda o conteo fila por fila sobre 750,000 registros.
5. Calcule métricas complejas y frecuencia por cliente en Snowflake, luego exporte solo resultados agregados a Excel.
6. Cierre otras aplicaciones, asegure al menos 20 GB de espacio libre y use Excel de 64 bits.

### Problema 2: Los resultados de Excel y Snowflake no coinciden

**Síntomas**

- El conteo de transacciones es diferente.
- Las tasas de descuento o devolución no coinciden.
- Los percentiles muestran diferencias.
- La probabilidad de recurrencia es distinta entre herramientas.

**Causa probable**

Las diferencias suelen deberse a filtros distintos, campos nulos tratados de forma diferente, condiciones incorrectas para devoluciones, uso de filas en lugar de clientes únicos o percentiles aproximados en Snowflake.

**Solución**

1. Confirme que Excel y Snowflake usan la misma tabla o vista curada.
2. Compare fechas mínimas, máximas y número total de transacciones en ambas herramientas.
3. Revise que Excel no tenga filtros activos en la tabla o tabla dinámica.
4. Compruebe que `ConDescuento` usa exactamente la misma regla que SQL: descuento mayor que cero.
5. Compruebe que `DevolucionFlag` interpreta correctamente el valor real de devolución.
6. Para recurrencia, valide que ambos cálculos usan clientes únicos y frecuencia mayor o igual a dos.
7. Distinga diferencias normales de percentiles: `APPROX_PERCENTILE` en Snowflake puede diferir levemente de `PERCENTIL.INC` en Excel.
8. Redondee únicamente para presentación; compare valores sin redondear durante la validación.

## Limpieza

1. Guarde y cierre el libro de Excel:

```text
C:\DAF\Batch_01\03_descriptive\03_analisis_descriptivo_ventas_clientes.xlsx
```

2. Verifique que los siguientes entregables estén presentes:

```text
C:\DAF\Batch_01\sql\03_analisis_descriptivo_ventas.sql
C:\DAF\Batch_01\03_descriptive\03_analisis_descriptivo_ventas_clientes.xlsx
C:\DAF\Batch_01\03_descriptive\03_diccionario_y_alcance.md
C:\DAF\Batch_01\03_descriptive\03_validacion_excel_snowflake.md
```

3. Elimine únicamente archivos temporales locales que hayan sido creados para exportaciones intermedias y que no formen parte de los entregables aprobados.

4. No elimine, altere ni modifique objetos de los esquemas `RAW` o `CURATED`.

5. Cierre la sesión de Snowflake o cierre SnowSQL/Snowflake CLI:

```sql
EXIT;
```

6. Confirme que no se guardaron contraseñas, tokens o credenciales en archivos `.sql`, `.xlsx`, `.pbix` o archivos de texto del directorio de trabajo.

## Resumen

En esta práctica se analizó la población curada de transacciones de ventas del periodo 2024-2025, diferenciando la unidad transaccional de la unidad cliente. Se clasificaron variables, se calcularon estadísticas descriptivas, percentiles, frecuencias, proporciones, tasas y probabilidades comerciales con Snowflake y Excel.

También se validaron resultados entre ambas herramientas y se documentaron limitaciones esenciales: exclusiones por calidad, cobertura temporal, tamaños de segmentos, granularidad y la imposibilidad de establecer causalidad solo a partir de una correlación. Los hallazgos validados serán la base para seleccionar preguntas, comparaciones y visualizaciones en la Práctica 4.
