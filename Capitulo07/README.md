# Caso integral de desempeño comercial

## Metadatos

| Campo | Valor |
|---|---|
| Duración | 113 minutos |
| Complejidad | Difícil |
| Nivel de Bloom | Crear |

## Descripción General

En esta práctica integradora analizarás el desempeño comercial de NovaRetail utilizando las tres versiones equivalentes del dataset sintético RetailNova v1.0: Excel/CSV, Snowflake/CSV comprimido y Power BI/Parquet. Partirás del brief elaborado en la Práctica 6 para convertir una preocupación de negocio en una pregunta analítica medible, validarás la calidad de exactamente 750,000 transacciones y construirás un dashboard ejecutivo validado contra una vista preparada en Snowflake. Finalmente, comunicarás hallazgos, limitaciones, nivel de certeza y recomendaciones accionables sin atribuir causalidad no sustentada.

## Objetivos de Aprendizaje

- [ ] Definir un problema analítico, preguntas, métricas, dimensiones, periodo y criterios de éxito para el caso de desempeño comercial.
- [ ] Validar cobertura, unicidad, nulos, rangos temporales, importes, claves y cálculos derivados del dataset de 750,000 transacciones.
- [ ] Crear una vista analítica reproducible en Snowflake que conserve trazabilidad y documente reglas de limpieza sin modificar el esquema `RAW`.
- [ ] Construir un modelo y dashboard en Power BI con medidas DAX para ventas netas, margen, pedidos, ticket promedio, crecimiento y cumplimiento de meta.
- [ ] Comunicar recomendaciones priorizadas y diferenciarlas de hipótesis o afirmaciones causales no demostradas.

## Prerrequisitos

### Conocimientos requeridos

- Haber completado la Práctica 6 y contar con el archivo `M6_Brief_Interpretacion_[Apellido].xlsx`.
- Comprender la diferencia entre síntoma, hipótesis, hallazgo descriptivo y causalidad.
- Manejar consultas SQL con `SELECT`, `WHERE`, `GROUP BY`, `HAVING`, `JOIN`, `CASE`, agregaciones y funciones de fecha.
- Conocer importación de datos y transformaciones básicas con Power Query.
- Conocer creación de medidas DAX y visualizaciones básicas en Power BI Desktop.
- Comprender fórmulas, tablas dinámicas y validación de datos en Excel.

### Accesos y archivos requeridos

- Acceso de lectura al rol `DAF_ANALYST_ROLE`, al warehouse `DAF_LAB_WH` y a la base `DATA_ANALYTICS_FOUNDATIONS`.
- Permiso de creación de vistas en el esquema `ANALYTICS` o un esquema personal asignado.
- Los siguientes archivos descargados en `C:\DAF\Batch_01\00_source\`:

| Archivo | Uso principal |
|---|---|
| `M6_Brief_Interpretacion_[Apellido].xlsx` | Punto de partida para hipótesis y pregunta prioritaria |
| `retailnova_sales_v1_0_excel.csv` | Validación mediante Excel y Power Query |
| `retailnova_sales_v1_0_snowflake.csv.gz` | Referencia para carga o validación en Snowflake |
| `retailnova_sales_v1_0_powerbi.parquet` | Fuente de alto rendimiento para Power BI |
| Diccionario de datos o métricas RetailNova v1.0 | Confirmación de significado de campos y unidades |

> **Regla de protección:** no ejecutes `UPDATE`, `DELETE`, `TRUNCATE`, `DROP` ni `ALTER` sobre objetos del esquema `DATA_ANALYTICS_FOUNDATIONS.RAW`. El objeto fuente es inmutable: `DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1`.

## Entorno de Laboratorio

### Hardware recomendado

| Recurso | Mínimo | Recomendado |
|---|---:|---:|
| Memoria RAM | 16 GB | 32 GB |
| Espacio libre SSD | 20 GB | 25 GB |
| Procesador | Intel Core i5 10.ª gen. / AMD Ryzen 5 equivalente | 4 núcleos o superior |
| Pantalla | 1920 × 1080 | 1920 × 1080 o superior |
| Internet | 20 Mbps estable | 20 Mbps o superior |

### Software

| Herramienta | Versión de referencia | Uso |
|---|---|---|
| Windows 11 Pro | 23H2 | Sistema operativo |
| Microsoft Excel para Microsoft 365 | 64 bits | Revisión del CSV y Power Query |
| Snowflake Snowsight | Release 9.30.1 | SQL y exploración |
| SnowSQL / Snowflake CLI | SnowSQL 1.3.2 / CLI 3.6.0 | Ejecución opcional de scripts |
| Power BI Desktop | 2.146.1454.0, 64 bits | Modelo, DAX y dashboard |
| Python | 3.12.1 con pandas 2.2.0 | Validación opcional de conteos |
| Apache Parquet / pyarrow | 15.0.0 | Lectura del archivo Parquet |

### Preparación de carpetas

1. Abre **Windows PowerShell**.
2. Ejecuta los siguientes comandos:

```powershell
$root = "C:\DAF\Batch_01"

New-Item -ItemType Directory -Force -Path `
"$root\00_source", `
"$root\01_brief", `
"$root\02_quality", `
"$root\03_descriptive", `
"$root\04_exploration", `
"$root\05_dashboard", `
"$root\sql"
```

3. Copia el brief de la Práctica 6 a la carpeta `01_brief` y conserva el original sin modificaciones.

```powershell
Copy-Item `
"C:\DAF\Batch_01\00_source\M6_Brief_Interpretacion_[Apellido].xlsx" `
"C:\DAF\Batch_01\01_brief\"
```

4. Crea los archivos de trabajo iniciales:

```powershell
New-Item -ItemType File -Force `
"C:\DAF\Batch_01\02_quality\M7_Control_Calidad_[Apellido].xlsx", `
"C:\DAF\Batch_01\03_descriptive\M7_Resumen_Descriptivo_[Apellido].xlsx", `
"C:\DAF\Batch_01\04_exploration\M7_Hallazgos_[Apellido].xlsx", `
"C:\DAF\Batch_01\sql\M7_Desempeno_Comercial_[Apellido].sql"
```

> Sustituye `[Apellido]` por tu apellido sin espacios, tildes ni caracteres especiales. Ejemplo: `M7_Desempeno_Comercial_Garcia.pbix`.

## Instrucciones Paso a Paso

### Paso 1: Revisar el brief y definir el problema analítico

**Objetivo:** transformar la preocupación general sobre ventas en un problema analítico concreto, medible y orientado a una decisión.

1. Abre `C:\DAF\Batch_01\01_brief\M6_Brief_Interpretacion_[Apellido].xlsx`.
2. Revisa las hipótesis, segmentos identificados, hallazgos preliminares, limitaciones y preguntas abiertas documentadas en la Práctica 6.
3. Selecciona una pregunta prioritaria. Debe relacionarse con el desempeño comercial y poder responderse con las dimensiones disponibles: fecha, región, canal, categoría, producto y segmento de cliente.
4. Evita formular preguntas que presupongan una causa. Por ejemplo, no uses: “¿Qué campaña causó la caída?” si el dataset no tiene variables de campaña ni un diseño causal.
5. Registra en una hoja nueva llamada `Problema_Analitico` del archivo `M7_Hallazgos_[Apellido].xlsx` la siguiente estructura.

#### Plantilla de definición del problema

| Elemento | Definición del estudiante |
|---|---|
| Situación observada | Ejemplo: disminución o bajo desempeño de ventas netas en el último trimestre disponible |
| Audiencia | Dirección comercial, gerencia regional o responsables de canal |
| Decisión potencial | Priorizar revisión comercial, de surtido, promociones o seguimiento de segmentos |
| Pregunta analítica prioritaria | ¿Cómo cambió la venta neta en el último trimestre frente al periodo comparable y en qué segmentos se concentra la variación? |
| Unidad de análisis | Transacción identificada por `SALE_ID` |
| Periodo de estudio | 2024-01-01 a 2025-12-31; comparación trimestral, mensual o interanual equivalente |
| Métrica principal | Ventas netas |
| Métricas complementarias | Pedidos, ticket promedio, margen, % margen, meta, % cumplimiento |
| Dimensiones | Fecha, región, canal, categoría, producto, segmento de cliente |
| Criterio de éxito | Identificar segmentos que expliquen una proporción relevante de la variación y presentar evidencia verificable |

6. Formula entre dos y cuatro hipótesis investigables. Etiquétalas como hipótesis, no como conclusiones.

#### Ejemplo de hipótesis correctamente formuladas

| ID | Hipótesis | Evidencia requerida |
|---|---|---|
| H1 | La variación de ventas netas se concentra en uno o más canales. | Comparación de ventas, pedidos y ticket por canal y periodo |
| H2 | El deterioro de margen se concentra en categorías con mayor descuento relativo. | % margen y descuento relativo por categoría |
| H3 | Los segmentos de clientes tienen patrones de compra diferentes entre periodos. | Ventas, pedidos y ticket por segmento y cohorte |

**Resultado esperado:** un problema analítico documentado, una pregunta prioritaria medible y un conjunto de hipótesis que puedan confirmarse, descartarse o permanecer como evidencia insuficiente.

**Verificación:** confirma que la pregunta incluye métrica, periodo de comparación, dimensiones y propósito de decisión. Verifica además que no contiene expresiones causales como “debido a”, “causó” o “provocó” sin evidencia adicional.

---

### Paso 2: Confirmar archivos, cobertura y equivalencia entre formatos

**Objetivo:** verificar que las versiones CSV, CSV comprimido y Parquet pertenecen al mismo dataset oficial antes de aplicar filtros o transformaciones.

1. En PowerShell, verifica que los archivos estén disponibles:

```powershell
Get-ChildItem "C:\DAF\Batch_01\00_source" |
Select-Object Name, Length, LastWriteTime
```

2. Cuenta las filas del CSV de Excel utilizando Python. Este método evita abrir completamente el archivo en Excel antes de conocer su volumen:

```powershell
python -c "import csv; p=r'C:\DAF\Batch_01\00_source\retailnova_sales_v1_0_excel.csv'; print('Filas de datos:', sum(1 for _ in csv.DictReader(open(p, encoding='utf-8-sig'))))"
```

3. Cuenta las filas del archivo comprimido destinado a Snowflake:

```powershell
python -c "import csv,gzip; p=r'C:\DAF\Batch_01\00_source\retailnova_sales_v1_0_snowflake.csv.gz'; print('Filas de datos:', sum(1 for _ in csv.DictReader((x.decode('utf-8') for x in gzip.open(p,'rb')))))"
```

4. Cuenta las filas y revisa las columnas del Parquet:

```powershell
python -c "import pandas as pd; p=r'C:\DAF\Batch_01\00_source\retailnova_sales_v1_0_powerbi.parquet'; df=pd.read_parquet(p); print('Filas de datos:', len(df)); print('Columnas:', ', '.join(df.columns))"
```

5. Registra en la hoja `Control_Versiones` de `M7_Control_Calidad_[Apellido].xlsx` los resultados.

| Fuente | Formato | Conteo esperado antes de filtrar | Conteo observado | Estado |
|---|---|---:|---:|---|
| Excel | CSV | 750,000 |  |  |
| Snowflake | CSV.GZ / tabla RAW | 750,000 |  |  |
| Power BI | Parquet | 750,000 |  |  |

6. Si algún archivo no contiene 750,000 registros, no continúes con interpretaciones comparativas. Documenta la discrepancia, verifica si existe una descarga incompleta y solicita el archivo oficial correcto.

**Resultado esperado:** las tres versiones del dataset muestran exactamente 750,000 registros antes de cualquier filtrado.

**Verificación:** los tres conteos deben ser `750000`. La lista de columnas debe contener al menos `SALE_ID`, un campo de fecha, importes de ventas, descuento, costo, margen y dimensiones comerciales. Consulta el diccionario oficial para confirmar los nombres exactos.

---

### Paso 3: Inspeccionar la fuente y ejecutar controles de calidad en Snowflake

**Objetivo:** validar la estructura, unicidad, completitud, rangos, claves y consistencia matemática del objeto fuente sin modificarlo.

1. Abre Snowsight o una sesión de SnowSQL con la conexión `da_foundations_lab`.
2. Configura el contexto de trabajo:

```sql
USE ROLE DAF_ANALYST_ROLE;
USE WAREHOUSE DAF_LAB_WH;
USE DATABASE DATA_ANALYTICS_FOUNDATIONS;
USE SCHEMA RAW;
```

3. Inspecciona la estructura real de la fuente. No asumas nombres de columnas distintos de los entregados por el diccionario.

```sql
DESC TABLE DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1;

SELECT *
FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1
LIMIT 20;
```

4. Guarda las consultas de este paso en:

```text
C:\DAF\Batch_01\sql\M7_Desempeno_Comercial_[Apellido].sql
```

5. Ejecuta el control básico de conteo y unicidad. Si el campo `SALE_ID` es texto, se evalúa del mismo modo:

```sql
SELECT
    COUNT(*) AS total_registros,
    COUNT(DISTINCT SALE_ID) AS sale_id_distintos,
    COUNT(*) - COUNT(DISTINCT SALE_ID) AS sale_id_duplicados
FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1;
```

6. Ejecuta el control de nulos. Ajusta los nombres de campo solo si el diccionario usa nombres distintos.

```sql
SELECT
    COUNT_IF(SALE_ID IS NULL) AS nulos_sale_id,
    COUNT_IF(SALE_DATE IS NULL) AS nulos_sale_date,
    COUNT_IF(SALES_AMOUNT IS NULL) AS nulos_sales_amount,
    COUNT_IF(DISCOUNT_AMOUNT IS NULL) AS nulos_discount_amount,
    COUNT_IF(COST_AMOUNT IS NULL) AS nulos_cost_amount,
    COUNT_IF(MARGIN_AMOUNT IS NULL) AS nulos_margin_amount,
    COUNT_IF(REGION IS NULL OR TRIM(REGION) = '') AS nulos_region,
    COUNT_IF(CHANNEL IS NULL OR TRIM(CHANNEL) = '') AS nulos_channel,
    COUNT_IF(CATEGORY IS NULL OR TRIM(CATEGORY) = '') AS nulos_category,
    COUNT_IF(CUSTOMER_ID IS NULL OR TRIM(CUSTOMER_ID) = '') AS nulos_customer_id,
    COUNT_IF(PRODUCT_ID IS NULL OR TRIM(PRODUCT_ID) = '') AS nulos_product_id
FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1;
```

7. Valida el rango temporal y la distribución anual:

```sql
SELECT
    MIN(SALE_DATE) AS fecha_minima,
    MAX(SALE_DATE) AS fecha_maxima,
    COUNT(DISTINCT SALE_DATE) AS dias_con_registros
FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1;

SELECT
    YEAR(SALE_DATE) AS anio,
    COUNT(*) AS transacciones,
    COUNT(DISTINCT SALE_ID) AS pedidos
FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1
GROUP BY YEAR(SALE_DATE)
ORDER BY anio;
```

8. Revisa valores negativos, descuentos no plausibles y registros monetarios extremos. Un valor atípico no debe eliminarse automáticamente: primero se cuantifica y documenta.

```sql
SELECT
    COUNT_IF(SALES_AMOUNT < 0) AS ventas_brutas_negativas,
    COUNT_IF(DISCOUNT_AMOUNT < 0) AS descuentos_negativos,
    COUNT_IF(COST_AMOUNT < 0) AS costos_negativos,
    COUNT_IF(DISCOUNT_AMOUNT > SALES_AMOUNT) AS descuento_mayor_que_venta,
    COUNT_IF((SALES_AMOUNT - DISCOUNT_AMOUNT) < 0) AS venta_neta_negativa,
    MIN(SALES_AMOUNT) AS venta_bruta_minima,
    MAX(SALES_AMOUNT) AS venta_bruta_maxima,
    MIN(MARGIN_AMOUNT) AS margen_minimo,
    MAX(MARGIN_AMOUNT) AS margen_maximo
FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1;
```

9. Evalúa la consistencia de la identidad de cálculo definida para este laboratorio:

\[
\text{Ventas netas} = \text{SALES\_AMOUNT} - \text{DISCOUNT\_AMOUNT}
\]

\[
\text{Margen calculado} = \text{Ventas netas} - \text{COST\_AMOUNT}
\]

```sql
SELECT
    COUNT(*) AS total_registros,
    COUNT_IF(
        ABS(
            COALESCE(MARGIN_AMOUNT, 0) -
            (
                COALESCE(SALES_AMOUNT, 0) -
                COALESCE(DISCOUNT_AMOUNT, 0) -
                COALESCE(COST_AMOUNT, 0)
            )
        ) > 0.01
    ) AS registros_margen_inconsistente
FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1;
```

10. Revisa valores de dimensiones y claves potencialmente no coincidentes o vacías:

```sql
SELECT 'REGION' AS dimension, COALESCE(NULLIF(TRIM(REGION), ''), '[NULO/VACIO]') AS valor, COUNT(*) AS registros
FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1
GROUP BY 1, 2

UNION ALL

SELECT 'CHANNEL', COALESCE(NULLIF(TRIM(CHANNEL), ''), '[NULO/VACIO]'), COUNT(*)
FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1
GROUP BY 1, 2

UNION ALL

SELECT 'CATEGORY', COALESCE(NULLIF(TRIM(CATEGORY), ''), '[NULO/VACIO]'), COUNT(*)
FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1
GROUP BY 1, 2
ORDER BY dimension, registros DESC;
```

11. En la hoja `Hallazgos_Calidad` del archivo de control, registra cada regla, resultado, impacto y decisión.

| Regla | Resultado | Impacto potencial | Decisión documentada |
|---|---:|---|---|
| Conteo de registros |  | Cobertura |  |
| Duplicados de `SALE_ID` |  | Sobreconteo de pedidos/ventas |  |
| Nulos en dimensiones |  | Segmentación incompleta |  |
| Rango de fechas |  | Comparabilidad temporal |  |
| Venta neta negativa |  | Interpretación de devoluciones o errores |  |
| Margen inconsistente |  | Confiabilidad de indicadores |  |

**Resultado esperado:** un diagnóstico reproducible de calidad que confirme o cuantifique excepciones sin alterar ni eliminar silenciosamente registros del esquema `RAW`.

**Verificación:** el conteo total debe ser `750000`. El rango esperado del dataset oficial es desde `2024-01-01` hasta `2025-12-31`. Todo hallazgo de nulos, duplicados, importes extremos o inconsistencias debe aparecer documentado con una decisión explícita.

---

### Paso 4: Crear la vista analítica trazable en Snowflake

**Objetivo:** construir una vista reutilizable con campos derivados, indicadores de calidad y reglas estandarizadas para análisis y Power BI.

1. Consulta el diccionario de datos para confirmar si el campo de meta se llama `SALES_TARGET_AMOUNT`, `TARGET_AMOUNT` u otro nombre equivalente.
2. Usa el objeto analítico asignado por el curso. En el entorno de referencia, crea la vista en:

```text
DATA_ANALYTICS_FOUNDATIONS.ANALYTICS.VW_SALES_ANALYTICS_V1
```

3. Si tu instrucción institucional exige el nombre `ANALYTICS_DB.COMMERCIAL.VW_SALES_ANALYTICS_V1`, sustituye únicamente el nombre destino de la vista; conserva la consulta y la fuente `RAW` autorizada.
4. Si no tienes privilegios de creación, usa el sufijo indicado por el instructor, por ejemplo `VW_SALES_ANALYTICS_V1_GARCIA`, o solicita la creación del objeto sin intentar modificar `RAW`.
5. Ejecuta el siguiente SQL. Si el diccionario usa otro nombre para la meta, reemplaza solo `SALES_TARGET_AMOUNT`.

```sql
USE ROLE DAF_ANALYST_ROLE;
USE WAREHOUSE DAF_LAB_WH;
USE DATABASE DATA_ANALYTICS_FOUNDATIONS;
USE SCHEMA ANALYTICS;

CREATE OR REPLACE VIEW DATA_ANALYTICS_FOUNDATIONS.ANALYTICS.VW_SALES_ANALYTICS_V1 AS
SELECT
    SALE_ID,
    SALE_DATE::DATE AS SALE_DATE,
    CUSTOMER_ID,
    PRODUCT_ID,

    NULLIF(TRIM(REGION), '') AS REGION,
    NULLIF(TRIM(CHANNEL), '') AS CHANNEL,
    NULLIF(TRIM(CATEGORY), '') AS CATEGORY,
    NULLIF(TRIM(CUSTOMER_SEGMENT), '') AS CUSTOMER_SEGMENT,

    COALESCE(SALES_AMOUNT, 0) AS SALES_AMOUNT,
    COALESCE(DISCOUNT_AMOUNT, 0) AS DISCOUNT_AMOUNT,
    COALESCE(COST_AMOUNT, 0) AS COST_AMOUNT,
    COALESCE(MARGIN_AMOUNT, 0) AS MARGIN_AMOUNT,
    COALESCE(SALES_TARGET_AMOUNT, 0) AS SALES_TARGET_AMOUNT,

    COALESCE(SALES_AMOUNT, 0) - COALESCE(DISCOUNT_AMOUNT, 0) AS NET_SALES_AMOUNT,

    DATE_TRUNC('MONTH', SALE_DATE)::DATE AS MONTH_START_DATE,
    DATE_TRUNC('QUARTER', SALE_DATE)::DATE AS QUARTER_START_DATE,
    YEAR(SALE_DATE) AS SALE_YEAR,
    QUARTER(SALE_DATE) AS SALE_QUARTER,
    MONTH(SALE_DATE) AS SALE_MONTH_NUMBER,

    CASE
        WHEN SALE_ID IS NULL THEN 'SALE_ID_NULO'
        WHEN SALE_DATE IS NULL THEN 'FECHA_NULA'
        WHEN SALES_AMOUNT IS NULL THEN 'VENTA_NULA'
        WHEN SALES_AMOUNT - COALESCE(DISCOUNT_AMOUNT, 0) < 0 THEN 'VENTA_NETA_NEGATIVA'
        WHEN ABS(
            COALESCE(MARGIN_AMOUNT, 0) -
            (
                COALESCE(SALES_AMOUNT, 0) -
                COALESCE(DISCOUNT_AMOUNT, 0) -
                COALESCE(COST_AMOUNT, 0)
            )
        ) > 0.01 THEN 'MARGEN_INCONSISTENTE'
        WHEN NULLIF(TRIM(REGION), '') IS NULL THEN 'REGION_NULA'
        WHEN NULLIF(TRIM(CHANNEL), '') IS NULL THEN 'CANAL_NULO'
        WHEN NULLIF(TRIM(CATEGORY), '') IS NULL THEN 'CATEGORIA_NULA'
        ELSE 'VALIDO'
    END AS QUALITY_STATUS,

    CASE
        WHEN SALES_AMOUNT IS NULL OR SALES_AMOUNT = 0 THEN NULL
        ELSE COALESCE(DISCOUNT_AMOUNT, 0) / SALES_AMOUNT
    END AS DISCOUNT_PCT,

    CASE
        WHEN COALESCE(SALES_AMOUNT, 0) - COALESCE(DISCOUNT_AMOUNT, 0) = 0 THEN NULL
        ELSE COALESCE(MARGIN_AMOUNT, 0) /
             (COALESCE(SALES_AMOUNT, 0) - COALESCE(DISCOUNT_AMOUNT, 0))
    END AS MARGIN_PCT
FROM DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1;
```

6. La vista no elimina registros. En cambio, expone el campo `QUALITY_STATUS` para que el análisis pueda excluir, incluir o cuantificar registros según una regla visible y auditable.
7. Valida la vista:

```sql
SELECT
    COUNT(*) AS total_registros,
    COUNT(DISTINCT SALE_ID) AS sale_id_distintos,
    MIN(SALE_DATE) AS fecha_minima,
    MAX(SALE_DATE) AS fecha_maxima
FROM DATA_ANALYTICS_FOUNDATIONS.ANALYTICS.VW_SALES_ANALYTICS_V1;

SELECT
    QUALITY_STATUS,
    COUNT(*) AS registros,
    ROUND(100.0 * COUNT(*) / SUM(COUNT(*)) OVER (), 4) AS porcentaje
FROM DATA_ANALYTICS_FOUNDATIONS.ANALYTICS.VW_SALES_ANALYTICS_V1
GROUP BY QUALITY_STATUS
ORDER BY registros DESC;
```

8. Documenta las reglas aplicadas en la hoja `Reglas_Limpieza`. Incluye la regla, justificación, cantidad afectada y efecto analítico.

**Resultado esperado:** una vista analítica con 750,000 registros, campos de fecha estandarizados, ventas netas derivadas, métricas porcentuales e indicadores transparentes de calidad.

**Verificación:** la vista debe conservar el mismo conteo de registros que la tabla `RAW`. Comprueba que `NET_SALES_AMOUNT = SALES_AMOUNT - DISCOUNT_AMOUNT` en una muestra de registros:

```sql
SELECT
    SALE_ID,
    SALES_AMOUNT,
    DISCOUNT_AMOUNT,
    NET_SALES_AMOUNT,
    COST_AMOUNT,
    MARGIN_AMOUNT,
    QUALITY_STATUS
FROM DATA_ANALYTICS_FOUNDATIONS.ANALYTICS.VW_SALES_ANALYTICS_V1
LIMIT 20;
```

---

### Paso 5: Realizar el análisis exploratorio y contrastar hipótesis

**Objetivo:** identificar tendencias, variaciones segmentadas, excepciones y evidencia para evaluar las hipótesis provenientes de la Práctica 6.

1. Calcula el desempeño mensual por canal. Filtra explícitamente registros válidos para el análisis principal, pero conserva el conteo de excluidos en el control de calidad.

```sql
SELECT
    MONTH_START_DATE,
    CHANNEL,
    SUM(NET_SALES_AMOUNT) AS ventas_netas,
    SUM(MARGIN_AMOUNT) AS margen,
    COUNT(DISTINCT SALE_ID) AS pedidos,
    DIV0(SUM(NET_SALES_AMOUNT), COUNT(DISTINCT SALE_ID)) AS ticket_promedio,
    DIV0(SUM(MARGIN_AMOUNT), SUM(NET_SALES_AMOUNT)) AS margen_pct
FROM DATA_ANALYTICS_FOUNDATIONS.ANALYTICS.VW_SALES_ANALYTICS_V1
WHERE QUALITY_STATUS = 'VALIDO'
GROUP BY MONTH_START_DATE, CHANNEL
ORDER BY MONTH_START_DATE, CHANNEL;
```

2. Evalúa la variación interanual por región y canal:

```sql
WITH resumen AS (
    SELECT
        SALE_YEAR,
        REGION,
        CHANNEL,
        SUM(NET_SALES_AMOUNT) AS ventas_netas,
        SUM(MARGIN_AMOUNT) AS margen,
        COUNT(DISTINCT SALE_ID) AS pedidos
    FROM DATA_ANALYTICS_FOUNDATIONS.ANALYTICS.VW_SALES_ANALYTICS_V1
    WHERE QUALITY_STATUS = 'VALIDO'
    GROUP BY SALE_YEAR, REGION, CHANNEL
)
SELECT
    REGION,
    CHANNEL,
    SUM(IFF(SALE_YEAR = 2024, ventas_netas, 0)) AS ventas_2024,
    SUM(IFF(SALE_YEAR = 2025, ventas_netas, 0)) AS ventas_2025,
    DIV0(
        SUM(IFF(SALE_YEAR = 2025, ventas_netas, 0)) -
        SUM(IFF(SALE_YEAR = 2024, ventas_netas, 0)),
        SUM(IFF(SALE_YEAR = 2024, ventas_netas, 0))
    ) AS crecimiento_interanual_pct,
    SUM(IFF(SALE_YEAR = 2024, pedidos, 0)) AS pedidos_2024,
    SUM(IFF(SALE_YEAR = 2025, pedidos, 0)) AS pedidos_2025
FROM resumen
GROUP BY REGION, CHANNEL
ORDER BY crecimiento_interanual_pct ASC;
```

3. Analiza categorías. El objetivo es localizar contribuciones, no buscar artificialmente un único responsable.

```sql
SELECT
    CATEGORY,
    SUM(NET_SALES_AMOUNT) AS ventas_netas,
    SUM(MARGIN_AMOUNT) AS margen,
    DIV0(SUM(MARGIN_AMOUNT), SUM(NET_SALES_AMOUNT)) AS margen_pct,
    AVG(DISCOUNT_PCT) AS descuento_promedio_pct,
    COUNT(DISTINCT SALE_ID) AS pedidos,
    DIV0(SUM(NET_SALES_AMOUNT), COUNT(DISTINCT SALE_ID)) AS ticket_promedio
FROM DATA_ANALYTICS_FOUNDATIONS.ANALYTICS.VW_SALES_ANALYTICS_V1
WHERE QUALITY_STATUS = 'VALIDO'
GROUP BY CATEGORY
ORDER BY ventas_netas DESC;
```

4. Revisa segmentos de clientes y cohortes anuales. Si no existe `CUSTOMER_SEGMENT`, utiliza la dimensión equivalente indicada en el diccionario.

```sql
SELECT
    SALE_YEAR,
    CUSTOMER_SEGMENT,
    SUM(NET_SALES_AMOUNT) AS ventas_netas,
    SUM(MARGIN_AMOUNT) AS margen,
    COUNT(DISTINCT SALE_ID) AS pedidos,
    COUNT(DISTINCT CUSTOMER_ID) AS clientes_activos,
    DIV0(SUM(NET_SALES_AMOUNT), COUNT(DISTINCT SALE_ID)) AS ticket_promedio
FROM DATA_ANALYTICS_FOUNDATIONS.ANALYTICS.VW_SALES_ANALYTICS_V1
WHERE QUALITY_STATUS = 'VALIDO'
GROUP BY SALE_YEAR, CUSTOMER_SEGMENT
ORDER BY SALE_YEAR, ventas_netas DESC;
```

5. Si existe una meta por transacción, analiza cumplimiento por región y canal:

```sql
SELECT
    REGION,
    CHANNEL,
    SUM(NET_SALES_AMOUNT) AS ventas_netas,
    SUM(SALES_TARGET_AMOUNT) AS meta_ventas,
    DIV0(SUM(NET_SALES_AMOUNT), SUM(SALES_TARGET_AMOUNT)) AS cumplimiento_meta_pct
FROM DATA_ANALYTICS_FOUNDATIONS.ANALYTICS.VW_SALES_ANALYTICS_V1
WHERE QUALITY_STATUS = 'VALIDO'
GROUP BY REGION, CHANNEL
ORDER BY cumplimiento_meta_pct;
```

6. Registra los hallazgos en la hoja `Matriz_Evidencia` de `M7_Hallazgos_[Apellido].xlsx`.

| Hipótesis o pregunta | Evidencia observada | Comparación temporal | Nivel de certeza | Estado |
|---|---|---|---|---|
| H1 |  | Periodos equivalentes | Alto / medio / bajo | Confirmatoria / insuficiente / no confirmatoria |
| H2 |  | Periodos equivalentes | Alto / medio / bajo | Confirmatoria / insuficiente / no confirmatoria |
| Hallazgo nuevo |  |  | Alto / medio / bajo | Descriptivo |

7. Usa lenguaje proporcional a la evidencia:
   - “Se observa una concentración de la variación en…”
   - “El patrón es consistente con la hipótesis…”
   - “La evidencia disponible es insuficiente para atribuir causalidad…”
   - “Se recomienda validar con datos adicionales de inventario, promociones o campañas…”

**Resultado esperado:** una matriz de evidencia que conecte hipótesis, consultas, resultados, nivel de certeza y limitaciones.

**Verificación:** cada afirmación incluida en tu análisis debe poder rastrearse a una consulta SQL, una medida DAX o una tabla documentada. Ningún hallazgo debe declarar causalidad sin variables o diseño analítico que la respalden.

---

### Paso 6: Validar una versión equivalente en Excel mediante Power Query

**Objetivo:** usar Excel como control complementario de trazabilidad y confirmar que el CSV contiene la cobertura esperada.

1. Abre Excel y crea un libro nuevo llamado:

```text
C:\DAF\Batch_01\02_quality\M7_Validacion_Excel_[Apellido].xlsx
```

2. Selecciona **Datos > Obtener datos > Desde archivo > Desde texto/CSV**.
3. Selecciona:

```text
C:\DAF\Batch_01\00_source\retailnova_sales_v1_0_excel.csv
```

4. En la vista previa, verifica codificación UTF-8, delimitador correcto y encabezados.
5. Selecciona **Transformar datos**.
6. En Power Query:
   1. Confirma que `SALE_ID` tenga tipo texto o entero consistente.
   2. Asigna tipo fecha a `SALE_DATE`.
   3. Asigna tipo decimal fijo o número decimal a importes monetarios.
   4. No elimines filas por nulos, duplicados o valores extremos durante esta validación.
   5. Selecciona **Inicio > Mantener filas > Mantener filas superiores** solo si necesitas una muestra para inspección visual; no sustituyas el conteo de la fuente completa por el de la muestra.
7. En la parte inferior de Power Query, confirma el conteo aproximado de 750,000 filas.
8. Selecciona **Cerrar y cargar en...** y elige **Solo crear conexión** y **Agregar estos datos al modelo de datos**.
9. Crea una hoja `Control` y registra:
   - Conteo de filas importadas.
   - Fecha mínima y máxima.
   - Número de valores distintos de `SALE_ID`, si el modelo o una tabla dinámica permite calcularlo.
   - Lista de columnas disponibles.
10. Guarda el libro.

**Resultado esperado:** un archivo Excel que demuestre que el CSV corresponde al dataset de referencia y que no se realizaron exclusiones silenciosas.

**Verificación:** el conteo de la consulta Power Query debe ser 750,000 y debe coincidir con Snowflake y Parquet antes de filtros.

---

### Paso 7: Construir el modelo analítico y medidas DAX en Power BI

**Objetivo:** importar la fuente de alto rendimiento, construir una tabla calendario y crear medidas que respeten las definiciones de negocio.

1. Abre Power BI Desktop.
2. Selecciona **Obtener datos > Parquet**.
3. Selecciona el archivo:

```text
C:\DAF\Batch_01\00_source\retailnova_sales_v1_0_powerbi.parquet
```

4. En Power Query:
   1. Renombra la consulta como `Ventas`.
   2. Confirma los tipos de datos.
   3. Configura `SALE_DATE` como fecha.
   4. Configura identificadores como texto cuando aplique.
   5. Configura importes como número decimal fijo.
   6. Agrega una columna personalizada de ventas netas solo si no existe en Parquet:

```powerquery
[SALES_AMOUNT] - [DISCOUNT_AMOUNT]
```

   7. Renombra la columna como `NET_SALES_AMOUNT`.
   8. No elimines registros de calidad dudosa sin documentar la regla. Si aplicarás un filtro `QUALITY_STATUS = "VALIDO"`, hazlo visible y consistente con el análisis de Snowflake.
5. Selecciona **Cerrar y aplicar**.
6. En la vista **Modelo**, crea la tabla calendario:

```DAX
Calendario =
ADDCOLUMNS(
    CALENDAR(DATE(2024, 1, 1), DATE(2025, 12, 31)),
    "Año", YEAR([Date]),
    "Mes Número", MONTH([Date]),
    "Mes", FORMAT([Date], "MMMM"),
    "Año Mes", FORMAT([Date], "YYYY-MM"),
    "Trimestre", "T" & FORMAT([Date], "Q")
)
```

7. Marca `Calendario` como tabla de fechas:
   - Selecciona la tabla `Calendario`.
   - Selecciona **Herramientas de tabla > Marcar como tabla de fechas**.
   - Elige `Calendario[Date]`.
8. Crea una relación de uno a varios entre `Calendario[Date]` y `Ventas[SALE_DATE]`.
9. Crea las medidas siguientes. Sustituye `SALES_TARGET_AMOUNT` por el nombre real de la columna de meta si fuera necesario.

```DAX
Ventas netas =
SUM(Ventas[NET_SALES_AMOUNT])
```

```DAX
Margen =
SUM(Ventas[MARGIN_AMOUNT])
```

```DAX
% margen =
DIVIDE([Margen], [Ventas netas])
```

```DAX
Pedidos =
DISTINCTCOUNT(Ventas[SALE_ID])
```

```DAX
Ticket promedio =
DIVIDE([Ventas netas], [Pedidos])
```

```DAX
Meta de ventas =
SUM(Ventas[SALES_TARGET_AMOUNT])
```

```DAX
% cumplimiento de meta =
DIVIDE([Ventas netas], [Meta de ventas])
```

```DAX
Ventas netas año anterior =
CALCULATE(
    [Ventas netas],
    SAMEPERIODLASTYEAR(Calendario[Date])
)
```

```DAX
Crecimiento interanual % =
DIVIDE(
    [Ventas netas] - [Ventas netas año anterior],
    [Ventas netas año anterior]
)
```

```DAX
Contribución % ventas =
DIVIDE(
    [Ventas netas],
    CALCULATE([Ventas netas], ALLSELECTED(Ventas))
)
```

10. Formatea:
    - Ventas netas, margen, ticket y meta como moneda.
    - `% margen`, `% cumplimiento de meta`, `Crecimiento interanual %` y `Contribución % ventas` como porcentaje con uno o dos decimales.
11. Crea una medida de control para validar la cobertura cargada:

```DAX
Registros cargados =
COUNTROWS(Ventas)
```

**Resultado esperado:** un modelo con una tabla de hechos `Ventas`, tabla calendario relacionada y medidas reutilizables para el dashboard.

**Verificación:** muestra tarjetas temporales para `Registros cargados`, `Pedidos` y `Ventas netas`. `Registros cargados` debe ser 750,000 antes de aplicar filtros de calidad. El valor de ventas netas agregado debe ser consistente con el resultado de Snowflake para el mismo alcance y filtros.

---

### Paso 8: Construir el dashboard de desempeño comercial

**Objetivo:** crear un dashboard ejecutivo que permita identificar desempeño, tendencias y segmentos relevantes para la decisión comercial.

1. Guarda el archivo como:

```text
C:\DAF\Batch_01\05_dashboard\M7_Desempeno_Comercial_[Apellido].pbix
```

2. Crea la página **Resumen ejecutivo**.
3. Agrega los siguientes segmentadores:
   - `Calendario[Date]` o rango de fechas.
   - `Ventas[REGION]`.
   - `Ventas[CHANNEL]`.
   - `Ventas[CATEGORY]`.

4. Incluye las siguientes tarjetas:
   - `Ventas netas`
   - `Margen`
   - `% margen`
   - `Pedidos`
   - `% cumplimiento de meta`

5. Agrega una visualización de tendencia:
   - Eje: `Calendario[Año Mes]`.
   - Valor: `Ventas netas`.
   - Serie o leyenda opcional: `CHANNEL`.
   - Ordena correctamente por fecha, no por texto de mes.

6. Agrega una comparación de desempeño:
   - Barras agrupadas por `REGION` o `CHANNEL`.
   - Valor principal: `Ventas netas`.
   - Tooltip o valor secundario: `% margen` o `% cumplimiento de meta`.

7. Agrega una visualización de categorías o productos:
   - Barras por `CATEGORY`.
   - Valor: `Ventas netas`.
   - Tooltip: `Margen`, `% margen`, `Ticket promedio`, `Contribución % ventas`.
   - Si usas productos, limita la visualización a un Top N razonable y documenta el criterio.

8. Crea la página **Diagnóstico**.
9. Incluye como mínimo:
   - Matriz por región, canal y categoría.
   - Tendencia mensual de ventas netas y pedidos.
   - Dispersión o barras que comparen descuento relativo y `% margen` por categoría, si la calidad y el volumen lo permiten.
   - Tabla de detalle filtrable con `SALE_ID`, fecha, región, canal, categoría, ventas netas, margen y `QUALITY_STATUS`.

10. Configura interacciones visuales para que los filtros de periodo, región, canal y categoría afecten las visualizaciones relevantes.
11. Incluye un cuadro de texto visible con:
    - Periodo cubierto: `2024-01-01 a 2025-12-31`.
    - Fuente: RetailNova v1.0, dataset sintético.
    - Nota metodológica: “Los resultados describen asociaciones y distribución del desempeño; no prueban causalidad.”

**Resultado esperado:** un archivo PBIX con dos páginas, filtros requeridos, tarjetas, tendencia mensual, comparaciones segmentadas y tabla de detalle.

**Verificación:** aplica un filtro de región y confirma que cambien tarjetas, tendencia, comparación y detalle. Aplica un filtro de periodo equivalente entre 2024 y 2025 y verifica que la medida de crecimiento interanual responda coherentemente.

---

### Paso 9: Validar resultados, elaborar el brief ejecutivo y registrar una mejora

**Objetivo:** asegurar consistencia entre Snowflake y Power BI, comunicar resultados con nivel de certeza y preparar la exposición final.

1. Selecciona un alcance de validación común, por ejemplo:
   - Todo el periodo.
   - Solo `QUALITY_STATUS = "VALIDO"`.
   - Una región, un canal y una categoría específicos.
2. Ejecuta en Snowflake una consulta de validación:

```sql
SELECT
    COUNT(*) AS registros_validos,
    COUNT(DISTINCT SALE_ID) AS pedidos,
    ROUND(SUM(NET_SALES_AMOUNT), 2) AS ventas_netas,
    ROUND(SUM(MARGIN_AMOUNT), 2) AS margen,
    ROUND(DIV0(SUM(MARGIN_AMOUNT), SUM(NET_SALES_AMOUNT)), 4) AS margen_pct
FROM DATA_ANALYTICS_FOUNDATIONS.ANALYTICS.VW_SALES_ANALYTICS_V1
WHERE QUALITY_STATUS = 'VALIDO';
```

3. Reproduce el mismo alcance en Power BI y compara:
   - Registros o pedidos.
   - Ventas netas.
   - Margen.
   - Porcentaje de margen.
4. Acepta diferencias únicamente por redondeo. Investiga diferencias materiales antes de presentar resultados.
5. Elabora una presentación de máximo cinco diapositivas o un brief de una página equivalente. Guarda el entregable en `05_dashboard`.

#### Estructura sugerida para el brief ejecutivo

| Sección | Contenido requerido |
|---|---|
| 1. Contexto y pregunta | Problema, audiencia, periodo y decisión potencial |
| 2. Desempeño general | Ventas netas, margen, pedidos, ticket y cumplimiento de meta |
| 3. Dónde se concentra la variación | Canal, región, categoría o segmento prioritario |
| 4. Evidencia y certeza | Hipótesis confirmatorias, insuficientes y hallazgos nuevos |
| 5. Recomendaciones y próximos pasos | Acciones priorizadas, límites y datos adicionales requeridos |

6. Redacta entre dos y tres recomendaciones usando el formato siguiente:

| Prioridad | Recomendación | Evidencia | Nivel de certeza | Factibilidad | Próximo paso |
|---|---|---|---|---|---|
| Alta | Revisar desempeño del segmento identificado | Variación observada en ventas, pedidos o margen | Medio | Alta | Validar inventario, promociones o disponibilidad |
| Media | Profundizar categoría con deterioro de margen | % margen y descuentos relativos | Medio | Media | Analizar precio, costo y promociones |
| Baja | Diseñar experimento o seguimiento | Evidencia descriptiva, no causal | Bajo | Media | Definir grupo de control o medición adicional |

7. Realiza una exposición de tres a cinco minutos.
8. Solicita retroalimentación sobre claridad, trazabilidad, diseño del dashboard y uso apropiado de causalidad.
9. Registra una mejora prioritaria en la hoja `Mejora_Prioritaria`:

| Retroalimentación recibida | Mejora prioritaria | Acción concreta | Fecha objetivo |
|---|---|---|---|
|  |  |  |  |

**Resultado esperado:** dashboard validado, brief ejecutivo conciso, recomendaciones trazables y una mejora prioritaria registrada.

**Verificación:** cada recomendación debe citar una métrica, segmento y periodo. Las conclusiones deben distinguir explícitamente entre evidencia confirmatoria, evidencia insuficiente y hallazgos descriptivos nuevos.

## Validación y Pruebas

Completa la siguiente lista antes de entregar la práctica.

| Prueba | Criterio de aceptación | Evidencia |
|---|---|---|
| Directorio de trabajo | Existen las carpetas `00_source` a `05_dashboard` y `sql` | Explorador de archivos o PowerShell |
| Brief previo | El archivo de Práctica 6 está disponible y fue revisado | Hoja `Problema_Analitico` |
| Conteo CSV | 750,000 registros antes de filtrar | Registro de control |
| Conteo CSV.GZ / Snowflake | 750,000 registros antes de filtrar | Consulta SQL |
| Conteo Parquet | 750,000 registros antes de filtrar | Python o Power BI |
| Unicidad | Conteo de `SALE_ID` documentado; discrepancias explicadas | Consulta SQL |
| Rango temporal | Fechas entre 2024-01-01 y 2025-12-31 | Consulta SQL |
| Calidad | Nulos, negativos, extremos e inconsistencias evaluados | `M7_Control_Calidad_[Apellido].xlsx` |
| Trazabilidad | No hubo modificaciones destructivas en `RAW` | Script SQL |
| Vista analítica | Vista creada o versión con sufijo autorizada | Consulta a `VW_SALES_ANALYTICS_V1` |
| Reglas de limpieza | Existen indicadores de calidad; no hay exclusiones silenciosas | Campo `QUALITY_STATUS` |
| Modelo Power BI | Calendario relacionado con ventas y tabla marcada como fecha | Vista Modelo |
| DAX | Medidas de ventas, margen, pedidos, ticket, crecimiento y meta creadas | Panel Campos |
| Dashboard | Dos páginas con filtros, tarjetas, tendencia, comparación y detalle | Archivo PBIX |
| Consistencia | Snowflake y Power BI coinciden para el mismo alcance | Tabla de validación |
| Comunicación | Brief de máximo cinco diapositivas o una página | Archivo final |
| Causalidad | No se afirman causas sin evidencia o diseño apropiado | Revisión de narrativa |

### Criterios de entrega

Entrega los siguientes archivos, conservando el patrón de nombres solicitado:

```text
C:\DAF\Batch_01\02_quality\M7_Control_Calidad_[Apellido].xlsx
C:\DAF\Batch_01\03_descriptive\M7_Resumen_Descriptivo_[Apellido].xlsx
C:\DAF\Batch_01\04_exploration\M7_Hallazgos_[Apellido].xlsx
C:\DAF\Batch_01\05_dashboard\M7_Desempeno_Comercial_[Apellido].pbix
C:\DAF\Batch_01\05_dashboard\M7_Brief_Ejecutivo_[Apellido].pptx
C:\DAF\Batch_01\sql\M7_Desempeno_Comercial_[Apellido].sql
```

> Si entregas un brief de una página en lugar de una presentación, utiliza un formato autorizado por el instructor, por ejemplo PDF: `M7_Brief_Ejecutivo_[Apellido].pdf`.

## Solución de Problemas

### Problema 1: Power BI se queda sin memoria, tarda demasiado o no puede cargar el Parquet

**Síntomas:** Power BI Desktop se vuelve lento, muestra mensajes de falta de memoria, no responde durante “Aplicar cambios” o falla al abrir el archivo Parquet de 750,000 transacciones.

**Causa probable:** el equipo tiene memoria disponible insuficiente, existen otras aplicaciones que consumen RAM, se han creado columnas innecesarias en Power Query o se está intentando cargar varias copias completas de la tabla.

**Corrección:**

1. Cierra Excel, navegadores con muchas pestañas y aplicaciones no necesarias.
2. Confirma que usas Power BI Desktop de 64 bits.
3. Importa únicamente las columnas requeridas para el dashboard; no elimines campos que debas conservar para trazabilidad sin documentarlo.
4. Evita duplicar la consulta `Ventas` como tablas cargadas adicionales.
5. Desactiva la carga de consultas auxiliares si las creaste para pruebas.
6. Guarda el PBIX en un disco SSD local y reinicia Power BI antes de volver a cargar.
7. Si el problema persiste, utiliza la vista de Snowflake como fuente alternativa según las instrucciones del docente y conserva la validación de equivalencia.

### Problema 2: Los totales de Snowflake y Power BI no coinciden

**Síntomas:** ventas netas, margen, pedidos o porcentaje de cumplimiento muestran valores diferentes entre la consulta SQL y el dashboard.

**Causa probable:** se están usando filtros distintos, `SALE_DATE` no tiene el mismo tipo, Power BI incluye registros no válidos mientras Snowflake los excluye, se usa `COUNTROWS` en lugar de `DISTINCTCOUNT(SALE_ID)`, o la definición de ventas netas no es consistente.

**Corrección:**

1. Establece exactamente el mismo periodo, región, canal, categoría y condición de calidad en ambas herramientas.
2. Confirma que Power BI usa `Ventas[SALE_DATE]` relacionada con `Calendario[Date]`.
3. Verifica que la medida de pedidos sea:

```DAX
Pedidos = DISTINCTCOUNT(Ventas[SALE_ID])
```

4. Verifica que ventas netas utilice la misma fórmula que Snowflake:

```DAX
Ventas netas = SUM(Ventas[NET_SALES_AMOUNT])
```

5. Revisa que `NET_SALES_AMOUNT` corresponda a `SALES_AMOUNT - DISCOUNT_AMOUNT`.
6. Compara primero un subconjunto pequeño y controlado, por ejemplo una región, un canal y un mes.
7. Documenta cualquier diferencia residual de redondeo; no aceptes diferencias materiales sin explicación.

## Limpieza

1. Guarda todos los archivos finales y confirma que se encuentren en las carpetas requeridas.
2. Cierra Power BI Desktop y Excel para liberar bloqueos de archivos.
3. No elimines, alteres ni modifiques objetos del esquema `RAW`.
4. Si creaste una vista personal o de práctica y el instructor autoriza su eliminación, ejecuta únicamente sobre tu propio objeto:

```sql
DROP VIEW IF EXISTS DATA_ANALYTICS_FOUNDATIONS.ANALYTICS.VW_SALES_ANALYTICS_V1_[APELLIDO];
```

5. No ejecutes esta instrucción sobre una vista compartida del curso sin autorización explícita.
6. Elimina únicamente archivos temporales locales que hayas creado manualmente y que no formen parte de los entregables.
7. Cierra sesión de Snowflake y no guardes credenciales en scripts SQL, archivos PBIX, libros Excel ni archivos de texto.

## Resumen

En esta práctica construiste un flujo integral de analítica comercial sobre RetailNova v1.0, un dataset sintético masivo de exactamente 750,000 transacciones. Definiste una pregunta analítica a partir del brief previo, verificaste consistencia entre CSV, Snowflake y Parquet, documentaste controles de calidad y creaste una vista analítica trazable sin modificar la fuente inmutable.

También creaste un modelo de Power BI con calendario, medidas DAX y un dashboard de resumen y diagnóstico. El resultado final debe permitir a una audiencia comercial identificar dónde se concentra el desempeño observado, distinguir evidencia descriptiva de explicaciones causales no demostradas y priorizar próximos pasos realistas basados en métricas verificables.
