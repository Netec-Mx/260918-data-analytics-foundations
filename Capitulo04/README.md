# Exploración del desempeño comercial

## Metadatos

| Duration | Complexity | Bloom level |
|---|---|---|
| 98 minutos | Difícil | Aplicar |

## Descripción General

En esta práctica aplicarás el flujo analítico **preguntar, explorar, comparar, validar y concluir** sobre el dataset artificial masivo **Ventas Retail LATAM (2026.1)**, con aproximadamente 750,000 transacciones entre 2024-01-01 y 2025-12-31. Partirás de la fuente curada creada en la Práctica 2 y de los resultados descriptivos de la Práctica 3 para identificar señales comerciales relevantes por periodo, región, canal, categoría, subcategoría y tipo de cliente.

El resultado principal será un guion analítico validado que contenga al menos cinco señales de negocio, tres hallazgos priorizados y una lista de KPI y visualizaciones que servirán como base para el dashboard de la Práctica 5.

## Objetivos de Aprendizaje

- [ ] Formular preguntas analíticas medibles que definan métrica, periodo, segmento, comparación y decisión esperada.
- [ ] Comparar ventas netas, margen bruto, unidades, clientes, transacciones y ticket promedio por mes y trimestre.
- [ ] Segmentar el desempeño comercial por región, canal, categoría, subcategoría y tipo de cliente.
- [ ] Identificar y documentar al menos cinco señales comerciales, incluyendo cambios temporales, desviaciones, margen, recurrencia y anomalías.
- [ ] Validar cada señal mediante consultas de control antes de redactar conclusiones y recomendaciones.

## Prerrequisitos

**Conocimientos requeridos**

- Haber completado las Prácticas 1, 2 y 3.
- Comprender agregaciones SQL con `SUM`, `COUNT`, `COUNT(DISTINCT ...)`, `AVG`, `GROUP BY` y `HAVING`.
- Comprender comparaciones temporales, variación absoluta y variación porcentual.
- Conocer el uso básico de tablas dinámicas, filtros y gráficos en Microsoft Excel.
- Comprender que una observación no equivale a una causa confirmada.
- Saber distinguir entre ventas netas, margen bruto, unidades, clientes, transacciones y ticket promedio.

**Archivos y accesos requeridos**

- Archivo `C:\DAF\Batch_01\03_descriptive\03_analisis_descriptivo_ventas_clientes.xlsx`.
- Archivo `C:\DAF\Batch_01\sql\03_estadistica_descriptiva.sql`.
- Acceso de lectura al esquema `DATA_ANALYTICS_FOUNDATIONS.CURATED`.
- Rol Snowflake `DAF_ANALYST_ROLE`.
- Warehouse `DAF_LAB_WH`.
- Conexión Snowflake CLI `da_foundations_lab` o acceso mediante Snowsight.
- Fuente curada disponible como:

```text
DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
```

> Si tu organización utiliza un sufijo individual, reemplaza el nombre de la tabla por tu objeto asignado, por ejemplo: `VENTAS_TRANSACCIONES_CURADAS_2026_1_STUDENT01`.

## Entorno de Laboratorio

| Componente | Configuración de referencia |
|---|---|
| Sistema operativo | Windows 11 Pro, 64 bits |
| Memoria RAM | 16 GB mínimo; 32 GB recomendados |
| Almacenamiento | 20–25 GB libres en SSD |
| Excel | Microsoft Excel para Microsoft 365, 64 bits |
| Base de datos | Snowflake |
| Cliente Snowflake | Snowflake CLI 3.6.0, SnowSQL o Snowsight |
| Dataset | Ventas Retail LATAM (2026.1), artificial, aproximadamente 750,000 transacciones |
| Periodo de datos | 2024-01-01 a 2025-12-31 |

1. Abre PowerShell o Símbolo del sistema y crea las carpetas de trabajo si aún no existen:

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

2. Crea un archivo SQL para esta práctica:

```powershell
New-Item -ItemType File -Force `
-Path "C:\DAF\Batch_01\sql\04_exploracion_desempeno_comercial.sql"
```

3. Crea un libro de trabajo para documentar el análisis:

```powershell
New-Item -ItemType File -Force `
-Path "C:\DAF\Batch_01\04_exploration\04_exploracion_desempeno_comercial.xlsx"
```

4. Conéctate a Snowflake mediante CLI:

```bash
snow connection test --connection da_foundations_lab
snow sql --connection da_foundations_lab -q "SELECT CURRENT_USER(), CURRENT_ROLE(), CURRENT_WAREHOUSE();"
```

> No almacenes contraseñas, tokens ni credenciales en archivos `.sql`, `.xlsx`, `.pbix` o scripts.

## Instrucciones Paso a Paso

### Paso 1: Preparar el espacio de trabajo y confirmar la fuente

**Objective:** Confirmar que el entorno de análisis utiliza la tabla curada correcta, el rol autorizado y el periodo esperado del dataset.

**Instructions:**

1. Abre el archivo `C:\DAF\Batch_01\sql\04_exploracion_desempeno_comercial.sql`.
2. Agrega y ejecuta las siguientes instrucciones de contexto:

```sql
USE ROLE DAF_ANALYST_ROLE;
USE WAREHOUSE DAF_LAB_WH;
USE DATABASE DATA_ANALYTICS_FOUNDATIONS;
USE SCHEMA CURATED;

SELECT
    CURRENT_USER() AS usuario,
    CURRENT_ROLE() AS rol,
    CURRENT_WAREHOUSE() AS warehouse,
    CURRENT_DATABASE() AS base_datos,
    CURRENT_SCHEMA() AS esquema;
```

3. Confirma la estructura de la fuente curada. Ajusta el nombre si utilizas un objeto con sufijo personal:

```sql
DESCRIBE TABLE DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1;
```

4. Identifica los nombres reales de las columnas equivalentes a los siguientes conceptos:

| Concepto analítico | Nombre esperado o equivalente |
|---|---|
| Identificador de transacción | `ID_TRANSACCION` |
| Identificador de pedido | `ID_PEDIDO` |
| Fecha | `FECHA` o `FECHA_TRANSACCION` |
| Cliente | `ID_CLIENTE` |
| Región | `REGION` |
| Canal | `CANAL` |
| Categoría | `CATEGORIA` |
| Subcategoría | `SUBCATEGORIA` |
| Tipo de cliente | `TIPO_CLIENTE` o `SEGMENTO_CLIENTE` |
| Producto | `ID_PRODUCTO` o `PRODUCTO` |
| Unidades | `CANTIDAD` o `UNIDADES` |
| Venta neta | `VENTAS_NETAS` o `IMPORTE_NETO` |
| Margen bruto | `MARGEN_BRUTO` |
| Descuento | `DESCUENTO` o `PORCENTAJE_DESCUENTO` |
| Devolución | `ES_DEVOLUCION`, `TIPO_TRANSACCION` o columna equivalente |

5. Ejecuta el perfil mínimo de la fuente. Si algún nombre de columna difiere, reemplázalo antes de continuar:

```sql
SELECT
    COUNT(*) AS transacciones,
    COUNT(DISTINCT ID_TRANSACCION) AS transacciones_distintas,
    COUNT(DISTINCT ID_PEDIDO) AS pedidos_distintos,
    COUNT(DISTINCT ID_CLIENTE) AS clientes_distintos,
    MIN(FECHA) AS fecha_minima,
    MAX(FECHA) AS fecha_maxima,
    SUM(VENTAS_NETAS) AS ventas_netas_total,
    SUM(MARGEN_BRUTO) AS margen_bruto_total
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1;
```

6. Abre el archivo de la Práctica 3 y revisa los estadísticos previamente obtenidos. Úsalos como referencia, no como sustituto de las validaciones de esta práctica.

**Expected output:**

- El contexto de Snowflake muestra el rol `DAF_ANALYST_ROLE`, warehouse `DAF_LAB_WH`, base de datos `DATA_ANALYTICS_FOUNDATIONS` y esquema `CURATED`.
- La fuente contiene aproximadamente 750,000 transacciones.
- El rango de fechas cubre desde 2024-01-01 hasta 2025-12-31.
- Quedan documentados los nombres reales de las columnas que utilizarás.

**Verification:**

- El conteo de transacciones debe estar entre 500,000 y 1,000,000.
- `fecha_minima` debe ser `2024-01-01` y `fecha_maxima` debe ser `2025-12-31`, salvo que la documentación oficial indique una actualización controlada.
- No ejecutes `UPDATE`, `DELETE`, `TRUNCATE`, `DROP` ni `ALTER` sobre ningún objeto del esquema `RAW`.

---

### Paso 2: Formular la pregunta analítica y definir KPI comparables

**Objective:** Convertir la solicitud operativa de ventas en preguntas medibles y documentar definiciones consistentes de KPI.

**Instructions:**

1. Abre `C:\DAF\Batch_01\04_exploration\04_exploracion_desempeno_comercial.xlsx`.
2. Crea una hoja llamada `01_Preguntas_KPI`.
3. Registra la siguiente pregunta principal:

> ¿Qué variaciones mensuales y trimestrales de ventas netas, margen bruto, unidades, clientes y ticket promedio requieren atención comercial durante 2025, qué regiones, canales, categorías, subcategorías o tipos de cliente explican dichas variaciones y qué acciones deben priorizarse?

4. Registra al menos tres preguntas secundarias:

   - ¿Qué meses de 2025 presentan las mayores variaciones intermensuales en ventas netas y margen bruto?
   - ¿Qué región, canal o categoría se desvía más del desempeño total de la empresa?
   - ¿Qué productos combinan alto volumen de unidades con bajo margen porcentual?
   - ¿Qué segmentos tienen mayor recurrencia de compra?
   - ¿Existen descuentos atípicos o ventas excesivamente concentradas en pocos clientes?

5. Crea una tabla de definiciones de KPI como la siguiente. Ajusta los nombres de campo a tu fuente.

| KPI | Fórmula | Denominador o unidad | Regla de comparabilidad |
|---|---|---|---|
| Ventas netas | `SUM(VENTAS_NETAS)` | Moneda | Mismo filtro de fecha y tratamiento de devoluciones |
| Margen bruto | `SUM(MARGEN_BRUTO)` | Moneda | Mismo periodo y población que ventas netas |
| Margen % | `SUM(MARGEN_BRUTO) / NULLIF(SUM(VENTAS_NETAS), 0)` | Ventas netas | No promediar porcentajes de filas |
| Unidades | `SUM(CANTIDAD)` | Unidades | Revisar si devoluciones generan cantidades negativas |
| Clientes compradores | `COUNT(DISTINCT ID_CLIENTE)` | Clientes | Aplicar mismo criterio de inclusión de devoluciones |
| Transacciones | `COUNT(DISTINCT ID_TRANSACCION)` | Transacciones | Verificar duplicados |
| Pedidos | `COUNT(DISTINCT ID_PEDIDO)` | Pedidos | Usar si un pedido puede tener varias líneas |
| Ticket promedio | `SUM(VENTAS_NETAS) / NULLIF(COUNT(DISTINCT ID_PEDIDO), 0)` | Pedido | No usar promedio simple de líneas |
| Descuento promedio ponderado | `SUM(DESCUENTO_MONTO) / NULLIF(SUM(VENTAS_BRUTAS), 0)` | Ventas brutas | Preferible a promediar porcentajes por línea |

6. Anota una decisión potencial asociada a cada pregunta. Por ejemplo: revisar estrategia promocional, priorizar reposición, investigar margen de productos o revisar dependencia de clientes.

**Expected output:**

- Una hoja con una pregunta principal, al menos tres preguntas secundarias y una tabla de KPI.
- Las métricas incluyen definición, denominador y regla de comparabilidad.

**Verification:**

- Cada pregunta debe indicar como mínimo una métrica, un periodo o comparación, una dimensión de segmentación y una posible decisión.
- El ticket promedio debe usar pedidos distintos o la unidad transaccional definida oficialmente, no un promedio simple de importes de líneas.

---

### Paso 3: Construir la exploración mensual y trimestral

**Objective:** Obtener una vista temporal comparable de las métricas comerciales principales e identificar cambios intermensuales y trimestrales.

**Instructions:**

1. En el archivo SQL, crea una consulta mensual. Si `FECHA` contiene una hora, conserva `DATE_TRUNC` para agrupar correctamente:

```sql
WITH mensual AS (
    SELECT
        DATE_TRUNC('MONTH', FECHA) AS mes,
        SUM(VENTAS_NETAS) AS ventas_netas,
        SUM(MARGEN_BRUTO) AS margen_bruto,
        SUM(CANTIDAD) AS unidades,
        COUNT(DISTINCT ID_CLIENTE) AS clientes,
        COUNT(DISTINCT ID_PEDIDO) AS pedidos,
        COUNT(DISTINCT ID_TRANSACCION) AS transacciones,
        SUM(VENTAS_NETAS) / NULLIF(COUNT(DISTINCT ID_PEDIDO), 0) AS ticket_promedio,
        SUM(MARGEN_BRUTO) / NULLIF(SUM(VENTAS_NETAS), 0) AS margen_pct
    FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
    GROUP BY DATE_TRUNC('MONTH', FECHA)
),
comparacion AS (
    SELECT
        *,
        LAG(ventas_netas) OVER (ORDER BY mes) AS ventas_mes_anterior,
        LAG(margen_bruto) OVER (ORDER BY mes) AS margen_mes_anterior,
        LAG(unidades) OVER (ORDER BY mes) AS unidades_mes_anterior,
        LAG(clientes) OVER (ORDER BY mes) AS clientes_mes_anterior,
        LAG(ticket_promedio) OVER (ORDER BY mes) AS ticket_mes_anterior
    FROM mensual
)
SELECT
    mes,
    ventas_netas,
    margen_bruto,
    unidades,
    clientes,
    pedidos,
    transacciones,
    ticket_promedio,
    margen_pct,
    ventas_netas - ventas_mes_anterior AS variacion_ventas_abs,
    (ventas_netas / NULLIF(ventas_mes_anterior, 0)) - 1 AS variacion_ventas_pct,
    (margen_bruto / NULLIF(margen_mes_anterior, 0)) - 1 AS variacion_margen_pct,
    (unidades / NULLIF(unidades_mes_anterior, 0)) - 1 AS variacion_unidades_pct,
    (clientes / NULLIF(clientes_mes_anterior, 0)) - 1 AS variacion_clientes_pct,
    (ticket_promedio / NULLIF(ticket_mes_anterior, 0)) - 1 AS variacion_ticket_pct
FROM comparacion
ORDER BY mes;
```

2. Ejecuta una comparación trimestral:

```sql
WITH trimestral AS (
    SELECT
        DATE_TRUNC('QUARTER', FECHA) AS trimestre,
        SUM(VENTAS_NETAS) AS ventas_netas,
        SUM(MARGEN_BRUTO) AS margen_bruto,
        SUM(CANTIDAD) AS unidades,
        COUNT(DISTINCT ID_CLIENTE) AS clientes,
        COUNT(DISTINCT ID_PEDIDO) AS pedidos,
        SUM(VENTAS_NETAS) / NULLIF(COUNT(DISTINCT ID_PEDIDO), 0) AS ticket_promedio
    FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
    GROUP BY DATE_TRUNC('QUARTER', FECHA)
)
SELECT
    trimestre,
    ventas_netas,
    margen_bruto,
    unidades,
    clientes,
    pedidos,
    ticket_promedio,
    (ventas_netas / NULLIF(LAG(ventas_netas) OVER (ORDER BY trimestre), 0)) - 1 AS variacion_ventas_pct_t_t,
    (margen_bruto / NULLIF(LAG(margen_bruto) OVER (ORDER BY trimestre), 0)) - 1 AS variacion_margen_pct_t_t
FROM trimestral
ORDER BY trimestre;
```

3. Exporta los resultados a Excel o cópialos a una hoja llamada `02_Tendencia`.
4. Crea una tabla dinámica o gráfico de líneas con:
   - Eje: mes.
   - Valores: ventas netas y margen bruto.
   - Segundo gráfico: clientes y ticket promedio.
5. Marca los tres mayores incrementos y las tres mayores caídas intermensuales de ventas netas.

**Expected output:**

- Una tabla mensual de 24 meses.
- Una tabla trimestral de 8 trimestres.
- Dos visualizaciones de tendencia temporal.
- Una lista inicial de meses con variaciones relevantes.

**Verification:**

- La primera fila temporal no debe interpretarse como variación, porque no tiene periodo anterior.
- La suma de ventas mensuales debe coincidir con las ventas totales de la fuente, sujeto a redondeo.
- Verifica que todas las métricas se calculan sobre el mismo rango temporal antes de compararlas.

---

### Paso 4: Comparar desempeño por región, canal y categoría

**Objective:** Localizar los segmentos que explican una proporción relevante del desempeño comercial o de sus variaciones.

**Instructions:**

1. Ejecuta el resumen por región, canal y categoría para 2025:

```sql
SELECT
    REGION,
    CANAL,
    CATEGORIA,
    SUM(VENTAS_NETAS) AS ventas_netas,
    SUM(MARGEN_BRUTO) AS margen_bruto,
    SUM(MARGEN_BRUTO) / NULLIF(SUM(VENTAS_NETAS), 0) AS margen_pct,
    SUM(CANTIDAD) AS unidades,
    COUNT(DISTINCT ID_CLIENTE) AS clientes,
    COUNT(DISTINCT ID_PEDIDO) AS pedidos,
    SUM(VENTAS_NETAS) / NULLIF(COUNT(DISTINCT ID_PEDIDO), 0) AS ticket_promedio,
    COUNT(DISTINCT ID_TRANSACCION) AS transacciones
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
WHERE FECHA >= '2025-01-01'
  AND FECHA < '2026-01-01'
GROUP BY REGION, CANAL, CATEGORIA
ORDER BY ventas_netas DESC;
```

2. Ejecuta la comparación de cada región contra el total nacional de 2025:

```sql
WITH region_2025 AS (
    SELECT
        REGION,
        SUM(VENTAS_NETAS) AS ventas_netas,
        SUM(MARGEN_BRUTO) AS margen_bruto,
        SUM(CANTIDAD) AS unidades,
        COUNT(DISTINCT ID_CLIENTE) AS clientes
    FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
    WHERE FECHA >= '2025-01-01'
      AND FECHA < '2026-01-01'
    GROUP BY REGION
),
total AS (
    SELECT
        SUM(ventas_netas) AS ventas_total,
        SUM(margen_bruto) AS margen_total,
        SUM(unidades) AS unidades_total,
        SUM(clientes) AS clientes_suma_regiones
    FROM region_2025
)
SELECT
    r.REGION,
    r.ventas_netas,
    r.margen_bruto / NULLIF(r.ventas_netas, 0) AS margen_pct,
    r.unidades,
    r.clientes,
    r.ventas_netas / NULLIF(t.ventas_total, 0) AS participacion_ventas_pct,
    r.margen_bruto / NULLIF(t.margen_total, 0) AS participacion_margen_pct
FROM region_2025 r
CROSS JOIN total t
ORDER BY r.ventas_netas DESC;
```

3. En Excel, crea una hoja llamada `03_Segmentos`.
4. Construye dos tablas dinámicas:
   - **Tabla dinámica 1:** filas `REGION`; columnas `CANAL`; valores ventas netas, margen bruto y margen %.
   - **Tabla dinámica 2:** filas `CATEGORIA`; columnas `TIPO_CLIENTE`; valores ventas netas, unidades, clientes y ticket promedio.
5. Identifica una región o canal con una desviación relevante. Una desviación puede ser:
   - Alta participación en ventas, pero baja participación en margen.
   - Margen porcentual inferior al total.
   - Caída mensual superior a la caída total.
   - Ticket promedio significativamente distinto del promedio general.

**Expected output:**

- Tablas comparativas por región, canal, categoría y tipo de cliente.
- Una primera señal de desviación frente al total.
- Evidencia numérica de participación y margen por segmento.

**Verification:**

- No sumes clientes distintos entre regiones para representar el total nacional, porque un mismo cliente podría comprar en varias regiones.
- Para margen %, utiliza `SUM(MARGEN_BRUTO) / SUM(VENTAS_NETAS)`; no uses `AVG(margen_pct)` de grupos pequeños.
- Las comparaciones deben usar el mismo periodo, definición de ventas netas y tratamiento de devoluciones.

---

### Paso 5: Identificar productos de alto volumen y bajo margen

**Objective:** Detectar productos o subcategorías cuya combinación de volumen y margen pueda requerir revisión comercial, de precio o de costos.

**Instructions:**

1. Calcula los indicadores de producto para 2025:

```sql
WITH producto AS (
    SELECT
        ID_PRODUCTO,
        PRODUCTO,
        CATEGORIA,
        SUBCATEGORIA,
        SUM(CANTIDAD) AS unidades,
        SUM(VENTAS_NETAS) AS ventas_netas,
        SUM(MARGEN_BRUTO) AS margen_bruto,
        SUM(MARGEN_BRUTO) / NULLIF(SUM(VENTAS_NETAS), 0) AS margen_pct,
        COUNT(DISTINCT ID_PEDIDO) AS pedidos,
        COUNT(DISTINCT ID_CLIENTE) AS clientes
    FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
    WHERE FECHA >= '2025-01-01'
      AND FECHA < '2026-01-01'
    GROUP BY ID_PRODUCTO, PRODUCTO, CATEGORIA, SUBCATEGORIA
),
umbrales AS (
    SELECT
        PERCENTILE_CONT(0.75) WITHIN GROUP (ORDER BY unidades) AS p75_unidades,
        PERCENTILE_CONT(0.25) WITHIN GROUP (ORDER BY margen_pct) AS p25_margen_pct
    FROM producto
)
SELECT
    p.*,
    u.p75_unidades,
    u.p25_margen_pct
FROM producto p
CROSS JOIN umbrales u
WHERE p.unidades >= u.p75_unidades
  AND p.margen_pct <= u.p25_margen_pct
ORDER BY p.unidades DESC, p.margen_pct ASC;
```

2. Selecciona al menos tres productos o una subcategoría con alto volumen y bajo margen.
3. Ejecuta una consulta de detalle mensual para uno de los productos identificados:

```sql
SELECT
    DATE_TRUNC('MONTH', FECHA) AS mes,
    ID_PRODUCTO,
    PRODUCTO,
    SUM(CANTIDAD) AS unidades,
    SUM(VENTAS_NETAS) AS ventas_netas,
    SUM(MARGEN_BRUTO) AS margen_bruto,
    SUM(MARGEN_BRUTO) / NULLIF(SUM(VENTAS_NETAS), 0) AS margen_pct,
    AVG(DESCUENTO) AS descuento_promedio_simple
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
WHERE ID_PRODUCTO = '<ID_PRODUCTO_SELECCIONADO>'
  AND FECHA >= '2025-01-01'
  AND FECHA < '2026-01-01'
GROUP BY DATE_TRUNC('MONTH', FECHA), ID_PRODUCTO, PRODUCTO
ORDER BY mes;
```

4. Documenta la señal como una observación, no como una causa. Ejemplo:

> El producto X se ubica en el cuartil superior de unidades y en el cuartil inferior de margen porcentual durante 2025. Requiere revisión de precio, costo, descuentos y devoluciones antes de recomendar una modificación comercial.

5. Registra la señal en una hoja llamada `04_Senales`.

**Expected output:**

- Una lista de productos de alto volumen y bajo margen.
- Una señal documentada con volumen, ventas, margen y periodo.
- Un producto seleccionado para validación posterior.

**Verification:**

- Confirma que el margen bajo no se debe a ventas netas cercanas a cero.
- Revisa si el producto tiene cantidades negativas, devoluciones o descuentos excepcionalmente altos.
- No afirmes que un descuento “causó” el margen bajo sin contrastar descuentos, costos u otras variables disponibles.

---

### Paso 6: Analizar recurrencia de clientes y concentración de ventas

**Objective:** Identificar segmentos con alta recurrencia y evaluar si las ventas dependen de pocos clientes.

**Instructions:**

1. Calcula la recurrencia por tipo de cliente y canal:

```sql
WITH cliente_segmento AS (
    SELECT
        TIPO_CLIENTE,
        CANAL,
        ID_CLIENTE,
        COUNT(DISTINCT ID_PEDIDO) AS pedidos_cliente,
        SUM(VENTAS_NETAS) AS ventas_cliente
    FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
    WHERE FECHA >= '2025-01-01'
      AND FECHA < '2026-01-01'
    GROUP BY TIPO_CLIENTE, CANAL, ID_CLIENTE
)
SELECT
    TIPO_CLIENTE,
    CANAL,
    COUNT(*) AS clientes,
    AVG(pedidos_cliente) AS pedidos_promedio_por_cliente,
    MEDIAN(pedidos_cliente) AS pedidos_mediana_por_cliente,
    SUM(IFF(pedidos_cliente >= 2, 1, 0)) AS clientes_recurrentes,
    SUM(IFF(pedidos_cliente >= 2, 1, 0)) / NULLIF(COUNT(*), 0) AS tasa_recorrencia,
    SUM(ventas_cliente) AS ventas_netas
FROM cliente_segmento
GROUP BY TIPO_CLIENTE, CANAL
ORDER BY tasa_recorrencia DESC, ventas_netas DESC;
```

2. Calcula la concentración de ventas por cliente dentro de cada región:

```sql
WITH ventas_cliente AS (
    SELECT
        REGION,
        ID_CLIENTE,
        SUM(VENTAS_NETAS) AS ventas_cliente
    FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
    WHERE FECHA >= '2025-01-01'
      AND FECHA < '2026-01-01'
    GROUP BY REGION, ID_CLIENTE
),
ranking AS (
    SELECT
        REGION,
        ID_CLIENTE,
        ventas_cliente,
        ROW_NUMBER() OVER (
            PARTITION BY REGION
            ORDER BY ventas_cliente DESC
        ) AS rango_cliente
    FROM ventas_cliente
)
SELECT
    REGION,
    SUM(ventas_cliente) AS ventas_region,
    SUM(IFF(rango_cliente <= 10, ventas_cliente, 0)) AS ventas_top_10_clientes,
    SUM(IFF(rango_cliente <= 10, ventas_cliente, 0))
        / NULLIF(SUM(ventas_cliente), 0) AS concentracion_top_10_pct
FROM ranking
GROUP BY REGION
ORDER BY concentracion_top_10_pct DESC;
```

3. Registra una señal de recurrencia. Ejemplo:

> El segmento [tipo de cliente/canal] presenta la mayor tasa de recurrencia y un número suficiente de clientes, por lo que puede ser prioritario para acciones de fidelización o venta cruzada.

4. Registra una señal de concentración. Ejemplo:

> La región [nombre] concentra un porcentaje elevado de sus ventas en los diez clientes principales. Esta dependencia puede representar una oportunidad de gestión de cuentas y un riesgo comercial que debe monitorearse.

5. Agrega ambas señales a la hoja `04_Senales`.

**Expected output:**

- Una tabla de recurrencia por tipo de cliente y canal.
- Una tabla de concentración de ventas por región.
- Dos señales documentadas: una de recurrencia y una de concentración.

**Verification:**

- Define “cliente recurrente” de forma explícita: en esta práctica, dos o más pedidos distintos durante 2025.
- Verifica que `ID_PEDIDO` no sea nulo antes de interpretar recurrencia.
- No uses el promedio de pedidos como único indicador; compara también la mediana y el número de clientes del segmento.

---

### Paso 7: Detectar descuentos atípicos y validar las cinco señales

**Objective:** Ejecutar consultas de control para comprobar que las señales identificadas no son producto de filtros erróneos, periodos incompletos, duplicados, devoluciones o denominadores inconsistentes.

**Instructions:**

1. Identifica descuentos atípicos. Si `DESCUENTO` representa porcentaje, utiliza esta consulta:

```sql
SELECT
    REGION,
    CANAL,
    CATEGORIA,
    SUBCATEGORIA,
    COUNT(*) AS transacciones,
    AVG(DESCUENTO) AS descuento_promedio,
    PERCENTILE_CONT(0.95) WITHIN GROUP (ORDER BY DESCUENTO) AS descuento_p95,
    MAX(DESCUENTO) AS descuento_maximo,
    SUM(VENTAS_NETAS) AS ventas_netas,
    SUM(MARGEN_BRUTO) / NULLIF(SUM(VENTAS_NETAS), 0) AS margen_pct
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
WHERE FECHA >= '2025-01-01'
  AND FECHA < '2026-01-01'
GROUP BY REGION, CANAL, CATEGORIA, SUBCATEGORIA
ORDER BY descuento_p95 DESC;
```

2. Si existe una columna de importe bruto y descuento monetario, calcula el descuento ponderado:

```sql
SELECT
    CATEGORIA,
    CANAL,
    SUM(DESCUENTO_MONTO) / NULLIF(SUM(VENTAS_BRUTAS), 0) AS descuento_ponderado_pct,
    SUM(VENTAS_NETAS) AS ventas_netas,
    SUM(MARGEN_BRUTO) / NULLIF(SUM(VENTAS_NETAS), 0) AS margen_pct
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
WHERE FECHA >= '2025-01-01'
  AND FECHA < '2026-01-01'
GROUP BY CATEGORIA, CANAL
ORDER BY descuento_ponderado_pct DESC;
```

3. Ejecuta el control de cobertura temporal para los meses asociados a señales de caída o crecimiento:

```sql
SELECT
    DATE_TRUNC('MONTH', FECHA) AS mes,
    COUNT(DISTINCT FECHA) AS dias_con_registros,
    MIN(FECHA) AS fecha_minima_mes,
    MAX(FECHA) AS fecha_maxima_mes,
    COUNT(*) AS transacciones,
    COUNT(DISTINCT ID_TRANSACCION) AS transacciones_distintas
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
GROUP BY DATE_TRUNC('MONTH', FECHA)
ORDER BY mes;
```

4. Ejecuta el control de duplicados:

```sql
SELECT
    ID_TRANSACCION,
    COUNT(*) AS registros_por_transaccion
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
GROUP BY ID_TRANSACCION
HAVING COUNT(*) > 1
ORDER BY registros_por_transaccion DESC
LIMIT 100;
```

5. Ejecuta un control de devoluciones. Ajusta la condición según la estructura real de la tabla:

```sql
SELECT
    DATE_TRUNC('MONTH', FECHA) AS mes,
    COUNT(*) AS registros,
    SUM(IFF(ES_DEVOLUCION = TRUE, 1, 0)) AS registros_devolucion,
    SUM(IFF(ES_DEVOLUCION = TRUE, VENTAS_NETAS, 0)) AS ventas_netas_devolucion,
    SUM(VENTAS_NETAS) AS ventas_netas_totales
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
GROUP BY DATE_TRUNC('MONTH', FECHA)
ORDER BY mes;
```

6. En la hoja `04_Senales`, crea una tabla con las siguientes columnas:

| ID | Señal | Métrica y comparación | Segmento | Evidencia | Consulta de control | Resultado de validación | Nivel de confianza |
|---|---|---|---|---|---|---|---|

7. Documenta obligatoriamente estas cinco señales:

| ID | Señal mínima requerida |
|---|---|
| S1 | Crecimiento o caída intermensual de ventas, margen, unidades, clientes o ticket |
| S2 | Región, canal o categoría con desviación frente al total |
| S3 | Producto o subcategoría de alto volumen y bajo margen |
| S4 | Segmento con alta recurrencia |
| S5 | Anomalía de descuentos o concentración de ventas en pocos clientes |

8. Para cada señal, registra al menos una consulta de control que revise uno o varios de estos elementos: conteo de transacciones, cobertura de fechas, devoluciones, duplicados, filtros y denominadores.

**Expected output:**

- Cinco señales documentadas.
- Cinco validaciones asociadas.
- Una señal de descuento atípico o de concentración de ventas.
- Un nivel de confianza preliminar para cada señal.

**Verification:**

- Una señal queda como **validada** solo si la métrica, el periodo, los filtros y los conteos son consistentes.
- Si hay meses con cobertura incompleta, registra la limitación y evita comparar el mes incompleto como si fuera un mes completo.
- Si se detectan duplicados, determina si representan líneas válidas de pedido o duplicación no esperada antes de modificar cualquier interpretación.

---

### Paso 8: Priorizar tres hallazgos y redactar conclusiones prudentes

**Objective:** Convertir las señales validadas en hallazgos priorizados, accionables y con un nivel explícito de certeza.

**Instructions:**

1. Crea una hoja llamada `05_Priorizacion`.
2. Copia las cinco señales validadas y asígnales una puntuación de 1 a 5 en tres criterios:

| Criterio | Pregunta de evaluación |
|---|---|
| Impacto potencial | ¿Qué tan relevante es el efecto económico, de margen, riesgo o clientes? |
| Confianza en los datos | ¿La señal superó los controles de fecha, duplicados, devoluciones y denominadores? |
| Accionabilidad | ¿Existe una acción comercial, operativa o analítica razonable? |

3. Calcula la prioridad total:

```text
Prioridad total = Impacto potencial + Confianza en los datos + Accionabilidad
```

4. Selecciona los tres hallazgos con mayor prioridad. Si existe empate, prioriza primero mayor impacto y luego mayor confianza.
5. Redacta cada hallazgo con la siguiente estructura:

> **Hallazgo:** qué ocurrió y en qué segmento o periodo.  
> **Evidencia:** métricas, comparación y magnitud.  
> **Interpretación:** explicación plausible, sin afirmar causalidad no comprobada.  
> **Recomendación:** acción sugerida.  
> **Certeza y límites:** controles realizados y aspectos pendientes de confirmar.

6. Utiliza esta plantilla:

| Prioridad | Hallazgo | Evidencia cuantitativa | Interpretación prudente | Recomendación | Confianza y límites |
|---:|---|---|---|---|---|
| 1 |  |  |  |  |  |
| 2 |  |  |  |  |  |
| 3 |  |  |  |  |  |

7. Evita afirmaciones como “la campaña causó la caída” si el dataset no contiene evidencia directa de campañas, tráfico web, inventario o precios históricos. Sustituye por expresiones como:
   - “es consistente con”.
   - “se concentra en”.
   - “requiere contraste con”.
   - “sugiere una oportunidad de revisión”.
   - “no permite atribuir causalidad por sí solo”.

**Expected output:**

- Tres hallazgos priorizados.
- Cada hallazgo contiene evidencia, recomendación, nivel de confianza y límites.
- Las conclusiones diferencian hechos observados de interpretaciones plausibles.

**Verification:**

- Los tres hallazgos deben provenir de señales con validación documentada.
- Ninguna recomendación debe depender de una causalidad que el dataset no demuestra.
- Al menos un hallazgo debe incluir una acción operativa o comercial concreta, como revisar descuentos, margen, reposición, estrategia de canal o gestión de clientes.

---

### Paso 9: Preparar el guion de negocio y los insumos para el dashboard

**Objective:** Convertir el análisis exploratorio en requisitos trazables para las visualizaciones y KPI de la Práctica 5.

**Instructions:**

1. Crea una hoja llamada `06_Guion_Dashboard`.
2. Registra un mensaje ejecutivo de no más de tres oraciones que resuma los tres hallazgos priorizados.
3. Crea la lista de KPI requeridos para el dashboard:

| KPI | Visualización sugerida | Segmentadores requeridos | Fuente o consulta |
|---|---|---|---|
| Ventas netas | Tarjeta y línea mensual | Fecha, región, canal | Consulta mensual |
| Margen bruto y margen % | Tarjeta y barras por categoría | Fecha, región, categoría | Consulta de segmentos |
| Unidades | Tarjeta y barras por producto | Fecha, categoría, subcategoría | Consulta de producto |
| Clientes compradores | Tarjeta y línea mensual | Fecha, tipo de cliente, canal | Consulta mensual |
| Ticket promedio | Tarjeta y línea mensual | Fecha, región, canal | Consulta mensual |
| Recurrencia | Barras por tipo de cliente/canal | Fecha, tipo de cliente | Consulta de recurrencia |
| Concentración de clientes | Barras por región | Fecha, región | Consulta de concentración |
| Descuento y margen | Dispersión o tabla de excepción | Fecha, canal, categoría | Consulta de descuentos |

4. Define al menos cuatro visualizaciones para la Práctica 5:
   - Tendencia mensual de ventas netas y margen bruto.
   - Comparación de ventas y margen por región o canal.
   - Desempeño por categoría/subcategoría.
   - Tabla o gráfico de productos de alto volumen y bajo margen.
   - Opcional: recurrencia, concentración o descuentos atípicos.

5. Guarda el libro como:

```text
C:\DAF\Batch_01\04_exploration\04_exploracion_desempeno_comercial.xlsx
```

6. Guarda el script SQL final como:

```text
C:\DAF\Batch_01\sql\04_exploracion_desempeno_comercial.sql
```

**Expected output:**

- Un guion ejecutivo breve.
- Una lista de KPI con definición y visualización propuesta.
- Al menos cuatro visualizaciones especificadas para el dashboard.
- Dos entregables guardados en las rutas requeridas.

**Verification:**

- Cada visualización propuesta debe responder a una pregunta o hallazgo del análisis.
- Cada KPI debe tener una definición consistente con la hoja `01_Preguntas_KPI`.
- Los hallazgos del guion deben ser trazables a una consulta SQL o tabla dinámica documentada.

## Validación y Pruebas

Completa la siguiente lista antes de entregar la práctica.

| Prueba | Criterio de aprobación |
|---|---|
| Fuente correcta | Se utilizó una tabla del esquema `CURATED`, no se modificó ningún objeto de `RAW`. |
| Volumen y periodo | La fuente contiene entre 500,000 y 1,000,000 transacciones y cubre el periodo oficial 2024-01-01 a 2025-12-31. |
| Comparabilidad temporal | Las comparaciones mensuales y trimestrales usan periodos completos y la misma definición de KPI. |
| Ventas netas | La suma mensual coincide con el total de ventas netas de la fuente, considerando redondeo. |
| Ticket promedio | Se calcula como ventas netas entre pedidos distintos, o según la unidad oficial documentada. |
| Margen porcentual | Se calcula mediante razón de sumas, no mediante promedio de porcentajes de filas. |
| Segmentación | Se analizaron región, canal, categoría, subcategoría y tipo de cliente. |
| Cinco señales | Existen cinco señales, incluyendo las cinco categorías requeridas. |
| Consultas de control | Cada señal tiene una validación de conteos, fechas, devoluciones, duplicados, filtros o denominadores. |
| Conclusiones | Se priorizaron tres hallazgos y se diferenciaron hechos, interpretación, recomendación y límites. |
| Trazabilidad | El libro Excel y el script SQL contienen referencias claras a las consultas y resultados utilizados. |
| Preparación de dashboard | Se definieron KPI, segmentadores y al menos cuatro visualizaciones para la Práctica 5. |

**Criterio de entrega:** la práctica se considera completa si todas las pruebas anteriores están aprobadas, o si las excepciones están documentadas con su impacto analítico y una acción pendiente.

## Solución de Problemas

### Problema 1: Los resultados mensuales muestran una caída extrema en el último mes o el número de días es inferior al esperado

**Síntomas**

- El último mes presenta ventas, clientes y transacciones mucho menores que los meses anteriores.
- La consulta de cobertura muestra menos días con registros que los demás meses.
- La variación intermensual parece inusualmente negativa.

**Causa probable**

El periodo puede estar incompleto por una carga parcial, un filtro de fecha incorrecto o una conversión inadecuada de una columna de fecha y hora.

**Solución**

1. Ejecuta la consulta de cobertura temporal del Paso 7.
2. Confirma `MIN(FECHA)`, `MAX(FECHA)` y el número de días con registros para el mes afectado.
3. Revisa que el filtro utilice límites completos, por ejemplo:

```sql
WHERE FECHA >= '2025-08-01'
  AND FECHA < '2025-09-01'
```

4. Si el mes está incompleto, no lo uses para una comparación mensual definitiva. Documenta la limitación y utiliza el último mes completo disponible.

### Problema 2: El ticket promedio o el margen porcentual no coincide entre Excel y Snowflake

**Síntomas**

- Excel muestra un ticket promedio diferente al resultado SQL.
- El margen porcentual de una tabla dinámica parece demasiado alto o bajo.
- Los totales cambian al agregar o quitar segmentos.

**Causa probable**

Se está usando `AVERAGE` sobre valores ya calculados por línea, se están contando filas en lugar de pedidos distintos, o se están aplicando filtros distintos en Excel y Snowflake.

**Solución**

1. En Snowflake, calcula el ticket como:

```sql
SUM(VENTAS_NETAS) / NULLIF(COUNT(DISTINCT ID_PEDIDO), 0)
```

2. Calcula el margen porcentual como:

```sql
SUM(MARGEN_BRUTO) / NULLIF(SUM(VENTAS_NETAS), 0)
```

3. En Excel, usa campos calculados o medidas que reproduzcan estas razones de sumas.
4. Compara filtros de fecha, región, canal, devoluciones y categorías entre ambas herramientas.
5. Verifica si una transacción puede contener varias líneas; en ese caso, `COUNT(*)` no representa necesariamente pedidos.

## Limpieza

1. Guarda y cierra el archivo Excel:

```text
C:\DAF\Batch_01\04_exploration\04_exploracion_desempeno_comercial.xlsx
```

2. Guarda el script SQL final:

```text
C:\DAF\Batch_01\sql\04_exploracion_desempeno_comercial.sql
```

3. Cierra las consultas, sesiones o pestañas de Snowsight que ya no necesites.
4. No elimines ni modifiques objetos en `DATA_ANALYTICS_FOUNDATIONS.RAW`.
5. No elimines la tabla o vista curada común utilizada por el curso.
6. Si generaste archivos temporales de exportación, conserva únicamente los que sean necesarios como evidencia de tu análisis en:

```text
C:\DAF\Batch_01\04_exploration\
```

7. Elimina copias temporales que contengan datos exportados fuera de la ruta oficial del laboratorio, según las políticas de tu institución.

## Resumen

En esta práctica aplicaste un flujo completo de exploración comercial: formulaste preguntas analíticas, exploraste KPI temporales, comparaste segmentos, validaste resultados y redactaste conclusiones prudentes. También identificaste señales de variación intermensual, desviación regional o por canal, productos de alto volumen y bajo margen, recurrencia de clientes y anomalías de descuento o concentración.

Los entregables de esta práctica son la base analítica de la siguiente actividad: un dashboard de desempeño comercial en Power BI Desktop que deberá mantener trazabilidad con la fuente curada y con las definiciones de KPI validadas.

**Entregables finales**

- `C:\DAF\Batch_01\04_exploration\04_exploracion_desempeno_comercial.xlsx`
- `C:\DAF\Batch_01\sql\04_exploracion_desempeno_comercial.sql`

**Recursos de consulta**

- [Guía de análisis exploratorio de datos del NIST](https://www.itl.nist.gov/div898/handbook/eda/eda.htm)
- [Documentación de funciones de ventana de Snowflake](https://docs.snowflake.com/en/sql-reference/functions-analytic)
- [Documentación de tablas dinámicas de Microsoft Excel](https://support.microsoft.com/excel)
