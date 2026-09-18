# Dashboard básico de desempeño de ventas

## Metadatos

| Duration | Complexity | Bloom level |
|---|---|---|
| 137 minutos | Difícil | Create |

## Descripción General

En esta práctica construirás un dashboard básico de desempeño de ventas en Power BI Desktop usando la tabla curada oficial de Snowflake o, cuando no sea posible la conexión directa, un extracto CSV curado autorizado. El dashboard responderá preguntas analíticas priorizadas en las Prácticas 1 y 4 mediante KPI, tendencias temporales, comparaciones categóricas, composición por canal y detalle transaccional agregado.

El trabajo incluye revisión de tipos de datos, creación de una tabla calendario, medidas DAX, configuración de interacciones y validación cruzada contra Snowflake y Excel. El resultado no debe ser una colección de gráficos: debe comunicar tres hallazgos priorizados, con contexto temporal, cobertura, filtros y nivel de certeza apropiados.

## Objetivos de Aprendizaje

- [ ] Conectar Power BI Desktop a la fuente curada de ventas y verificar campos, tipos de datos y relaciones básicas.
- [ ] Crear una tabla calendario entre 2024-01-01 y 2025-12-31 y relacionarla correctamente con la fecha de venta.
- [ ] Implementar medidas DAX para ventas netas, margen bruto, margen porcentual, unidades, transacciones, clientes únicos y ticket promedio.
- [ ] Diseñar un dashboard que use visualizaciones adecuadas para KPI, tendencia, comparación, composición y detalle.
- [ ] Validar al menos cinco KPI o agregados contra Snowflake y los controles de Excel antes de entregar el archivo PBIX.

## Prerrequisitos

### Conocimientos requeridos

Debes haber completado las Prácticas 1 a 4 y disponer de los siguientes insumos:

- Brief analítico elaborado en la Práctica 1, especialmente la pregunta de negocio priorizada, las definiciones de KPI y el período de análisis.
- Consulta o tabla curada generada en la Práctica 2.
- Validaciones de calidad, reglas aplicadas y limitaciones identificadas en la Práctica 3.
- Hallazgos, hipótesis, narrativa y recomendaciones preliminares de la Práctica 4.
- Fundamentos de Power BI Desktop: carga de datos, paneles **Datos**, **Visualizaciones** y **Filtros**, y creación de medidas DAX.
- Fundamentos de Excel para revisar controles agregados.
- Fundamentos de Snowflake SQL para ejecutar consultas `SELECT`.

### Accesos y archivos requeridos

Debes confirmar antes de iniciar:

- Acceso de lectura a `DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1`.
- Rol `DAF_ANALYST_ROLE` y warehouse `DAF_LAB_WH`, o los equivalentes asignados por el instructor.
- Conexión Snowflake CLI llamada `da_foundations_lab`, si usarás terminal para validar.
- Power BI Desktop instalado, versión 64 bits.
- Archivo de Excel de controles de prácticas anteriores, si fue entregado por el instructor.
- Extracto CSV curado autorizado, solo si no es posible usar la conexión directa a Snowflake.
- Espacio disponible en `C:\DAF\Batch_01\05_dashboard\` para el archivo PBIX, exportaciones y capturas.

> **Importante:** El dataset es artificial y contiene aproximadamente 750,000 transacciones, cerca de 120,000 clientes y aproximadamente 1,500 productos entre el 2024-01-01 y el 2025-12-31. No contiene datos personales reales. No ejecutes `UPDATE`, `DELETE`, `TRUNCATE`, `DROP` ni `ALTER` sobre objetos del esquema `RAW`.

## Entorno de Laboratorio

### Directorios de trabajo

Abre PowerShell y valida o crea la estructura obligatoria:

```powershell
New-Item -ItemType Directory -Force -Path `
  C:\DAF\Batch_01\00_source, `
  C:\DAF\Batch_01\01_brief, `
  C:\DAF\Batch_01\02_quality, `
  C:\DAF\Batch_01\03_descriptive, `
  C:\DAF\Batch_01\04_exploration, `
  C:\DAF\Batch_01\05_dashboard, `
  C:\DAF\Batch_01\sql
```

Verifica la estructura:

```powershell
Get-ChildItem C:\DAF\Batch_01\
```

### Recursos técnicos

| Componente | Requisito de referencia |
|---|---|
| Sistema operativo | Windows 11 Pro de 64 bits |
| Memoria RAM | 16 GB mínimo; 32 GB recomendados |
| Espacio libre | 20–25 GB en SSD |
| Procesador | Intel Core i5 10.ª generación / AMD Ryzen 5 5000 Series o superior |
| Resolución | 1920 × 1080 píxeles o superior |
| Power BI Desktop | 2.138.1004.0, 64 bits, o versión equivalente aprobada |
| Snowflake | Snowsight o Snowflake CLI 3.6.0 / SnowSQL 1.3.2 |
| Fuente principal | `DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1` |
| Fuente alternativa | Extracto CSV curado autorizado por el instructor |

### Convención de archivos

Usa los siguientes nombres, reemplazando `APELLIDO_NOMBRE` por tu identificador:

| Entregable | Ruta sugerida |
|---|---|
| Archivo Power BI | `C:\DAF\Batch_01\05_dashboard\Lab05_Dashboard_Ventas_APELLIDO_NOMBRE.pbix` |
| Consultas de validación | `C:\DAF\Batch_01\sql\Lab05_validacion_dashboard_APELLIDO_NOMBRE.sql` |
| Evidencia de validación | `C:\DAF\Batch_01\05_dashboard\Lab05_validacion_APELLIDO_NOMBRE.xlsx` |
| Captura del dashboard | `C:\DAF\Batch_01\05_dashboard\Lab05_dashboard_APELLIDO_NOMBRE.png` |

> **Seguridad:** No guardes contraseñas institucionales en archivos `.sql`, `.pbix`, `.xlsx`, archivos de texto ni capturas. Usa el mecanismo de autenticación institucional solicitado por Snowflake y Power BI.

## Instrucciones Paso a Paso

### Paso 1: Revisar el brief, los hallazgos y las definiciones de KPI

**Objetivo:** Traducir los insumos de las prácticas anteriores en requisitos explícitos para el dashboard.

**Instrucciones**

1. Abre el brief de la Práctica 1 ubicado en `C:\DAF\Batch_01\01_brief\`.
2. Abre los resultados descriptivos y exploratorios de las Prácticas 3 y 4.
3. Identifica la pregunta analítica priorizada. Debe contener, como mínimo:
   - Métrica principal.
   - Dimensión de comparación.
   - Período.
   - Segmento, si aplica.
   - Decisión que busca apoyar.
4. Escribe en una hoja de trabajo o documento de apoyo tres hallazgos que el dashboard deberá comunicar. Formula cada uno sin asumir el resultado antes de validarlo.
5. Define los KPI que mostrarás y confirma su fórmula con las definiciones acordadas previamente.
6. Usa la siguiente matriz como plantilla. Complétala con tus decisiones antes de construir visuales.

| Pregunta de negocio | Métrica | Dimensión | Período | Visual seleccionado | Decisión apoyada |
|---|---|---|---|---|---|
| ¿Cuál es el desempeño comercial del período seleccionado? | Ventas netas, margen, unidades | Total filtrado | Seleccionado por usuario | Tarjetas KPI | Seguimiento ejecutivo |
| ¿Cómo evolucionan las ventas netas mensualmente? | Ventas netas | Mes | 2024–2025 o período filtrado | Línea | Detectar tendencia o estacionalidad |
| ¿Qué región o categoría concentra el desempeño? | Ventas netas o margen bruto | Región o categoría | Período filtrado | Barras horizontales | Priorizar revisión comercial |
| ¿Qué proporción aporta cada canal? | Ventas netas | Canal | Período filtrado | Barras apiladas o dona simple | Evaluar mezcla comercial |

7. Evita formular títulos conclusivos como “La región Norte es la mejor” antes de validar los resultados. Durante el diseño utiliza títulos neutrales, por ejemplo: “Ventas netas por región — período seleccionado”.

**Resultado esperado**

Dispones de una matriz que conecta preguntas de negocio, métricas, dimensiones, visuales y decisiones. Los visuales planeados tienen una justificación analítica clara.

**Verificación**

Confirma que puedes responder “sí” a las siguientes preguntas:

- ¿Cada visual responde una pregunta de negocio concreta?
- ¿La métrica coincide con la definición aprobada en prácticas anteriores?
- ¿El período y la población analizada son explícitos?
- ¿Los tres hallazgos esperados pueden comunicarse sin extrapolar más allá del dataset artificial?

---

### Paso 2: Inspeccionar la fuente curada y registrar el mapeo de campos

**Objetivo:** Confirmar la estructura de la tabla curada antes de conectarla a Power BI.

**Instrucciones**

1. Inicia sesión en Snowsight o usa Snowflake CLI con la conexión institucional.
2. Configura el contexto de trabajo:

```sql
USE ROLE DAF_ANALYST_ROLE;
USE WAREHOUSE DAF_LAB_WH;
USE DATABASE DATA_ANALYTICS_FOUNDATIONS;
USE SCHEMA CURATED;
```

3. Inspecciona las columnas disponibles:

```sql
DESC TABLE DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1;
```

4. Obtén una muestra de registros sin modificar la fuente:

```sql
SELECT *
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
LIMIT 20;
```

5. Confirma el total de filas y cobertura temporal. Ajusta los nombres de columnas si la tabla curada usa nombres distintos:

```sql
SELECT
    COUNT(*) AS transacciones,
    MIN(FECHA_VENTA) AS fecha_minima,
    MAX(FECHA_VENTA) AS fecha_maxima,
    COUNT(DISTINCT CLIENTE_ID) AS clientes_unicos
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1;
```

6. Registra el mapeo real de campos. Los nombres siguientes son convenciones de esta guía; debes usar los nombres reales de tu tabla:

| Uso analítico | Nombre esperado o equivalente posible | Tipo esperado |
|---|---|---|
| Identificador de transacción | `TRANSACCION_ID` | Texto o número entero |
| Fecha de venta | `FECHA_VENTA` | Fecha o fecha/hora |
| Cliente | `CLIENTE_ID` | Texto o número entero |
| Región | `REGION` | Texto |
| Canal | `CANAL` | Texto |
| Categoría | `CATEGORIA` | Texto |
| Producto | `PRODUCTO` o `PRODUCTO_NOMBRE` | Texto |
| Unidades | `UNIDADES` o `CANTIDAD` | Número entero |
| Ventas netas | `VENTAS_NETAS`, `IMPORTE_NETO` o equivalente | Decimal fijo / moneda |
| Costo | `COSTO_TOTAL` o equivalente | Decimal fijo / moneda |
| Margen bruto | `MARGEN_BRUTO` si existe | Decimal fijo / moneda |

7. Guarda las consultas usadas en:

```text
C:\DAF\Batch_01\sql\Lab05_validacion_dashboard_APELLIDO_NOMBRE.sql
```

> **Criterio de trazabilidad:** La fuente utilizada en Power BI debe ser la tabla curada aprobada o un CSV derivado de esa misma versión curada. No sustituyas el dataset por otro archivo, por una tabla RAW ni por datos copiados manualmente.

**Resultado esperado**

Conoces los nombres reales de los campos, su granularidad y el intervalo de fechas disponible. Has confirmado que la fuente es curada y que cubre el período oficial esperado.

**Verificación**

La consulta de cobertura debe mostrar un intervalo que incluya fechas desde `2024-01-01` hasta `2025-12-31`, salvo que la Práctica 2 haya documentado exclusiones justificadas. Si aparecen fechas fuera del período, regístralo como una observación de calidad y utiliza el intervalo validado en el dashboard.

---

### Paso 3: Conectar Power BI Desktop a la fuente curada

**Objetivo:** Cargar la fuente validada en Power BI Desktop conservando trazabilidad y tipos de datos correctos.

**Instrucciones**

1. Abre Power BI Desktop.
2. Selecciona **Archivo > Guardar como**.
3. Guarda el archivo inicialmente como:

```text
C:\DAF\Batch_01\05_dashboard\Lab05_Dashboard_Ventas_APELLIDO_NOMBRE.pbix
```

4. Selecciona **Obtener datos > Más > Base de datos > Snowflake**.
5. En **Servidor**, escribe el nombre de cuenta proporcionado por el instructor, por ejemplo:

```text
<ORG_ACCOUNT>
```

6. En **Almacén**, escribe:

```text
DAF_LAB_WH
```

7. Selecciona **Importar** como modo de conectividad, a menos que el instructor haya indicado DirectQuery.
8. Autentícate usando tu cuenta institucional. No incluyas credenciales dentro de un archivo de consulta.
9. En el Navegador, selecciona:

```text
DATA_ANALYTICS_FOUNDATIONS
  > CURATED
    > VENTAS_TRANSACCIONES_CURADAS_2026_1
```

10. Selecciona **Transformar datos** en lugar de cargar inmediatamente.
11. En Power Query, renombra la consulta a:

```text
Ventas Curadas
```

12. Si no puedes usar el conector Snowflake, utiliza la alternativa autorizada:
    1. Selecciona **Obtener datos > Texto/CSV**.
    2. Abre el CSV curado entregado por el instructor.
    3. Selecciona **Transformar datos**.
    4. Renombra la consulta a `Ventas Curadas`.
    5. Registra en una nota del dashboard que se usó un extracto curado autorizado, incluyendo fecha de extracción si está disponible.

13. Revisa los tipos de datos en Power Query. Corrige únicamente los tipos necesarios:
    - `FECHA_VENTA`: **Fecha**.
    - `TRANSACCION_ID`, `CLIENTE_ID`: **Texto** o número entero según su definición original.
    - `UNIDADES`: **Número entero**.
    - Importes monetarios: **Número decimal fijo**.
    - Región, canal, categoría y producto: **Texto**.
14. Elimina columnas solamente si no son necesarias y si su eliminación no impide el detalle ni las validaciones. Para esta práctica es preferible conservar los campos analíticos principales.
15. Selecciona **Cerrar y aplicar**.

**Resultado esperado**

El modelo de Power BI contiene una tabla llamada `Ventas Curadas`, proveniente de Snowflake o de un CSV curado autorizado, con tipos de datos consistentes.

**Verificación**

En la vista **Datos**, revisa que:

- `FECHA_VENTA` no tenga ícono de texto.
- Los importes no aparezcan como texto.
- Los valores de región, canal y categoría sean legibles.
- La cantidad de filas cargada sea consistente con el control de Snowflake, considerando cualquier filtro documentado de la fuente curada.

---

### Paso 4: Crear la tabla calendario y la relación del modelo

**Objetivo:** Implementar una dimensión de fechas continua para analizar tendencias mensuales y aplicar filtros temporales correctamente.

**Instrucciones**

1. En Power BI Desktop, selecciona la vista **Modelo**.
2. Selecciona **Modelado > Nueva tabla**.
3. Crea la siguiente tabla calendario:

```dax
Calendario =
ADDCOLUMNS(
    CALENDAR(DATE(2024, 1, 1), DATE(2025, 12, 31)),
    "Año", YEAR([Date]),
    "Mes Número", MONTH([Date]),
    "Mes", FORMAT([Date], "MMMM"),
    "Año-Mes", FORMAT([Date], "YYYY-MM"),
    "Año-Mes Orden", YEAR([Date]) * 100 + MONTH([Date]),
    "Trimestre", "T" & FORMAT([Date], "Q"),
    "Año-Trimestre", FORMAT([Date], "YYYY") & "-T" & FORMAT([Date], "Q")
)
```

4. En la vista **Datos**, selecciona la columna `Calendario[Mes]`.
5. En **Herramientas de columna > Ordenar por columna**, selecciona `Mes Número`.
6. Selecciona `Calendario[Año-Mes]`.
7. En **Ordenar por columna**, selecciona `Año-Mes Orden`.
8. En la vista **Modelo**, crea una relación:
   - Tabla origen: `Calendario`.
   - Columna: `Date`.
   - Tabla destino: `Ventas Curadas`.
   - Columna: `FECHA_VENTA`.
   - Cardinalidad: **Uno a varios (1:*)**.
   - Dirección de filtro cruzado: **Única**, desde `Calendario` hacia `Ventas Curadas`.
   - Relación activa: **Sí**.
9. Selecciona la tabla `Calendario` y usa **Herramientas de tabla > Marcar como tabla de fechas**. Elige la columna `Date`.
10. Revisa que no exista más de una relación activa entre las mismas tablas.

> Si `FECHA_VENTA` incluye hora, crea o usa una columna con tipo **Fecha** en Power Query antes de relacionarla. Una fecha/hora con horas distintas puede impedir coincidencias con la tabla calendario.

**Resultado esperado**

Existe una tabla calendario continua de 731 días para 2024 y 2025, relacionada de forma activa con la tabla de ventas.

**Verificación**

Inserta temporalmente un visual de tabla con `Calendario[Año-Mes]` y una medida de ventas. Deben aparecer meses cronológicamente ordenados, sin orden alfabético incorrecto como “abril, agosto, diciembre”.

---

### Paso 5: Crear y formatear las medidas DAX básicas

**Objetivo:** Centralizar los cálculos del dashboard en medidas reutilizables y coherentes con las definiciones acordadas.

**Instrucciones**

1. En la vista **Modelo** o **Datos**, selecciona la tabla `Ventas Curadas`.
2. Selecciona **Modelado > Nueva medida**.
3. Crea las medidas siguientes. Sustituye los nombres de columnas por los nombres reales identificados en el Paso 2.

```dax
Ventas Netas =
SUM('Ventas Curadas'[VENTAS_NETAS])
```

4. Crea la medida de unidades:

```dax
Unidades =
SUM('Ventas Curadas'[UNIDADES])
```

5. Crea la medida de transacciones:

```dax
Transacciones =
DISTINCTCOUNT('Ventas Curadas'[TRANSACCION_ID])
```

6. Crea la medida de clientes únicos:

```dax
Clientes Únicos =
DISTINCTCOUNT('Ventas Curadas'[CLIENTE_ID])
```

7. Si la tabla curada incluye una columna de margen bruto validada, crea:

```dax
Margen Bruto =
SUM('Ventas Curadas'[MARGEN_BRUTO])
```

8. Si no existe `MARGEN_BRUTO`, pero existen ventas netas y costo total, crea:

```dax
Margen Bruto =
SUM('Ventas Curadas'[VENTAS_NETAS])
    - SUM('Ventas Curadas'[COSTO_TOTAL])
```

9. Crea el margen porcentual. Usa `DIVIDE` para evitar errores cuando las ventas netas sean cero:

```dax
Margen % =
DIVIDE([Margen Bruto], [Ventas Netas], 0)
```

10. Crea el ticket promedio. La definición de esta práctica es ventas netas divididas entre transacciones únicas:

```dax
Ticket Promedio =
DIVIDE([Ventas Netas], [Transacciones], 0)
```

11. Crea una medida de cobertura para mostrar el contexto temporal filtrado:

```dax
Cobertura Seleccionada =
VAR FechaInicio =
    MIN('Calendario'[Date])
VAR FechaFin =
    MAX('Calendario'[Date])
RETURN
    FORMAT(FechaInicio, "dd/MM/yyyy")
        & " a "
        & FORMAT(FechaFin, "dd/MM/yyyy")
```

12. Formatea las medidas:
    - `Ventas Netas`, `Margen Bruto`, `Ticket Promedio`: **Moneda** con dos decimales, usando la moneda definida en el dataset o el brief.
    - `Margen %`: **Porcentaje** con uno o dos decimales.
    - `Unidades`, `Transacciones`, `Clientes Únicos`: número entero con separador de miles.
13. Si trabajas con varias monedas, no sumes montos de monedas distintas sin conversión documentada. En ese caso, usa un filtro de moneda o la métrica monetaria normalizada definida en la fuente curada.

**Resultado esperado**

El modelo incluye siete medidas centrales y una medida de cobertura. Los cálculos responden a los filtros de fecha, región, canal y categoría.

**Verificación**

Crea un visual de tabla temporal con las medidas. Aplica un filtro por una región y confirma que todas las cifras cambian de forma coherente. Quita el filtro y confirma que vuelven al total general.

---

### Paso 6: Diseñar la página del dashboard según la pregunta analítica

**Objetivo:** Crear una página de dashboard que priorice las preguntas de negocio y reduzca la carga cognitiva.

**Instrucciones**

1. Renombra la primera página del informe como:

```text
Desempeño de ventas
```

2. Configura el tamaño de página en **Formato > Configuración del lienzo > 16:9**.
3. Inserta un cuadro de texto como encabezado. Usa un título orientado a la decisión y no a la herramienta, por ejemplo:

```text
Desempeño comercial: ventas, margen y mezcla por segmento
```

4. Debajo del título, agrega un subtítulo o cuadro de texto:

```text
Fuente: tabla curada Ventas Retail LATAM 2026.1 | Cobertura oficial: 01/01/2024–31/12/2025 | Datos sintéticos
```

5. Crea cuatro segmentadores:
   - `Calendario[Date]`, preferiblemente tipo **Entre**.
   - `Ventas Curadas[REGION]`.
   - `Ventas Curadas[CANAL]`.
   - `Ventas Curadas[CATEGORIA]`.

6. Ubica los segmentadores en una franja superior o lateral, manteniendo suficiente espacio para los visuales principales.
7. Inserta tarjetas KPI para:
   - `Ventas Netas`.
   - `Margen Bruto`.
   - `Margen %`.
   - `Transacciones`.
   - Opcionalmente, `Ticket Promedio` o `Clientes Únicos` si el diseño conserva legibilidad.
8. Configura las tarjetas:
   - Unidades de visualización: **Millones** o **Miles**, según corresponda.
   - Decimales: consistentes entre visuales.
   - Títulos claros, por ejemplo “Ventas netas”.
   - Evita usar una unidad abreviada diferente en cada tarjeta.
9. Inserta un gráfico de líneas:
   - Eje X: `Calendario[Año-Mes]`.
   - Eje Y: `[Ventas Netas]`.
   - Título: `Evolución mensual de ventas netas — período seleccionado`.
   - Orden: ascendente por `Año-Mes Orden`.
10. Inserta un gráfico de barras horizontales:
    - Eje Y: `REGION` o `CATEGORIA`, según la dimensión priorizada en tu brief.
    - Eje X: `[Ventas Netas]` o `[Margen Bruto]`.
    - Ordena de mayor a menor por la medida.
    - Título: `Ventas netas por región — período seleccionado` o equivalente.
11. Inserta un visual de composición por canal:
    - Preferencia: barras apiladas para comparar aportes con precisión.
    - Categoría o eje: `CANAL`.
    - Valor: `[Ventas Netas]`.
    - Título: `Composición de ventas netas por canal`.
    - Si usas gráfico de dona, limítalo a pocos canales y activa etiquetas legibles; no lo uses si existen muchas categorías.
12. Inserta una tabla de detalle agregada con:
    - `REGION`
    - `CANAL`
    - `CATEGORIA`
    - `[Ventas Netas]`
    - `[Margen Bruto]`
    - `[Margen %]`
    - `[Unidades]`
    - `[Transacciones]`
13. Ordena la tabla por `[Ventas Netas]` de mayor a menor.
14. Agrega una tarjeta pequeña o cuadro de texto que muestre el período aplicado. Puedes usar la medida `[Cobertura Seleccionada]`.
15. Mantén una paleta de colores consistente:
    - Un color principal para ventas.
    - Un color secundario para margen.
    - Colores de canal estables entre visuales.
    - Contraste suficiente entre texto y fondo.
16. Evita fondos decorativos, gráficos 3D, exceso de bordes, más de una métrica no comparable en el mismo eje y títulos genéricos como “Gráfico 1”.

**Resultado esperado**

La página contiene segmentadores, KPI, tendencia temporal, comparación por categoría o región, composición por canal y tabla de detalle. Cada visual responde una pregunta definida en el Paso 1.

**Verificación**

Pide a una persona compañera que observe el dashboard durante 30 segundos y responda:

1. ¿Cuál es el período analizado?
2. ¿Cuál es la métrica principal?
3. ¿Qué visual muestra la tendencia?
4. ¿Qué dimensión permite identificar concentración comercial?
5. ¿Dónde puede filtrar región, canal y categoría?

Si no puede responder, mejora títulos, ubicación, formato o notas de contexto.

---

### Paso 7: Configurar interacciones, filtros y narrativa de hallazgos

**Objetivo:** Garantizar que los visuales se filtren de manera intencional y que el dashboard comunique evidencia útil para la decisión.

**Instrucciones**

1. Selecciona el gráfico de barras por región o categoría.
2. En la cinta de opciones, selecciona **Formato > Editar interacciones**.
3. Para cada visual, decide si debe:
   - Filtrarse.
   - Resaltarse.
   - No cambiar.
4. Configura como regla general:
   - Segmentadores: deben filtrar todos los KPI y visuales analíticos.
   - Gráfico de barras: puede filtrar la línea, composición y tabla.
   - Gráfico de composición por canal: puede filtrar tabla y KPI.
   - Tarjetas KPI: deben responder a todos los segmentadores.
5. Evita que un visual de composición filtre accidentalmente la serie temporal si ello dificulta interpretar la tendencia. Si es necesario, cambia la interacción a **Ninguno**.
6. Inserta una sección de texto denominada “Hallazgos priorizados” en la parte inferior o en un panel lateral.
7. Redacta tres hallazgos después de revisar los resultados validados. Cada hallazgo debe incluir:
   - Métrica y dimensión.
   - Período.
   - Comparación relevante.
   - Nivel de certeza o limitación cuando corresponda.
   - Implicación operativa o pregunta de seguimiento.

8. Usa esta estructura de redacción:

```text
Hallazgo 1: Durante [período], [segmento] registró [métrica] de [valor], equivalente a [comparación]. Esto sugiere [implicación], sujeto a la cobertura y reglas de calidad documentadas.

Hallazgo 2: La evolución mensual de [métrica] muestra [patrón observado] entre [fecha inicial] y [fecha final]. El patrón requiere [acción o análisis adicional] antes de atribuir causalidad.

Hallazgo 3: [Canal/categoría/región] aportó [proporción o valor] de [métrica] en [período]. Se recomienda [acción concreta] o validar [variable adicional].
```

9. No afirmes causalidad sin evidencia. Por ejemplo, evita “las ventas cayeron por la campaña” si el dataset no contiene información que permita demostrar esa relación.
10. Agrega una nota metodológica visible:

```text
Nota: resultados calculados sobre datos sintéticos curados. Los valores responden a los filtros activos y no deben interpretarse como desempeño real de una empresa.
```

11. Guarda el archivo PBIX.

**Resultado esperado**

Las interacciones son coherentes y los tres hallazgos del dashboard están respaldados por visuales y métricas validadas.

**Verificación**

Prueba los siguientes escenarios:

- Selecciona una región y confirma que KPI, línea, composición y tabla responden como se espera.
- Selecciona un canal y confirma que no se producen filtros contradictorios.
- Borra todas las selecciones y verifica que el dashboard vuelve al total general.
- Lee cada hallazgo y localiza el visual que aporta la evidencia correspondiente.

---

### Paso 8: Validar métricas contra Snowflake y Excel

**Objetivo:** Confirmar que las cifras del dashboard coinciden con la fuente curada y los controles de Excel antes de finalizar el entregable.

**Instrucciones**

1. En Snowflake, ejecuta la siguiente consulta global. Ajusta los nombres reales de columnas:

```sql
SELECT
    COUNT(*) AS filas,
    COUNT(DISTINCT TRANSACCION_ID) AS transacciones,
    COUNT(DISTINCT CLIENTE_ID) AS clientes_unicos,
    SUM(UNIDADES) AS unidades,
    SUM(VENTAS_NETAS) AS ventas_netas,
    SUM(MARGEN_BRUTO) AS margen_bruto,
    DIV0(SUM(MARGEN_BRUTO), SUM(VENTAS_NETAS)) AS margen_pct,
    DIV0(SUM(VENTAS_NETAS), COUNT(DISTINCT TRANSACCION_ID)) AS ticket_promedio,
    MIN(FECHA_VENTA) AS fecha_minima,
    MAX(FECHA_VENTA) AS fecha_maxima
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1;
```

2. Si `MARGEN_BRUTO` no existe, reemplázalo por el cálculo validado:

```sql
SUM(VENTAS_NETAS) - SUM(COSTO_TOTAL) AS margen_bruto
```

3. Ejecuta una validación mensual para comprobar la tendencia del gráfico de línea:

```sql
SELECT
    DATE_TRUNC('MONTH', FECHA_VENTA) AS mes,
    SUM(VENTAS_NETAS) AS ventas_netas,
    SUM(MARGEN_BRUTO) AS margen_bruto,
    COUNT(DISTINCT TRANSACCION_ID) AS transacciones
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
GROUP BY 1
ORDER BY 1;
```

4. Ejecuta una validación por región o categoría, según el gráfico que hayas construido:

```sql
SELECT
    REGION,
    SUM(VENTAS_NETAS) AS ventas_netas,
    SUM(MARGEN_BRUTO) AS margen_bruto,
    COUNT(DISTINCT TRANSACCION_ID) AS transacciones
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
GROUP BY REGION
ORDER BY ventas_netas DESC;
```

5. Ejecuta una validación por canal:

```sql
SELECT
    CANAL,
    SUM(VENTAS_NETAS) AS ventas_netas,
    DIV0(
        SUM(VENTAS_NETAS),
        SUM(SUM(VENTAS_NETAS)) OVER ()
    ) AS participacion_ventas
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1
GROUP BY CANAL
ORDER BY ventas_netas DESC;
```

6. En Power BI, elimina todos los filtros. Copia o registra los valores de:
   - Ventas netas.
   - Margen bruto.
   - Margen porcentual.
   - Unidades.
   - Transacciones.
   - Clientes únicos.
   - Ticket promedio.
7. En Excel, abre el archivo de controles de la práctica anterior o crea una hoja de control a partir del extracto curado autorizado. Confirma, como mínimo:
   - Total de ventas netas.
   - Total de unidades.
   - Conteo de transacciones.
   - Cobertura mínima y máxima.
   - Un agregado por región, categoría o canal.
8. Documenta los resultados en `Lab05_validacion_APELLIDO_NOMBRE.xlsx` usando la siguiente tabla:

| Control | Snowflake | Power BI | Excel | Diferencia absoluta | Estado |
|---|---:|---:|---:|---:|---|
| Ventas netas totales |  |  |  |  |  |
| Margen bruto total |  |  |  |  |  |
| Margen porcentual |  |  |  |  |  |
| Unidades totales |  |  |  |  |  |
| Transacciones únicas |  |  |  |  |  |
| Clientes únicos |  |  |  |  |  |
| Ticket promedio |  |  |  |  |  |
| Ventas de una región/categoría |  |  |  |  |  |
| Ventas de un mes |  |  |  |  |  |

9. Para valores monetarios, utiliza una tolerancia máxima de redondeo de `0.01` cuando las tres herramientas trabajen con el mismo nivel de precisión. Para conteos enteros, la diferencia debe ser `0`.
10. Si existe discrepancia:
    - Confirma que los filtros de Power BI estén eliminados.
    - Confirma que Power BI utiliza la tabla curada correcta.
    - Revisa si Excel usa la misma versión del extracto.
    - Revisa los tipos de datos.
    - Revisa si la medida usa `COUNTROWS` cuando debería usar `DISTINCTCOUNT`.
    - Revisa si la tabla calendario dejó fechas sin relación.
11. Corrige la causa, actualiza el modelo y repite la validación.
12. Guarda las consultas SQL y la evidencia Excel.

**Resultado esperado**

Al menos cinco KPI o agregados están validados contra Snowflake y Excel. Las diferencias no explicadas han sido resueltas antes de la entrega.

**Verificación**

La tabla de evidencia debe contener al menos cinco filas con estado **Conforme**. Los KPI principales mostrados en tarjetas deben coincidir con Snowflake y Excel para el mismo conjunto de filtros.

## Validación y Pruebas

Completa esta lista antes de entregar:

| Criterio | Prueba | Resultado esperado |
|---|---|---|
| Fuente trazable | Revisar origen de `Ventas Curadas` en Power Query | Snowflake CURATED o CSV curado autorizado |
| Protección de RAW | Revisar script SQL | No contiene sentencias de modificación sobre `RAW` |
| Cobertura temporal | Comparar `MIN` y `MAX` de fecha | Intervalo validado y documentado |
| Modelo | Revisar relación Calendario–Ventas Curadas | Relación activa 1:* y filtro en una dirección |
| Orden temporal | Revisar eje del gráfico de línea | Meses en orden cronológico |
| KPI | Revisar medidas DAX | Fórmulas coherentes con la especificación |
| Formato | Revisar moneda, porcentaje y unidades | Formato consistente y comprensible |
| Segmentadores | Aplicar región, canal, categoría y fechas | Visuales responden según las interacciones definidas |
| Comparación categórica | Revisar gráfico de barras | Orden descendente y título con métrica/período |
| Composición | Revisar canal | Categorías legibles y proporción interpretable |
| Narrativa | Leer hallazgos | Tres hallazgos sustentados, sin causalidad no demostrada |
| Validación cruzada | Comparar Snowflake, Power BI y Excel | Cinco o más controles conformes |
| Nota de cobertura | Revisar encabezado o pie | Fuente, período y naturaleza sintética visibles |
| Entregables | Revisar directorio `05_dashboard` | PBIX, evidencia Excel, SQL y captura presentes |

Como prueba final, realiza esta secuencia:

1. Quita todos los filtros y toma nota de los KPI generales.
2. Filtra un mes específico y una región.
3. Ejecuta en Snowflake una consulta equivalente con `WHERE`.
4. Compara ventas netas, transacciones y margen.
5. Vuelve a quitar filtros.
6. Confirma que el dashboard retorna a los valores globales validados.
7. Exporta una captura de la página completa y guárdala con la convención establecida.

## Solución de Problemas

### Problema 1: Los KPI de Power BI no coinciden con Snowflake

**Síntomas**

- Las ventas netas o transacciones del dashboard son diferentes a las de Snowflake.
- La diferencia persiste incluso cuando aparentemente no hay filtros.
- La línea mensual muestra valores distintos a la consulta agrupada por mes.

**Causa probable**

Existe un filtro activo no visible, el campo de fecha no está relacionado correctamente con `Calendario`, se usó `COUNTROWS` en lugar de `DISTINCTCOUNT`, o Power BI está conectado a una fuente distinta de la tabla curada validada.

**Corrección**

1. Abre el panel **Filtros** y elimina filtros de visual, página e informe.
2. Revisa la relación activa entre `Calendario[Date]` y `Ventas Curadas[FECHA_VENTA]`.
3. Confirma que `FECHA_VENTA` es de tipo fecha y no texto o fecha/hora incompatible.
4. Revisa que la medida de transacciones use:

```dax
DISTINCTCOUNT('Ventas Curadas'[TRANSACCION_ID])
```

5. En Power Query, revisa el paso **Origen** para confirmar la tabla o el CSV utilizado.
6. Actualiza los datos y repite las consultas de validación con exactamente el mismo período y filtros.

### Problema 2: La conexión Snowflake desde Power BI falla o la carga es demasiado lenta

**Síntomas**

- Power BI muestra un error de autenticación, servidor no encontrado o falta de permisos.
- El Navegador no muestra la base de datos o la tabla CURATED.
- La carga de aproximadamente 750,000 transacciones tarda demasiado o Power BI deja de responder.

**Causa probable**

La cuenta, warehouse, rol o método de autenticación no son correctos; el warehouse puede estar suspendido; o el equipo tiene memoria disponible insuficiente para cargar y transformar el dataset.

**Corrección**

1. Confirma con el instructor el valor correcto de `<ORG_ACCOUNT>`, el rol asignado y el warehouse.
2. Prueba acceso en Snowsight con el mismo usuario y ejecuta:

```sql
USE ROLE DAF_ANALYST_ROLE;
USE WAREHOUSE DAF_LAB_WH;
SELECT COUNT(*)
FROM DATA_ANALYTICS_FOUNDATIONS.CURATED.VENTAS_TRANSACCIONES_CURADAS_2026_1;
```

3. Si Snowsight funciona pero Power BI no, borra permisos guardados en **Archivo > Opciones y configuración > Configuración de origen de datos**, vuelve a autenticarte y no selecciones credenciales incorrectas en caché.
4. Cierra aplicaciones que consuman memoria, especialmente Excel con libros grandes, navegadores con muchas pestañas y otras instancias de Power BI.
5. Usa el CSV curado autorizado únicamente si el instructor confirma que la conexión directa no está disponible.
6. No reduzcas arbitrariamente el dataset para “hacerlo cargar” si ello impide validar los KPI contra la fuente completa.

## Limpieza

1. Guarda el archivo PBIX final:

```text
C:\DAF\Batch_01\05_dashboard\Lab05_Dashboard_Ventas_APELLIDO_NOMBRE.pbix
```

2. Guarda el archivo de evidencia:

```text
C:\DAF\Batch_01\05_dashboard\Lab05_validacion_APELLIDO_NOMBRE.xlsx
```

3. Guarda el script SQL de validación:

```text
C:\DAF\Batch_01\sql\Lab05_validacion_dashboard_APELLIDO_NOMBRE.sql
```

4. Exporta o captura la página del dashboard y guárdala como:

```text
C:\DAF\Batch_01\05_dashboard\Lab05_dashboard_APELLIDO_NOMBRE.png
```

5. Cierra Power BI Desktop después de verificar que el archivo fue guardado.
6. Cierra sesiones de Snowflake o terminales que ya no necesites.
7. No elimines, alteres ni modifiques objetos del esquema `RAW`.
8. No elimines la tabla curada común. Si creaste un objeto personal autorizado en prácticas anteriores, consérvalo o elimínalo únicamente si el instructor lo solicita expresamente.

## Resumen

En esta práctica conectaste Power BI Desktop a una fuente curada y trazable del dataset sintético masivo Ventas Retail LATAM 2026.1. Creaste un modelo con tabla calendario, medidas DAX básicas y visualizaciones diseñadas para responder preguntas específicas sobre desempeño, tendencia, concentración y composición de ventas.

El entregable final debe demostrar que cada visual tiene un propósito analítico, que los filtros y títulos dan contexto suficiente y que los KPI han sido validados contra Snowflake y Excel. Un dashboard correcto no afirma más de lo que los datos permiten: comunica resultados del período y segmentos seleccionados, explicita la cobertura y distingue entre un patrón observado, una hipótesis y una conclusión sustentada.
