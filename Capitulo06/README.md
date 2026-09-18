# Interpretación ejecutiva de un dashboard

## Metadatos

| Duration | Complexity | Bloom level |
|---|---|---|
| 76 minutos | Media | Aplicar |

## Descripción General

En esta práctica explorarás críticamente el dashboard preconstruido `M6_Dashboard_Interpretacion.pbix` y registrarás evidencia en el libro `M6_Resumen_Ejecutivo.xlsx`. No construirás visualizaciones, medidas ni consultas SQL: el foco es interpretar correctamente los indicadores disponibles, evaluar la validez de las comparaciones y evitar conclusiones causales no sustentadas.

El resultado será el archivo `M6_Brief_Interpretacion_[Apellido].xlsx`, que documentará datos literales, observaciones comparables, hallazgos potenciales, limitaciones y tres mensajes ejecutivos. Este brief será una entrada obligatoria para la Práctica 7, donde los hallazgos se validarán y ampliarán con SQL, Excel y Power BI.

> **Contexto del dataset:** RetailNova v1.0 es un dataset artificial masivo con exactamente 750,000 transacciones de ventas entre `2023-01-01` y `2025-12-31`. Incluye dimensiones de fecha, producto, categoría, canal, región, cliente, vendedor, segmento y campaña. No contiene datos personales reales.

## Objetivos de Aprendizaje

- [ ] Distinguir entre dato, observación, hallazgo, insight y recomendación dentro de un dashboard comercial.
- [ ] Interpretar ventas, margen, pedidos, ticket promedio, crecimiento interanual y cumplimiento de meta con contexto temporal y segmentación.
- [ ] Evaluar si una comparación es válida considerando filtros, periodos, granularidad, escalas, objetivos y denominadores.
- [ ] Clasificar conclusiones como sustentadas, parcialmente sustentadas o no sustentadas según la evidencia disponible.
- [ ] Redactar mensajes ejecutivos breves con evidencia, implicación de negocio, acción recomendada y nivel de certeza explícito.

## Prerrequisitos

### Conocimientos requeridos

Antes de iniciar, debes poder:

- Interpretar filas, columnas, tablas, campos, registros y agregaciones.
- Leer tarjetas KPI, gráficos de barras, gráficos de líneas, tablas y segmentadores.
- Aplicar y quitar filtros básicos en Power BI Desktop.
- Usar libros de Excel, tablas y celdas de texto.
- Diferenciar los siguientes niveles analíticos:
  - **Dato:** valor registrado o literal visible.
  - **Observación:** comparación o patrón visible en los datos.
  - **Hallazgo:** conclusión cuantificada y contextualizada, respaldada por evidencia.
  - **Insight:** interpretación relevante para una decisión, vinculada con una acción o investigación.
  - **Recomendación:** acción propuesta, proporcionada al nivel de certeza disponible.

### Acceso y archivos requeridos

Debes contar con:

- Power BI Desktop `2.146.1454.0` o versión institucional compatible.
- Microsoft Excel para Microsoft 365, versión `2508 Build 16.0.19127.20192` o equivalente.
- Permiso de escritura en `C:\DAF\Batch_01\`.
- El paquete de archivos de la práctica, que debe incluir:
  - `M6_Dashboard_Interpretacion.pbix`
  - `M6_Resumen_Ejecutivo.xlsx`
  - `Diccionario_Metricas_RetailNova_v1.0.pdf` o archivo equivalente.
- Acceso local al dashboard preconstruido. No se requiere conexión a Snowflake durante esta práctica.

> **Importante:** no sustituyas el dashboard por otra versión, otro PBIX ni un archivo creado por ti. Las conclusiones de esta práctica deben ser trazables a la versión 1.0 de RetailNova entregada por el curso.

## Entorno de Laboratorio

### Recursos de hardware recomendados

| Recurso | Mínimo | Recomendado |
|---|---:|---:|
| Memoria RAM | 16 GB | 32 GB |
| Espacio libre en SSD | 20 GB | 25 GB |
| Resolución de pantalla | 1920 × 1080 | 1920 × 1080 o superior |
| Procesador | Intel Core i5 / AMD Ryzen 5, 4 núcleos | Intel Core i5 10.ª generación / Ryzen 5 5000 o superior |
| Internet | 20 Mbps estable | 20 Mbps o superior |

### Software y archivos

| Componente | Uso en la práctica |
|---|---|
| Windows 11 Pro | Sistema operativo de referencia |
| Power BI Desktop | Exploración crítica del dashboard preconstruido |
| Microsoft Excel | Elaboración del brief ejecutivo |
| RetailNova v1.0 | Fuente canónica artificial de 750,000 transacciones |
| Diccionario de métricas | Validación de definiciones, denominadores y reglas de cálculo |

### Preparación de carpetas

Abre **Windows PowerShell** y ejecuta el siguiente comando para confirmar o crear la estructura obligatoria del curso:

```powershell
$base = "C:\DAF\Batch_01"
"00_source","01_brief","02_quality","03_descriptive","04_exploration","05_dashboard","sql" |
ForEach-Object { New-Item -ItemType Directory -Force -Path (Join-Path $base $_) }
```

Copia los archivos proporcionados a las siguientes ubicaciones:

| Archivo | Ubicación requerida |
|---|---|
| `M6_Dashboard_Interpretacion.pbix` | `C:\DAF\Batch_01\05_dashboard\` |
| `M6_Resumen_Ejecutivo.xlsx` | `C:\DAF\Batch_01\01_brief\` |
| Diccionario RetailNova v1.0 | `C:\DAF\Batch_01\00_source\` |

Comprueba la presencia de los archivos:

```powershell
Get-ChildItem `
  "C:\DAF\Batch_01\05_dashboard\M6_Dashboard_Interpretacion.pbix", `
  "C:\DAF\Batch_01\01_brief\M6_Resumen_Ejecutivo.xlsx", `
  "C:\DAF\Batch_01\00_source\" |
Select-Object FullName, Length, LastWriteTime
```

> **Regla de seguridad:** no escribas contraseñas institucionales, credenciales de Snowflake ni información personal en el archivo PBIX, el libro Excel o las celdas del brief.

## Instrucciones Paso a Paso

### Paso 1: Preparar el espacio de trabajo y abrir los archivos

**Objetivo:** verificar que se utilizan los archivos oficiales y preparar una copia de trabajo del brief.

**Instrucciones:**

1. Abre el Explorador de archivos y navega a:
   ```text
   C:\DAF\Batch_01\01_brief\
   ```
2. Crea una copia de `M6_Resumen_Ejecutivo.xlsx`.
3. Renombra la copia con el formato:
   ```text
   M6_Brief_Interpretacion_[Apellido].xlsx
   ```
   Por ejemplo:
   ```text
   M6_Brief_Interpretacion_Garcia.xlsx
   ```
4. Abre el archivo renombrado en Excel.
5. Revisa las hojas disponibles. Si la plantilla ya contiene secciones, conserva su estructura. Si una sección no existe, créala siguiendo las instrucciones de los pasos posteriores.
6. Abre Power BI Desktop.
7. Selecciona **Archivo > Abrir informe** y abre:
   ```text
   C:\DAF\Batch_01\05_dashboard\M6_Dashboard_Interpretacion.pbix
   ```
8. Espera a que el informe cargue completamente. No pulses **Actualizar** salvo que el instructor lo indique.
9. En Excel, completa el encabezado del brief con tu nombre, apellido, fecha de realización y nombre del archivo PBIX utilizado.

**Expected output:**

- Power BI Desktop muestra el dashboard RetailNova preconstruido.
- Existe un archivo de trabajo individual en `C:\DAF\Batch_01\01_brief\`.
- El brief identifica al estudiante y al archivo fuente utilizado.

**Verification:**

- Confirma que el nombre del archivo Excel sigue exactamente el patrón requerido.
- Confirma que la extensión del dashboard es `.pbix`.
- Confirma que no has modificado el archivo original `M6_Resumen_Ejecutivo.xlsx`.

---

### Paso 2: Confirmar el contexto analítico y las definiciones de KPI

**Objetivo:** establecer el alcance de cada visualización antes de interpretar cualquier valor.

**Instrucciones:**

1. En Power BI, identifica la página principal del dashboard y revisa los segmentadores, filtros visibles y filtros aplicados a nivel de página o informe.
2. Registra en la sección **Contexto del dashboard** del brief:
   - Fecha de actualización visible en el dashboard, si está disponible.
   - Periodo seleccionado.
   - Moneda reportada.
   - Región, canal, categoría, segmento, campaña u otros filtros activos.
   - Si la página muestra todos los datos o una selección.
3. Si existe un panel de filtros, expándelo y revisa:
   - **Filtros en este objeto visual**.
   - **Filtros en esta página**.
   - **Filtros en todas las páginas**.
4. Anota explícitamente los filtros que puedan afectar una comparación. Por ejemplo:
   - `Canal = Digital`
   - `Región = Norte`
   - `Año = 2025`
   - `Segmento = Corporativo`
5. Abre el diccionario de métricas RetailNova v1.0 y localiza las definiciones de:
   - Ventas.
   - Margen.
   - Pedidos.
   - Ticket promedio.
   - Crecimiento interanual.
   - Cumplimiento de meta.
6. Para cada KPI, registra en el brief:
   - Nombre literal del indicador.
   - Definición según el diccionario.
   - Unidad: moneda, porcentaje, número de pedidos u otra.
   - Denominador, cuando aplique.
   - Periodo de comparación, cuando aplique.
7. No asumas que “ventas” equivale a ingresos netos, facturación, ventas con impuestos o ventas sin devoluciones. Usa únicamente la definición oficial del diccionario.
8. Si la meta no está definida en el dashboard ni en el diccionario, registra:
   ```text
   Meta: no visible o no definida en la evidencia disponible.
   ```

#### Modelo de registro para el brief

| Elemento | Registro esperado |
|---|---|
| Periodo analizado | Ejemplo: 01/01/2025 a 31/12/2025; confirmar en dashboard |
| Comparación temporal | Ejemplo: mismo periodo del año anterior; confirmar definición |
| Moneda | Ejemplo: USD, MXN o moneda indicada por el dashboard |
| Filtros globales | Registrar todos los filtros activos |
| KPI: Ticket promedio | Definición oficial, unidad y denominador |
| KPI: Cumplimiento de meta | Fórmula o definición oficial disponible |

**Expected output:**

- El brief contiene una ficha de contexto completa.
- Cada KPI clave tiene una definición documentada o una limitación explícita si la definición no está disponible.

**Verification:**

- Comprueba que puedes responder: “¿Qué periodo mide este indicador?”, “¿en qué moneda?”, “¿qué filtros lo afectan?” y “¿cómo se calcula?”.
- Si no puedes responder una de estas preguntas, no continúes con interpretaciones; registra la limitación y consulta al instructor o al diccionario.

---

### Paso 3: Registrar datos literales visibles sin interpretarlos

**Objetivo:** separar los valores observables del dashboard de cualquier explicación o recomendación.

**Instrucciones:**

1. En el brief, localiza o crea una hoja llamada **Datos literales**.
2. Recorre las tarjetas KPI, gráficos de líneas, barras, tablas y matrices del dashboard.
3. Registra al menos:
   - Seis valores de tarjetas KPI.
   - Cuatro valores temporales de una serie o gráfico de tendencia.
   - Cuatro valores de un desglose por segmento, canal, región o categoría.
   - La meta visible, si existe.
4. Para cada registro, anota:
   - Visual de origen.
   - Métrica.
   - Valor literal.
   - Unidad.
   - Periodo.
   - Filtros activos.
   - Página del dashboard.
5. Conserva el formato visible. Por ejemplo, si el dashboard presenta `12.4 %`, no lo conviertas a `0.124` sin indicarlo.
6. Usa expresiones descriptivas, no explicativas:
   - Correcto: “La tarjeta Ventas muestra 4.2 M”.
   - Incorrecto: “Las ventas son bajas por problemas en el canal”.
7. Si un gráfico permite mostrar información al pasar el cursor, utiliza el tooltip para registrar el valor exacto de fechas, categorías o segmentos relevantes.
8. No copies valores redondeados como si fueran exactos si el tooltip muestra más precisión. Registra la precisión disponible y especifica la fuente.

#### Tabla sugerida: Datos literales

| ID | Visual de origen | Métrica/campo | Valor literal | Unidad | Periodo | Filtros activos |
|---|---|---|---:|---|---|---|
| D-01 | Tarjeta “Ventas” | Ventas | Registrar valor | Moneda | Registrar periodo | Registrar filtros |
| D-02 | Tarjeta “Margen” | Margen | Registrar valor | Moneda o % | Registrar periodo | Registrar filtros |
| D-03 | Línea temporal | Ventas mensuales | Registrar valor | Moneda | Mes específico | Registrar filtros |
| D-04 | Barras por canal | Pedidos | Registrar valor | Número | Periodo visible | Registrar filtros |

**Expected output:**

- Una tabla con evidencia literal trazable al dashboard.
- Los datos registrados no contienen interpretaciones causales ni recomendaciones.

**Verification:**

- Selecciona tres filas de la tabla y vuelve al dashboard.
- Verifica que otra persona podría encontrar el mismo valor usando el visual, periodo y filtros registrados.
- Comprueba que cada dato tiene unidad y contexto temporal.

---

### Paso 4: Convertir datos en observaciones comparables

**Objetivo:** formular observaciones descriptivas mediante comparaciones válidas y cuantificadas.

**Instrucciones:**

1. En el brief, crea o completa la sección **Observaciones comparables**.
2. Elige comparaciones que estén respaldadas por los visuales disponibles. Puedes utilizar:
   - Periodo actual frente a periodo anterior.
   - Mismo periodo del año anterior.
   - Resultado real frente a meta.
   - Canal frente a canal.
   - Región frente a región.
   - Categoría frente a categoría.
3. Antes de redactar cada observación, confirma los cinco criterios de comparabilidad:
   1. **Periodo:** ¿los periodos tienen duración equivalente?
   2. **Escala:** ¿ambos valores usan la misma unidad y moneda?
   3. **Filtros:** ¿los segmentos comparados tienen el mismo contexto o el filtro está documentado?
   4. **Granularidad:** ¿se comparan métricas agregadas al mismo nivel?
   5. **Denominador:** ¿los porcentajes se calculan sobre poblaciones comparables?
4. Calcula variaciones solamente cuando la definición lo permita. Para una variación porcentual:
   ```text
   Variación % = ((Valor actual - Valor de referencia) / Valor de referencia) × 100
   ```
5. Distingue entre porcentaje y puntos porcentuales:
   - Si una tasa cambia de `20 %` a `25 %`, el cambio es `+5 puntos porcentuales`.
   - La variación relativa es `+25 %`.
6. Redacta al menos cinco observaciones. Cada una debe incluir:
   - Métrica.
   - Periodo o segmentos comparados.
   - Magnitud de la diferencia.
   - Filtro o alcance relevante.
7. Evita usar verbos causales como “provocó”, “generó” o “demuestra”.
8. Cuando una comparación no sea válida, regístrala como una comparación descartada e indica por qué.

#### Ejemplos de redacción responsable

| Nivel | Redacción adecuada |
|---|---|
| Observación temporal | “Las ventas del periodo seleccionado fueron inferiores a las del mismo periodo del año anterior en el contexto de filtros vigente.” |
| Observación por segmento | “El canal con mayor volumen de pedidos en el periodo visible fue [canal], según el gráfico segmentado.” |
| Observación frente a meta | “El cumplimiento de meta mostrado fue de [valor], equivalente a una brecha de [valor] frente al objetivo, si la definición de meta aplica al mismo periodo.” |
| Comparación no válida | “No se compara el margen mensual con el margen anual, porque los periodos y la granularidad no son equivalentes.” |

#### Tabla sugerida: Observaciones comparables

| ID | Observación | Comparación | Evidencia usada | ¿Comparable? | Motivo o condición |
|---|---|---|---|---|---|
| O-01 | Redactar observación | Actual vs. año anterior | D-01, D-03 | Sí/No | Justificar |
| O-02 | Redactar observación | Canal A vs. Canal B | D-04 y visual | Sí/No | Justificar |

**Expected output:**

- Al menos cinco observaciones cuantificadas y contextualizadas.
- Al menos una comparación descartada o condicionada cuando el dashboard no proporcione contexto suficiente.

**Verification:**

- Revisa cada observación y confirma que contiene una referencia explícita: periodo, meta, segmento o línea base.
- Elimina o corrige cualquier frase que interprete una causa sin evidencia adicional.

---

### Paso 5: Contrastar contra meta, tiempo y segmentos relevantes

**Objetivo:** priorizar observaciones usando referencias de negocio y segmentación adecuada.

**Instrucciones:**

1. Identifica cuáles de tus observaciones pueden contrastarse con una meta visible.
2. Para cada KPI con meta, registra:
   - Resultado real.
   - Meta.
   - Brecha absoluta.
   - Brecha porcentual, si el denominador está definido.
   - Periodo de vigencia de la meta.
3. Comprueba que no estás comparando una meta anual con un resultado mensual, trimestral o parcial sin un ajuste explícito.
4. Identifica la comparación temporal que el dashboard ofrece:
   - Mes anterior.
   - Trimestre anterior.
   - Acumulado anual anterior.
   - Mismo periodo del año anterior.
5. No reemplaces una comparación interanual por una secuencial sin indicarlo. Ambas responden preguntas distintas:
   - Secuencial: cambio respecto al periodo inmediatamente anterior.
   - Interanual: cambio respecto al mismo periodo de otro año.
6. Revisa al menos dos segmentaciones relevantes para una variación importante. Por ejemplo:
   - Ventas por canal.
   - Margen por categoría.
   - Pedidos por región.
   - Ticket promedio por segmento.
7. Determina si la variación total se concentra en uno o varios segmentos. Si el dashboard no permite determinarlo, registra esa restricción.
8. Formula dos o tres **hallazgos potenciales**, solo cuando la evidencia incluya magnitud, contexto, segmentación y una comparación válida.
9. Redacta los hallazgos usando lenguaje calibrado:
   - “Los datos sugieren…”
   - “La variación se concentra visualmente en…”
   - “Se observa una asociación entre…”
   - “Requiere validación adicional con datos transaccionales o fuentes complementarias.”

#### Matriz de contraste sugerida

| KPI | Resultado | Referencia | Tipo de referencia | Brecha | Segmento revisado | Interpretación permitida |
|---|---:|---:|---|---:|---|---|
| Ventas | Registrar | Registrar | Meta / interanual / secuencial | Registrar | Canal o región | Observación o hallazgo potencial |
| Margen | Registrar | Registrar | Meta / interanual | Registrar | Categoría | Observación o hallazgo potencial |
| Pedidos | Registrar | Registrar | Interanual / secuencial | Registrar | Segmento | Observación |
| Ticket promedio | Registrar | Registrar | Comparación válida | Registrar | Canal | Observación |

**Expected output:**

- Una matriz que evidencia el contraste de KPI contra una referencia apropiada.
- Dos o tres hallazgos potenciales redactados con contexto y nivel de certeza adecuado.

**Verification:**

- Cada hallazgo potencial debe responder:
  - ¿Qué ocurrió?
  - ¿Dónde o en qué segmento ocurrió?
  - ¿Contra qué referencia se compara?
  - ¿Cuál es la magnitud?
  - ¿Qué evidencia falta para afirmar causalidad?
- Si falta alguna respuesta, reclasifica el texto como observación, no como hallazgo.

---

### Paso 6: Clasificar afirmaciones y documentar limitaciones

**Objetivo:** distinguir afirmaciones sustentadas de interpretaciones que requieren evidencia adicional.

**Instrucciones:**

1. En el brief, crea o completa la hoja **Evaluación de afirmaciones**.
2. Clasifica cada afirmación como:
   - **Sustentada:** el dashboard muestra evidencia directa, contexto y comparación válida.
   - **Parcialmente sustentada:** existe evidencia de una parte de la afirmación, pero faltan controles, definición, segmentación o información causal.
   - **No sustentada:** la afirmación no puede verificarse con el dashboard o atribuye causalidad sin evidencia.
3. Evalúa como mínimo las siguientes afirmaciones, adaptándolas a los valores reales observados:
   1. “El KPI de ventas del periodo está por debajo de la meta.”
   2. “La disminución de ventas fue causada por el canal digital.”
   3. “El canal con más pedidos es necesariamente el más rentable.”
   4. “El crecimiento interanual observado representa una mejora operativa sostenida.”
   5. “La categoría con menor margen debe eliminarse del catálogo.”
4. Para cada afirmación, indica:
   - Evidencia disponible en el dashboard.
   - Evidencia ausente.
   - Clasificación.
   - Redacción corregida o prudente.
5. Documenta al menos cuatro limitaciones del dashboard. Considera las siguientes posibilidades:
   - No incluye costos logísticos.
   - No incluye devoluciones registradas después del periodo.
   - No incluye inventario, quiebres de stock ni disponibilidad de producto.
   - No incluye precios de competidores.
   - No incluye datos de tráfico digital, conversiones web o inversión publicitaria.
   - No incluye campañas experimentales ni grupos de control.
   - No muestra la composición detallada de la meta.
   - No permite evaluar causalidad.
6. Explica cómo cada limitación afecta la interpretación. Por ejemplo:
   ```text
   La ausencia de costos logísticos impide concluir que un canal con mayor margen comercial tenga también mayor rentabilidad neta.
   ```
7. No uses “no sustentada” como sinónimo de “falsa”. Una afirmación no sustentada puede ser verdadera, pero no está demostrada por la evidencia disponible.

#### Tabla sugerida: Evaluación de afirmaciones

| ID | Afirmación propuesta | Evidencia disponible | Evidencia faltante | Clasificación | Redacción responsable |
|---|---|---|---|---|---|
| A-01 | El KPI está bajo meta | Tarjeta KPI y meta del mismo periodo | Ninguna o definir condición | Sustentada/Parcial | Redactar |
| A-02 | El canal digital causó la caída | Variación por canal | Diseño causal, campañas, inventario, precios | No sustentada | “La caída se concentra en…” |
| A-03 | Más pedidos implica más rentabilidad | Pedidos por canal | Costos, margen neto, devoluciones | No sustentada | “Conviene contrastar…” |

**Expected output:**

- Una evaluación de al menos cinco afirmaciones.
- Un inventario de cuatro o más limitaciones interpretativas.
- Redacciones alternativas que expresan incertidumbre de manera profesional.

**Verification:**

- Confirma que ninguna afirmación causal se clasifica como sustentada si el dashboard solo muestra asociación temporal o segmentación.
- Confirma que cada limitación explica su efecto sobre una decisión potencial.

---

### Paso 7: Redactar tres mensajes ejecutivos

**Objetivo:** comunicar resultados de forma breve, accionable y proporcional a la evidencia disponible.

**Instrucciones:**

1. Crea o completa la hoja **Mensajes ejecutivos**.
2. Redacta tres mensajes, cada uno de entre 60 y 100 palabras.
3. Cada mensaje debe seguir la estructura obligatoria:
   1. **Evidencia:** qué KPI cambió, valor, referencia, periodo y segmento.
   2. **Implicación:** por qué importa para el negocio.
   3. **Acción:** qué debe revisarse, priorizarse, decidirse o validarse.
   4. **Nivel de certeza:** certeza alta, media o limitada, con una razón.
4. Usa un mensaje para una situación de desempeño global, uno para una segmentación relevante y uno para una oportunidad o riesgo asociado a una meta.
5. Utiliza lenguaje proporcional:
   - Para certeza limitada: “los datos sugieren”, “se observa una asociación”, “conviene validar”.
   - Para evidencia directa: “el dashboard muestra”, “el indicador se ubica”, “el resultado visible es”.
6. Incluye una acción verificable. Evita recomendaciones vagas como “mejorar las ventas”.
7. No propongas acciones irreversibles basadas únicamente en evidencia descriptiva. Por ejemplo, no recomendar eliminar una categoría sin validar rentabilidad neta, inventario, devoluciones y demanda.
8. Menciona las limitaciones relevantes cuando afecten de forma importante la recomendación.

#### Plantilla de mensaje ejecutivo

> **Evidencia:** Durante [periodo], [KPI] fue de [valor], equivalente a [variación o brecha] frente a [referencia], con mayor concentración en [segmento si aplica].  
> **Implicación:** Este resultado puede afectar [meta, margen, capacidad comercial, retención u otra prioridad].  
> **Acción:** Se recomienda [acción o validación concreta] antes de [momento de decisión].  
> **Nivel de certeza:** [Alta/Media/Limitada], porque [evidencia disponible y limitación].

#### Ejemplo de estructura sin valores predefinidos

> Durante el periodo seleccionado, el cumplimiento de meta mostrado por el dashboard fue de [valor], con una brecha de [valor] respecto al objetivo definido para el mismo alcance. La diferencia se observa principalmente en [segmento], según el desglose disponible. Se recomienda revisar ese segmento en la siguiente práctica mediante consultas transaccionales y contraste con margen, pedidos y devoluciones. El nivel de certeza es medio: la brecha frente a meta es visible, pero el dashboard no permite atribuir su causa a una campaña, precio o disponibilidad.

**Expected output:**

- Tres mensajes ejecutivos completos.
- Cada mensaje contiene evidencia, implicación, acción y nivel de certeza.

**Verification:**

- Revisa que todos los números citados aparezcan en la sección de datos literales u observaciones.
- Comprueba que cada acción sea una investigación, priorización o decisión coherente con la evidencia.
- Sustituye “causó”, “demuestra” o “garantiza” cuando no exista evidencia causal.

---

### Paso 8: Revisar trazabilidad, guardar y preparar la entrega

**Objetivo:** asegurar que el brief sea reutilizable y verificable en la Práctica 7.

**Instrucciones:**

1. Revisa que el brief incluya las siguientes secciones o equivalentes:
   - Contexto del dashboard.
   - Datos literales.
   - Observaciones comparables.
   - Matriz de contraste.
   - Evaluación de afirmaciones.
   - Limitaciones.
   - Mensajes ejecutivos.
2. Verifica que cada cifra utilizada en una observación, hallazgo o mensaje pueda rastrearse a:
   - Un visual.
   - Un periodo.
   - Un conjunto de filtros.
   - Una definición de KPI.
3. Revisa ortografía, unidades y formatos porcentuales.
4. Confirma que no hay credenciales, contraseñas ni datos personales en el libro.
5. Guarda el archivo con el nombre exacto:
   ```text
   M6_Brief_Interpretacion_[Apellido].xlsx
   ```
6. Guarda el archivo en:
   ```text
   C:\DAF\Batch_01\01_brief\
   ```
7. Cierra Excel y vuelve a abrir el archivo para verificar que no está corrupto.
8. Cierra Power BI Desktop sin sobrescribir el dashboard original, salvo instrucción explícita del instructor.

**Expected output:**

- Un brief individual, guardado y legible.
- Un archivo listo para ser usado como entrada en la Práctica 7.

**Verification:**

Ejecuta el siguiente comando en PowerShell:

```powershell
Get-Item "C:\DAF\Batch_01\01_brief\M6_Brief_Interpretacion_[Apellido].xlsx" |
Select-Object Name, Length, LastWriteTime
```

Sustituye `[Apellido]` por tu apellido real. Confirma que el archivo existe, tiene tamaño mayor que 0 KB y una fecha de modificación correspondiente a la sesión actual.

## Validación y Pruebas

Completa la siguiente lista de control antes de entregar el brief:

| Criterio de validación | Resultado esperado | Estado |
|---|---|---|
| Archivo correcto | Existe `M6_Brief_Interpretacion_[Apellido].xlsx` en `01_brief` | Pendiente/Correcto |
| Fuente correcta | Se utilizó `M6_Dashboard_Interpretacion.pbix` | Pendiente/Correcto |
| Periodo documentado | El periodo analizado está registrado | Pendiente/Correcto |
| Filtros documentados | Se registraron filtros globales, de página y relevantes del visual | Pendiente/Correcto |
| Definiciones KPI | Ventas, margen, pedidos, ticket, crecimiento y meta tienen definición o limitación documentada | Pendiente/Correcto |
| Datos literales | Hay al menos 14 registros trazables de tarjetas, series o segmentaciones | Pendiente/Correcto |
| Observaciones | Hay al menos cinco observaciones comparables y cuantificadas | Pendiente/Correcto |
| Comparabilidad | Cada comparación revisa periodo, escala, filtros, granularidad y denominador | Pendiente/Correcto |
| Hallazgos potenciales | Hay dos o tres, con contexto y sin atribución causal indebida | Pendiente/Correcto |
| Afirmaciones | Se clasificaron al menos cinco como sustentadas, parcialmente sustentadas o no sustentadas | Pendiente/Correcto |
| Limitaciones | Se documentaron cuatro o más limitaciones relevantes | Pendiente/Correcto |
| Mensajes ejecutivos | Hay tres mensajes con evidencia, implicación, acción y certeza | Pendiente/Correcto |
| Lenguaje responsable | No se afirma causalidad sin evidencia causal | Pendiente/Correcto |

Como prueba final, selecciona uno de tus mensajes ejecutivos y responde estas preguntas:

1. ¿Qué valor exacto respalda el mensaje?
2. ¿En qué visual se observa?
3. ¿Cuál era el periodo y filtro aplicable?
4. ¿Contra qué referencia se compara?
5. ¿Qué parte del mensaje es evidencia y qué parte es recomendación?
6. ¿Qué información adicional sería necesaria para demostrar causalidad?

Si no puedes responder alguna pregunta, corrige el mensaje o reduce su nivel de certeza.

## Solución de Problemas

### Problema 1: Los KPI o gráficos aparecen en blanco, con error o con valores distintos a los esperados

**Síntomas:** Power BI muestra visuales vacíos, mensajes de error, un aviso de actualización pendiente o cifras distintas después de pulsar **Actualizar**.

**Causa probable:** el dashboard fue abierto con una versión incorrecta, se alteraron filtros, falta una ruta local del paquete de práctica o se intentó actualizar una fuente que no forma parte del alcance de esta actividad.

**Solución:**

1. Cierra el archivo PBIX sin guardar cambios.
2. Verifica que abriste exactamente:
   ```text
   C:\DAF\Batch_01\05_dashboard\M6_Dashboard_Interpretacion.pbix
   ```
3. Vuelve a abrirlo y espera a que finalice la carga.
4. Restablece los filtros mediante el botón de restablecimiento de filtros, si está disponible.
5. No ejecutes **Actualizar** sin instrucciones del instructor.
6. Si el problema persiste, toma una captura del error y repórtalo junto con la versión de Power BI Desktop instalada.

### Problema 2: No es posible determinar si una comparación frente a meta o periodo anterior es válida

**Síntomas:** el dashboard muestra una meta, una variación o un porcentaje, pero no especifica claramente el periodo, el denominador, la moneda, la definición del KPI o los filtros aplicados.

**Causa probable:** el visual presenta información resumida y no contiene todos los metadatos necesarios para sostener la comparación.

**Solución:**

1. Revisa el panel de filtros del visual, la página y el informe.
2. Consulta el diccionario de métricas RetailNova v1.0.
3. Comprueba si el tooltip, el título del visual o una página de definiciones aporta el contexto faltante.
4. Si la información continúa ausente, no calcules ni afirmes una conclusión definitiva.
5. Registra la comparación como **parcialmente sustentada** o **no sustentada**, según corresponda.
6. Escribe una limitación explícita y formula una acción de validación para la Práctica 7.

## Limpieza

1. Guarda y cierra `M6_Brief_Interpretacion_[Apellido].xlsx`.
2. Cierra Power BI Desktop sin guardar cambios sobre `M6_Dashboard_Interpretacion.pbix`, excepto si el instructor solicitó una copia específica.
3. No elimines los archivos fuente ni el diccionario de métricas.
4. Conserva el brief en:
   ```text
   C:\DAF\Batch_01\01_brief\
   ```
5. Mantén el dashboard original en:
   ```text
   C:\DAF\Batch_01\05_dashboard\
   ```
6. No borres las carpetas obligatorias de `C:\DAF\Batch_01\`, ya que se reutilizarán en prácticas posteriores.

## Resumen

En esta práctica interpretaste un dashboard comercial sin confundir valores visibles con conclusiones de negocio. Registraste datos literales, elaboraste observaciones comparables, contrastaste KPI con metas, periodos y segmentos, y evaluaste el nivel de sustento de afirmaciones analíticas.

El entregable `M6_Brief_Interpretacion_[Apellido].xlsx` debe contener evidencia trazable, limitaciones explícitas y tres mensajes ejecutivos con lenguaje proporcional al nivel de certeza. Recuerda el principio central de la práctica: una caída o incremento observado puede sugerir una asociación, pero no demuestra por sí mismo una causa. En la Práctica 7 validarás y ampliarás estos resultados con SQL, Excel y Power BI.
