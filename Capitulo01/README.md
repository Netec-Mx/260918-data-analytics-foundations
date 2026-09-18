# Convertir solicitudes operativas en preguntas analíticas

## Metadatos

| Elemento | Valor |
|---|---|
| Duración | 75 minutos |
| Complejidad | Media |
| Nivel de Bloom | Aplicar |

## Descripción General

En esta práctica transformarás solicitudes operativas generales del liderazgo comercial en preguntas analíticas que puedan responderse con evidencia verificable. Usarás el diccionario de datos del dataset sintético **Ventas Retail LATAM 2026.1**, con entre 500,000 y 1,000,000 de transacciones según la versión distribuida; la versión oficial de referencia contiene 750,000 transacciones entre el 2024-01-01 y el 2025-12-31.

El producto final será un brief analítico reutilizable en Excel y PDF. Este documento funcionará como contrato analítico para definir qué datos deberán limpiarse en la Práctica 2, qué métricas se describirán en la Práctica 3, qué hipótesis se explorarán en la Práctica 4 y qué indicadores formarán parte del dashboard de la Práctica 5.

## Objetivos de Aprendizaje

Al finalizar la práctica, podrás:

- [ ] Distinguir entre una solicitud operativa, una pregunta analítica, un hallazgo, un insight y una decisión.
- [ ] Formular preguntas de negocio específicas, medibles, acotadas temporalmente y orientadas a una decisión comercial.
- [ ] Definir métricas principales y auxiliares, dimensiones, segmentos, granularidad y periodos de comparación.
- [ ] Registrar hipótesis iniciales como explicaciones verificables, sin confundirlas con conclusiones.
- [ ] Crear y exportar un brief analítico que asegure trazabilidad para las prácticas posteriores.

## Prerrequisitos

### Conocimientos requeridos

Debes poder:

- Interpretar los conceptos de fila, columna, tabla, campo, registro, métrica y dimensión.
- Leer tablas y reportes de ventas sencillos.
- Diferenciar la cadena analítica: **dato → información → hallazgo → insight → decisión**.
- Reconocer que una variación observada no prueba automáticamente una causa.
- Usar funciones básicas de Excel, filtros y formato de tablas.
- Comprender fundamentos equivalentes a Excel Intermediate con Copilot y SQL Essentials con Snowflake.

### Acceso requerido

Antes de iniciar, confirma que tienes acceso a:

- La carpeta compartida del curso.
- La plantilla de brief analítico proporcionada por el instructor.
- El diccionario de datos **Ventas Retail LATAM 2026.1**.
- La versión para Excel o CSV del dataset sintético, únicamente para consulta de estructura si el instructor la habilitó.
- Microsoft Excel para Microsoft 365 con permiso para guardar archivos XLSX y exportar PDF.
- La ruta local obligatoria `C:\DAF\Batch_01\`.

> **Importante:** esta práctica no requiere consultar Snowflake ni cargar datos en Power BI Desktop. Sin embargo, el brief debe considerar que el mismo dataset tendrá versiones equivalentes en Excel, Snowflake y Power BI. Los campos definidos aquí deberán ser trazables en las siguientes prácticas.

## Entorno de Laboratorio

### Estructura obligatoria de carpetas

Trabaja exclusivamente dentro de la siguiente ruta:

```text
C:\DAF\Batch_01\
```

La estructura esperada es:

```text
C:\DAF\Batch_01\
├── 00_source\
├── 01_brief\
├── 02_quality\
├── 03_descriptive\
├── 04_exploration\
├── 05_dashboard\
└── sql\
```

### Configuración de hardware y software

| Componente | Requisito de referencia |
|---|---|
| Sistema operativo | Windows 11 Pro 23H2 o equivalente |
| Memoria RAM | 16 GB mínimo; 32 GB recomendados para prácticas posteriores |
| Espacio libre | 20 a 25 GB libres en SSD |
| Excel | Microsoft Excel para Microsoft 365, 64 bits |
| Dataset | Ventas Retail LATAM 2026.1, artificial y sin datos personales reales |
| Conectividad | Internet estable de al menos 20 Mbps para acceder a recursos del curso |

### Preparación inicial de carpetas

1. Abre **Símbolo del sistema** o **Windows PowerShell**.
2. Ejecuta el siguiente comando:

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

3. Copia desde la carpeta compartida del curso los siguientes archivos hacia `C:\DAF\Batch_01\00_source\`:

```text
Diccionario_Datos_Ventas_Retail_LATAM_2026_1.xlsx
Plantilla_Brief_Analitico_Ventas.xlsx
```

4. Si el instructor distribuyó un archivo de datos local, cópialo también a `00_source`. No modifiques ni sobrescribas el archivo original.

> **Regla de seguridad y trazabilidad:** no almacenes credenciales institucionales, contraseñas temporales ni cadenas de conexión en el brief, el archivo Excel ni el PDF. En prácticas posteriores, el objeto fuente Snowflake será inmutable: `DATA_ANALYTICS_FOUNDATIONS.RAW.VENTAS_TRANSACCIONES_2026_1`.

## Instrucciones Paso a Paso

### Paso 1: Preparar el espacio de trabajo y los archivos base

**Objetivo:** Crear una estructura de trabajo trazable y abrir los documentos necesarios sin modificar los archivos fuente.

**Instrucciones:**

1. Confirma que existe la ruta `C:\DAF\Batch_01\01_brief\`.
2. Abre el Explorador de archivos y verifica que los archivos de fuente estén disponibles en:

   ```text
   C:\DAF\Batch_01\00_source\
   ```

3. Abre el archivo `Plantilla_Brief_Analitico_Ventas.xlsx` desde `00_source`.
4. Guarda inmediatamente una copia de trabajo con el nombre exacto:

   ```text
   C:\DAF\Batch_01\01_brief\01_brief_analitico_ventas.xlsx
   ```

5. Abre en una segunda ventana el archivo `Diccionario_Datos_Ventas_Retail_LATAM_2026_1.xlsx`.
6. Revisa el nombre de las hojas del diccionario y localiza, como mínimo, la información sobre:
   - Identificador de transacción.
   - Fecha de transacción o fecha de venta.
   - Producto y categoría.
   - Cliente o identificador de cliente.
   - Región, país, ciudad, canal o sucursal, según estén disponibles.
   - Cantidad, precio, descuento, importe de venta, costo, margen o devolución, según correspondan al diccionario.
   - Estado de la transacción, si existe.
7. En el archivo de brief, completa los metadatos iniciales:
   - Nombre del estudiante.
   - Fecha de elaboración.
   - Versión del dataset: `Ventas Retail LATAM 2026.1`.
   - Cobertura temporal de referencia: `2024-01-01 a 2025-12-31`.
   - Fuente: dataset sintético oficial del curso.
   - Estado del documento: `Borrador inicial`.

**Resultado esperado:**

Existe un archivo de trabajo llamado `01_brief_analitico_ventas.xlsx` en la carpeta `01_brief`. El archivo fuente y el diccionario permanecen intactos en `00_source`.

**Verificación:**

- Comprueba en el Explorador de archivos que el archivo tenga extensión `.xlsx`.
- Verifica que la ruta sea exactamente `C:\DAF\Batch_01\01_brief\01_brief_analitico_ventas.xlsx`.
- Confirma que no estás editando directamente la plantilla original ubicada en `00_source`.
- Revisa que el brief identifique correctamente que el dataset es artificial y que no contiene datos personales reales.

---

### Paso 2: Clasificar las solicitudes operativas del liderazgo comercial

**Objetivo:** Diferenciar una solicitud general de información de una pregunta analítica que pueda orientar una decisión.

**Instrucciones:**

1. En la hoja del brief destinada a solicitudes o contexto, registra las siguientes solicitudes operativas simuladas:

   | ID | Solicitud operativa del liderazgo comercial |
   |---|---|
   | S1 | “Necesitamos saber por qué bajaron las ventas.” |
   | S2 | “¿Qué productos debemos impulsar?” |
   | S3 | “¿Cómo van las regiones?” |

2. Para cada solicitud, identifica por qué todavía no es una pregunta analítica suficiente. Usa criterios como los siguientes:
   - No especifica una métrica.
   - No define periodo actual ni periodo de comparación.
   - No indica región, canal, categoría, cliente u otro segmento.
   - No define qué decisión se tomará con la respuesta.
   - Usa términos ambiguos como “bajaron”, “impulsar” o “cómo van”.
   - Puede inducir una conclusión causal sin evidencia suficiente.

3. Completa una columna llamada **Ambigüedades o información faltante**. Registra al menos tres elementos por solicitud.

4. Completa una columna llamada **Decisión potencial que podría apoyar**. Formula una decisión concreta, no un análisis. Por ejemplo:
   - Ajustar la asignación comercial o promocional.
   - Priorizar productos para una campaña.
   - Revisar desempeño regional y reasignar recursos.
   - Solicitar una investigación adicional antes de actuar.

5. Usa la siguiente tabla como referencia para registrar tu análisis inicial:

   | ID | Solicitud operativa | Ambigüedades o información faltante | Decisión potencial |
   |---|---|---|---|
   | S1 | Necesitamos saber por qué bajaron las ventas | No define métrica, periodo, comparación ni segmento; presupone una causa | Determinar si se requiere intervención comercial, operativa o mayor investigación |
   | S2 | ¿Qué productos debemos impulsar? | No define criterio de “impulsar”, objetivo, canal, horizonte ni presupuesto | Priorizar productos para promoción, reposición o revisión de portafolio |
   | S3 | ¿Cómo van las regiones? | No especifica indicador, objetivo, periodo, comparación ni nivel geográfico | Reasignar recursos comerciales o profundizar en regiones con desempeño atípico |

6. Añade una nota metodológica en el brief:

   > Una solicitud operativa expresa una necesidad o intención de negocio. Una pregunta analítica define qué se medirá, para qué decisión, en qué periodo, para qué población o segmento y bajo qué condiciones se interpretará el resultado.

**Resultado esperado:**

El brief contiene las tres solicitudes originales, sus ambigüedades documentadas y una decisión potencial asociada a cada una.

**Verificación:**

- Cada solicitud tiene al menos tres ambigüedades registradas.
- Ninguna solicitud se presenta todavía como una conclusión.
- Las decisiones potenciales empiezan con un verbo de acción, por ejemplo: priorizar, reasignar, revisar, investigar o ajustar.
- El texto distingue claramente la necesidad comercial de la evidencia que deberá producirse.

---

### Paso 3: Convertir las solicitudes en preguntas analíticas medibles

**Objetivo:** Redactar una pregunta analítica para cada solicitud utilizando métricas, dimensiones, periodos y criterios de comparación explícitos.

**Instrucciones:**

1. Crea una sección llamada **Preguntas analíticas propuestas**.
2. Para cada solicitud, redacta una pregunta que incluya explícitamente:
   - La métrica principal.
   - El periodo de análisis.
   - El periodo o referencia de comparación.
   - Una o más dimensiones de corte.
   - El segmento de interés, cuando aplique.
   - La finalidad de decisión.

3. Evita formular preguntas que supongan causalidad. Por ejemplo, no escribas:

   > “¿Qué causó la caída de ventas?”

   Esta pregunta supone que existe una caída confirmada y que el dataset contiene todas las variables necesarias para probar una causa.

4. Utiliza formulaciones como:
   - “¿En qué medida…?”
   - “¿Qué variación se observa…?”
   - “¿Qué segmentos concentran…?”
   - “¿Qué productos presentan…?”
   - “¿Qué diferencias existen…?”
   - “¿Qué evidencia disponible es consistente con…?”

5. Registra una versión inicial de las siguientes preguntas, ajustando los nombres de campos a los existentes en el diccionario:

   | ID | Pregunta analítica propuesta |
   |---|---|
   | P1 | ¿Cuál fue la variación mensual de las ventas netas y del número de transacciones durante 2025 frente al mismo mes de 2024, por región, canal y categoría de producto, para identificar los segmentos que requieren revisión comercial? |
   | P2 | ¿Qué productos y categorías concentran mayor venta neta, margen bruto, unidades vendidas y tasa de devolución durante los últimos seis meses disponibles de 2025, por canal y región, para priorizar productos candidatos a promoción, reposición o revisión? |
   | P3 | ¿Cómo se compara el desempeño de cada región en ventas netas, margen bruto, transacciones, clientes activos y variación interanual durante 2025 frente a 2024, para orientar la asignación de recursos comerciales? |

6. Debajo de cada pregunta, identifica sus componentes en una tabla como esta:

   | Componente | P1: Ejemplo de respuesta |
   |---|---|
   | Decisión que apoya | Determinar qué segmentos requieren revisión comercial prioritaria |
   | Métrica principal | Ventas netas |
   | Métricas auxiliares | Número de transacciones, unidades, descuento, devoluciones, margen bruto |
   | Periodo actual | Meses de 2025 |
   | Periodo comparativo | Mismo mes de 2024 |
   | Dimensiones | Mes, región, canal, categoría, producto |
   | Granularidad del análisis | Mes × región × canal × categoría |
   | Criterio de lectura | Variación absoluta, variación porcentual y contribución a la variación total |

7. Revisa que las métricas de tu pregunta estén soportadas por el diccionario. Si una métrica no existe directamente, documenta cómo se calcularía. Ejemplos:

   | Métrica | Definición propuesta | Validación requerida |
   |---|---|---|
   | Ventas brutas | Suma del importe antes de descuentos y devoluciones, si los campos existen | Confirmar definición en el diccionario |
   | Ventas netas | Ventas brutas menos descuentos y devoluciones, según reglas del dataset | Verificar si ya existe un campo de importe neto |
   | Tasa de devolución | Importe o unidades devueltas / importe o unidades vendidas | Confirmar tratamiento de devoluciones |
   | Margen bruto | Ventas netas menos costo | Confirmar disponibilidad y unidad monetaria del costo |
   | Clientes activos | Conteo distinto de identificadores de cliente con al menos una transacción válida | Confirmar campo de cliente y criterios de transacción válida |

8. No inventes nombres de campos. En una columna llamada **Campo confirmado en diccionario**, escribe el nombre exacto solo después de validarlo en el archivo de diccionario. Si un campo no existe, marca:

   ```text
   No disponible / requiere validación
   ```

**Resultado esperado:**

El brief contiene tres preguntas analíticas concretas. Cada pregunta define una métrica, una comparación temporal, dimensiones de análisis y la decisión que pretende apoyar.

**Verificación:**

- Las tres preguntas pueden responderse con una tabla, una consulta o una visualización definida.
- Las preguntas no contienen términos ambiguos sin definición, como “mejor”, “mucho”, “poco” o “bien”.
- Cada pregunta especifica un periodo y, cuando corresponde, un comparativo.
- Ninguna pregunta afirma que una variable “causó” otra sin un diseño o evidencia causal.
- Los nombres de campos usados coinciden con el diccionario o están marcados como pendientes de validación.

---

### Paso 4: Priorizar una pregunta y definir el contrato analítico

**Objetivo:** Seleccionar una pregunta prioritaria y convertirla en un contrato analítico que guíe las prácticas posteriores.

**Instrucciones:**

1. Crea una matriz de priorización para P1, P2 y P3.
2. Evalúa cada pregunta con una escala de 1 a 5, donde:
   - `1` = bajo o poco favorable.
   - `3` = intermedio.
   - `5` = alto o muy favorable.

3. Usa los siguientes criterios:

   | Criterio | Pregunta orientadora |
   |---|---|
   | Impacto de negocio | ¿La respuesta puede apoyar una decisión comercial relevante? |
   | Disponibilidad de datos | ¿Los campos requeridos existen o son razonablemente derivables? |
   | Claridad de la métrica | ¿La métrica principal puede definirse sin ambigüedad? |
   | Comparabilidad temporal | ¿Existe un periodo comparable suficiente? |
   | Viabilidad en el curso | ¿Puede analizarse y visualizarse en las prácticas posteriores? |
   | Riesgo de interpretación | ¿Puede comunicarse sin hacer afirmaciones no sustentadas? |

4. Calcula el puntaje total de cada pregunta. Puedes usar una fórmula como:

```excel
=SUM(B2:G2)
```

5. Selecciona una pregunta prioritaria. Para esta práctica se recomienda priorizar **P1**, porque permite establecer una base reutilizable de desempeño comercial por periodo, región, canal y categoría. Sin embargo, puedes elegir P2 o P3 si justificas la decisión y si los campos requeridos están disponibles.

6. Crea una sección llamada **Contrato analítico de la pregunta priorizada** y completa los siguientes apartados:

   | Elemento | Contenido requerido |
   |---|---|
   | Pregunta priorizada | Redacción final de la pregunta |
   | Decisión que apoya | Decisión concreta que podría tomar liderazgo comercial |
   | Usuario o audiencia | Liderazgo comercial, gerencia regional u otra audiencia definida |
   | Métrica principal | Nombre, definición, unidad y regla de cálculo |
   | Métricas auxiliares | Lista de métricas que ayudan a interpretar la principal |
   | Periodo actual | Fechas exactas de inicio y fin |
   | Periodo de comparación | Fechas exactas y criterio de comparabilidad |
   | Dimensiones de corte | Región, canal, categoría, producto, mes u otras |
   | Segmentos | Segmentos incluidos o excluidos |
   | Granularidad | Nivel mínimo de detalle requerido para responder |
   | Fuente de datos | Versión Excel, objeto Snowflake y modelo Power BI previstos |
   | Criterios de éxito | Condiciones que debe cumplir el análisis para ser útil |
   | Limitaciones conocidas | Restricciones de cobertura, calidad o interpretación |

7. Para la pregunta P1, un ejemplo de granularidad adecuada sería:

   ```text
   Una fila analítica por mes, región, canal y categoría de producto.
   ```

8. Registra explícitamente el principio de comparabilidad temporal. Por ejemplo:

   > La comparación principal será mes contra el mismo mes del año anterior. No se compararán meses parciales contra meses completos. Antes de calcular variaciones se validará que ambos periodos tengan cobertura equivalente de fechas y que se apliquen las mismas reglas de exclusión.

9. Define criterios de éxito medibles. Por ejemplo:
   - Las métricas principales tienen una definición documentada y reproducible.
   - Los campos requeridos están confirmados en el diccionario.
   - El análisis identifica variaciones por segmento, no solo un total general.
   - Todo porcentaje de variación tiene denominador válido y periodo comparable.
   - Las recomendaciones futuras diferenciarán evidencia observada, hipótesis e implicación de negocio.

**Resultado esperado:**

El brief contiene una pregunta priorizada, una justificación de priorización y un contrato analítico completo que puede ser utilizado por otra persona para repetir el análisis.

**Verificación:**

- La pregunta seleccionada tiene el puntaje más alto o una justificación documentada para una elección diferente.
- El periodo actual y el periodo comparativo incluyen fechas específicas.
- La granularidad no es ambigua.
- Los criterios de éxito pueden verificarse al final de las prácticas posteriores.
- La fuente se describe de forma trazable, sin incluir contraseñas ni credenciales.

---

### Paso 5: Documentar hipótesis, supuestos, restricciones y campos requeridos

**Objetivo:** Separar lo que los datos muestran de las explicaciones que deben validarse, e identificar riesgos de calidad que afectarán el análisis.

**Instrucciones:**

1. Crea una hoja o sección llamada **Hipótesis, supuestos y restricciones**.
2. Registra al menos tres hipótesis iniciales para la pregunta priorizada. Una hipótesis debe ser una explicación posible y verificable, no una afirmación definitiva.

3. Para cada hipótesis, completa los siguientes campos:

   | Campo | Descripción |
   |---|---|
   | ID de hipótesis | Identificador, por ejemplo H1 |
   | Hipótesis inicial | Explicación plausible que deberá evaluarse |
   | Evidencia disponible esperada | Campos o métricas que podrían examinarse |
   | Evidencia faltante | Información no disponible en el dataset o que requeriría otra fuente |
   | Método de contraste | Comparación, segmentación o análisis previsto |
   | Estado inicial | Pendiente de validación |

4. Usa como referencia hipótesis responsables para P1:

   | ID | Hipótesis inicial | Evidencia disponible esperada | Precaución |
   |---|---|---|---|
   | H1 | La disminución de ventas netas podría concentrarse en determinadas categorías o canales. | Ventas netas por mes, categoría y canal. | Una concentración no prueba la causa de la caída. |
   | H2 | Las devoluciones o descuentos podrían contribuir a una reducción de ventas netas en algunos segmentos. | Campos de devolución, descuento e importe neto, si existen. | Se debe validar la regla contable aplicada. |
   | H3 | La variación podría reflejar cambios en el número de transacciones, en el valor promedio por transacción o en ambos. | Identificador de transacción, ventas netas y fechas. | Requiere descomposición de métricas. |
   | H4 | Diferencias regionales podrían estar asociadas con la mezcla de productos o canales. | Región, categoría, producto y canal. | La asociación no demuestra causalidad. |

5. Documenta al menos cinco supuestos verificables. Ejemplos:
   - Cada identificador de transacción representa una transacción única o cuenta con reglas para deduplicación.
   - Las fechas están registradas en un formato válido y dentro de la cobertura esperada.
   - Los importes monetarios usan una unidad y moneda consistentes según el diccionario.
   - Los valores negativos tienen una semántica definida, por ejemplo devolución, nota de crédito o corrección.
   - El identificador de cliente puede utilizarse para conteos distintos si el diccionario lo autoriza.
   - Las categorías y regiones utilizan catálogos consistentes.

6. Registra restricciones y riesgos de calidad que deberán revisarse en la Práctica 2. Como mínimo, incluye:

   | Restricción o riesgo | Posible efecto en el análisis | Acción prevista |
   |---|---|---|
   | Registros duplicados | Sobreestimación de ventas, transacciones o clientes | Revisar unicidad del identificador de transacción y reglas de duplicado |
   | Fechas nulas, inválidas o futuras | Comparaciones temporales incorrectas | Validar rango entre 2024-01-01 y 2025-12-31, salvo documentación distinta |
   | Devoluciones | Distorsión de ventas netas y cantidades | Confirmar campos, signos y regla de tratamiento |
   | Importes nulos o negativos | Métricas inconsistentes | Revisar semántica y definir tratamiento documentado |
   | Segmentos incompletos | Resultados sesgados por región, canal o categoría | Medir valores nulos y definir categoría “No informado” solo si procede |
   | Cobertura temporal desigual | Variaciones porcentuales engañosas | Comparar periodos equivalentes y completos |
   | Catálogos inconsistentes | Fragmentación artificial de categorías o regiones | Normalizar valores solo en capa curada, nunca en RAW |

7. Crea una tabla de **Campos requeridos**. Usa los nombres exactos que aparezcan en el diccionario. La estructura mínima será:

   | Necesidad analítica | Campo confirmado en diccionario | Tipo esperado | Uso previsto | Obligatorio |
   |---|---|---|---|---|
   | Identificar una transacción | `[nombre exacto]` | Texto o entero | Conteo de transacciones y revisión de duplicados | Sí |
   | Ubicar temporalmente la venta | `[nombre exacto]` | Fecha | Filtros, agregación mensual y comparación interanual | Sí |
   | Medir ventas | `[nombre exacto]` | Decimal o moneda | Métrica principal o cálculo de ventas netas | Sí |
   | Segmentar por territorio | `[nombre exacto]` | Texto | Corte regional | Sí |
   | Segmentar por canal | `[nombre exacto]` | Texto | Corte por canal | Según pregunta |
   | Segmentar por producto | `[nombre exacto]` | Texto | Producto y categoría | Sí |
   | Identificar devoluciones | `[nombre exacto]` o `No disponible` | Decimal, entero o indicador | Ajuste o análisis auxiliar | Según disponibilidad |
   | Identificar cliente | `[nombre exacto]` o `No disponible` | Texto o entero | Clientes activos y segmentación | Según pregunta |

8. Agrega una nota de trazabilidad para prácticas posteriores:

   > La validación de campos, tipos de datos, duplicados, valores nulos, fechas inválidas y reglas de devoluciones se ejecutará sobre la versión preparada del dataset. En Snowflake no se modificarán objetos del esquema `RAW`; cualquier transformación se realizará en el esquema `CURATED` o en un objeto autorizado con sufijo de estudiante.

**Resultado esperado:**

El brief documenta hipótesis verificables, supuestos explícitos, restricciones de calidad y los campos requeridos para responder la pregunta priorizada.

**Verificación:**

- Hay al menos tres hipótesis y cada una está marcada como pendiente de validación.
- Ninguna hipótesis se presenta como hecho confirmado.
- Hay al menos cinco supuestos verificables.
- Se registran riesgos de duplicados, devoluciones, fechas futuras o inválidas y datos incompletos.
- Los campos requeridos utilizan nombres confirmados por el diccionario o están marcados como no disponibles.

---

### Paso 6: Redactar el mensaje analítico inicial y revisar el nivel de certeza

**Objetivo:** Comunicar el alcance del análisis sin convertir hipótesis en conclusiones ni confundir datos con insights.

**Instrucciones:**

1. Crea una sección llamada **Mensaje analítico inicial**.
2. Redacta un mensaje de entre 80 y 150 palabras para la audiencia comercial. El mensaje debe incluir:
   - La pregunta priorizada.
   - La decisión que se pretende apoyar.
   - El periodo de análisis y comparación.
   - La métrica principal.
   - Las principales dimensiones de segmentación.
   - Una advertencia sobre las hipótesis y limitaciones iniciales.

3. Usa expresiones que indiquen adecuadamente el nivel de certeza:
   - Para hechos a validar: “se medirá”, “se revisará”, “se comparará”.
   - Para asociaciones: “coincide con”, “podría estar relacionado con”, “es consistente con”.
   - Para límites: “no permite atribuir causalidad por sí solo”, “requiere validación adicional”.
   - Para acciones futuras: “se recomendará evaluar”, “se priorizará investigar”.

4. Evita expresiones no sustentadas como:
   - “La región Norte bajó por falta de inventario.”
   - “Los clientes dejaron de comprar por el precio.”
   - “El producto A es el peor producto.”
   - “La campaña debe eliminarse.”

5. Usa esta estructura recomendada:

   > Se analizará [métrica principal] para [periodo] frente a [comparativo], segmentada por [dimensiones], con el propósito de apoyar [decisión]. El análisis identificará variaciones y concentraciones relevantes, así como el comportamiento de métricas auxiliares como [métricas]. Las diferencias observadas permitirán formular hallazgos; cualquier explicación sobre sus causas se tratará como hipótesis y requerirá evidencia adicional. Antes de interpretar resultados se validarán [restricciones principales].

6. Añade una mini tabla que separe los niveles de razonamiento:

   | Nivel | Ejemplo aplicado al brief |
   |---|---|
   | Dato | Importe de una transacción individual, fecha, región y producto |
   | Información | Ventas netas mensuales por región y categoría |
   | Hallazgo posible | Una región presenta una variación interanual negativa mayor que el total |
   | Insight posible | La variación podría concentrarse en una combinación de canal y categoría; requiere contraste |
   | Decisión posible | Priorizar una revisión comercial o una acción focalizada después de validar la evidencia |

7. Revisa ortografía, consistencia de fechas, nombres de métricas y términos de negocio. Asegúrate de usar siempre la misma definición para “ventas netas”, “transacción”, “cliente activo” y otras métricas incluidas.

**Resultado esperado:**

El brief presenta un mensaje ejecutivo inicial que comunica el propósito del análisis, sus límites y el uso esperado de la evidencia.

**Verificación:**

- El mensaje no declara resultados que aún no han sido calculados.
- El mensaje diferencia claramente una posible explicación de un hecho observado.
- La tabla de niveles incluye dato, información, hallazgo, insight y decisión.
- El mensaje es comprensible para una persona de liderazgo comercial sin requerir conocimiento técnico avanzado.

---

### Paso 7: Finalizar, exportar y entregar el brief analítico

**Objetivo:** Completar un entregable reutilizable, legible y trazable en formatos XLSX y PDF.

**Instrucciones:**

1. Revisa que el archivo Excel contenga, como mínimo, las siguientes secciones o hojas:
   - Portada o metadatos.
   - Solicitudes operativas y ambigüedades.
   - Preguntas analíticas propuestas.
   - Matriz de priorización.
   - Contrato analítico de la pregunta priorizada.
   - Hipótesis, supuestos y restricciones.
   - Campos requeridos.
   - Mensaje analítico inicial.

2. Aplica formato profesional:
   - Títulos claros.
   - Encabezados visibles.
   - Ajuste de texto en celdas extensas.
   - Filtros en tablas con múltiples registros.
   - Fechas en formato consistente, preferiblemente `aaaa-mm-dd`.
   - Métricas monetarias con formato de moneda solo si la moneda está confirmada en el diccionario.
   - Porcentajes con uno o dos decimales cuando corresponda.

3. Comprueba que no existan credenciales, contraseñas, tokens, correos personales no requeridos ni datos no relacionados con la práctica.
4. Guarda el libro de Excel.
5. Exporta el archivo a PDF desde Excel:
   1. Selecciona **Archivo > Exportar > Crear documento PDF/XPS**.
   2. Selecciona la carpeta:

      ```text
      C:\DAF\Batch_01\01_brief\
      ```

   3. Asigna el nombre exacto:

      ```text
      01_brief_analitico_ventas.pdf
      ```

   4. En **Opciones**, confirma que se exporten las hojas activas o el libro completo según la estructura creada.
   5. Verifica la orientación de páginas. Usa orientación horizontal para tablas anchas.
   6. Publica o guarda el PDF.

6. Abre el PDF exportado y verifica que:
   - No existan tablas cortadas.
   - Los encabezados sean legibles.
   - Las páginas estén ordenadas.
   - El contenido del PDF coincida con la versión final del Excel.

7. Actualiza el estado del brief a:

   ```text
   Listo para Práctica 2: validación de calidad de datos
   ```

**Resultado esperado:**

La carpeta `C:\DAF\Batch_01\01_brief\` contiene los dos entregables requeridos:

```text
01_brief_analitico_ventas.xlsx
01_brief_analitico_ventas.pdf
```

**Verificación:**

- Ambos archivos existen en la carpeta correcta.
- El PDF se abre sin errores y contiene todas las secciones relevantes.
- La pregunta priorizada, las métricas, los campos requeridos y las restricciones son consistentes entre XLSX y PDF.
- El documento no afirma causas como hechos sin evidencia.
- El brief puede ser usado como entrada para la Práctica 2 sin requerir aclaraciones fundamentales.

## Validación y Pruebas

Completa la siguiente lista de control antes de entregar:

| Criterio de validación | Resultado esperado | Estado |
|---|---|---|
| Archivo Excel creado | Existe `01_brief_analitico_ventas.xlsx` en `01_brief` | ☐ |
| Archivo PDF creado | Existe `01_brief_analitico_ventas.pdf` en `01_brief` | ☐ |
| Solicitudes registradas | Incluye S1, S2 y S3 | ☐ |
| Ambigüedades identificadas | Cada solicitud tiene al menos tres elementos faltantes o ambiguos | ☐ |
| Preguntas formuladas | Existen tres preguntas analíticas medibles | ☐ |
| Periodo definido | Cada pregunta incluye fechas o una referencia temporal verificable | ☐ |
| Comparabilidad definida | La pregunta priorizada define un periodo comparativo equivalente | ☐ |
| Decisión definida | La pregunta priorizada apoya una decisión comercial concreta | ☐ |
| Métrica principal definida | Incluye definición, unidad y regla de cálculo o validación pendiente | ☐ |
| Métricas auxiliares definidas | Apoyan la interpretación de la métrica principal | ☐ |
| Segmentación definida | Incluye dimensiones y segmentos relevantes | ☐ |
| Granularidad definida | Indica el nivel mínimo de detalle analítico | ☐ |
| Hipótesis documentadas | Incluye al menos tres hipótesis pendientes de validación | ☐ |
| Supuestos documentados | Incluye al menos cinco supuestos verificables | ☐ |
| Restricciones documentadas | Considera duplicados, devoluciones, fechas y datos incompletos | ☐ |
| Campos trazables | Los nombres se confirmaron en el diccionario o se marcaron como pendientes | ☐ |
| Nivel de certeza adecuado | No confunde hallazgos esperados con insights o causas confirmadas | ☐ |
| Trazabilidad futura | El brief indica su uso en calidad, descriptivos, exploración y dashboard | ☐ |

Como prueba final, responde estas preguntas sin abrir el diccionario:

1. ¿Qué decisión pretende apoyar la pregunta priorizada?
2. ¿Cuál es la métrica principal y cómo se definirá?
3. ¿Qué periodo se compara contra cuál?
4. ¿Qué dimensiones se usarán para segmentar?
5. ¿Qué riesgos de calidad se deben validar antes de calcular resultados?
6. ¿Qué hipótesis requieren evidencia adicional?

Si alguna respuesta no es clara, actualiza el brief antes de entregarlo.

## Solución de Problemas

### Problema 1: La pregunta analítica sigue siendo demasiado amplia o no permite definir una visualización

**Síntomas:** La pregunta contiene expresiones como “cómo van las ventas”, “qué región está mejor” o “qué productos impulsar”, pero no define métrica, periodo, comparación ni decisión. Diferentes compañeros interpretarían la misma pregunta de maneras distintas.

**Causa:** La solicitud operativa se copió casi literalmente sin convertirla en un problema analítico medible y acotado.

**Solución:**

1. Identifica la decisión que se tomará con la respuesta.
2. Define una métrica principal única, por ejemplo ventas netas.
3. Agrega un periodo actual y un comparativo equivalente.
4. Define dimensiones, como región, canal, categoría o producto.
5. Establece la granularidad, por ejemplo mes × región × canal.
6. Reescribe la pregunta con la estructura: “¿Cuál es la variación de [métrica] en [periodo] frente a [comparativo], por [dimensiones], para apoyar [decisión]?”

---

### Problema 2: El brief afirma causas que el dataset no puede demostrar

**Síntomas:** El documento contiene afirmaciones como “las ventas bajaron por inventario”, “los clientes dejaron de comprar por precio” o “la campaña causó la caída”, pero no hay variables suficientes, método causal ni evidencia adicional documentada.

**Causa:** Se confundió un hallazgo potencial con un insight confirmado o una explicación causal.

**Solución:**

1. Cambia afirmaciones causales por hipótesis verificables.
2. Usa lenguaje de certeza apropiado: “podría estar asociado”, “coincide con”, “requiere validación”.
3. Documenta qué campos del dataset podrían aportar evidencia.
4. Registra qué información externa sería necesaria, por ejemplo inventario, campañas, precios históricos o datos de competencia.
5. Mantén separados los apartados de evidencia, hallazgo, insight posible y decisión.
6. Indica que el dataset de ventas puede mostrar patrones y asociaciones, pero no prueba por sí solo causalidad.

## Limpieza

1. Guarda y cierra `01_brief_analitico_ventas.xlsx`.
2. Confirma que el PDF final está disponible y puede abrirse.
3. Mantén los archivos entregables en:

   ```text
   C:\DAF\Batch_01\01_brief\
   ```

4. No elimines la plantilla ni el diccionario de datos ubicados en `00_source`.
5. No modifiques archivos fuente del dataset sintético.
6. Si creaste archivos temporales, borradores duplicados o exportaciones de prueba, elimínalos o muévelos fuera de `01_brief` para evitar confusión durante la evaluación.
7. Conserva el brief final, ya que será una entrada obligatoria para las prácticas de calidad, análisis descriptivo, exploración y dashboard.

## Resumen

En esta práctica convertiste solicitudes comerciales generales en preguntas analíticas medibles y orientadas a decisiones. Definiste una pregunta priorizada con métricas, dimensiones, segmentos, granularidad, periodos comparables y criterios de éxito.

También documentaste hipótesis, supuestos, restricciones y campos requeridos. Esta documentación evita que las siguientes prácticas produzcan análisis desconectados de la necesidad de negocio o conclusiones no sustentadas. Recuerda la secuencia profesional: los datos se organizan como información; la comparación permite identificar hallazgos; el contexto ayuda a formular insights; y la evidencia sustenta decisiones bajo un nivel de certeza explícito.

El entregable final para esta práctica es:

```text
C:\DAF\Batch_01\01_brief\01_brief_analitico_ventas.xlsx
C:\DAF\Batch_01\01_brief\01_brief_analitico_ventas.pdf
```

En la siguiente práctica validarás la calidad de los campos, fechas, duplicados, devoluciones y valores incompletos identificados en este brief.
