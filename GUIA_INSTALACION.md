# Guía de instalación — Reportes POB y Comedor

Esta guía asume Excel para Microsoft 365 de 64 bits en una PC donde no podés
instalar software. No hace falta instalar nada: todo se hace con Power
Query, que ya viene incluido en Excel (pestaña **Datos**).

## 0. Antes de empezar: subir los archivos de origen a OneDrive/SharePoint

1. En tu OneDrive for Business (o un sitio de SharePoint, es el mismo
   conector) corporativo, creá una carpeta para los archivos de fichadas
   (por ejemplo `Documents/POB/`).
2. Subí ahí los archivos `POB*.xlsx` (todos los que tengas acumulados). El
   nombre tiene que empezar con `POB` y terminar en `.xlsx`.
3. Anotá la **URL raíz del sitio** (`URL_Sitio_OneDrive`) y la **Ruta de
   carpeta** (`Ruta_Carpeta`). **Ojo:** no es simplemente "copiar la URL de
   la barra del navegador cuando estás mirando la carpeta" — esa URL trae de
   más (la vista de la carpeta, parámetros como `?FolderCTID=...&id=...`) y
   la consulta la va a rechazar con un error como
   *"The input URL is invalid. Please provide a URL to the root of a
   SharePoint site..."*. Hacelo así:

   - **Si es tu OneDrive personal** (`.../personal/tuusuario/...`): la URL
     raíz termina justo después de `/personal/tuusuario/`, por ejemplo
     `https://tuempresa-my.sharepoint.com/personal/jperez_tuempresa_com/`.
     Todo lo que sigue (`/Documents/POB/`) es la Ruta de carpeta.
   - **Si es un sitio de SharePoint** (`.../sites/NombreDelSitio/...`, como
     una biblioteca de equipo): la URL raíz termina justo después de
     `/sites/NombreDelSitio/`, por ejemplo
     `https://tuempresa.sharepoint.com/sites/BASEDEDATOS-CONTROL/`. Todo lo
     que sigue es la Ruta de carpeta — pero **no** copies el `id=...` de la
     URL (viene codificado con `%2F` en vez de `/`); mejor seguí el paso 4.
4. Si no estás seguro de qué poner en `Ruta_Carpeta`, dejala vacía primero,
   completá sólo `URL_Sitio_OneDrive` (la raíz) y refrescá la consulta
   `02_Origen` en el Editor de Poder Query: te va a listar TODOS los
   archivos del sitio, con una columna `Folder Path` que muestra el path
   real y sin codificar de cada carpeta (por ejemplo
   `/sites/BASEDEDATOS-CONTROL/Documents partages/CONTROL DE CATERING/POB/TEST`).
   Copiá de ahí un fragmento que identifique sólo tu carpeta (no hace falta
   el path completo ni las mayúsculas exactas, la consulta busca por
   coincidencia parcial) — por ejemplo `POB/TEST` — y pegalo en
   `Ruta_Carpeta`.

5. **Paso obligatorio, una sola vez por libro**: desactivá el chequeo de
   niveles de privacidad para este Excel. La consulta `02_Origen` usa un
   valor que viene de otra consulta (`01_Parametros`, que lee la hoja
   Config) para conectarse a SharePoint, y Power Query bloquea eso por
   defecto con un error tipo *"Formula.Firewall: ... references other
   queries or steps..."*. Para permitirlo:
   - `Datos > Obtener datos > Opciones de consulta` (o `Archivo > Opciones
     y configuración > Opciones`).
   - En el panel izquierdo, bajo **"Libro actual"** (no "Global"), entrá a
     **Privacidad**.
   - Elegí **"Ignorar siempre los niveles de privacidad (esto podría
     exponer datos confidenciales)"**.
   - Esto es seguro acá: todo lo que se combina es interno de tu empresa
     (la hoja Config del propio libro y tu sitio de SharePoint/OneDrive
     corporativo), no hay ningún origen externo de por medio.

Vas a necesitar estos dos datos (URL y Ruta de carpeta) para completar la
hoja Config, y el ajuste de privacidad del punto 5 para que las consultas
puedan refrescar sin errores.

## 1. Abrir la plantilla

1. Abrí `Plantilla_Reportes_POB.xlsx`.
2. Si Excel te pide habilitar contenido/macros al abrirlo, aceptá (la
   plantilla en sí no trae macros; sólo hace falta si vas a usar el Camino A).
3. Andá a la hoja **Config** y completá:
   - `tParametros` → `URL_Sitio_OneDrive` con la URL del paso 0.
   - `tParametros` → `Ruta_Carpeta` con la ruta del paso 0.
   - Revisá `tLectores`: ya viene precargada con los 100 lectores de la
     muestra. Si tu operación tiene lectores nuevos que no están en la
     lista, agregalos con su Tipo (`INGRESO`, `EGRESO`, `COMEDOR` o
     `INTERNO`) — mientras no lo hagas, el lector nuevo se trata como
     `INTERNO` automáticamente y queda listado en `Control_Calidad`
     ("Lectores sin clasificar") para que lo notes.
   - Revisá `tHorasFoto`, `tFranjas`, `tTurnosComedor` y `tEventosValidos`:
     ya vienen con los valores del prompt original; ajustalos si tu
     operación necesita otros horarios.

## Camino A — Con la macro (recomendado)

1. Guardá `Plantilla_Reportes_POB.xlsx` como **Excel habilitado para
   macros** (`.xlsm`): `Archivo > Guardar como`, tipo "Libro de Excel
   habilitado para macros".
2. Abrí el editor de VBA: `Alt+F11` (o `Desarrollador > Visual Basic`; si no
   ves la pestaña Desarrollador, activala en `Archivo > Opciones > Personalizar
   cinta de opciones` y tildá "Programador").
3. Click derecho sobre el nombre del libro (`VBAProject (Plantilla_Reportes_POB.xlsm)`)
   en el panel izquierdo → **`Importar archivo...`** → elegí `vba/Instalar_Reportes.bas`.
   **Importante: usá "Importar archivo", NO crees un módulo en blanco y
   pegues el texto adentro.** El archivo empieza con una línea
   (`Attribute VB_Name = "Instalar_Reportes"`) que VBA sólo interpreta bien
   al importar el archivo; si la pegás como texto en un módulo ya creado,
   tira "Error de compilación: Error de sintaxis" al querer ejecutar la
   macro. Si ya te pasó eso: borrá ese módulo (click derecho → `Quitar
   Módulo1` → "No" a exportar) y volvé a hacerlo con `Importar archivo...`.
4. Cerrá el editor de VBA, volvé a Excel.
5. Con el libro activo, corré la macro: `Alt+F8`, elegí `Instalar_Reportes`,
   `Ejecutar`.
6. La primera vez que se conecte a OneDrive, Excel te va a pedir iniciar
   sesión: usá tu cuenta corporativa. Marcá "conectar siempre con esta
   cuenta" para no tener que repetirlo en cada refresco.
7. La macro va a:
   - Crear las 12 consultas Power Query.
   - Cargar `POB_Diario`, `POB_Horario`, `POB_Franjas`, `Comedor` y
     `Control_Calidad` como tablas en sus hojas, y el detalle en la hoja
     oculta `Estadias`.
   - Armar un dashboard básico en la hoja `Dashboard` (tablas dinámicas,
     segmentaciones de Segmento y Empresa, y un gráfico por reporte).
8. Al terminar te avisa con un mensaje. Si dice que el dashboard no se pudo
   armar del todo, no pasa nada: las consultas y las tablas ya están
   cargadas igual (lo que suele fallar es sólo la parte visual, por
   diferencias entre versiones de Excel). Segui con la sección **3. Armar
   el dashboard a mano** más abajo para completar esa parte.
9. **Paso manual obligatorio, incluso con la macro**: click derecho sobre la
   segmentación "Segmento" → `Configuración de segmentación de datos` →
   tildá **"Selección única"**. Excel no da una forma confiable de
   automatizar esto por VBA para segmentaciones normales (no basadas en
   modelo de datos), así que la macro no lo hace; sin este paso, se puede
   seleccionar `AP` y `LLY` a la vez y mezclar esas filas con `TOTAL` sin
   darte cuenta.
10. Podés volver a correr la macro las veces que quieras (por ejemplo si
    cambiaste algo en Config): reemplaza lo que ya existía, no duplica nada
    (el paso 9 sí hay que repetirlo si la segmentación se recreó).

## Camino B — Manual (sin macro)

### B.1 Crear las consultas

Repetí esto para cada archivo de `pq/`, **en este orden**: `01_Parametros`,
`02_Origen`, `03_Limpieza`, `04_Empresas`, `05_Validos`, `06_Estadias`,
`07_fnExpandirGrupos`, `08_POB_Diario`, `09_POB_Horario`, `10_POB_Franjas`,
`11_Comedor`, `12_ControlCalidad`:

1. `Datos > Obtener datos > Desde otras fuentes > Consulta en blanco`.
2. En el panel de consultas (a la derecha), click derecho sobre la consulta
   nueva → `Editor avanzado`.
3. Borrá todo lo que haya y pegá el contenido completo del archivo `.pq`
   correspondiente (abrilo con el Bloc de notas, copiá todo).
4. `Listo`. Si Power Query pide configurar el nivel de privacidad de la
   fuente de datos, elegí "Organizacional" para el sitio de OneDrive.
5. Click derecho sobre la consulta en el panel → `Cambiar nombre` → poné
   exactamente el mismo nombre del archivo (por ejemplo `01_Parametros`,
   sin la extensión). Esto es importante: las consultas siguientes hacen
   referencia a las anteriores por ese nombre exacto.
6. Para las consultas `01` a `07` (todas menos las últimas 5): dejalas como
   "Sólo conexión" (no hace falta cargarlas a ninguna hoja). Si Power Query
   las cargó solas a una hoja nueva, click derecho sobre la consulta →
   `Cargar en... > Sólo crear conexión`.

### B.2 Cargar los reportes como tabla

Para `08_POB_Diario`, `09_POB_Horario`, `10_POB_Franjas`, `11_Comedor` y
`12_ControlCalidad`:

1. Click derecho sobre la consulta → `Cargar en...`
2. Elegí `Tabla` → `Hoja de cálculo existente` → señalá la celda A1 de la
   hoja correspondiente (`POB_Diario`, `POB_Horario`, `POB_Franjas`,
   `Comedor`, `Control_Calidad` — ya vienen creadas y vacías en la
   plantilla).

Para `06_Estadias` (el detalle de auditoría): cargalo igual como tabla en la
hoja oculta `Estadias` (tenés que desocultarla primero: click derecho sobre
cualquier pestaña de hoja → `Mostrar` → `Estadias`; podés volver a ocultarla
después con click derecho → `Ocultar`).

### B.3 Armar el dashboard a mano

En la hoja `Dashboard`:

1. **Tablas dinámicas**: por cada tabla cargada (`POB_Diario`, `POB_Horario`,
   `POB_Franjas`, `Comedor`), seleccioná la tabla → `Insertar > Tabla
   dinámica` → `Hoja de cálculo existente` → elegí una celda libre en
   `Dashboard`. Armá cada una así:
   - Filas: `Fecha` (para POB_Diario), `Hora` (POB_Horario), `Franja`
     (POB_Franjas) o `Turno` (Comedor).
   - Valores: `Presentes_en_el_dia` y `Maximo_simultaneo` (POB_Diario) o
     `POB`/`Personas_unicas` según corresponda.
   - Filtros: arrastrá `Segmento` y `Empresa` al área de Filtros.
   - En el filtro de Segmento, dejalo en `TOTAL`; en Empresa, en `TODAS`.
2. **Segmentaciones**: con cualquiera de las tablas dinámicas seleccionada,
   `Analizar tabla dinámica > Insertar Segmentación de datos` → tildá
   `Segmento` y `Empresa` (podés hacerlo en dos pasos, una por campo).
   - Click derecho sobre la segmentación de `Segmento` → `Configuración de
     segmentación de datos` → tildá "Selección única" (para no mezclar
     filas `TOTAL` con `AP`/`LLY`).
   - Con cada segmentación seleccionada, `Conexiones de informes` (botón en
     la cinta, pestaña Segmentación) → tildá las 4 tablas dinámicas para que
     una sola segmentación controle todas.
3. **Escala de tiempo**: seleccioná la tabla dinámica de `POB_Diario` →
   `Insertar > Escala de tiempo` → `Fecha`. Conectala a las demás dinámicas
   igual que las segmentaciones (`Conexiones de informes`).
4. **Gráficos**: con cada tabla dinámica seleccionada, `Insertar > Gráfico
   dinámico` y elegí el tipo:
   - POB_Diario → Líneas (Presentes y Máximo).
   - POB_Horario → Líneas.
   - POB_Franjas → Columnas.
   - Comedor → Columnas apiladas.
   - Para el "top 10 empresas por POB" (gráfico de barras): armá una quinta
     tabla dinámica sobre `POB_Diario`, con `Empresa` en Filas y
     `Presentes_en_el_dia` en Valores, Segmento=TOTAL en Filtro, ordená de
     mayor a menor y aplicá un filtro de "Top 10" sobre el campo Empresa.

## 2. Actualizar los datos

Cada vez que subas archivos `POB*.xlsx` nuevos a la carpeta de OneDrive:

`Datos > Actualizar todo` (o `Ctrl+Alt+F5`).

Con ~90.000 filas por mes, el refresco completo debería tardar unos minutos
(las consultas usan `Table.Buffer` y agrupaciones para no recorrer los datos
más de lo necesario, pero el paso de "armar estadías" y las fotos por hora
son el costo principal). Si notás que tarda demasiado, revisá primero que
no haya quedado ningún archivo viejo/corrupto en la carpeta de OneDrive.

## 3. Agregar lectores nuevos

Si se instala un lector nuevo en planta:

1. Andá a `Config > tLectores`.
2. Agregá una fila con el nombre EXACTO del lector (tal como aparece en la
   columna "Reader Description" de los archivos de origen) y su Tipo.
3. `Datos > Actualizar todo`.

Hasta que lo agregues, ese lector se trata como `INTERNO` automáticamente y
aparece listado en `Control_Calidad` bajo "Lectores sin clasificar", así que
no se te va a pasar por alto.

## 4. Agregar lectores (usuarios) nuevos del reporte

Para que otra persona pueda abrir y actualizar este Excel:

1. Compartí el archivo `.xlsx`/`.xlsm` desde OneDrive/SharePoint con permiso
   de edición.
2. Esa persona necesita, además, permiso de **lectura** sobre la carpeta de
   origen de las fichadas en OneDrive (la de `Ruta_Carpeta`), aunque no
   edite el Excel: Power Query se conecta con las credenciales de quien
   tiene el Excel abierto.
3. La primera vez que esa persona actualice, Excel le va a pedir loguearse
   con su cuenta corporativa para autorizar la conexión a OneDrive (igual
   que el paso 7 del Camino A).

## 5. Solución de problemas frecuentes

| Error | Causa | Solución |
|---|---|---|
| `Error de compilación: Error de sintaxis` al ejecutar la macro (VBA) | Pegaste el código del `.bas` como texto en un módulo en blanco en vez de importarlo. La línea `Attribute VB_Name = ...` sólo es válida al importar. | Borrá el módulo, usá `Importar archivo...` (ver Camino A, paso 3). |
| `DataFormat.Error: The input URL is invalid. Please provide a URL to the root of a SharePoint site...` | `URL_Sitio_OneDrive` tiene la URL completa de la vista de la carpeta (con `?FolderCTID=...&id=...`), no la raíz del sitio. | Cortá la URL como se explica en la sección 0, puntos 3 y 4. |
| `Formula.Firewall: ... references other queries or steps, so it may not directly access a data source` | Power Query bloquea por defecto que una consulta use un valor de otra consulta (acá, `01_Parametros`) para conectarse a un origen de datos. | Sección 0, punto 5: ignorar los niveles de privacidad para este libro. |
| Un lector nuevo no se clasifica como esperás | Todavía no está en `tLectores`, o el nombre no coincide exactamente (revisá espacios dobles). | Sección 3 de esta guía. |
| Las cifras no coinciden con `validacion/resultados_esperados.xlsx` | Puede ser una diferencia de criterio ya documentada, o un problema de carga. | `validacion/COMO_VALIDAR.md`, sección 4. |
