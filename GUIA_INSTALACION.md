# Guía de instalación — Reportes POB y Comedor

Esta guía asume Excel para Microsoft 365 de 64 bits en una PC donde no podés
instalar software. No hace falta instalar nada: todo se hace con Power
Query, que ya viene incluido en Excel (pestaña **Datos**).

## 0. Antes de empezar: subir los archivos de origen a OneDrive

1. En tu OneDrive for Business corporativo, creá una carpeta para los
   archivos de fichadas (por ejemplo `Documents/POB/`).
2. Subí ahí los archivos `POB*.xlsx` (todos los que tengas acumulados). El
   nombre tiene que empezar con `POB` y terminar en `.xlsx`.
3. Anotá la **URL de tu sitio de OneDrive**. Se obtiene así:
   - Abrí OneDrive en el navegador (onedrive.com, con tu cuenta corporativa).
   - Andá a la carpeta donde subiste los archivos.
   - Copiá la URL del navegador hasta la parte `.../personal/tuusuario/` (sin
     lo que sigue después). Por ejemplo:
     `https://tuempresa-my.sharepoint.com/personal/jperez_tuempresa_com/`
   - La parte que sigue (por ejemplo `/Documents/POB/`) es la **Ruta de
     carpeta**.

Vas a necesitar estos dos datos para completar la hoja Config.

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
3. Click derecho sobre el nombre del libro en el panel izquierdo →
   `Insertar > Módulo`.
4. Abrí `vba/Instalar_Reportes.bas` con el Bloc de notas, copiá todo el
   contenido y pegalo en el módulo nuevo (o `Archivo > Importar archivo...`
   dentro del editor de VBA y elegí el `.bas` directamente).
5. Cerrá el editor de VBA, volvé a Excel.
6. Con el libro activo, corré la macro: `Alt+F8`, elegí `Instalar_Reportes`,
   `Ejecutar`.
7. La primera vez que se conecte a OneDrive, Excel te va a pedir iniciar
   sesión: usá tu cuenta corporativa. Marcá "conectar siempre con esta
   cuenta" para no tener que repetirlo en cada refresco.
8. La macro va a:
   - Crear las 12 consultas Power Query.
   - Cargar `POB_Diario`, `POB_Horario`, `POB_Franjas`, `Comedor` y
     `Control_Calidad` como tablas en sus hojas, y el detalle en la hoja
     oculta `Estadias`.
   - Armar un dashboard básico en la hoja `Dashboard` (tablas dinámicas,
     segmentaciones de Segmento y Empresa, y un gráfico por reporte).
9. Al terminar te avisa con un mensaje. Si dice que el dashboard no se pudo
   armar del todo, no pasa nada: las consultas y las tablas ya están
   cargadas igual (lo que suele fallar es sólo la parte visual, por
   diferencias entre versiones de Excel). Segui con la sección **3. Armar
   el dashboard a mano** más abajo para completar esa parte.
10. Podés volver a correr la macro las veces que quieras (por ejemplo si
    cambiaste algo en Config): reemplaza lo que ya existía, no duplica nada.

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
