# Reportes POB y Comedor desde fichadas RFID

Modelo de reportes de personal a bordo (POB) y uso de comedor para Aguada
Pichana (AP) y Loma Las Yeguas (LLY), armado en Power Query (Excel) a partir
de las fichadas de control de acceso, con una implementación de referencia
en Python para validar los números antes de confiar en el Excel.

## Qué hay en este repositorio

| Carpeta/archivo | Contenido |
|---|---|
| `pq/*.pq` | Las 12 consultas M, en el orden en que hay que crearlas. |
| `Plantilla_Reportes_POB.xlsx` | Libro base con la hoja Config (tablas con nombre) y las hojas de salida vacías. |
| `vba/Instalar_Reportes.bas` | Macro opcional que instala todo automáticamente. |
| `GUIA_INSTALACION.md` | Paso a paso para dejarlo andando en tu Excel corporativo. |
| `validacion/` | Implementación de referencia en Python (pandas) + pruebas + resultados esperados. |
| `validacion/COMO_VALIDAR.md` | Qué cifras comparar entre tu Excel y la referencia. |
| `generar_plantilla.py` / `generar_macro.py` | Scripts que generan la plantilla y la macro a partir de `validacion/config.py` y `pq/*.pq` (por si hay que regenerarlos después de un cambio). |
| `data/POB.xlsx` | Muestra real usada para diseñar y probar todo esto (88.008 fichadas, 1–25/9/2026). |

## Cómo funciona, en criollo

1. **Origen**: Power Query lee todos los `POB*.xlsx` de una carpeta de OneDrive
   for Business (conexión directa, sin descargar nada a mano) y apila la hoja
   "04 - Fichadas entre fechas" de cada uno.
2. **Limpieza**: saca duplicados, arma la fecha+hora real de cada fichada,
   limpia el nombre del lector, y a cada persona le asigna un identificador
   (`IGG`, o `CRED_<credencial>` si no tiene IGG) y su última empresa conocida.
3. **Clasificación de lectores**: cada lector es INGRESO, EGRESO, COMEDOR o
   INTERNO (tabla `tLectores`, editable). Un lector nuevo que el usuario
   todavía no clasificó se trata como INTERNO y queda listado en Control de
   Calidad.
4. **Estadías**: una máquina de estados recorre las fichadas válidas de cada
   persona en orden y arma "estadías" (Segmento, Inicio, Fin). De ahí salen
   todos los reportes de POB, contando siempre personas distintas (nunca
   sumando filas).
5. **Comedor**: cuenta personas distintas con marca en un molinete de
   comedor dentro de cada turno (`tTurnosComedor`), extremos inclusivos.

Todo lo editable (URL de OneDrive, horas de foto, franjas, turnos, eventos
válidos, clasificación de lectores) vive en la hoja **Config**, sin tocar el
código M.

## Supuestos y decisiones tomadas

El prompt original dejaba algunos puntos abiertos. Estas son las decisiones
que tomé, con su razón, para que las puedas revisar:

1. **Empresa de la persona**: se calcula sobre TODAS las fichadas (válidas e
   inválidas), no sólo las que otorgaron acceso. Razón: el dato de empresa
   viaja con la credencial incluso en un rechazo (badge inactivo, etc.), y
   descartar esas filas sólo perdería información sin necesidad.
2. **INGRESO estando ya adentro en otro segmento**: el prompt dice
   textualmente "sigue adentro y se actualiza el segmento" (sin el "se
   mueve: sale... y queda..." que sí usa para INTERNO/COMEDOR). Para que los
   reportes por segmento sean consistentes, lo implementé igual que el caso
   INTERNO: cierra la estadía del segmento viejo y abre una nueva en el
   nuevo, en ese mismo minuto. Si preferís otro criterio, es un cambio
   acotado en `validacion/logica_pob.py::construir_estadias` y en
   `pq/06_Estadias.pq`.
3. **Regla de fin de mes "en vivo"**: se evalúa contra el momento en que se
   corre el cálculo (ahora), no contra la última fichada cargada. Así, el
   ejemplo del prompt ("entró el 10, no marcó nada más → se descarta el
   día 1") funciona apenas se cruza el 1° del mes siguiente, exista o no ya
   el archivo de ese mes cargado — igual que se comportaría el sistema real
   si lo estuvieran mirando en vivo. Consecuencia práctica: si corrés el
   Excel el mismo día 1 muy temprano y todavía no cargaste el archivo del
   mes que empieza, la persona igual se descarta (el archivo, cuando
   llegue, no va a tener marcas de ella en ese hueco de todos modos).
4. **Segmento del comedor**: sale de la columna `Segmento` de la propia
   fichada (ya viene 1 a 1 con AP/LLY para los lectores de comedor en la
   muestra), no de una tabla de mapeo aparte.
5. **Filas en cero no se generan**: si una combinación Fecha/Hora/Segmento/
   Empresa no tuvo a nadie presente, esa fila simplemente no aparece (ni en
   Python ni en M). Para una tabla dinámica esto no es un problema.
6. **Sort ante empates de minuto**: como la resolución es de a minuto, dos
   fichadas de la misma persona en el mismo minuto se procesan en el orden
   en que aparecen en el archivo de origen (no hay manera de saber cuál fue
   "antes" con ese nivel de precisión).
7. **Clasificación de `tLectores` corregida contra datos reales**: la
   precarga original (basada sólo en la muestra) clasificaba como `INTERNO`
   varios lectores que en la operación real son accesos válidos:
   - `AP - PlantaAcc Puerta 4/6/9` y `AP - PlantaAccPeato-`: hay personal
     (por ejemplo, de la empresa Víctor Contreras) que entra y sale
     **exclusivamente** por estos accesos secundarios, no por la portería
     principal. Ahora son INGRESO/EGRESO reales.
   - Toda la familia `LLY - Proc - ...` (KM 0, Porton - Mpl, Porton 1 Zona
     GEKs, Porton 2 Zona Tanque, Puerta 1 Acceso Peatonal): mismo caso para
     Loma Las Yeguas.
   - Las puertas de emergencia (`MP_...`): en el uso real sólo se marcan al
     evacuar, nunca para entrar. Tratarlas como `INTERNO` dejaba a la gente
     contada como "presente" indefinidamente después de una evacuación real
     (se verificó un caso concreto: un grupo marcó varias puertas de
     emergencia de LLY una noche y, sin esta corrección, seguían apareciendo
     presentes 9 días después). Ahora se tratan como `EGRESO`.

   Esta corrección se validó contra un archivo de producción real,
   comparando día por día y yacimiento por yacimiento contra una foto de
   las 13hs conocida de antemano: los 14 puntos de control (7 días × 2
   yacimientos) pasaron de un error de 26%-75% a un error de -7%/+10%
   (la mayoría dentro de ±5%). El detalle de esa validación puntual no
   quedó en el repo (se hizo sobre un archivo de producción, no sobre la
   muestra), pero el criterio adoptado sí: ver `validacion/config.py`.
   **Si tu operación tiene otros lectores "secundarios" parecidos que no
   están en esta lista, agregalos a mano en `tLectores` de la hoja Config.**

## Limitaciones conocidas

- **Arranque en frío**: alguien que ya estaba adentro antes de la primera
  fichada disponible en los archivos cargados no se detecta hasta que vuelve
  a marcar. No hay forma de saberlo con los datos disponibles.
- **Identidad por credencial**: si a una misma persona alguna fichada le
  vino sin IGG (se identifica como `CRED_<credencial>`) y otras sí lo
  tienen, esa persona puede aparecer bajo dos identificadores distintos. En
  la muestra esto pasa con 14 credenciales. Se ve reflejado indirectamente
  en Control de Calidad ("Marcas sin IGG identificadas por credencial" y
  "Personas con más de una credencial"), pero no se corrige automáticamente
  porque no hay una regla no ambigua para unificarlas.
- **Credencial reasignada a otra persona**: si una credencial vieja se le
  da de baja a alguien y se reasigna a otra persona más adelante, y ambas
  fichadas vinieron sin IGG, el sistema las trata como la misma persona
  (mismo `CRED_<credencial>`). En la muestra hay 37 credenciales asociadas a
  más de un IGG a lo largo del período, lo que sugiere que este escenario
  puede darse.
- **Estadías de duración cero**: por la resolución de minuto, dos fichadas
  de la misma persona en el mismo minuto (p. ej. un movimiento rápido entre
  dos lectores) pueden generar una estadía con Inicio = Fin. No afecta
  ningún cálculo (un intervalo vacío no cuenta en ninguna foto), pero podés
  verlas en la hoja oculta `Estadias` si audita el detalle.
- **Nombre de empresa "TOTAL AUSTRAL" / "TOTAL"**: hay una empresa real
  llamada "TOTAL" en los datos (además de "TOTAL AUSTRAL"). No colisiona con
  la etiqueta "TOTAL" que usan los reportes para el segmento combinado
  (son columnas distintas: Segmento vs. Empresa), pero puede confundir a
  simple vista en una tabla dinámica.
- **Rendimiento del script Python**: `validacion/referencia.py` no está
  optimizado para velocidad (tarda cerca de 1:30 min con el mes de muestra);
  el objetivo de "menos de 3 minutos" del prompt aplica a la consulta M en
  Excel (que sí usa `Table.Buffer` y agrupaciones vectorizadas), no a este
  script de referencia.

## Empezar

1. Seguí `GUIA_INSTALACION.md` para dejar el Excel funcionando (con o sin
   la macro).
2. Corré `python3 validacion/referencia.py` para generar
   `validacion/resultados_esperados.xlsx` a partir de `data/POB.xlsx`.
3. Seguí `validacion/COMO_VALIDAR.md` para comparar tu Excel contra esos
   resultados.
