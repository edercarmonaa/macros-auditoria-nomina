# macros-auditoria-nomina

Automatizan la reorganización por lotes de archivos Excel de nómina: mueven y
eliminan columnas, ajustan encabezados y, en PartFed, asignan cuentas por banco.

## El problema

Preparar varios archivos con una misma disposición de columnas exige repetir
cortes, inserciones y eliminaciones en Excel. En archivos de nómina, un movimiento
incorrecto puede dejar datos bajo el encabezado equivocado y obliga a revisar cada
libro manualmente. Estas macros codifican dos secuencias concretas, `FONE` y `PartFed`.

Los nombres de procedimientos y campos permiten identificar ese uso. El código
no especifica el sistema de origen, la institución usuaria ni un formato oficial
de intercambio; no se presupone compatibilidad con cualquier exportación de nómina.

## La solución

La persona elige una carpeta y ejecuta la macro correspondiente. Se abre cada
`.xlsx`, se transforma su primera hoja y se guarda el archivo. Los nombres de los
archivos completados quedan en `Hoja1` del libro que contiene las macros.

**El procesamiento sobrescribe las entradas. Trabajar siempre con copias.** Los
archivos procesados antes de un error ya pueden haberse guardado. Repetir una macro
sobre su propia salida vuelve a mover y eliminar columnas.

## Funcionalidades principales

- Selección de carpeta y recorrido de archivos `.xlsx`, sin subcarpetas.
- `FONE`: eliminación de C, movimientos de bloques y columnas, encabezado `IMP_DEP`
  y eliminación de O:LE en la disposición resultante.
- `PartFed`: movimientos e inserciones, encabezados de empleado, fechas y periodos,
  y búsqueda de la cuenta de banco en Z a partir del banco en AA.
- Configuración local de cinco cuentas para PartFed, sin cuentas incrustadas.
- Registro local de nombres de archivos completados.
- Omisión de archivos de bloqueo `~$`, rechazo de libros ya abiertos o de solo
  lectura y restauración del estado de Excel tras errores.
- Apertura sin actualización automática de vínculos y con eventos y macros de los
  archivos abiertos deshabilitados durante el proceso.

Los movimientos son secuenciales: cada posición se refiere a la hoja después de
los cambios anteriores. La secuencia exacta está en los módulos `.bas`.

## ¿Qué mejora este proyecto?

Reduce la repetición manual de una transformación conocida y aplica el mismo orden
de operaciones a una carpeta. La configuración externa evita distribuir cuentas
reales dentro del código. Esto no sustituye la revisión de resultados contables.

## ¿Para quién está pensado?

Para personas que preparan archivos de nómina en Excel y conocen las disposiciones
que esperan `FONE` y `PartFed`. No hay usuarios, roles ni autenticación de aplicación.

## Tecnologías utilizadas

- VBA y el modelo de objetos de Microsoft Excel.
- Libros locales `.xlsm` como contenedores y archivos `.xlsx` como entradas/salidas.
- PowerShell para cargar la configuración de PartFed al iniciar Excel.
- Git y un comprobador opcional en Python 3 con biblioteca estándar.

No hay servidor, frontend web, base de datos, dependencias de paquetes de la
aplicación, Docker ni GitHub Actions.

## Requisitos

- Excel de escritorio para Windows con VBA y `Application.FileDialog` disponibles.
  El proyecto no fija una versión mínima ni incluye una matriz de versiones probadas.
- Permiso para ejecutar las macros revisadas conforme a las políticas de tu equipo.
  No habilitar globalmente todas las macros.
- PowerShell para el lanzador de PartFed y Git para clonar/revisar el repositorio.
- Python 3 solo para la comprobación estática; no se necesita para ejecutar VBA.
- Archivos de entrada con la disposición que espera cada módulo.

Excel web no ejecuta estos módulos. macOS no está soportado por esta preparación:
el código usa rutas de Windows y `Environ$` para la configuración.

## Instalación

No se ha asignado una URL pública. Cuando exista, clonar desde PowerShell:

```powershell
$URL_REPOSITORIO = Read-Host 'URL del repositorio que deseas clonar'
git clone $URL_REPOSITORIO macros-auditoria-nomina
cd macros-auditoria-nomina
New-Item -ItemType Directory -Force local
```

Si ya tienes esta carpeta, omite la clonación y trabaja desde ella.

1. En Excel, crear un libro **nuevo y vacío**, con una hoja llamada exactamente
   `Hoja1`. Guardarlo como `local\Macros.xlsm`, en formato habilitado para macros.
2. Abrir el editor de VBA con `Alt+F11`. Usar **Archivo > Importar archivo** para
   importar `fone.bas` y `PartFed.bas`. Los módulos tienen nombres distintos y pueden
   convivir en el mismo libro. No pegar las líneas `Attribute` en la ventana de código.
3. Seleccionar **Depurar > Compilar VBAProject** y resolver cualquier error antes
   de ejecutar. Guardar el libro.
4. No importar versiones antiguas ni utilizar como distribución los `.xlsm`
   originales que pudieran existir localmente. Están ignorados y conservan contenido
   privado y VBA anterior; los módulos de este repositorio son la fuente revisada.

No hay instalación de paquetes, migraciones ni proceso de build automatizado.

## Configuración

`FONE` no requiere variables. `PartFed` requiere las siguientes:

| Variable | Uso |
| --- | --- |
| `PARTFED_CUENTA_SERFIN` | Cuenta asociada a SERFIN |
| `PARTFED_CUENTA_BANCOMER` | Cuenta asociada a BANCOMER |
| `PARTFED_CUENTA_BANAMEX` | Cuenta asociada a BANAMEX |
| `PARTFED_CUENTA_BANORTE` | Cuenta asociada a BANORTE |
| `PARTFED_CUENTA_HSBC` | Cuenta asociada a HSBC |

```powershell
Copy-Item .env.example .env
notepad .env
```

Sustituir los cinco placeholders en `.env` por cuentas locales autorizadas de 10 a
18 dígitos, sin comillas, espacios ni separadores. Se preservan los ceros iniciales.
La longitud solo valida la forma; no acredita titularidad ni validez bancaria.
Los nombres de los bancos se mantienen como en el código original.

VBA **no lee `.env` automáticamente**: usa el entorno del proceso Excel. El script
`scripts/Start-PartFed.ps1` valida el archivo, configura el entorno de su proceso,
inicia Excel y restaura el entorno de PowerShell. No cambia variables permanentes.
Cierra todas las instancias de Excel antes de usarlo. El `.env` solo se interpreta
como pares `NOMBRE=valor`; no se ejecuta su contenido.

## Ejecutar el proyecto

Para `FONE`, abrir el libro nuevo `local\Macros.xlsm` directamente en Excel.

Para `PartFed`, desde la raíz del proyecto en PowerShell:

```powershell
.\scripts\Start-PartFed.ps1 -Libro .\local\Macros.xlsm
```

Excel debe estar instalado y registrado para iniciar `excel.exe`. Si la política
del equipo impide ejecutar scripts, solicitar el mecanismo autorizado al administrador;
el proyecto no cambia esa política.

En Excel, pulsar `Alt+F8`, seleccionar `FONE` o `PartFed` y ejecutar. Elegir únicamente
la carpeta de **copias de trabajo** correspondiente a esa macro.

## Uso

1. Conservar los originales fuera de la carpeta de trabajo.
2. Preparar una carpeta de copias `.xlsx` del mismo formato; cerrar esas copias.
3. Abrir el libro de macros, con configuración si se utilizará PartFed.
4. Ejecutar una macro y seleccionar la carpeta de copias.
5. Revisar las columnas, encabezados, importes, fechas y cuentas de los resultados.
   En PartFed, un banco sin correspondencia detiene el proceso sin guardar ese archivo.
6. Revisar el registro en `Hoja1` y guardar el libro de macros si se desea conservarlo.

El registro comienza en A1 y puede sobrescribir filas previas; no es una bitácora
acumulativa. Los archivos de salida de PartFed contienen las cuentas configuradas:
son datos privados y deben permanecer fuera del repositorio.

## Estructura del proyecto

```text
PartFed.bas                  Transformación PartFed y asignación bancaria
fone.bas                     Transformación FONE
.env.example                 Cinco variables, solo placeholders
.gitignore                   Exclusión de datos y configuración privada
.gitattributes               Finales de línea para módulos y scripts
scripts/Start-PartFed.ps1     Inicio de Excel con configuración local
scripts/check_publication.py Revisión estática de archivos publicables
docs/PRUEBAS.md               Verificación funcional manual
docs/AUDITORIA.md             Hallazgos y alcance de la preparación
local/                       Libros creados por el usuario, ignorados
```

No se incluyen capturas ni libros con información de nómina.

## Seguridad

No almacenar cuentas privadas, contraseñas, tokens ni claves en los módulos o en la
documentación. Mantener `.env` local y usar `.env.example` exclusivamente como plantilla.
Los libros, entradas, salidas y registros de Excel están ignorados por privacidad.
`.gitignore` no impide que alguien los agregue con `git add -f`: revisar siempre el índice.

No se detectaron credenciales de acceso en la auditoría local. Las cuentas de banco
se tratan como información financiera privada, sin clasificarlas como contraseñas.
Los originales locales siguen conteniendo datos privados y no deben compartirse.

Antes de compartir un archivo o captura, revisar datos y metadatos. No se ha definido
un canal privado de reporte de seguridad: acordarlo con el responsable antes de enviar
evidencia sensible; evitar publicar datos privados en issues.

## Pruebas

Desde la raíz, con Python 3:

```sh
python3 scripts/check_publication.py
git status --short --untracked-files=all
git ls-files
```

El comprobador no muestra los valores detectados. Revisa tanto archivos rastreados
como nuevos no ignorados. No reemplaza un escáner completo ni examina historial.

La compilación y las pruebas en Excel están descritas en [docs/PRUEBAS.md](docs/PRUEBAS.md).
No se ejecutaron en el entorno de preparación, por lo que no se afirma validación
funcional completa. No hay suite automatizada de integración, linter VBA ni build configurados.

## Estado del proyecto

Código de automatización existente en preparación inicial para publicación. La revisión
estática se completó; falta aceptación funcional de los módulos corregidos en Excel.

## Limitaciones actuales

- Transformaciones por posición, sin validación completa de encabezados o esquema.
- Operaciones destructivas: guardan sobre cada entrada y no crean copias de seguridad.
- Sin transacción para toda la carpeta ni reversión de archivos ya guardados.
- Procesamiento de la primera hoja; presupone hojas de cálculo sin protección.
- `PartFed` necesita que no exista una hoja `Bancos` en la entrada.
- Solo cinco bancos y sus nombres históricos; sin normalización de nombres.
- El rango usado puede incluir filas vacías formateadas y causar búsquedas sin coincidencia.
- Se conservan encabezados históricos como `PERIORICIDAD_PAGO` y `PERIODOD_FIN`;
  no se corrigieron sin una especificación del formato destino.
- No se valida la legalidad, exactitud contable ni compatibilidad con un sistema externo.

## Próximas mejoras

Propuestas, todavía sin implementar: documentar esquemas de entrada/salida, añadir
fixtures ficticios aprobados por el responsable del formato, validar encabezados antes
de modificar y permitir un directorio de salida separado. No son funcionalidades actuales.

## Contribuciones

Crear un fork y una rama, realizar cambios acotados y ejecutar el comprobador estático.
Para cambios VBA, compilar y completar las pruebas relevantes en Excel con datos
ficticios. Abrir un pull request explicando el cambio y las comprobaciones realizadas.
No adjuntar libros reales, `.env`, cuentas privadas ni capturas con datos personales.

## Licencia

Este proyecto todavía no incluye un archivo de licencia. El titular debe decidirla
antes de ofrecer condiciones explícitas de uso y redistribución.
