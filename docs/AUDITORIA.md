# Preparación para repositorio público

Fecha: 20 de octubre de 2026. Alcance: todos los archivos originales de la carpeta,
incluidos los dos contenedores Excel, su XML y todos sus módulos VBA extraídos.
La auditoría inicial se realizó antes de modificar archivos del proyecto.

## 1. Problema real

Reorganización repetitiva de columnas en archivos Excel de nómina. Las macros
aplican dos secuencias concretas por carpeta y PartFed asigna cuentas según banco.
No hay evidencia de un sistema destino ni de una especificación oficial del formato.

## 2. Funcionalidades verificadas en código

Selección de carpeta, recorrido de `.xlsx`, cortes/inserciones/eliminaciones,
encabezados, búsqueda bancaria de PartFed, guardado sobre la entrada y registro
de nombres de archivos en `Hoja1`. Se verificó su presencia en código, no su ejecución.

## 3. Auditoría de seguridad

| Severidad | Hallazgo inicial | Resultado |
| --- | --- | --- |
| CRÍTICO | Ninguno identificado | No se encontraron claves privadas ni credenciales de acceso |
| ALTO | Cinco cuentas bancarias en `PartFed.bas` y VBA de `MacroPartFed.xlsm` | Retiradas del módulo público; original binario excluido |
| MEDIO | Metadatos de autor, ruta local y nombres de archivos procesados en ambos `.xlsm` | Ambos libros se conservan únicamente locales e ignorados |
| MEDIO | Comillas inválidas en ambos bucles; sin recuperación del estado de Excel ante errores | Comillas ASCII y manejo de error/restauración añadidos |
| MEDIO | Guardado destructivo y búsquedas sin control de coincidencia | Se documentan copias obligatorias; se detiene PartFed ante bancos sin coincidencia |
| BAJO | Sin exclusiones ni documentación de instalación/configuración | Añadidos `.gitignore`, plantilla y documentación |
| INFORMATIVO | No existía repositorio Git | Inicializado localmente, sin commits, remoto ni archivos agregados al índice |
| FALSO POSITIVO | oletools marcó `Open`, cadenas hexadecimales y una cadena Base64 | Apertura de libros y texto del VBA revisados; no identificados como credenciales ni eliminados por esa coincidencia |

Las cuentas de banco son datos financieros privados, no credenciales de
autenticación. El código maneja nombres/números de empleado, pero no se encontraron
registros de empleados en los archivos entregados. No se publican valores ni autores.

## 4. Cambios realizados

- Módulos ASCII importables con nombres distintos y `Option Explicit`.
- Cinco cuentas sustituidas por variables del proceso, con validación de dígitos.
- Cuentas de referencia como texto para preservar ceros iniciales.
- Corrección de `Do While`, declaración de `f` y contadores de tipo `Long`.
- Selección explícita de la primera hoja de la entrada, en vez de su última hoja activa.
- Apertura sin actualizar vínculos, con eventos y macros entrantes deshabilitados.
- Rechazo de archivos ya abiertos, de solo lectura y, en PartFed, con hoja `Bancos` previa.
- Omisión de archivos `~$`; cierre del archivo actual sin guardar ante fallos.
- Restauración de alertas, eventos, actualización de pantalla y seguridad de automatización.
- Búsqueda bancaria por fila, cálculo antes de pegar valores y autollenado solo
  cuando hay más de una fila de datos; detención si existe `#N/A`.
- `.env.example`, lanzador PowerShell, comprobador estático, README y lista de pruebas.
- `.gitattributes` con CRLF para VBA y PowerShell al extraer desde Git.

Se conservaron 30 instrucciones de columnas/encabezados de FONE y 116 de PartFed,
comparándolas en orden con el código de los libros originales. Esto no comprueba
la corrección contable ni sustituye una prueba en Excel.

## 5. Secretos eliminados

No se identificaron secretos de acceso. Se retiró esta información privada del
conjunto publicable:

| Severidad | Archivo | Tipo | Corrección |
| --- | --- | --- | --- |
| ALTO | `PartFed.bas` | Cinco cuentas bancarias | Variables de entorno sin valores predeterminados |
| ALTO | `MacroPartFed.xlsm` | Copia incrustada de las cuentas | Libro ignorado; no usar para distribución |
| MEDIO | Ambos `.xlsm` | Metadatos, ruta y nombres de archivos | Originales locales excluidos; crear un contenedor vacío |

Los originales binarios no se han saneado ni actualizado. Siguen conteniendo
información privada y código anterior; no deben adjuntarse ni forzarse al índice.

## 6. Variables de entorno

| Variable | Origen | Nueva ubicación |
| --- | --- | --- |
| `PARTFED_CUENTA_SERFIN` | Literal de cuenta en PartFed | Entorno de Excel, cargado desde `.env` local |
| `PARTFED_CUENTA_BANCOMER` | Literal de cuenta en PartFed | Entorno de Excel, cargado desde `.env` local |
| `PARTFED_CUENTA_BANAMEX` | Literal de cuenta en PartFed | Entorno de Excel, cargado desde `.env` local |
| `PARTFED_CUENTA_BANORTE` | Literal de cuenta en PartFed | Entorno de Excel, cargado desde `.env` local |
| `PARTFED_CUENTA_HSBC` | Literal de cuenta en PartFed | Entorno de Excel, cargado desde `.env` local |

## 7. `.gitignore`

Excluye configuración privada, claves/certificados, nombres habituales de archivos
de credenciales, todos los contenedores/datos Excel, CSV/TSV, backups, bases locales,
archivos comprimidos, ejecutables y temporales. Mantiene publicable `.env.example`
y los módulos fuente. Las exclusiones de Excel/CSV también contemplan mayúsculas.

## 8. `.env.example`

Contiene únicamente las cinco variables anteriores con `your_account_here`.
No existe `.env` real en la entrega. VBA usa `Environ$`; el lanzador carga el archivo
local, sin ejecutarlo ni imprimir valores. FONE no necesita configuración.

## 9. Archivos retirados del seguimiento

Ninguno: antes de esta preparación no existía `.git` y no había archivos rastreados.
El índice sigue vacío; ambos `.xlsm` quedan ignorados desde el inicio.

## 10. Archivos grandes

| Archivo | Tamaño original | Acción | Historial |
| --- | --- | --- | --- |
| `MacroPartFed.xlsm` | 23,4 KiB | Conservar local, ignorar por privacidad | No existía |
| `MacroFone.xlsm` | 18,6 KiB | Conservar local, ignorar por privacidad | No existía |

No se encontraron archivos cercanos a 100 MiB ni necesidad de Git LFS.

## 11. Datos personales

Se excluyeron del conjunto publicable cuentas, identidades de autor, ruta de
usuario y nombres de archivos guardados en los libros. No se inventaron sustitutos
personales ni se copiaron valores a documentación. Los originales se preservaron
localmente para evitar pérdida de información del usuario.

## 12. Historial Git

No había historial para auditar. El repositorio inicializado tiene cero commits,
cero archivos rastreados y ningún remoto configurado. No hay objetos históricos
que limpiar ni credenciales históricas identificadas para rotación. No se hizo
commit, push, publicación ni cambio de servicios externos.

## 13. Docker

No existe; no se añadieron contenedores ni configuración de producción artificial.

## 14. GitHub Actions

No existe; no se añadieron workflows ni se requieren GitHub Secrets.

## 15. Frontend

No hay frontend web ni variables públicas de frameworks. La interfaz es Excel.
Los archivos de salida de PartFed contienen cuentas y deben tratarse como privados.

## 16. README

Documenta el trabajo manual sustituido, usuarios que preparan nómina, operaciones
verificadas, importación de módulos en un libro nuevo, configuración, ejecución y
limitaciones. No inventa URL de repositorio, versión mínima de Excel ni licencia.
Incluye flujo de clonación interactivo para cuando exista una URL.

## 17. Validación

| Comprobación | Resultado |
| --- | --- |
| Inventario completo de cuatro archivos originales | Completado |
| XML, relaciones y metadatos de ambos Excel | Completado; sin relaciones externas |
| Extracción de todos los módulos con oletools 0.60.2 | Completada; módulos de hoja/libro sin procedimientos adicionales |
| Revisión manual de fuentes y VBA incrustado | Completada; sin credenciales de acceso identificadas |
| Comparación de secuencias de columnas/encabezados | Coinciden con los originales |
| Ausencia de valores privados originales en archivos publicables | Comprobada sin imprimir valores |
| Comprobador local de publicación | Ejecutado sin hallazgos en el conjunto publicable |
| Pruebas aisladas del comprobador | Detectó cuatro exposiciones sintéticas sin imprimir valores, rechazó una plantilla inválida y aceptó el conjunto limpio |
| `.env.example` y exclusiones Git | Comprobadas |
| Enlaces locales de la documentación | Todas las rutas enlazadas existen |
| Historial y tamaño | Cero commits y ningún archivo mayor de 100 MiB |
| Compilación VBA y ejecución de macros | Pendientes: no hay Excel para Windows en este entorno |
| Lanzador PowerShell | Revisado estáticamente; no ejecutado, PowerShell no disponible |
| Tests/build/lint existentes | No existían; no hay build ni linter VBA configurado |

oletools se descargó exclusivamente como herramienta temporal fuera del proyecto.
No se añadió como dependencia de ejecución. No se utilizaron Gitleaks ni TruffleHog.

## 18. Estado Git final

Diez archivos nuevos sin seguimiento: los dos módulos, README, `.gitignore`,
`.gitattributes`, `.env.example`, dos documentos y dos scripts. No hay archivos
staged ni commits. Los dos libros originales permanecen ignorados. Los módulos
fueron modificados en disco, pero Git los muestra como nuevos porque no había
repositorio anterior.

## 19. Acciones manuales

1. Revisar los cambios y confirmar el esquema correcto de entrada y salida.
2. Crear un libro limpio e importar las fuentes revisadas; compilar y completar
   [las pruebas en Excel](PRUEBAS.md), incluido el lanzador PowerShell.
3. Configurar cuentas autorizadas únicamente en `.env` local; no hay credenciales
   de acceso identificadas que requieran rotación automática o manual.
4. Decidir licencia y comprobar autorización para publicar el código.
5. Revisar `git status`, el contenido que se agregue al índice y el resultado del
   comprobador antes de hacer personalmente el commit y, posteriormente, publicar.

## 20. Riesgos pendientes

- La ejecución de los módulos modificados y del lanzador no está validada en Windows.
- No se dispone de una especificación completa ni fixtures del esquema de entrada.
- Los originales `.xlsm` locales conservan datos privados y VBA anterior.
- El guardado es destructivo y no transaccional; requiere copias de trabajo.
- El repositorio aún no tiene licencia.

La revisión estática no garantiza ausencia absoluta de secretos o defectos. La
aceptación funcional requiere intervención del responsable en Excel para Windows.
