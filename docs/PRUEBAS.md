# Validación antes de publicar

La comprobación automática revisa los archivos rastreados y los no ignorados,
patrones de exposición, tamaño, exclusiones y placeholders. No compila VBA,
no inspecciona todo tipo de secreto y no valida reglas contables.

```sh
python3 scripts/check_publication.py
git status --short --untracked-files=all
git ls-files
```

En Windows puede utilizarse `python` si ese es el nombre del ejecutable instalado.

## Pruebas en Excel para Windows

Ejecutar exclusivamente con libros nuevos y datos ficticios; nunca con los originales
locales de este proyecto ni con la única copia de datos de nómina.

1. Crear el libro según el README, importar los dos módulos y seleccionar
   **Depurar > Compilar VBAProject**. No debe haber errores.
2. Crear una carpeta de trabajo con copias ficticias `.xlsx` que respeten el formato
   real de cada proceso. Marcar cada columna de origen para verificar los movimientos.
   El proyecto no incluye especificación ni fixtures completos de esos formatos.
3. Ejecutar `FONE`: confirmar eliminación de C, traslado de columnas, encabezado
   `IMP_DEP`, eliminación de O:LE en la disposición resultante y registro en `Hoja1`.
   Las posiciones se aplican secuencialmente, sobre columnas ya movidas.
4. Ejecutar `PartFed` con cinco cuentas ficticias de 10 a 18 dígitos en un `.env`
   local. Incluir una cuenta con cero inicial. Comprobar que Z contiene la cuenta
   correcta según el banco de AA, como texto; que no quedan fórmulas de búsqueda ni
   hoja temporal `Bancos`; y revisar todos los encabezados y movimientos del módulo.
5. Repetir con una fila de datos, solo encabezados y varias filas. Verificar que no
   falla el autollenado con una sola fila. El rango usado de Excel puede incluir
   filas formateadas vacías; comprobarlo en los archivos reales.
6. Probar un banco desconocido: debe cancelar sin guardar el archivo actual.
   Confirmar que los archivos completados anteriormente conservan sus cambios.
7. Probar variables ausentes o inválidas: `PartFed` debe detenerse antes de modificar
   entradas. `FONE` no necesita variables.
8. Probar cancelación del selector, carpeta vacía, archivo abierto, solo lectura,
   hoja protegida, falta de `Hoja1` y existencia previa de `Bancos` para `PartFed`.
   Confirmar mensajes sin datos privados y cierre sin guardar ante errores.
9. Crear un archivo de bloqueo `~$prueba.xlsx`: no debe procesarse.
10. Antes y después de éxito, cancelación y error, comparar `DisplayAlerts`,
    `ScreenUpdating`, `EnableEvents` y `AutomationSecurity` de Excel. Deben restaurarse.
11. Revisar los resultados en una segunda apertura, especialmente importes, fechas,
    ceros iniciales y datos fuera del rango transformado. No ejecutar dos veces
    sobre el mismo archivo: las transformaciones no son idempotentes.
12. Confirmar que `.env`, entradas, salidas y libros `.xlsm` aparecen como ignorados
    y que ningún nombre de persona o de archivo privado entró en documentación.

## Resultado de esta preparación

Se ejecutaron el análisis estático local, el análisis del VBA incrustado con
oletools y la comparación de las secuencias de movimientos de columnas contra los
libros originales. No se ejecutaron macros ni PowerShell, ni se compiló VBA:
el entorno de preparación no dispone de Excel para Windows ni de PowerShell.
La aceptación funcional queda pendiente de esta lista manual.
