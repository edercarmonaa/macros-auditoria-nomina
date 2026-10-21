param(
    [Parameter(Mandatory = $true)]
    [string]$Libro,
    [string]$Configuracion = (Join-Path $PSScriptRoot '..\.env')
)

$ErrorActionPreference = 'Stop'
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    throw 'Cierre Excel para que la nueva instancia reciba la configuracion.'
}
$bookPath = (Resolve-Path -LiteralPath $Libro).Path
if ([IO.Path]::GetExtension($bookPath) -ine '.xlsm') {
    throw 'Seleccione un libro .xlsm creado con los modulos revisados.'
}
$required = @(
    'PARTFED_CUENTA_SERFIN', 'PARTFED_CUENTA_BANCOMER',
    'PARTFED_CUENTA_BANAMEX', 'PARTFED_CUENTA_BANORTE', 'PARTFED_CUENTA_HSBC'
)
$values = @{}
foreach ($line in Get-Content -LiteralPath $Configuracion) {
    $entry = $line.Trim()
    if (!$entry -or $entry.StartsWith('#')) { continue }
    $pair = $entry.Split(@('='), 2)
    if ($pair.Count -ne 2) { throw 'Configuracion invalida: utilice NOMBRE=valor.' }
    $key = $pair[0].Trim()
    $value = $pair[1].Trim()
    if ($key -cnotin $required -or $values.ContainsKey($key)) {
        throw 'Configuracion invalida: variable desconocida o duplicada.'
    }
    if ($value -notmatch '^[0-9]{10,18}$') {
        throw 'Las cinco cuentas deben contener entre 10 y 18 digitos.'
    }
    $values[$key] = $value
}
if ($values.Count -ne $required.Count) { throw 'Faltan cuentas en la configuracion.' }

$previous = @{}
try {
    foreach ($key in $required) {
        $previous[$key] = [Environment]::GetEnvironmentVariable($key, 'Process')
        [Environment]::SetEnvironmentVariable($key, $values[$key], 'Process')
    }
    Start-Process -FilePath 'excel.exe' -ArgumentList ('"{0}"' -f $bookPath)
}
finally {
    foreach ($key in $previous.Keys) {
        [Environment]::SetEnvironmentVariable($key, $previous[$key], 'Process')
    }
}
