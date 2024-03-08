###################################################################
#Script Name	: Crear Responsiva                                                                                 
#Description	: Crea una carpeta con el hostname del equipo, asi como prepara la responsiva y el checklist necesario como documentacion al entregar un equipo en MAVER                     
#Args           : 
#Author       	: Uriel Aguirre                                              
#Email         	: aaguirre@maver.com.mx                     
###################################################################

function Send-Change {
    param (
        $Campo,
        $datos
    )

    $documento.Content.Find.Execute($campo, $true, $true, $false, $false, $false, $true, 0, $true, $datos)
}

function Check-enviroment {
    $LocalizacionActual = (Get-Location)
    $CarpetaResponsivas = "" + $LocalizacionActual + "\Responsiva\"

    if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
        $Message = "El modulo ImportExcel No esta instalado."
        Write-Host $Message -ForegroundColor Red
        try{
            Install-Module -Name ImportExcel -Force -Scope CurrentUser -AllowClobber
        } catch {
            $Message = "El modulo ImportExcel No se pudo instalar, probablemente por permisos."
            Write-Host $Message -ForegroundColor Red
            return $false
        }
        Clear-Host
        if ((Get-Module -Name ImportExcel -ListAvailable)) {
            $Message = "El modulo ImportExcel fue instalado con exito."
            Write-Host $Message -ForegroundColor Green
        } else {
            $Message = "El modulo ImportExcel No se pudo instalar, probablemente por permisos."
            Write-Host $Message -ForegroundColor Red
            return $false
        }
    } else {
        $Message = "El modulo ImportExcel ya esta instalado."
        Write-Host $Message -ForegroundColor Green
    }

    if (Test-Path $CarpetaResponsivas){
        $Message = "Carpeta de responsivas existe."
        Write-Host $Message -ForegroundColor Green
    } else {
        $Message = "La carpeta de responsivas no existe."
        Write-Host $Message -ForegroundColor Red

        try {
            New-Item $CarpetaResponsivas -itemType Directory
        } catch {
            $Message = "La carpeta de responsivas no pudo ser creada."
            Write-Host $Message -ForegroundColor Red
            return
        }
            $Message = "La carpeta de responsivas fue creada."
            Write-Host $Message -ForegroundColor Green
            return $true
    }
    return $true
}

function New-Responsiva {
    param (
        $Info
    )

    if ([string]::IsNullOrEmpty($Info.HostName)) {
        Write-Host "Error: El valor de Fecha esta vacio." -ForegroundColor Red
        exit
    }
    if ([string]::IsNullOrEmpty($Info.NumeroSeriePC)) { 
        Write-Host "Error: El valor de Numero Serie PC esta vacio." -ForegroundColor Red
        exit
    }

    if ([string]::IsNullOrEmpty($Info.Nombre)) {
        Write-Host "Error: El valor de Nombre esta vacio." -ForegroundColor Red
        exit
    }

    $LocalizacionActual = (Get-Location)
    $CarpetaResponsivas = "" + $LocalizacionActual + "\Responsiva\"
    $CarpetaUsuario = $CarpetaResponsivas + $Info.HostName + "\"
    if(-not (Test-Path $CarpetaUsuario)){
        New-Item $CarpetaUsuario -itemType Directory
    }
    $PlantillaWord = "" + $LocalizacionActual + "\PNO-SIT-01-F02 CARTA RESPONSIVA DE EQUIPO V3 - copia.docx"
    $CopiaplantillaWord = "" + $CarpetaUsuario + $Info.NumeroSeriePC + " - " + $Info.Nombre + "_SF.docx"
    Copy-Item $PlantillaWord -Destination $CopiaplantillaWord
    $archivoPDFSalida = "" + $CarpetaUsuario + "\" + $Info.NumeroSeriePC + " - " + $Info.Nombre + "_SF.pdf"


    $word = New-Object -ComObject Word.Application
    $documento = $word.Documents.Open($CopiaplantillaWord)
    
    if ([string]::IsNullOrEmpty($Info.Fecha)) {
        Write-Host "Error: El valor de Fecha esta vacio." -ForegroundColor Red
        exit
    }
    Send-Change "{{FechaCorta}}" $Info.Fecha

    if ([string]::IsNullOrEmpty($Info.FechaLarga)) {
        Write-Host "Error: El valor de Fecha Larga esta vacio." -ForegroundColor Red
        exit
    }
    Send-Change "{{FechaLarga}}" $Info.FechaLarga

    Send-Change "{{Anexo}}" $Info.Anexo

    if ([string]::IsNullOrEmpty($Info.ReferenciaFisica)) {
        Write-Host "Error: El valor de Referencia Fisica esta vacio." -ForegroundColor Red
        exit
    }
    Send-Change "{{ReferenciaFisica}}" $Info.ReferenciaFisica

    Send-Change "{{NombreDeUsuario1}}" $Info.Nombre
    Send-Change "{{NombreDeUsuario2}}" $Info.Nombre
    Send-Change "{{NombreDeUsuario3}}" $Info.Nombre

    if ([string]::IsNullOrEmpty($Info.Departamento)) {
        Write-Host "Error: El valor de Departamento esta vacio." -ForegroundColor Red
        exit
    }
    Send-Change "{{DepartamentoUsuario}}" $Info.Departamento

    if ([string]::IsNullOrEmpty($Info.UnidadNegocioUsuario)) {
        Write-Host "Error: El valor de Unidad Negocio Usuario esta vacio." -ForegroundColor Red
        exit
    }
    Send-Change "{{UnidadNegocioUsuario}}" $Info.UnidadNegocioUsuario

    if ([string]::IsNullOrEmpty($Info.NombreGerente)) {
        Write-Host "Error: El valor de Nombre Gerente esta vacio." -ForegroundColor Red
        exit
    }
    Send-Change "{{NombreGerente}}" $Info.NombreGerente

    if ([string]::IsNullOrEmpty($Info.PuestoGerente)) {
        Write-Host "Error: El valor de Puesto Gerente esta vacio." -ForegroundColor Red
        exit
    }
    Send-Change "{{PuestoGerente}}" $Info.PuestoGerente

    if ([string]::IsNullOrEmpty($Info.MarcaPC)) {
        Write-Host "Error: El valor de Marca PC esta vacio." -ForegroundColor Red
        exit
    }
    Send-Change "{{MarcaPC}}" $Info.MarcaPC

    if ([string]::IsNullOrEmpty($Info.ModeloPC)) {
        Write-Host "Error: El valor de Modelo PC esta vacio." -ForegroundColor Red
        exit
    }
    Send-Change "{{ModeloPC}}" $Info.ModeloPC

    Send-Change "{{NumeroSeriePC}}" $Info.NumeroSeriePC

    if ([string]::IsNullOrEmpty($Info.MarcaMonitor)) {
        Write-Host "Error: El valor de Marca Monitor esta vacio." -ForegroundColor Red
        exit
    }
    Send-Change "{{MarcaMonitor}}" $Info.MarcaMonitor

    if ([string]::IsNullOrEmpty($Info.ModeloMonitor)) {
        Write-Host "Error: El valor de Modelo Monitor esta vacio." -ForegroundColor Red
        exit
    }
    Send-Change "{{ModeloMonitor}}" $Info.ModeloMonitor

    if ([string]::IsNullOrEmpty($Info.NumeroSerieMonitor)) {
        Write-Host "Error: El valor de Numero Serie Monitor esta vacio." -ForegroundColor Red
        exit
    }
    Send-Change "{{NumeroSerieMonitor}}" $Info.NumeroSerieMonitor


    $documento.Save()
    $documento.Close($false)

    $word.Quit()

    $WordApp = New-Object -ComObject Word.Application
    $WordApp.Visible = $false
    $WordApp.Application.DisplayAlerts = 0

    $doc = $WordApp.Documents.Open($CopiaplantillaWord)
    $doc.SaveAs([ref]$archivoPDFSalida, [ref]17)

    $doc.Close($false)
    $WordApp.Quit()

    Clear-Host

    Write-Host "El archivo PDF se ha generado con exito en: $archivoPDFSalida" -ForegroundColor Green
}

function Show-MenuHoja {
    
    for($i=1; $i -lt 10; $i++){
        try {
            $Workbook.Sheets[$i].index
        } 
        catch {
            Break
        }
    }
    $index = $i -1

    Write-Host "`n======================="
    for($i=1; $i -ile $index; $i++){
        $label = "||     " + $i + "). " + $Workbook.Sheets[$i].name + "      ||"
        Write-Host $label
    }
    Write-Host "=======================`n"
    while ($Selected -eq $null) {
        $input = Read-Host "Selecciona una hoja "
        if ($input -match '^\d+$' -and [int]$input -gt 0 -and [int]$input -le $index) {
            $Selected = [int]$input
        } else {
            Write-Host "Seleccion no válida. Introduce un número válido."
        }
    }

    Clear-Host

    return $Selected
}

function Show-MenuEquipo {
    param (
        $hojaSeleccionada
    )

    $sheet = $workbook.Sheets.Item($hojaSeleccionada)

    $Equipos = @()

    for($i=3; $i -lt 100; $i++){
        try {
            if($sheet.Cells.Item($i,2).Value2 -notlike $Null){
                $indice = $sheet.Cells.Item($i,2).Value2
                $NombreDeUsuario = $sheet.Cells.Item($i,3).Value2
                $MarcaPC = $sheet.Cells.Item($i,4).Value2
                $ModeloPC = $sheet.Cells.Item($i,5).Value2
                $NumeroSeriePC = $sheet.Cells.Item($i,6).Value2
                $MarcaMonitor = $sheet.Cells.Item($i,7).Value2
                $ModeloMonitor = $sheet.Cells.Item($i,8).Value2
                $NumeroSerieMonitor = $sheet.Cells.Item($i,9).Value2
                $Anexo = $sheet.Cells.Item($i,10).Value2
                $ReferenciaFisica = $sheet.Cells.Item($i,11).Value2
                $UnidadNegocioUsuario = $sheet.Cells.Item($i,12).Value2
                $DepartamentoUsuario = $sheet.Cells.Item($i,13).Value2
                $PuestoUsuario = $sheet.Cells.Item($i,14).Value2
                $NombreGerente = $sheet.Cells.Item($i,15).Value2
                $PuestoGerente = $sheet.Cells.Item($i,16).Value2
                $FechaCorta = Get-Date([DateTime]::FromOADate($sheet.Cells.Item($i,17).Value2)) -Format "dd/MM/yyyy"
                $FechaLarga = "Tlaquepaque, Jalisco a " + (Get-Date([DateTime]::FromOADate($sheet.Cells.Item($i,17).Value2)) -Format "dd") + " de " + (Get-Date([DateTime]::FromOADate($sheet.Cells.Item($i,17).Value2)) -Format "MMMM") + " de " + (Get-Date([DateTime]::FromOADate($sheet.Cells.Item($i,17).Value2)) -Format "yyyy")
                $HostName = $sheet.Cells.Item($i,18).Value2

                $Equipos += [PSCustomObject]@{
                    Indice = $indice
                    Nombre = $NombreDeUsuario
                    Fecha = $FechaCorta
                    FechaLarga = $FechaLarga
                    Anexo = $Anexo
                    ReferenciaFisica = $ReferenciaFisica
                    Departamento = $DepartamentoUsuario
                    UnidadNegocioUsuario = $UnidadNegocioUsuario
                    NombreGerente = $NombreGerente
                    PuestoGerente = $PuestoGerente
                    MarcaPC = $MarcaPC
                    ModeloPC = $ModeloPC
                    NumeroSeriePC = $NumeroSeriePC
                    MarcaMonitor = $MarcaMonitor
                    ModeloMonitor = $ModeloMonitor
                    NumeroSerieMonitor = $NumeroSerieMonitor
                    HostName = $HostName
                    PuestoUsuario = $PuestoUsuario
                }
            } else {
                Break
            }
        } 
        catch {
            Break
        }
    }
    $Indice = $i -3

    Write-Host "Existen $Indice Registros"

   Write-Host "`n====================================================================="
    for($i=0; $i -ile $Indice -1; $i++){
        $label = "||     " + $Equipos[$i].Indice + "). " + $Equipos[$i].Nombre + " " + $Equipos[$i].Fecha + "      ||"
        Write-Host $label
    }
    Write-Host "=====================================================================`n"

    Write-Host "Existen " $Indice " Registros"

    $Selected = $null
    while ($Selected -eq $null) {
        $input = Read-Host "Selecciona una Equipo "
        if ($input -match '^\d+$' -and [int]$input -gt 0 -and [int]$input -le $Indice) {
            $Selected = [int]$input
        } else {
            Write-Host "Seleccion no válida. Introduce un número válido."
        }
    }

    Clear-Host

    return $Equipos[$Selected - 1]
}

function New-CheckList {
    param (
        $Info
    )

    Add-Type -Path "itextsharp.dll"
    Add-Type -Path "BouncyCastle.Crypto.dll"

    $LocalizacionActual = (Get-Location)
    $CarpetaResponsivas = "" + $LocalizacionActual + "\Responsiva\"
    $CarpetaUsuario = $CarpetaResponsivas + $Info.HostName + "\"
    $PdfFileInput = "" + $LocalizacionActual +"\Formato de cheklist V 3.2 Formulario.pdf"
    $PdfFileOutput = "" + $CarpetaUsuario + $Info.NumeroSeriePC + "_" + $Info.HostName + ".pdf"

    $PdfReader = New-Object iTextSharp.text.pdf.PdfReader($PdfFileInput)
    $PdfStamper = New-Object iTextSharp.text.pdf.PdfStamper($PdfReader, [System.IO.File]::Create($PdfFileOutput))

    $PdfFields = @{
        "Nombre del Ingeniero que preparó el equipo" = "Andres Uriel Aguirre Ocampo"
        "Fecha de formateo" = $Info.Fecha
        "Nombre del usuario" = $Info.Nombre
        Numserie = $Info.NumeroSeriePC
        "Nombre del equipo" = $Info.HostName
        "Puesto del usuario" = $Info.PuestoUsuario
        "Depto Usuario" = $Info.DepartamentoUsuario
        "Fecha de entrega" = $Info.Fecha
        "Nombre del equipo_2" = $Info.HostName
    }

    ForEach ($PdfField in $PdfFields.GetEnumerator()) {
        $PdfStamper.AcroFields.SetField($PdfField.Key, $PdfField.Value)
        $PdfStamper.AcroFields.SetFieldProperty($PdfField.Key, "setfflags", [iTextSharp.text.pdf.PdfFormField]::FF_READ_ONLY, 0)
    }

    $PdfStamper.Close()
    $PdfReader.Close()



    Write-Host "El archivo PDF se ha generado con exito en: $PdfFileOutput" -ForegroundColor Green

}

function Set-Informacion{
    param (
        $Info
    )
}

if(Check-enviroment){
    try {
        $UbicacionActual = Get-Location
        $RutaExcel = "" + $UbicacionActual + "\Equipos Nuevos.xlsx"
        $Excel = New-Object -ComObject Excel.Application
        $Workbook = $Excel.Workbooks.Open($RutaExcel)

        $hojaSeleccionada = (Show-MenuHoja)
        $EquipoSeleccionado = $Null
        $EquipoSeleccionado = Show-MenuEquipo $hojaSeleccionada[2]
        $Excel.Quit()
        $EquipoSeleccionado
        New-Responsiva $EquipoSeleccionado $PlantillaWord $archivoPDFSalida

        New-CheckList $EquipoSeleccionado

        #TODO Actualizar Excel
        Set-Informacion $EquipoSeleccionado
    } catch {
        Write-Host "Se produjo un error: $_"
        if ($Excel -ne $null) {
            $Excel.Quit()
        }
    }
} else {
    $Message = "Existen problemas de entorno y no fue posible ejecutar el script."
    Write-Host $Message -ForegroundColor Red
    exit
}
