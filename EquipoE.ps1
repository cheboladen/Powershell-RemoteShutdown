 <#

.SYNOPSIS
Importa el archivo que contiene los datos, para seleccionar los equipos que se desean apagar.

.DESCRIPTION
EquipoE.ps1 importa el archivo, lo analiza y se basa en su contenido para ejecutar la orden de apagado.
Después, devuelve una frase de confirmación de en que equipo se ejecuto y de si fue con exito o no.
A continuación, guarda en el archivo logerror.txt el error sucedido si es que hubiera alguno.

.PARAMETER Fichero $Fichero
Ruta del fichero .CSV a analizar.

.INPUTS
Fichero CSV a analizar

.OUTPUTS
Nos dice en que equipos se ejecuto la orden de apagado de entre los proporcionados por el csv, notificando su exito o su fracaso en dicha tarea.
Escribe en el fichero Log.txt los eventos según sean ERROR, SUCCESS o INFO.

ERROR: No se consiguió ejecutar la operación.
SUCCESS: Se ejecutó la operación.
INFO: Rendimiento en milisegundos.

#>

#Funccion para apagar los ordenadores remotamente.

function ApagarOrdenadores {
    [cmdletbinding()]
    param (
            [parameter(ValueFromPipeline)]
            [String] $Fichero)
    process {
        $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $Log = ".\Eventos.log"

        #Importar CSV
        $DatosCSV = Import-csv $Fichero -Delimiter ";" -Header Apellidos, Nombre, Computername, MAC | Select-Object -Skip 1
        $Timestamp = Get-Date
        "$Timestamp | INFO | Ejecutando operaciones utilizando $Fichero" | Out-File -Encoding utf8 $Log -Append

        #Validar CSV
        $PatronHostname = [Regex]::new('A[0-9]{2}W[[0-9]{2}')
        $PatronMAC = [Regex]::new('([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})')
        $ValidarCSV = $DatosCSV | Where-Object {($_.Computername -match $PatronHostname) -and ($_.MAC -match $PatronMAC)}

        #Comprobar Calidad de Código
        $EquipoE = ".\EquipoE.ps1"
        $Debug = Invoke-ScriptAnalyzer -Path $EquipoE
        $ErroresDebug = $Debug.Count
        $Timestamp = Get-Date
        "$Timestamp | DEBUG | Errores de calidad: $ErroresDebug" | Out-File -Encoding utf8 $Log -Append

        ForEach ($pc in $ValidarCSV.Computername) {

            $Timestamp = Get-Date
                If (Test-Connection -BufferSize 32 -Count 1 -Computername $pc -Quiet){

                    Invoke-Command -Computername $pc -ScriptBlock {Get-NetAdapter | Sort-Object -Property MacAddress | format-table -Property MacAddress}
                    ($DatosCSV.MAC) | foreach-object {Get-CimInstance -ClassName Win32_networkadapterconfiguration -Filter "IPEnabled='True'" -ComputerName $pc -ErrorAction SilentlyContinue | Select-Object -Property MACAddress | Select-string -Pattern ($_) -AllMatches}

                    Try {
                        $session = New-CimSession -ComputerName $pc
                        $Query = "Select * From Win32_OperatingSystem"
                        Invoke-CimMethod -Query $Query -MethodName Shutdown -CimSession $Session
                    }
                    Catch [System.Management.Automation.ItemNotFoundException] {
                        $_.Exception | Out-File -Encoding utf8 $Log -Append
                        "$Timestamp | ERROR | $pc no se encuentra en la lista de ordenadores." | Out-File -Encoding utf8 $Log -Append
                    }
                    Catch [System.Management.Automation.ParameterBindingException] {
                        $_.Exception | Out-File -Encoding utf8 $Log -Append
                        "$Timestamp | ERROR | No se ha podido conectar con el ordenador $pc." | Out-File -Encoding utf8 $Log -Append
                    }
                    Catch {
                        $_.Exception | Out-File -Encoding utf8 $Log -Append
                        "$Timestamp | ERROR | Error desconocido producido en $pc." | Out-File -Encoding utf8 $Log -Append
                    }
                    Finally {
                        "$Timestamp | SUCCESS | Orden de apagado ejecutada en $pc" | Out-File -Encoding utf8 $Log -Append
                    }
                        Remove-CimSession -CimSession $Session
                }
                else {
                    "$Timestamp | ERROR | No se pudo conectar con $pc" | Out-File -Encoding utf8 $Log -Append
                }
        }
        $Rendimiento = $Stopwatch.Elapsed.Milliseconds
        $Stopwatch.Stop()
        $Ejecuciones = $DatosCSV.Count
        "$Timestamp | TRACE | Ejecuciones: $Ejecuciones | Tiempo: $Rendimiento" | Out-File -Encoding utf8 $Log -Append
    }
}