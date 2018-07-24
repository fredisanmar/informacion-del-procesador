set WshShell = WScript.CreateObject("WScript.Shell")
Dim Msg, Style, Title, Response, MyString
Msg = "¿quieres ver la informacion de tu procesador?"    
Style = vbOkCancel    
Title = "informacion procesador"    

Response = MsgBox(Msg, Style, Title)
If Response = vbOk Then    
   On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")

For Each objItem in colItems
        echo
        MsgBox "descripcion: " & objItem.Description &vbnewline& "nombre: "&objItem.Name&vbnewline&"Manufactura: "&objItem.Manufacturer&vbnewline&"bus de direccion: " & objItem.AddressWidth& " bits" &vbnewline& "bus de datos: " & objItem.DataWidth&" bits"&vbnewline&"velocidad maxima de reloj: "&objItem.MaxClockSpeed&"MHz" &vbnewline&"firmware de virtualizacion habilitado: "&objItem.VirtualizationFirmwareEnabled&vbnewline &"voltage: "&objItem.CurrentVoltage&"V"&vbnewline&"Utilizado: "&objItem.LoadPercentage&"%"&vbnewline&     "arquitectura: " & objItem.Architecture & vbnewline & "disponibilidad: " & objItem.Availability & vbnewline & "estado de CPU : " & objItem.CpuStatus &vbnewline& "velocidad de reloj actual: " & objItem.CurrentClockSpeed &vbnewline&   "ID de dispositivo: " & objItem.DeviceID &vbnewline& "reloj externo: " & objItem.ExtClock &vbnewline& "familia: " & objItem.Family &vbnewline& "tamaño de caché L2: " & objItem.L2CacheSize &vbnewline& "velocidad de caché L2: " & objItem.L2CacheSpeed &vbnewline& "nivel: " & objItem.Level &vbnewline& "ID de dispositivo PNP: "&objItem.PNPDeviceID&vbnewline&"ID de procesador: "&objItem.ProcessorId&vbnewline& "tipo de procesador: " & objItem.ProcessorType &vbnewline& "Revision: " & objItem.Revision &vbnewline& "Rol: " & objItem.Role &vbnewline& "designacion de socket: " & objItem.SocketDesignation &vbnewline& "informacion de estado: " & objItem.StatusInfo &vbnewline& "Stepping: " & objItem.Stepping &vbnewline& "ID unico: " & objItem.UniqueId &vbnewline& "metodo de actualizacion: " & objItem.UpgradeMethod &vbnewline& "Version: " & objItem.Version

Next
quit= "Ok"   
Else    
    MyString = "Cancel"    
End If