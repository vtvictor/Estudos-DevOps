strcomputer = inputbox("Digite o nome do Computador ou o IP")
if strcomputer = "" then
    wscript.quit
else

'ping it!
Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery _
    ("select * from Win32_PingStatus where address = '" & strcomputer & "'")
For Each objStatus in objPing
    If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then
        'request timed out
        msgbox(strcomputer & " did not reply" & vbcrlf & vbcrlf & _
    "Verifique o nome ou o ip e tente novamente")
    else
        'Quem está ai?
        set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & _
    strComputer & "\root\cimv2")
        Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
        For Each objComputer in colSettings
            msgbox("Nome do Computador: " & objComputer.Name & vbcrlf & "Usuario Logado : " & _
    objcomputer.username  & vbcrlf & "Dominio: " & objComputer.Domain)
        Next
    end if
next
end if