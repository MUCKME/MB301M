Dim RK1, RK2
Dim VS,Version
On Error Resume Next
Set oExcel = CreateObject("Excel.Application")
VS = oExcel.Version
If Err or VS = "" Then 
	Version = InputBox("Excel Version Number(4 digits), ex.2003, 2007, 2010 or 2013.","Input Excel Version" , "20xx", 10000, 8500)
	if Version = "" then wscript.quit
	Select Case Trim(Version)
   		Case "2000"
			VS = "9.0"
   		Case "2003"
			VS = "11.0"
   		Case "2007"
			VS = "12.0"
   		Case "2010"
			VS = "14.0"
   		Case "2013"
			VS = "15.0"
   		Case else
			Msgbox "Version can not be indentified."
			wscript.quit
	End Select
end if

RK1 = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & VS & "\Excel\Security\VBAWarnings"
RK2 = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & VS & "\Excel\Security\Level"

Var = MsgBox("Macro Settings:" & chr(10) & chr(10) & "	Do you want to enable macro automatically£¿", 48 + 3 + 512, "Confirm")
Select Case Var
   Case vbYes
   	Call WDReg(RK1, 1, "REG_DWORD")
   	Call WDReg(RK2, 1, "REG_DWORD")
   	'msgbox "Already set to: enable macro automatically!"
   Case vbNo
   	Call WDReg(RK1, 2, "REG_DWORD")
   	Call WDReg(RK2, 2, "REG_DWORD")
   	'msgbox "Already set to: enabel macro manually by user self!"
   Case vbCancel
	wscript.quit
end Select
wscript.quit

Sub WDReg(strkey, Value, ValueType)
    Dim oWshell
    Set oWshell = CreateObject("WScript.Shell")
    If ValueType = "" Then
        oWshell.RegWrite strkey, Value
    Else
        oWshell.RegWrite strkey, Value, ValueType
    End If
    Set oWshell = Nothing
End Sub