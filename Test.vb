[PCOMM SCRIPT HEADER]
Language = VBSCRIPT
DESCRIPTION=
[PCOMM SCRIPT SOURCE]
Option Explicit
autECLSession.SetConnectionByName (ThisSessionName)
Rem This line calls the macro subroutine
subSub1_
Global objConnection1
Global objrecordset1
Global Password as string
Global UserName as string

Sub subSub1_()


call ReadConfig	
   autECLSession.autECLOIA.WaitForAppAvailable
   autECLSession.autECLOIA.WaitForInputReady
   autECLSession.autECLPS.SendKeys UserName       'Login User name
   autECLSession.autECLOIA.WaitForInputReady
   autECLSession.autECLPS.SendKeys "[tab]"
   autECLSession.autECLOIA.WaitForInputReady
   autECLSession.autECLPS.SendKeys Password    'Login Password
   end Sub
Sub ReadConfig()
Const adOpenStatic = 3
    Const adLockOptimistic = 3
    Const adCmdText = 512
    Set objConnection1 = CreateObject("ADODB.Connection")
    Set objrecordset = CreateObject("ADODB.Recordset")
    objConnection1.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & ITDataFile_Path & ";" & _
    "Extended Properties=""Excel 8.0;HDR=Yes;"";"
	objrecordset2.Open "Select * FROM [" & Sheet1 & "$]", _
    objConnection2, adOpenStatic, adLockOptimistic, adCmdText
	UserName = objrecordset.Fields.Item("USERNAME")
	password = objrecordset.Fields.Item("PASSWORD")
	MsgBox ("User Name is" &UserName)
	end Sub

			

   
