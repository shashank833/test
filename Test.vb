[PCOMM SCRIPT HEADER]
Language = VBSCRIPT
DESCRIPTION=
[PCOMM SCRIPT SOURCE]
Option Explicit
autECLSession.SetConnectionByName (ThisSessionName)
Global login, password As String
Rem This line calls the macro subroutine
subSub1_
Global objConnection1
Global objrecordset1

Sub subSub1_()
Call Excel_conn
	
   autECLSession.autECLOIA.WaitForAppAvailable
   autECLSession.autECLOIA.WaitForInputReady
   autECLSession.autECLPS.SendKeys login       'Login User name
   autECLSession.autECLOIA.WaitForInputReady
   autECLSession.autECLPS.SendKeys "[tab]"
   autECLSession.autECLOIA.WaitForInputReady
   autECLSession.autECLPS.SendKeys password    'Login Password
   