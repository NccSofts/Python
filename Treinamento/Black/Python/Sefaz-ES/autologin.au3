# AutoIt Version: 3.3.10.2
# Author:         Mike Kovacic

 # Script Function:
 #   Set windows 7 to auto login using provided credentials.

  #      * This must be compiled as x64 to work.

#ce ---------------------------------------------------------------------------

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
Global $MuhApp = GUICreate("Set Autologon", 309, 127, 212, 132)
GLOBAL $Username = GUICtrlCreateInput("helpdesk", 8, 16, 153, 21)
Global $Password = GUICtrlCreateInput("Kdru$3k2oQxk2#", 8, 48, 153, 21)
Global $Domain = GUICtrlCreateInput(@ComputerName, 8, 80, 153, 21)
Global $Set = GUICtrlCreateButton("Set Autologon", 176, 24, 123, 25)
Global $REmove = GUICtrlCreateButton("Remove Autologin", 176, 56, 123, 25)
GUISetState(@SW_SHOW)
While 1
    $nMsg = GUIGetMsg()
    Switch $nMsg
        Case $GUI_EVENT_CLOSE
            Exit
        Case $set
            setautoadminlogin()
        Case $remove
            removeautologin()
    EndSwitch
WEnd
Func setautoadminlogin()
$a = 0
    If RegWrite("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa", "LmCompatabilityLevel", "REG_DWORD", "2") then $a += 1
    If RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "AutoAdminLogon", "REG_SZ", "1") then $a += 1
    If RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "ForceAutoLogon", "REG_SZ", "1") then $a += 1
    If RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "DefaultUserName", "REG_SZ", guictrlread($username)) then $a += 1
    If RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "DefaultPassword", "REG_SZ", guictrlread($Password)) then $a += 1
    If RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "DefaultDomainName", "REG_SZ", guictrlread($Domain)) then $a += 1
    If $a  = 6 then
        MsgBox(64,"Success!","Autologon has been set.")
    else
        MsgBox(16,"Oops!","Autologon could not be set. Make sure you are a local Admin.")
    endif
EndFunc   ;==>setautoadminlogin
Func removeautologin()
$b = 0
    If RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "AutoAdminLogon", "REG_SZ", "0") then $b += 1
    If RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "ForceAutoLogon", "REG_SZ", "0") then $b += 1
    If RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "DefaultUserName") then $b += 1
    If RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "DefaultPassword") then $b += 1
    If RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "DefaultDomainName") then $b += 1
    If RegDelete("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa", "LmCompatabilityLevel") then $b += 1
If $b  = 6 then
MsgBox(64,"Success!","Autologon has been removed.")
else
MsgBox(16,"Oops!","Autologon could not be removed. Either Autologin was not set or you are not a local Admin.")
endif
EndFunc   ;==>removeautologin