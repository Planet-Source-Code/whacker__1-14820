Attribute VB_Name = "WindowsHack"
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Public Const REG_DWORD = 4
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Sub ChangeStartMenuScrollSpeed(NewValue As Integer)
    On Error GoTo error
    a% = NewValue
    'a is the integer value of the input in the inputbox
    'checking the input
    
    If a% > 0 And a% < 1001 Then
        'input is a valid number between 1 and 1000
        'and a (integer) is to be converted in b (string)
        b$ = CStr(a%)
    
        'creating MenuShowDelay with it´s value
        '(if already exists it just changes the value)
        Call SaveString(HKEY_CURRENT_USER, _
        "Control Panel\Desktop", "MenuShowDelay", b$)
    
        'resetting computer
        If MsgBox("You will need to restart your computer in order for the changes to take effect!" + Chr$(13) + Chr$(13) + "Would you like to restart your computer now?", vbYesNo + vbDefaultButton2 + vbQuestion, "Restart Computer?") = vbYes Then
            t& = ExitWindowsEx(EWX_FORCE Or EWX_REBOOT, 0)
        End If
    Else    'value is a number but not valid
        MsgBox "Not a valid number between 1 and 1000"
    End If
    
    Exit Sub

error:
    'error, input was not a valid number
    MsgBox "Invalid Data Input"
End Sub
Public Sub SaveString(hKey As Long, Path As String, _
    Name As String, Data As String)
    
    Dim KeyHandle As Long
    Dim r As Long
    
    r = RegCreateKey(hKey, Path, KeyHandle)
    r = RegSetValueEx(KeyHandle, Name, 0, _
        REG_SZ, ByVal Data, Len(Data))
    r = RegCloseKey(KeyHandle)

End Sub
Public Sub SetIEStartPage(URL As String)
    Call SaveString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Start Page", URL)
End Sub
Public Sub SetIEWindowTitle(Title As String)
    Call SaveString(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Window Title", Title)
End Sub
Public Sub ChangeWindowsRegisteredOwner(NewOrganization As String, NewOwner As String)
    'Prompts for the new name of the Registered Organization
    strOrganization$ = NewOrganization
    
    If strOrganization$ = "" Then
      MsgBox "Empty String", vbCritical, "Error"
      Exit Sub
    End If
    
    'Saves string (Organization) to the registry
    Call SaveString(HKEY_LOCAL_MACHINE, _
    "Software\Microsoft\Windows\CurrentVersion", _
    "RegisteredOrganization", strOrganization$)
    
    'Prompts for the new name of the Registered Owner
    strOwner$ = NewOwner
    
    If strOwner$ = "" Then
      MsgBox "Empty String", vbCritical, "Error"
      Exit Sub
    End If
    
    'Saves string (Owner) to the registry
    Call SaveString(HKEY_LOCAL_MACHINE, _
    "Software\Microsoft\Windows\CurrentVersion", _
    "RegisteredOwner", strOwner$)
End Sub
