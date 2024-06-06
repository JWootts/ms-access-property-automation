'written by jordan wootton (june 6th 2024) - proof of concept... garbage code zzzzzzz

Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As LongPtr) As Long
Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Const PROJECT_LOCATION As String = "C:\Users\JWootts\Desktop\Database3.accdb"
Private Const VBE_PASS As String = "password"

Public Function unlock_access_vbe()

    Dim app As New Access.Application
    Dim vbProj As Object
    Dim windowPtr As Long, pwFieldPtr As LongPtr, btnWindowPtr As LongPtr
    Dim strBuff As String, btnCaption As String
    
    app.OpenCurrentDatabase PROJECT_LOCATION, False, VBE_PASS
    Set vbProj = app.CodeProject.Application.VBE.VBprojects(1)
    
    vbProj.VBE.CommandBars(1).FindControl(ID:=2578, recursive:=True).Execute
    
    windowPtr = FindWindowA(vbNullString, vbProj.Name & " Password")
    
    If windowPtr <> 0 Then
    
        pwFieldPtr = FindWindowEx(windowPtr, ByVal 0&, "Edit", vbNullString)
        
        If pwFieldPtr <> 0 Then
            
            Call SendMessage(pwFieldPtr, &HC2, False, ByVal VBE_PASS)
            DoEvents
            
            btnWindowPtr = FindWindowEx(windowPtr, ByVal 0&, "Button", vbNullString)
            strBuff = String(GetWindowTextLength(btnWindowPtr) + 1, Chr$(0))
            GetWindowText btnWindowPtr, strBuff, Len(strBuff)
            btnCaption = strBuff
            
ReCheck:
            If Not btnCaption Like "OK*" Then
                If btnWindowPtr <> 0 Then
                    btnWindowPtr = FindWindowEx(windowPtr, btnWindowPtr, "Button", vbNullString) 'find next button pointer in grouping
                    strBuff = String(GetWindowTextLength(btnWindowPtr) + 1, Chr$(0))
                    GetWindowText btnWindowPtr, strBuff, Len(strBuff)
                    btnCaption = strBuff
                    GoTo ReCheck
                End If
            End If
            
            'press ok button to unlock vbe project
            Call SendMessage(btnWindowPtr, &HF5, 0, vbNullString) 'click
            
        End If
          
    End If
    
End Function


'called after unlock_access_vbe // project is open / unlocked
Public Function SetDevConstant(vbeProjectName As String, Optional val As Integer = 0)
    
    Dim windowPtr As Long, conditionalArgsEditPts As LongPtr, devBoxPtr As LongPtr
    
    windowPtr = FindWindowA(vbNullString, vbeProjectName + " - Project Properties")
    
    'find internal tab box
    conditionalArgsEditPts = FindWindowEx(windowPtr, ByVal 0&, "#32770", vbNullString)
    
    devBoxPtr = FindWindowEx(conditionalArgsEditPts, ByVal 0&, "Edit", vbNullString)
    devBoxPtr = FindWindowEx(conditionalArgsEditPts, ByVal devBoxPtr, "Edit", vbNullString)
    devBoxPtr = FindWindowEx(conditionalArgsEditPts, ByVal devBoxPtr, "Edit", vbNullString)
    devBoxPtr = FindWindowEx(conditionalArgsEditPts, ByVal devBoxPtr, "Edit", vbNullString)
    devBoxPtr = FindWindowEx(conditionalArgsEditPts, ByVal devBoxPtr, "Edit", vbNullString)
    
    Call SendMessage(devBoxPtr, &H1330&, 1, ByVal 0&) 'focus
    Call SendMessage(devBoxPtr, &HB1, False, ByVal -1&) 'clear
    Call SendMessage(devBoxPtr, &HC2, False, ByVal "DEV = " & val)
    
End Function

'called after SetDevConstant
Public Function SetProtectionLevel(vbeProjectName As String, Optional off As Boolean = False)

    Dim windowPtr As Long, tabsPtr As LongPtr, lockCheckPtr As LongPtr, conditionalArgsEditPts As LongPtr
    Dim editBoxPtr As LongPtr, okButtonPtr As LongPtr
    
    Dim pass As String: pass = IIf(off, "", VBE_PASS)
    
    windowPtr = FindWindowA(vbNullString, vbeProjectName + " - Project Properties")
    tabsPtr = FindWindowEx(windowPtr, ByVal 0&, "SysTabControl32", vbNullString)
    
    Call SendMessage(tabsPtr, &H1330&, 1, ByVal 0&) 'focus protection panel
    
    'find internal tab box
    conditionalArgsEditPts = FindWindowEx(windowPtr, ByVal 0&, "#32770", vbNullString)
    
    'click lock checkbox
    lockCheckPtr = FindWindowEx(conditionalArgsEditPts, ByVal 0&, "Button", vbNullString)
    lockCheckPtr = FindWindowEx(conditionalArgsEditPts, lockCheckPtr, "Button", vbNullString)
    
    'get current value
    nRetVal = SendMessage(lockCheckPtr, &HF2, 0&, ByVal 0&) 'get state
    
    'disable protection
    If off And nRetVal Then
        Call SendMessage(lockCheckPtr, &H1330&, 1, ByVal 0&) 'focus
        Call SendMessage(lockCheckPtr, &HF5, 0, vbNullString) 'click
    End If
    
    'enable protection
    If Not off And nRetVal <> 1 Then
        Call SendMessage(lockCheckPtr, &H1330&, 1, ByVal 0&) 'focus
        Call SendMessage(lockCheckPtr, &HF5, 0, vbNullString) 'click
    End If
    
    'remove password if need be
    editBoxPtr = FindWindowEx(conditionalArgsEditPts, ByVal 0&, "Edit", vbNullString)
    Call SendMessage(editBoxPtr, &H1330&, 1, ByVal 0&) 'focus
    Call SendMessage(editBoxPtr, &HB1, False, ByVal -1&) 'clear
    Call SendMessage(editBoxPtr, &HC2, False, ByVal pass) 'set password to nothing
    editBoxPtr = FindWindowEx(conditionalArgsEditPts, editBoxPtr, "Edit", vbNullString)
    Call SendMessage(editBoxPtr, &H1330&, 1, ByVal 0&) 'focus
    Call SendMessage(editBoxPtr, &HB1, False, ByVal -1&) 'clear
    Call SendMessage(editBoxPtr, &HC2, False, ByVal pass) 'set password to nothing
    
    'close
    okButtonPtr = FindWindowEx(windowPtr, ByVal 0&, "Button", vbNullString)
    Call SendMessage(okButtonPtr, &HF5, 0, vbNullString) 'click
    
End Function



