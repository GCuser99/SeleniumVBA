Attribute VB_Name = "tbUtils"
' ==========================================================================
' SeleniumVBA v3.1
' A Selenium wrapper for Edge, Chrome, Firefox, and IE written in Windows VBA based on JSon wire protocol.
'
' (c) GCUser99
'
' https://github.com/GCuser99/SeleniumVBA/tree/main
'
' ==========================================================================
Option Private Module
Option Explicit

Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, lpdwProcessId As Long) As Long
Private Declare PtrSafe Function FindWindowExA Lib "user32" (ByVal hwndParent As LongPtr, ByVal hwndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpIID As UUID) As LongPtr
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As LongPtr, ByVal dwId As Long, riid As Any, ppvObject As Object) As Long

Private Const IID_IDispatch As String = "{00020400-0000-0000-C000-000000000046}"
Private Const OBJID_NATIVEOM As Long = &HFFFFFFF0

Type UUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Sub test2()
    Debug.Print ThisLibFolderPath
    Debug.Print App.Path
    Debug.Print App.ModulePath
    Debug.Print App.Major & "." & App.Minor & "." & App.Build
    Debug.Print CallerApplicationObject Is Nothing
End Sub

Public Function ThisLibFolderPath() As String
    ThisLibFolderPath = App.Path
End Function

Public Function CallerApplicationObject() As Object
    Dim pid As Long
    Dim obj As Object
    
    'get the calling application process id
    pid = GetCurrentProcessId()
    
    'Check if we have an Excel application
    Set obj = FindAO(pid, Array("XLMAIN", "XLDESK", "EXCEL7"))
    If Not obj Is Nothing Then
        Set CallerApplicationObject = obj
        Exit Function
    End If
    
    'Check if we have an Access application
    Set obj = FindAO(pid, Array("OMAIN"))
    If Not obj Is Nothing Then
        Set CallerApplicationObject = obj
        Exit Function
    End If
    
    'Check if we have an PowerPoint application
    Set obj = FindAO(pid, Array("PPTFrameClass", "MDIClient", "mdiClass"))
    If Not obj Is Nothing Then
        Set CallerApplicationObject = obj
        Exit Function
    End If
    
    'Check if we have an Word application
    Set obj = FindAO(pid, Array("OpusApp", "_WwF", "_WwB", "_WwG"))
    If Not obj Is Nothing Then
        Set CallerApplicationObject = obj
        Exit Function
    End If
End Function

Private Function FindAO(ByVal targetPid As Long, winTree As Variant) As Object
    Dim obj As Object
    Dim hwndMain As LongPtr
    Dim hwndChild As LongPtr
    Dim hwndParent As LongPtr
    Dim i As Long
    Dim iid As UUID
    Dim winTreeLBound As Long
    Dim winTreeUBound As Long
    Dim thisPid As Long
    
    'construct iid that will be used in call to AccessibleObjectFromWindow
    Call IIDFromString(StrPtr(IID_IDispatch), iid)
    
    winTreeLBound = LBound(winTree)
    winTreeUBound = UBound(winTree)

    hwndMain = 0&
    Do
        'get next window from desktop
        hwndMain = FindWindowExA(0&, hwndMain, winTree(winTreeLBound), vbNullString)
        If hwndMain = 0& Then Exit Do
        Call GetWindowThreadProcessId(hwndMain, thisPid)
        If thisPid = targetPid Then
            'pid of this window matches target
            'now drill down levels to find the handle of the application window
            hwndParent = hwndMain
            hwndChild = hwndMain
            For i = winTreeLBound + 1 To winTreeUBound
                hwndChild = FindWindowExA(hwndParent, 0&, winTree(i), vbNullString)
                hwndParent = hwndChild
            Next i
            If AccessibleObjectFromWindow(hwndChild, OBJID_NATIVEOM, iid, obj) = 0& Then
                Set FindAO = obj.Application
                Exit Function
            End If
        End If
    Loop
End Function

Public Function ActiveVBAProjectFolderPath() As String
    'This returns the calling code project's parent document path. So if caller is from a project that references the SeleniumVBA Add-in
    'then this returns the path to the caller, not the Add-in (unless they are the same).
    'But be aware that if qc'ing this routine in Debug mode, the path to this SeleniumVBA project will be returned, which
    'may not be the caller's intended target if it resides in a different project.
    Dim fso As New FileSystemObject
    Dim strPath As String
    Dim oApp As Object
    
    strPath = vbNullString
    
    Set oApp = CallerApplicationObject
    
    If oApp Is Nothing Then
        'then in the twin basic ide?
        ActiveVBAProjectFolderPath = ThisLibFolderPath
        Exit Function
    End If

    Dim VBAIsTrusted As Boolean
    VBAIsTrusted = False
    On Error Resume Next
    VBAIsTrusted = (oApp.VBE.VBProjects.Count) > 0
    On Error GoTo 0

    If Not VBAIsTrusted Then
        MsgBox "Error: No Access to VB Project" & vbLf & vbLf & "File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust Access to VBA project object model", vbCritical
    End If

    'if the parent document holding the active vba project has not yet been saved, then Application.VBE.ActiveVBProject.Filename
    'will throw an error so trap and report below...
    
    On Error Resume Next
    strPath = oApp.VBE.ActiveVBProject.Filename
    On Error GoTo 0

    If strPath <> vbNullString Then
        strPath = fso.GetParentFolderName(strPath)
        ActiveVBAProjectFolderPath = strPath
    Else
        Err.Raise 1, "WebShared", "Error: Attempting to reference a folder/file path relative to the parent document location of this active code project - save the parent document first."
    End If
End Function