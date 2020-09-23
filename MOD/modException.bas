Attribute VB_Name = "modException"
' Nome del Progetto: Radio Streaming and Recorder v1.0.1b © 2010/2011
' ****************************************************************************************************
' Copyright © 2010 - 2011 Salvo Cortesiano - Società: http://www.netshadows.it/
' Tutti i diritti riservati, Indirizzo Internet: http://www.netshadows.it/
' Blog / Forum: http://www.netshadows.it/leombredellarete/forum
' ****************************************************************************************************
' Attenzione: Questo programma per computer è protetto dalle vigenti leggi sul copyright
' e sul diritto d'autore. Le riproduzioni non autorizzate di questo codice, la sua distribuzione
' la distribuzione anche parziale è considerata una violazione delle leggi, e sarà pertanto
' perseguita con l'estensione massima prevista dalla legge in vigore.
' ****************************************************************************************************

Option Explicit

Public PID As Long

Public AppendSeek As Boolean
Public sCheckDelete As Boolean
Public sCheckDivide As Boolean
Public schkSaveStation As Boolean
Public sWriteTagOfTrack As Boolean

Public ID3TagComments As String
Public ID3TagEncodedBy As String
Public ID3TagCopyrightInfo As String
Public ID3TagLanguages As String

' .... ["] Constant Quote
Public Const sQuote As String = """"

' .... Class INI
Public INI As New clsINI

Public ShutDown As New clsExit

' .... Open Application, URL, Files o Folders
Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)

Public Enum SW_SHOW_MODE
    SW_HIDE = 0
    SW_MAXIMIZE = 3
    SW_MINIMIZE = 6
    SW_NORMAL = 1
End Enum

' ... Init control's XP or Vista
Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public m_hMod As Long
'Public m_hMod2 As Long

'ITALIAN SORRY :(
' Di conseguenza possiamo risolvere questo problema semplicemente ignorandolo.
' L'unico problema in questo modo è che l'applicazione continua a inviare messaggi al sistema e danno origine
' alla nota finestra che invita a trasmettere le informazioni del Microsoft sul problema:
Public Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Public Const SEM_FAILCRITICALERRORS = &H1
Public Const SEM_NOGPFAULTERRORBOX = &H2
Public Const SEM_NOOPENFILEERRORBOX = &H8000&

' ... Exception Handler (Call the Stack)
Public Const MySEH_ERROR = 12345&

Public Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)
Public Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long

Private Const EXCEPTION_CONTINUE_EXECUTION = -1
Private Const EXCEPTION_CONTINUE_SEARCH = 0
Private Const EXCEPTION_EXECUTE_HANDLER = 1

Public Declare Sub DebugBreak Lib "kernel32" ()
Private m_bInIDE As Boolean

' .... Verify if the File exist
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Long

Private Const GENERIC_READ As Long = &H80000000
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Const MAX_PATH As Long = 260

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

' .... Parse Path of File
Public Enum Extract
    [Only_Extension] = 0
    [Only_FileName_and_Extension] = 1
    [Only_FileName_no_Extension] = 2
    [Only_Path] = 3
End Enum

Private stripMyString As String

' .... For the Connection
Private Const INTERNET_CONNECTION_MODEM As Long = &H1
Private Const INTERNET_CONNECTION_LAN As Long = &H2
Private Const INTERNET_CONNECTION_PROXY As Long = &H4
Private Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
Private Const INTERNET_CONNECTION_OFFLINE As Long = &H20
Public ConnectieType As String
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpSFlags As Long, ByVal dwReserved As Long) As Long

Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

' .... Hide App to Windows Task
Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Declare Function GetCurrentProcessId Lib "kernel32" () As Long

' .... Play Sound Resource
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long

Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4
Private Const SND_ALIAS = &H10000
Private Const SND_FILENAME = &H20000
Private Const SND_RESOURCE = &H40004
Private Const SND_ALIAS_ID = &H110000
Private Const SND_ALIAS_START = 0
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10
Private Const SND_VALID = &H1F
Private Const SND_NOWAIT = &H2000
Private Const SND_VALIDFLAGS = &H17201F
Private Const SND_RESERVED = &HFF000000
Private Const SND_TYPE_MASK = &H170007

Private Const WAVERR_BASE = 32
Private Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)
Private Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)
Private Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)
Private Const WAVERR_SYNC = (WAVERR_BASE + 3)
Private Const WAVERR_LASTERROR = (WAVERR_BASE + 3)

Private m_snd() As Byte

' .... Find and Close the Prev Instance of Application
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const PROCESS_TERMINATE As Long = &H1

' .... Constants that are used by the API
Public Const WM_CLOSE = &H10
'Public Const SYNCHRONIZE = &H100000 ' .... This const is OK?
Public Const INFINITE = &HFFFFFFFF
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Const GW_HWNDNEXT = 2
Public mWnd As Long
Private Sub InitControlsCtx()
 On Local Error GoTo ErrorHandler
      Dim iccex As tagInitCommonControlsEx
      With iccex
          .lngSize = LenB(iccex)
          .lngICC = ICC_USEREX_CLASSES
      End With
      InitCommonControlsEx iccex
Exit Sub
ErrorHandler:
    Err.Clear
End Sub

Public Function MyExceptionHandler(lpEP As Long) As Long
   Dim lRes As VbMsgBoxResult
   lRes = MsgBox("Exception Handler!" & vbCrLf & "Ignore, Close, or Call the Debugger?", _
   vbAbortRetryIgnore Or vbCritical, App.Title & "Exception Handler")
   Select Case lRes
      Case vbIgnore
         If InIDE Then
            Stop
            MyExceptionHandler = EXCEPTION_CONTINUE_EXECUTION
            On Error GoTo 0
            Err.Raise MySEH_ERROR, "MyExceptionHandler", "EXCEPTION-" & MySEH_ERROR
         Else
            MyExceptionHandler = EXCEPTION_CONTINUE_SEARCH
            Err.Raise MySEH_ERROR, "MyExceptionHandler", "EXCEPTION-" & MySEH_ERROR
         End If
       Case vbAbort
            MyExceptionHandler = EXCEPTION_EXECUTE_HANDLER
            Err.Raise MySEH_ERROR, "MyExceptionHandler", "EXCEPTION-" & MySEH_ERROR
       Case vbRetry
            MyExceptionHandler = EXCEPTION_CONTINUE_SEARCH
            Err.Raise MySEH_ERROR, "MyExceptionHandler", "EXCEPTION-" & MySEH_ERROR
   End Select
End Function

Public Sub Main()
    On Local Error GoTo ErrorHandler
    
    ' .... Verify the Instance
    If App.PrevInstance Then
        WriteErrorLogs -101010, "Application Runnig...", "Form_Load {Sub: Form_Load}", True, True
    If MsgBox("An hoter instance is runnig." & vbLf & vbLf & "You want to close the First instance before running a new instance?", _
        vbYesNo + vbQuestion, App.Title) = vbNo Then End
        If ForceClose = False Then
                MsgBox "Unable to Kill the current Process of {" & App.EXEName & "}!" _
                & vbCrLf & "Ok to Continue in a New Instance of the Program!!!", vbExclamation, App.Title
        End If
    End If
    
    ' .... Exception Handler' = (Call the stack)
    SetUnhandledExceptionFilter AddressOf MyExceptionHandler
    
    ' .... Subclass the SO
    SetErrorMode SEM_NOGPFAULTERRORBOX
    
    ' .... Load the Library {shell32.dll}
    m_hMod = LoadLibrary("shell32.dll")
    
    ' .... Init the Controls
    InitControlsCtx
    
    ' .... Load and show the Form
    Load frmStreamingRadio
    frmStreamingRadio.Show
    
Exit Sub
ErrorHandler:
    WriteErrorLogs Err.Number, Err.Description, "ModMain {Sub: Main}", True, True
        Err.Clear
    End
End Sub

Public Sub WriteErrorLogs(strErrNumber As Variant, strErrDescription As String, Optional strErrSource As String = "Unknow", _
                        Optional visError As Boolean = True, Optional errAppend As Boolean = True)
    
    Dim FileNum As Variant: Dim sFN As String
    
    On Error GoTo ErrorHandler
    
    Call PlaySoundResource(101)
    
    FileNum = FreeFile
    
    sFN = App.Path + "\" + App.EXEName + "_errs.log"
    
    If Dir$(sFN, vbNormal) = Empty Then
        Open sFN For Output As FileNum
            Print #FileNum, Tab(5); "Log Error Generate from [" & App.EXEName & "]..."
            Print #FileNum, Tab(5); Format(Now, "Long Date") & "/" & Time
            Print #FileNum, Tab(5); "----------------------------------------------------------------------------"
            Print #FileNum, Tab(5); ""
            Print #FileNum, Tab(5); ""
            Print #FileNum, Tab(5); "*/___ LOG STARTED..."
            Print #FileNum, Tab(5); ""
        Close FileNum
    End If
    
    If errAppend Then
        Open sFN For Append As FileNum
    Else
        Open sFN For Output As FileNum
    End If
    
        Print #FileNum, Tab(5); Format(Now, "Long Date") & "/" & Time
        Print #FileNum, Tab(5); "Error #" & CStr(strErrNumber)
        Print #FileNum, Tab(5); "Description: " & CStr(strErrDescription)
        Print #FileNum, Tab(5); "Source: " & CStr(strErrSource)
        Print #FileNum, Tab(5); ""
        Print #FileNum, Tab(5); ""
        Close FileNum
        
        If visError Then
            MsgBox "Error #" & CStr(strErrNumber) & "." & vbCrLf & "Description: " & CStr(strErrDescription) _
            & vbCrLf & "Source: " & CStr(strErrSource) & vbCrLf & vbCrLf & "For more info, see the Log file!", vbCritical, App.Title
        End If
        
    Exit Sub
    
ErrorHandler:
        MsgBox "Unexpected Error #" & Err.Number & "!" & vbCrLf & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub

Public Property Get InIDE() As Boolean
   Debug.Assert (pIsInIDE)
   InIDE = m_bInIDE
End Property

Public Property Get pIsInIDE() As Boolean
   m_bInIDE = True
   pIsInIDE = True
End Property

Public Sub DelayTime(ByVal Second As Long, Optional ByVal Refresh As Boolean = False)
    On Error Resume Next
    Dim Start As Date
    Start = Now
    Do
    If Refresh Then DoEvents
    Loop Until DateDiff("s", Start, Now) >= Second
End Sub

Public Function FileExists(sSource As String) As Boolean
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   hFile = FindFirstFile(sSource, WFD)
   FileExists = hFile <> INVALID_HANDLE_VALUE
   Call FindClose(hFile)
End Function

Public Function GetFilePath(ByVal FileName As String, strExtract As Extract) As String
    Select Case strExtract
        'Extract only extension of File
        Case 0
            GetFilePath = Mid$(FileName, InStrRev(FileName, ".", , vbTextCompare) + 1)
        'Extract only Filename and Extension
        Case 1
            GetFilePath = Mid$(FileName, InStrRev(FileName, "\") + 1, Len(FileName))
        'Extract only FileName
        Case 2
            GetFilePath = StripUndo(Mid$(FileName, InStrRev(FileName, "\", , vbTextCompare) + 1))
        'Extract only Path
        Case 3
            GetFilePath = Mid$(FileName, 1, InStrRev(FileName, "\", , vbTextCompare) - 1)
        End Select
End Function

Private Function StripUndo(ByVal FileName As String) As String
    Dim i As Integer
    Dim sTmp As String
    On Error Resume Next
sTmp = Mid$(FileName, i + 1, Len(FileName))
    For i = 1 To Len(sTmp)
      If Mid$(sTmp, i, 1) = "." Then
        Exit For
    Else
        stripMyString = Mid$(FileName, i + 2, Len(FileName))
    End If
Next
     StripUndo = Left(sTmp, i - 1)
End Function

Public Function StripLeft(strString As String, strChar As String, Optional sLeftsRight As Boolean = True) As String
  On Local Error Resume Next
  Dim i As Integer
    If sLeftsRight Then
        For i = 1 To Len(strString)
            If Mid$(strString, i, 1) = strChar Then
                    StripLeft = Mid$(strString, 1, i - 1)
                Exit For
            End If
        Next
    Else
        For i = (Len(strString)) To 1 Step -1
        If Mid$(strString, i, 1) = strChar Then
                StripLeft = Mid$(strString, i + 1, Len(strString) - i + 1)
            Exit For
        End If
    Next
End If
End Function

Public Function GetNetConnectString() As Boolean
   Dim dwFlags As Long: GetNetConnectString = False: ConnectieType = Empty
   On Local Error GoTo ErrorHandler
   If InternetGetConnectedState(dwFlags, 0&) Then
      If dwFlags And INTERNET_CONNECTION_LAN Then
        GetNetConnectString = True
        ConnectieType = "Connet via Lan!"
      End If
      If dwFlags And INTERNET_CONNECTION_PROXY Then
        GetNetConnectString = True
        ConnectieType = "Connet via Proxy Server!"
      End If
      If dwFlags And INTERNET_CONNECTION_MODEM Then
         GetNetConnectString = True
         ConnectieType = "Connet via Modem!"
      End If
      If dwFlags And INTERNET_CONNECTION_OFFLINE Then
        GetNetConnectString = False
        ConnectieType = "Connection Nothing!"
      End If
      If dwFlags And INTERNET_CONNECTION_MODEM_BUSY Then
        GetNetConnectString = False
        ConnectieType = "Connection Busy!"
      End If
   Else
        GetNetConnectString = False
        ConnectieType = "Connection NONE!"
   End If
Exit Function
ErrorHandler:
    GetNetConnectString = False
    ConnectieType = "Error to retrive Connection type: Error# " & Err.Number
    WriteErrorLogs Err.Number, Err.Description, "ModException {Function: GetNetConnectString}", False, True
End Function

Public Function sShutDown(ByVal exitWindows As EnumExitWindows) As Boolean
    On Local Error GoTo ErrorHandler
    If exitWindows = WE_SHUTDOWN Then
        ShutDown.exitWindows WE_SHUTDOWN
    ElseIf exitWindows = WE_REBOOT Then
        ShutDown.exitWindows WE_REBOOT
    ElseIf exitWindows = WE_LOGOFF Then
        ShutDown.exitWindows WE_LOGOFF
    ElseIf exitWindows = WE_POWEROFF Then
        ShutDown.exitWindows WE_POWEROFF
  End If
  sShutDown = True
Exit Function
ErrorHandler:
    WriteErrorLogs Err.Number, Err.Description, "modException {Function: sShutDown}", True, True
        sShutDown = False
    Err.Clear
End Function

Public Function ReplaceChars(StrIN As String, Optional StripChar As String = "", Optional ReplaceChar As String = "") As String
    Dim x As Integer
    On Error Resume Next
    x = 1
    If StripChar <> "" Then
        Do Until x <= 0 Or StripChar = ReplaceChar
            x = InStr(1, StrIN, StripChar)
            If x > 0 Then StrIN = Left$(StrIN, x - 1) & ReplaceChar & Right$(StrIN, Len(StrIN) - (x - 1) - Len(StripChar))
        DoEvents
        Loop
    Else
        For x = 1 To Len(StrIN)
            If x > Len(StrIN) Then Exit For
            If Asc(Mid$(StrIN, x, 1)) < 32 Or Asc(Mid$(StrIN, x, 1)) > 126 Then
                StrIN = Left$(StrIN, x - 1) & ReplaceChar & Right$(StrIN, Len(StrIN) - (x - 1) - 1)
                If ReplaceChar = "" Then x = x - 1
            End If
        Next
    End If
    ReplaceChars = StrIN
End Function

Public Function PlaySoundResource(ByVal SndID As Long) As Long
   Const flags = SND_MEMORY Or SND_ASYNC Or SND_NODEFAULT
   On Error GoTo ErrorHandler
   DoEvents
   m_snd = LoadResData(SndID, "WAV")
   PlaySoundResource = PlaySoundData(m_snd(0), 0, flags)
Exit Function
ErrorHandler:
    Err.Clear
End Function

Public Function InstanceToWnd(ByVal target_pid As Long) As Long
    Dim test_hwnd As Long: Dim test_pid As Long: Dim test_thread_id As Long
    On Local Error Resume Next
    ' .... Find the first window
    test_hwnd = FindWindow(ByVal 0&, ByVal 0&)
    Do While test_hwnd <> 0
        ' .... Check if the window isn't a child
        If GetParent(test_hwnd) = 0 Then
            ' .... Get the window's thread
            test_thread_id = GetWindowThreadProcessId(test_hwnd, test_pid)
            If test_pid = target_pid Then
                InstanceToWnd = test_hwnd
                Exit Do
            End If
        End If
        '.... Retrieve the next window
        test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    Loop
End Function

Public Function ForceClose() As Boolean
    Dim hProcess As Long
    
    On Local Error GoTo ErrorHandler
    
    If INI.GetKeyValue("PROCESSID", "PID") <> Empty Then _
    PID = INI.GetKeyValue("PROCESSID", "PID") Else PID = 0
    
    If PID = 0 Then
            ForceClose = False
        Exit Function
    Else
        ForceClose = True
    End If
    
    hProcess = OpenProcess(PROCESS_TERMINATE, 0, PID)
    TerminateProcess hProcess, 0&
    
Exit Function
ErrorHandler:
    WriteErrorLogs Err.Number, Err.Description, "FormMain {Function: ForceClose}", True, True
        ForceClose = False
    Err.Clear
End Function

Public Function CreateRes(ByVal resID As Integer, resType As String, OriginalName As String) As Boolean
    Dim sFile As String: Dim b() As Byte: Dim iFile As Integer
    On Error GoTo ErrorHandler
    
    b = LoadResData(resID, resType)
    sFile = App.Path + "\" + OriginalName
    iFile = FreeFile
    Open sFile For Binary Access Write Lock Read As #iFile
    Put #iFile, , b
    Close #iFile
    iFile = 0
    CreateRes = True: Exit Function
ErrorHandler:
    CreateRes = False: Err.Clear: Exit Function
End Function

Public Function WindowToProcessId(ByVal hwnd As Long) As Long
    Dim lpProc As Long
    Call GetWindowThreadProcessId(hwnd, lpProc)
    WindowToProcessId = lpProc
End Function

Public Function InstallLibrary(Optional sstate As SW_SHOW_MODE = SW_NORMAL) As Boolean
    On Local Error GoTo ErrorHandler
    ' .... FLAG to TRUE
    InstallLibrary = True
    
    ' .... Extract the Library
    If CreateRes(101, "DLL", "axvlc.dll") = False Then InstallLibrary = False
    If CreateRes(102, "DLL", "libvlc.dll") = False Then InstallLibrary = False
    If CreateRes(103, "DLL", "libvlccore.dll") = False Then InstallLibrary = False
    If CreateRes(101, "CMD", "vlcsilentinstall.cmd") = False Then InstallLibrary = False
    
    ' .... Create VLC Path
    Call MakeNewDir(frmStreamingRadio.SpecialFolders.SpecialFolderPath(CSIDL_PROGRAM_FILES) + "\VideoLAN\VLC")
    
    ' .... Move the Library to VLC Path
    Dim ret As Long
    
    ret = CopyFile(App.Path + "\axvlc.dll", frmStreamingRadio.SpecialFolders.SpecialFolderPath(CSIDL_PROGRAM_FILES) _
    + "\VideoLAN\VLC\axvlc.dll", True)
    
    If ret = 0 Then
        ret = CopyFile(App.Path + "\libvlc.dll", frmStreamingRadio.SpecialFolders.SpecialFolderPath(CSIDL_PROGRAM_FILES) _
        + "\VideoLAN\VLC\libvlc.dll", True)
    Else
        InstallLibrary = False
    End If
    
    If ret = 0 Then
        ret = CopyFile(App.Path + "\libvlccore.dll", frmStreamingRadio.SpecialFolders.SpecialFolderPath(CSIDL_PROGRAM_FILES) _
        + "\VideoLAN\VLC\libvlccore.dll", True)
    Else
        InstallLibrary = False
    End If
    
    If ret <> 0 Then InstallLibrary = False
    
    If Dir$(App.Path + "\vlcsilentinstall.cmd") <> Empty Then _
    ShellExecute 0&, vbNullString, App.Path + "\vlcsilentinstall.cmd", vbNullString, _
        "c:\", sstate
    
    Call DelayTime(3, True)
    MsgBox "VLC library installation success!", vbInformation, App.Title
    
    If Dir$(App.Path + "\vlcsilentinstall.cmd") <> Empty Then Call _
    Kill(App.Path + "\vlcsilentinstall.cmd")
Exit Function
ErrorHandler:
    WriteErrorLogs Err.Number, Err.Description, "modException {Function: InstallLibrary}", True, True
        InstallLibrary = False
    Err.Clear
End Function

Public Sub MakeNewDir(newDir As String)
    Dim NewLen As Integer: Dim DirLen As Integer: Dim maxLen As Integer
    NewLen = 4: maxLen = Len(newDir)
    If Right(newDir, 1) <> "\" Then
        newDir = newDir + "\": maxLen = maxLen + 1
    End If
    On Error GoTo DirError
MakeNext:
    DirLen = InStr(NewLen, newDir, "\"): MkDir Left(newDir, DirLen - 1): NewLen = DirLen + 1
    If NewLen >= maxLen Then Exit Sub
    GoTo MakeNext
DirError:
    Resume Next
End Sub
