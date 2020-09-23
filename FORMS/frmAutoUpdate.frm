VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmAutoUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Radio Streaming-AutoUpdate 1.0.2"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAutoUpdate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   615
      MouseIcon       =   "frmAutoUpdate.frx":3D52
      TabIndex        =   23
      Top             =   2310
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Height          =   165
      Left            =   120
      TabIndex        =   22
      Top             =   2745
      Width           =   5610
   End
   Begin VB.CommandButton cmdmore 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   75
      MouseIcon       =   "frmAutoUpdate.frx":3EA4
      TabIndex        =   8
      ToolTipText     =   "Show more..."
      Top             =   2310
      Width           =   450
   End
   Begin RadioStreamingRecorder.ProgressBar PB 
      Height          =   210
      Left            =   1845
      Top             =   2385
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   370
      FromColor       =   16744576
      ToColor         =   8388608
      BorderStyle     =   2
      BackColor       =   14215660
   End
   Begin VB.Timer tSetup 
      Enabled         =   0   'False
      Left            =   165
      Top             =   1275
   End
   Begin VB.CommandButton cmdControlla 
      Caption         =   "Download Update"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3270
      MouseIcon       =   "frmAutoUpdate.frx":3FF6
      TabIndex        =   7
      Top             =   195
      Width           =   2370
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check Update"
      Height          =   375
      Left            =   1005
      MouseIcon       =   "frmAutoUpdate.frx":4148
      TabIndex        =   3
      Top             =   195
      Width           =   2205
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5190
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label SourceLabel 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   105
      TabIndex        =   25
      Top             =   5925
      Width           =   5610
   End
   Begin VB.Label lblTot 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Label StatusLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "##"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   615
      Left            =   105
      TabIndex        =   21
      Top             =   5280
      Width           =   5625
   End
   Begin VB.Label RateLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1845
      TabIndex        =   20
      Top             =   3915
      Width           =   3960
   End
   Begin VB.Label TimeLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1875
      TabIndex        =   19
      Top             =   2910
      Width           =   3825
   End
   Begin VB.Label ToLabel 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   660
      Left            =   1590
      TabIndex        =   18
      Top             =   3270
      Width           =   4155
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Tempo Rimanente:"
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   105
      TabIndex        =   17
      Top             =   2910
      Width           =   1740
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Trasferimento:"
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   105
      TabIndex        =   16
      Top             =   3915
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Download Da:"
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   105
      TabIndex        =   15
      Top             =   3270
      Width           =   1365
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "File Size:"
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   105
      TabIndex        =   14
      Top             =   4365
      Width           =   1110
   End
   Begin VB.Label lblSize 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1350
      TabIndex        =   13
      Top             =   4365
      Width           =   4365
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Disk Free Size:"
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   105
      TabIndex        =   12
      Top             =   4635
      Width           =   1710
   End
   Begin VB.Label DiskFree 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1755
      TabIndex        =   11
      Top             =   4635
      Width           =   3990
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Disk Size:"
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   90
      TabIndex        =   10
      Top             =   4950
      Width           =   1770
   End
   Begin VB.Label TotalSize 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1905
      TabIndex        =   9
      Top             =   4950
      Width           =   3840
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   135
      Picture         =   "frmAutoUpdate.frx":429A
      Top             =   150
      Width           =   480
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Whats New?:"
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1635
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Dimension:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1590
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Data Update:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1665
   End
   Begin VB.Label Label3 
      Caption         =   "n.a"
      Height          =   600
      Left            =   1920
      TabIndex        =   2
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "n.a"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "n.a"
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   960
      Width           =   3735
   End
End
Attribute VB_Name = "frmAutoUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim CancelSearch As Boolean
Dim strSetupFile As String
Dim MyVersion As String

Dim Update As String
Dim T1() As String

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Const MAX_PATH = 260

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Long
End Type


Private Type WIN32_FIND_DATA
  dwFileAttributes  As Long
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

Private Const SW_NORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10
Private Const SW_SHOWNOACTIVATE As Long = 4
Private Const SW_SHOWNORMAL As Long = 1

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCheck_Click()
    ' .... Verifica la versione situata sul Server (HTTP)
    If CheckForUpDate("http://www.netshadows.it/RSRver.txt") = True Then
    End If
End Sub


Private Sub cmdControlla_Click()
' Scarica l'aggiornamento del Programma
    If cmdControlla.Caption = "Download Update" Then
        CancelSearch = False
        cmdControlla.Caption = "Cancel Update"
        cmdCancel.Enabled = False
        cmdCheck.Enabled = False
        '//* OPPURE QUESTI POSSIBILI DOWNLOAD:
        '           Vista_Icons_Pack.rar
        '           Pack_Vista_Inspirat_11.rar
        '           tavolozza colori.zip
        '           sfondo web.jpg
        '*\\ --------------------------------------
        
        Call DownloadFile("http://www.netshadows.it/RSRv102.exe", App.Path + "\RSRv103.exe", , , "RSRv103.exe")
    ElseIf cmdControlla.Caption = "Cancel Update" Then
        CancelSearch = True
        cmdControlla.Caption = "Download Update"
        StatusLabel = "Download del File Interrotto dall'Utente!"
        PB.Value = 0
        cmdCancel.Enabled = True
        cmdCheck.Enabled = True
        cmdControlla.Enabled = False
        If Dir(App.Path & "\" & strSetupFile) <> "" Then
        On Error Resume Next
            Close #3
            Kill (App.Path & "\" & strSetupFile)
        End If
    End If
End Sub


Private Sub cmdmore_Click()
    If cmdControlla.Caption = "Cancel Update" Then
    If cmdmore.Caption = "6" Then
        cmdmore.Caption = "5"
        cmdmore.ToolTipText = "Hide more..."
        frmAutoUpdate.Height = 5790
    ElseIf cmdmore.Caption = "5" Then
        frmAutoUpdate.Height = 3165
        cmdmore.Caption = "6"
        cmdmore.ToolTipText = "Show more..."
    End If
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyEscape Then cmdCancel = True
End Sub

Private Sub Form_Load()
    MyVersion = CalcVersion(App.Major & App.Minor)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmAutoUpdate = Nothing
End Sub

Private Sub tSetup_Timer()
    On Error Resume Next
    If CancelSearch = True Then
            tSetup.Enabled = False
            If Dir(App.Path + "\" + strSetupFile) <> "" Then Kill (App.Path + "\" & strSetupFile)
        Exit Sub
    End If
    If Dir(App.Path + "\" + strSetupFile) <> "" Then ShellExecute 0&, _
    vbNullString, App.Path + "\" + strSetupFile, vbNullString, App.Path, SW_SHOWNORMAL
    tSetup.Enabled = False
End Sub

Private Function CalcVersion(rlVersionNo As Long) As String
   Dim nMajor As Integer
   Dim nMinor As Integer
   On Error Resume Next
   DoEvents
   nMajor = CInt(rlVersionNo / &H10000)
   nMinor = CInt(rlVersionNo And &HFFFF&)
   CalcVersion = CStr(nMajor) & "." & LTrim(CStr(nMinor))
End Function

Private Function CheckForUpDate(strURL As String) As Boolean
    On Error GoTo ErrorHandler
    Update = Inet1.OpenURL(strURL)
    T1 = Split(Update, "#")
    If T1(0) > App.Revision Then
        CheckForUpDate = True
        If MsgBox("New Update available on the server:" & vbCrLf & vbCrLf & "New Version: " & T1(0) & vbCrLf & "Download file da: " & T1(1) _
            & Chr(13) & Chr(13) & "Download and Install New Version?", vbYesNo + vbInformation + vbDefaultButton1, "Confirm Update") = vbYes Then
            cmdControlla.Enabled = True
            cmdControlla = True
        End If
    ElseIf T1(0) <= App.Revision Then
            MsgBox "No Update available on the server!", vbExclamation, App.Title
        CheckForUpDate = False
    End If
Exit Function
ErrorHandler:
        CheckForUpDate = False
    Err.Clear
End Function

Private Function DiskFreeSpace(strDrive As String) As Double
Dim SecPerCluster  As Long: Dim BytesPerSec As Long
Dim NumbFreeClusters As Long: Dim TotNumbOfClusters As Long
Dim TotalFREE As Currency: Dim Total As Currency
Dim status As Long: Dim qw As Currency
Dim qe As Currency: Dim qr As Currency
Dim qt As Currency: Dim qa As Currency
Dim qFR As Currency: Dim qTOT As Currency
On Error Resume Next
DoEvents
status = GetDiskFreeSpace(strDrive, SecPerCluster, BytesPerSec, NumbFreeClusters, TotNumbOfClusters)
qw = SecPerCluster
qe = BytesPerSec
qr = NumbFreeClusters
qt = TotNumbOfClusters
qa = status
DoEvents
lblTot.Caption = qt * (qe * qw) ' Totale Spazio  del Disco
DiskFreeSpace = qr * (qe * qw)
End Function

Private Function DownloadFile(strURL As String, strDestination As String, Optional UserName As String = Empty, Optional password As String = Empty, Optional strFileName As String = Empty) As Boolean

Const CHUNK_SIZE As Long = 1024
Const ROLLBACK As Long = 4096

Dim bData() As Byte
Dim blnResume As Boolean
Dim intFile As Integer
Dim lngBytesReceived As Long
Dim lngFileLength As Long
Dim lngX
Dim sglLastTime As Single
Dim sglRate As Single
Dim sglTime As Single
Dim strFile As String
Dim strHeader As String
Dim strHost As String
Dim FileData As WIN32_FIND_DATA
Dim ftime As SYSTEMTIME
Dim fsize As WIN32_FIND_DATA
Dim tCreation As String
Dim tLastAccessTime As String
Dim tModTime As String
Dim tSize As String
Dim tVersion As String

strSetupFile = strFileName

On Local Error GoTo InternetErrorHandler
CancelSearch = False
strFile = ReturnFileOrFolder(strDestination, True)
strHost = ReturnFileOrFolder(strURL, True, True)
SourceLabel = Empty
TimeLabel = Empty
ToLabel = Empty
RateLabel = Empty
StartDownload:
If blnResume Then
    StatusLabel = "Resuming download..."
    lngBytesReceived = lngBytesReceived - ROLLBACK
    If lngBytesReceived < 0 Then lngBytesReceived = 0
Else
    StatusLabel = "Informazioni sul Fle..."
End If
DoEvents
With Inet1
    .url = strURL
    .UserName = UserName
    .password = password
    .Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
    While .StillExecuting
        DoEvents
        If CancelSearch Then GoTo ExitDownload
    Wend
    StatusLabel = "Salvataggio:"
    SourceLabel = FitText(SourceLabel, strHost & " da " & .RemoteHost)
    ToLabel = FitText(ToLabel, strDestination)
    strHeader = .GetHeader
    ' Info sul File
    FileData = Findfile(Inet1.url)
    fsize = Findfile(Inet1.url)
    ' Creato
    Call FileTimeToSystemTime(FileData.ftCreationTime, ftime)
    'tCreation = "Creato: " & CStr(Format(ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear, "Long Date") & " / " _
                & Format(ftime.wHour + 1 & "." & ftime.wMinute & "." & ftime.wSecond, "Long Time"))
    ' Ultima modifica
    Call FileTimeToSystemTime(FileData.ftLastWriteTime, ftime)
    'tLastAccessTime = "Ultima modifica: " & CStr(Format(ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear, "Long Date") & " / " _
                        & Format(ftime.wHour + 1 & "." & ftime.wMinute & "." & ftime.wSecond, "Long Time"))
    ' Ultimo accesso
    Call FileTimeToSystemTime(FileData.ftLastAccessTime, ftime)
    'tModTime = "Ultimo accesso: " & CStr(Format(ftime.wDay & "/" & ftime.wMonth & "/" & ftime.wYear, "Long Date") & " / " _
                & Format(ftime.wHour + 1 & "." & ftime.wMinute & "." & ftime.wSecond, "Long Time"))
    ' Dimensioni
    'tSize = "Dimensioni: " & FormatFileSize(fsize.nFileSizeHigh + fsize.nFileSizeLow)
    
    'MsgBox tCreation & vbCrLf & tLastAccessTime & vbCrLf & tModTime & vbCrLf & tSize _
    & vbCrLf & vbCrLf & frmDownload.Inet1.URL, vbInformation, "Info File"
    
End With
Select Case Mid(strHeader, 10, 3)
    Case "200"
        If blnResume Then
            Kill strDestination
            If MsgBox("Impossibile riesumare il Download." & vbCr & vbCr & "Vuoi comunque continuare?", _
                     vbExclamation + vbYesNo, "Resume Download") = vbYes Then
                    blnResume = False
                Else
                    CancelSearch = True
                    GoTo ExitDownload
                End If
            End If
    Case "206"  ' 206=Contenuto Parziale
    Case "204"
        MsgBox "Niente da scaricare!", vbInformation, "Nessun Download"
        CancelSearch = True
        GoTo ExitDownload
    Case "401"
        MsgBox "Autorizzazione (negata) Download del file fallito!", vbCritical, "Non autorizzato"
        CancelSearch = True
        GoTo ExitDownload
    Case "404"  ' File non trovato
        MsgBox "File " & """" & strFileName & """" & " non presente sul server o agiornamento non disponibile!", vbCritical, "File non trovato"
        CancelSearch = True
        GoTo ExitDownload
    Case vbCrLf
    MsgBox "Impossibile stabilire una connessione." & vbCr & vbCr & "Verificare le Impostazioni di connessione alla rete e riprovare", _
               vbExclamation, "Impossibile Connettersi"
        CancelSearch = True
        GoTo ExitDownload
    Case Else
        strHeader = Left(strHeader, InStr(strHeader, vbCr))
        If strHeader = Empty Then strHeader = "<nothing>"
        MsgBox "Il server ha risposto in questo modo:" & vbCr & vbCr & strHeader, vbCritical, "Errore Downloading File"
        CancelSearch = True
        GoTo ExitDownload
End Select
If blnResume = False Then
    sglLastTime = Timer - 1
    strHeader = Inet1.GetHeader("Content-Length")
    lngFileLength = Val(strHeader)
    If lngFileLength = 0 Then
        GoTo ExitDownload
    End If
End If
lblSize.Caption = FormatFileSize(lngFileLength)
DiskFree.Caption = FormatFileSize(DiskFreeSpace(Left(strDestination, InStr(strDestination, "\"))))
TotalSize.Caption = FormatFileSize(lblTot.Caption)
    Label1.Caption = Format(Now, "Long Date")
    Label2.Caption = lblSize.Caption
    Label3.Caption = "Bug solved for MP3 Tag's..."
If Mid(strDestination, 2, 2) = ":\" Then
    If DiskFreeSpace(Left(strDestination, InStr(strDestination, "\"))) < lngFileLength Then
        MsgBox "Non c'è abbastanza spazio su questo Disco per il Download del File." & vbCr & vbCr & "Provare a liberare un po di spazio e ritentare.", _
               vbCritical, "Spazio su Disco Insufficiente"
        GoTo ExitDownload
    End If
End If
DoEvents
If blnResume = False Then lngBytesReceived = 0
On Local Error GoTo FileErrorHandler
strHeader = ReturnFileOrFolder(strDestination, False)
If Dir(strHeader, vbDirectory) = Empty Then
    MkDir strHeader
End If
    intFile = FreeFile()
    Open strDestination For Binary Access Write As #3
    If blnResume Then Seek #3, lngBytesReceived + 1
    ' ProgressBar
    PB.Steps = FormatPercentage(lngFileLength)
    Do
    bData = Inet1.GetChunk(CHUNK_SIZE, icByteArray)
    Put #3, , bData
    If CancelSearch Then Exit Do
    lngBytesReceived = lngBytesReceived + UBound(bData, 1) + 1
    sglRate = lngBytesReceived / (Timer - sglLastTime)
    sglTime = (lngFileLength - lngBytesReceived) / sglRate
    TimeLabel = FormatTime(sglTime) & " (" & FormatFileSize(lngBytesReceived) & " su " & FormatFileSize(lngFileLength) & " copiati)"
    RateLabel = FormatFileSize(sglRate, "###.0") & "/sec"
    ' Avanzamento
   PB.Value = FormatPercentage(lngBytesReceived)
Loop While UBound(bData, 1) > 0
Close #3
    If CancelSearch <> True Then
        MsgBox "Download dell'aggiornamento completato! Verrà lanciato il Programma di (Setup), Ok per continuare!", vbInformation, App.Title
        strSetupFile = strFileName
        cmdControlla.Caption = "Download Update"
        cmdControlla.Enabled = False
        cmdCheck.Enabled = True
        cmdCancel.Enabled = True
        tSetup.Interval = 100
        tSetup.Enabled = True
        frmAutoUpdate.Height = 3165
        cmdmore.Caption = "6"
        cmdmore.ToolTipText = "Show more..."
    End If
ExitDownload:
If lngBytesReceived = lngFileLength And CancelSearch = False Then
    StatusLabel = "Download del File Completato!"
        cmdControlla.Caption = "Download Update"
        cmdControlla.Enabled = False
        cmdCheck.Enabled = True
        cmdCancel.Enabled = True
        frmAutoUpdate.Height = 3165
        cmdmore.Caption = "6"
        cmdmore.ToolTipText = "Show more..."
    PB.Value = 0
    Sleep (0.7)
    DownloadFile = True
    
    GoTo Cleanup
    
Else
    If CancelSearch = True Then
        StatusLabel = "Download del File Interrotto!"
        cmdControlla.Caption = "Download Update"
        cmdControlla.Enabled = False
        cmdCheck.Enabled = True
        cmdCancel.Enabled = True
        frmAutoUpdate.Height = 3165
        cmdmore.Caption = "6"
        cmdmore.ToolTipText = "Show more..."
        PB.Value = 0
    End If
    If Dir(strDestination) = Empty Then
        CancelSearch = True
    Else
        If CancelSearch = False Then
           If MsgBox("La connessione al server è stata resettata." & vbCr & vbCr & _
                      "Cliccare su ""Riprova"" per riprendere il Download." & _
                      vbCr & "(Tempo rimanente: " & FormatTime(sglTime) & ")" & vbCr & vbCr & _
                      "Cliccare su ""Annulla"" per interrompere il Download del File.", vbExclamation + vbRetryCancel, "Download Incompleto") = vbRetry Then
                    blnResume = True
                    GoTo StartDownload
            End If
        End If
    End If
    DownloadFile = False
End If
Cleanup:
Inet1.Cancel
cmdControlla.Caption = "Download Update"
cmdControlla.Enabled = False
 cmdCheck.Enabled = True
cmdCancel.Enabled = True
frmAutoUpdate.Height = 3165
cmdmore.Caption = "6"
cmdmore.ToolTipText = "Show more..."
PB.Value = 0
Sleep (0.2)
'Unload frmDownload
Exit Function
InternetErrorHandler:
    If Err.Number = 9 Then Resume Next
        If Err.Number = 52 Then
    Else
        MsgBox "Errore: " & Err.Description, vbCritical, "Errore Downloading File"
        If Dir(App.Path & "\" & strSetupFile) <> "" Then Kill (App.Path & "\" & strSetupFile)
        Err.Clear
    End If
    GoTo ExitDownload
FileErrorHandler:
    MsgBox "Impossibile scrivere il File sul Disco." & vbCr & vbCr & "Errore: " & Err.Number & ": " & Err.Description, vbCritical, "Error Downloading File"
    CancelSearch = True
    cmdControlla.Caption = "Download Update"
    cmdCancel.Caption = "Annulla"
    cmdControlla.Enabled = False
    cmdCheck.Enabled = True
    cmdCancel.Enabled = True
    frmAutoUpdate.Height = 3165
    cmdmore.Caption = "6"
    cmdmore.ToolTipText = "Show more..."
    If Dir(App.Path & "\" & strSetupFile) <> "" Then Kill (App.Path & "\" & strSetupFile)
    Err.Clear
    GoTo ExitDownload
End Function

Private Function Findfile(xstrfilename) As WIN32_FIND_DATA
    Dim Win32Data As WIN32_FIND_DATA: Dim plngFirstFileHwnd As Long: Dim plngRtn As Long
    On Error Resume Next
    plngFirstFileHwnd = FindFirstFile(xstrfilename, Win32Data)
    If plngFirstFileHwnd = 0 Then
        Findfile.cFileName = "Error"
    Else
    Findfile = Win32Data
End If
    plngRtn = FindClose(plngFirstFileHwnd)
End Function

Private Function FitText(ByRef Ctl As Control, ByVal strCtlCaption) As String
Dim lngCtlLeft As Long: Dim lngMaxWidth As Long
Dim lngTextWidth As Long: Dim lngX As Long
On Error Resume Next
lngCtlLeft = Ctl.Left
lngMaxWidth = Ctl.Width
lngTextWidth = Ctl.Parent.TextWidth(strCtlCaption)
lngX = (Len(strCtlCaption) \ 2) - 2
While lngTextWidth > lngMaxWidth And lngX > 3
DoEvents
DoEvents
    strCtlCaption = Left(strCtlCaption, lngX) & "..." & Right(strCtlCaption, lngX)
    lngTextWidth = Ctl.Parent.TextWidth(strCtlCaption)
    lngX = lngX - 1
DoEvents
DoEvents
Wend
FitText = strCtlCaption
End Function

Private Function FormatFileSize(ByVal dblFileSize As Double, Optional ByVal strFormatMask As String) As String
On Error Resume Next
Select Case dblFileSize
    Case 0 To 1023 ' Bytes
        FormatFileSize = Format(dblFileSize) & " bytes"
    Case 1024 To 1048575 ' KB
        If strFormatMask = Empty Then strFormatMask = "###0"
        FormatFileSize = Format(dblFileSize / 1024#, strFormatMask) & " KB"
    Case 1024# ^ 2 To 1073741823 ' MB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 2), strFormatMask) & " MB"
    Case Is > 1073741823# ' GB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 3), strFormatMask) & " GB"
End Select
End Function

Private Function FormatPercentage(ByVal dblFileSize As Double, Optional ByVal strFormatMask As String) As String
On Error Resume Next
Select Case dblFileSize
    Case 0 To 1023
        FormatPercentage = Format(dblFileSize)
    Case 1024 To 1048575
        If strFormatMask = Empty Then strFormatMask = "###0"
        FormatPercentage = Format(dblFileSize / 1024#, strFormatMask)
    Case 1024# ^ 2 To 1073741823
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatPercentage = Format(dblFileSize / (1024# ^ 2), strFormatMask)
    Case Is > 1073741823#
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatPercentage = Format(dblFileSize / (1024# ^ 3), strFormatMask)
End Select
End Function

Private Function FormatTime(ByVal sglTime As Single) As String
On Error Resume Next
Select Case sglTime
    Case 0 To 59
        FormatTime = Format(sglTime, "0") & " sec"
    Case 60 To 3599
        FormatTime = Format(Int(sglTime / 60), "#0") & " min " & Format(sglTime Mod 60, "0") & " sec"
    Case Else
        FormatTime = Format(Int(sglTime / 3600), "#0") & " hr " & Format(sglTime / 60 Mod 60, "0") & " min"
End Select
End Function

Private Function ReturnFileOrFolder(FullPath As String, ReturnFile As Boolean, Optional IsURL As Boolean = False) As String
Dim intDelimiterIndex As Integer
On Error Resume Next
intDelimiterIndex = InStrRev(FullPath, IIf(IsURL, "/", "\"))
If intDelimiterIndex = 0 Then
    ReturnFileOrFolder = FullPath
Else
    ReturnFileOrFolder = IIf(ReturnFile, Right(FullPath, Len(FullPath) - intDelimiterIndex), Left(FullPath, intDelimiterIndex))
End If
End Function

Private Sub Sleep(Seconds As Double)
   Dim TempTime As Double
   On Error Resume Next
   DoEvents
   DoEvents
   TempTime = Timer
   DoEvents
   DoEvents
   Do While Timer - TempTime < Seconds
      DoEvents
      DoEvents
      If Timer < TempTime Then
         TempTime = TempTime - 24# * 3600#
         DoEvents
      End If
      DoEvents
      DoEvents
   Loop
End Sub
