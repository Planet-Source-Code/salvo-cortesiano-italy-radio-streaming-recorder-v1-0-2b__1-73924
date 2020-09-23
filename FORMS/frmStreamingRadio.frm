VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{DC1B114B-9B18-4A24-A69D-51B73A811D94}#2.0#0"; "MP3Player.ocx"
Object = "{8E48C67F-FAAF-4143-864F-4316EC1B6DA7}#2.0#0"; "Downloaders.ocx"
Object = "{77A4B2E6-7C1B-439A-9450-F2C1FA1F637F}#1.0#0"; "ShouCastRip.ocx"
Begin VB.Form frmStreamingRadio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Radio Streaming and Recorder"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9525
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStreamingRadio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   9525
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   315
      Left            =   6510
      TabIndex        =   70
      Top             =   7680
      Width           =   1410
   End
   Begin VB.Timer tvcl 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4170
      Top             =   5205
   End
   Begin VB.CommandButton cmdOption 
      Caption         =   "Show &Option"
      Height          =   315
      Left            =   8025
      TabIndex        =   67
      Top             =   7665
      Width           =   1410
   End
   Begin Downloaders.Downloader Downloader 
      Left            =   7830
      Top             =   -300
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox picMP3List 
      BorderStyle     =   0  'None
      Height          =   6000
      Left            =   0
      ScaleHeight     =   6000
      ScaleWidth      =   4515
      TabIndex        =   52
      Top             =   0
      Visible         =   0   'False
      Width           =   4515
      Begin VB.Timer scrTime 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4170
         Top             =   4770
      End
      Begin VB.CommandButton cmdMP3Loop 
         Caption         =   "q"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   975
         TabIndex        =   60
         ToolTipText     =   "Loop No..."
         Top             =   5550
         Width           =   435
      End
      Begin VB.Timer tMP3 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4170
         Top             =   5610
      End
      Begin VB.CommandButton cmdMP3Stop 
         Caption         =   "n"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   525
         TabIndex        =   57
         ToolTipText     =   "MP3 Stop..."
         Top             =   5550
         Width           =   435
      End
      Begin VB.CommandButton cmdMP3Play 
         Caption         =   "4"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   75
         TabIndex        =   56
         ToolTipText     =   "MP3 Play..."
         Top             =   5550
         Width           =   435
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   1920
         Left            =   75
         Pattern         =   "*.mp3"
         System          =   -1  'True
         TabIndex        =   55
         Top             =   3525
         Width           =   4395
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         Height          =   2730
         Left            =   75
         TabIndex        =   54
         Top             =   765
         Width           =   4380
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   75
         TabIndex        =   53
         Top             =   375
         Width           =   4395
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Navigate to Folder streaming Files List"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   90
         TabIndex        =   61
         Top             =   75
         Width           =   4380
      End
      Begin VB.Label lblDuration 
         Caption         =   "Total 00:00"
         Height          =   225
         Left            =   2835
         TabIndex        =   59
         Top             =   5595
         Width           =   1575
      End
      Begin VB.Label lblPosition 
         Caption         =   "Time: 00:00"
         Height          =   210
         Left            =   1515
         TabIndex        =   58
         Top             =   5595
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdShowListMP3 
      Caption         =   "&Get MP3"
      Height          =   315
      Left            =   3330
      TabIndex        =   51
      Top             =   6075
      Width           =   1050
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   9495
      TabIndex        =   48
      Text            =   "smtp.netshadows.it"
      Top             =   -270
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picBackGround 
      BorderStyle     =   0  'None
      Height          =   5625
      Index           =   1
      Left            =   15
      ScaleHeight     =   5625
      ScaleWidth      =   4515
      TabIndex        =   41
      Top             =   -15
      Visible         =   0   'False
      Width           =   4515
      Begin VB.Frame Frame6 
         Caption         =   "Radio Station Link's"
         Height          =   5505
         Left            =   30
         TabIndex        =   42
         Top             =   75
         Width           =   4425
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   450
            Left            =   60
            ScaleHeight     =   450
            ScaleWidth      =   4275
            TabIndex        =   44
            Top             =   4965
            Width           =   4275
            Begin VB.CommandButton cmdSelectIp 
               Caption         =   "&Select this IP"
               Height          =   315
               Left            =   165
               TabIndex        =   49
               Top             =   90
               Width           =   1950
            End
            Begin VB.CommandButton cmdCloseBackground 
               Caption         =   "Close &List"
               Height          =   315
               Left            =   2835
               TabIndex        =   45
               Top             =   90
               Width           =   1410
            End
         End
         Begin VB.ListBox lstStationLinks 
            Height          =   4680
            Left            =   60
            TabIndex        =   43
            Top             =   300
            Width           =   4230
         End
      End
   End
   Begin VB.ListBox lstURLS 
      Height          =   270
      Left            =   30
      TabIndex        =   40
      Top             =   45
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&X"
      Height          =   255
      Left            =   9150
      TabIndex        =   35
      ToolTipText     =   "Delete the selected Station URL"
      Top             =   6615
      Width           =   285
   End
   Begin VB.CommandButton cmdAddStation 
      Caption         =   "&Add Station"
      Height          =   315
      Left            =   7980
      TabIndex        =   34
      Top             =   6060
      Width           =   1410
   End
   Begin VB.CommandButton cmdSaveToMP3 
      Caption         =   "&Save to MP3"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6435
      TabIndex        =   33
      Top             =   6060
      Width           =   1485
   End
   Begin VB.ComboBox cmbRadioStation 
      Height          =   330
      ItemData        =   "frmStreamingRadio.frx":3D52
      Left            =   105
      List            =   "frmStreamingRadio.frx":3D54
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   6570
      Width           =   9000
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "C&onnect to Url"
      Height          =   315
      Left            =   4485
      TabIndex        =   31
      Top             =   6060
      Width           =   1770
   End
   Begin VB.Frame Frame5 
      Caption         =   "Stats"
      Height          =   1320
      Left            =   4560
      TabIndex        =   24
      Top             =   4665
      Width           =   4905
      Begin VB.TextBox received 
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   3690
      End
      Begin VB.TextBox written 
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   930
         Width           =   3705
      End
      Begin VB.TextBox meta 
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   585
         Width           =   3705
      End
      Begin VB.Label Label2 
         Caption         =   "Received"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Written"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   975
         Width           =   990
      End
      Begin VB.Label Label4 
         Caption         =   "Meta-Info"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1035
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stream Informations"
      Height          =   2880
      Left            =   4560
      TabIndex        =   18
      Top             =   1740
      Width           =   4920
      Begin VB.TextBox txtStationTitle 
         Height          =   285
         Left            =   1620
         TabIndex        =   39
         Top             =   2010
         Width           =   3225
      End
      Begin VB.TextBox txtStationURL 
         Height          =   285
         Left            =   1620
         TabIndex        =   38
         Top             =   1635
         Width           =   3240
      End
      Begin VB.TextBox sbr 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1320
         Width           =   4755
      End
      Begin VB.TextBox sgenre 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   960
         Width           =   4740
      End
      Begin VB.TextBox surl 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   600
         Width           =   4755
      End
      Begin VB.TextBox sname 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   4740
      End
      Begin VB.TextBox notices 
         Height          =   450
         Left            =   4980
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   -135
         Width           =   1110
      End
      Begin VB.Label lblNewFileName 
         Alignment       =   2  'Center
         Caption         =   "n.a"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   105
         TabIndex        =   69
         ToolTipText     =   "New FileName..."
         Top             =   2580
         Width           =   4725
      End
      Begin VB.Label lblHoldFileName 
         Alignment       =   2  'Center
         Caption         =   "n.a"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   105
         TabIndex        =   68
         ToolTipText     =   "Hold FileName..."
         Top             =   2340
         Width           =   4725
      End
      Begin VB.Label Label5 
         Caption         =   "Station Title"
         Height          =   255
         Left            =   135
         TabIndex        =   37
         Top             =   2040
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "Station URL"
         Height          =   255
         Left            =   135
         TabIndex        =   36
         Top             =   1680
         Width           =   1290
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Song Information"
      Height          =   975
      Left            =   4560
      TabIndex        =   15
      Top             =   720
      Width           =   4920
      Begin VB.TextBox curl 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   600
         Width           =   4725
      End
      Begin VB.TextBox cname 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   4725
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Operations"
      Height          =   615
      Left            =   4560
      TabIndex        =   12
      Top             =   75
      Width           =   4920
      Begin VB.CheckBox spub 
         Caption         =   "Is Public?"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3360
         TabIndex        =   66
         Top             =   240
         Width           =   1515
      End
      Begin VB.CheckBox chkConnected 
         Caption         =   "Connected?"
         Enabled         =   0   'False
         Height          =   255
         Left            =   210
         TabIndex        =   14
         Top             =   240
         Width           =   1635
      End
      Begin VB.CheckBox chkScrittura 
         Caption         =   "Writing?"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1935
         TabIndex        =   13
         Top             =   240
         Width           =   1380
      End
      Begin MP3Player.MP3Play MP3 
         Height          =   690
         Left            =   2250
         TabIndex        =   62
         Top             =   -540
         Visible         =   0   'False
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   1217
      End
      Begin RadioStreamingRecorder.SpecialFolders SpecialFolders 
         Left            =   2940
         Top             =   -90
         _ExtentX        =   450
         _ExtentY        =   370
      End
   End
   Begin ComctlLib.Slider SDVolume 
      Height          =   210
      Left            =   1530
      TabIndex        =   9
      ToolTipText     =   "Volume"
      Top             =   5265
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   370
      _Version        =   327682
      SmallChange     =   5
      Max             =   100
      SelStart        =   100
      Value           =   100
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   3870
      Top             =   300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar prgBAR 
      Height          =   195
      Left            =   7800
      TabIndex        =   8
      Top             =   7005
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   3810
      Top             =   870
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin ComctlLib.TreeView TreeStationURL 
      Height          =   5100
      Left            =   30
      TabIndex        =   7
      Top             =   45
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   8996
      _Version        =   327682
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   75
      TabIndex        =   5
      Top             =   6075
      Width           =   1200
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   ";"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   525
      TabIndex        =   3
      ToolTipText     =   "Pause..."
      Top             =   5220
      Width           =   435
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "Play..."
      Top             =   5220
      Width           =   435
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "n"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1020
      TabIndex        =   1
      ToolTipText     =   "Stop..."
      Top             =   5220
      Width           =   435
   End
   Begin VB.CommandButton cmdStreaming 
      Caption         =   "&Open Streaming"
      Enabled         =   0   'False
      Height          =   300
      Left            =   1380
      TabIndex        =   0
      Top             =   6075
      Width           =   1860
   End
   Begin ShouCastRip.ShoutCastRip SCR 
      Left            =   8730
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Size of last Audio File:"
      Height          =   255
      Left            =   4710
      TabIndex        =   65
      Top             =   7275
      Width           =   2550
   End
   Begin VB.Label lblAudioSize 
      Caption         =   "n.a"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   7365
      TabIndex        =   64
      Top             =   7275
      Width           =   2100
   End
   Begin VB.Label lblFileDeleted 
      Caption         =   "n.a"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   60
      TabIndex        =   63
      Top             =   7815
      Width           =   6330
   End
   Begin VB.Label lblInfoStation 
      Alignment       =   2  'Center
      Caption         =   "n.a"
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   45
      TabIndex        =   50
      Top             =   5655
      Width           =   4410
   End
   Begin VB.Label lblSongPath 
      Height          =   180
      Left            =   4035
      TabIndex        =   47
      Top             =   0
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.Label lblSonTitle 
      Caption         =   "n.a"
      Height          =   225
      Left            =   60
      TabIndex        =   46
      Top             =   7515
      Width           =   6330
   End
   Begin VB.Label lblVolume 
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4020
      TabIndex        =   11
      ToolTipText     =   "Volume"
      Top             =   5250
      Width           =   450
   End
   Begin VB.Label lblVLCPlugin 
      Caption         =   "VLCPlugin Installed?"
      Height          =   255
      Left            =   60
      TabIndex        =   10
      Top             =   7275
      Width           =   4545
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmStreamingRadio.frx":3D56
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmStreamingRadio.frx":3F30
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblStationUrl 
      Caption         =   "Station URL's..."
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   6945
      Width           =   4815
   End
   Begin VB.Label lblInfoStreaming 
      Caption         =   "Ready..."
      Height          =   255
      Left            =   4935
      TabIndex        =   4
      Top             =   6945
      Width           =   2820
   End
End
Attribute VB_Name = "frmStreamingRadio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Option Compare Text

Dim MP3FileName As String
Dim pPos As Integer
Dim strLoop As Boolean

Dim OpenWeb As Integer

Dim acDrive As String

Dim IconSet As Long

Private Const SW_SHOWNORMAL As Long = 1

Dim sStation As String: Dim sURLS As String: Dim sHoldLink As String

Private dwComplete As Boolean

' .... ^|^
Private RadioStationURL As String
Private RadioStationName As String

' .... Array, meybe i use this?
Dim TempArray(20) As String

' .... No Close Form [X]
Private sCancel As Boolean

' .... Stations URLs
Private Const bSTATIONS As String = "http://205.188.215.225:8006-1FM The Chillout Lounge|" _
& "http://64.71.184.99:8010-Chillout DIGITALLY IMPORT|http://64.62.194.40:9010-JapanARadio " _
& "Japan's best music mix|http://205.188.215.225:8006-80s, 80s, 80s!  S K Y  F M  Hear your " _
& "cla|http://174.36.206.217:8485-Radio HSL  Hit Songs Lagataar|http://184.95.62.170:9002-KPOP " _
& "@ Big B Radio  The Only Hot Station fo|http://213.246.51.97:8029-LOVE CHINA  Chinese Hits|" _
& "http://72.26.204.28:7566-Cool Jazz  JAZZRADIOcom|http://205.188.215.231:8014-Lounge  " _
& "D I G I T A L L Y  I M P O R T E D|http://74.86.76.2:7258-All '80s & '90s  TheHitsus"

' .... Create Dir
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Private newDir As String

' .... Declaration to library of VLC
Private VLC As VLCPlugin2

' .... Retrive the connection and obtain the Radio Stations URLs
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long

Private sConnType As String * 255
Private sWholeData As String
Private bDataReceived As Boolean
Private sEndTag As String
Private sShoutCastURL As String

Private Type Station
    StationName As String
    ID As String
    BitRate As String
    Genre As String
    CurrentTrack As String
    ListenerCount As String
End Type

Private tStations(20000) As Station
Private sStationFileLoc As String

Private WithEvents m_frmSysTray As frmSysTray
Attribute m_frmSysTray.VB_VarHelpID = -1

Private Enum sSetIcon
    ICON_PROGRAM = 0
    ICON_NOTE = 1
    ICON_LOG = 2
    ICON_AUDIO = 3
    ICON_CDLIST = 4
    ICON_TODO = 5
    NOTHING_ICON = 6
End Enum

Private Sub cmbRadioStation_Click()
    On Local Error GoTo ErrorHandler
    txtStationTitle.Text = StripLeft(cmbRadioStation.List(cmbRadioStation.ListIndex), _
    "-", False) '/Station Name
    txtStationURL.Text = StripLeft(cmbRadioStation.List(cmbRadioStation.ListIndex), _
    "-", True) '/Station URL
    
    RadioStationURL = StripLeft(cmbRadioStation.List(cmbRadioStation.ListIndex), _
    "-", True) '/Station URL
    RadioStationName = StripLeft(cmbRadioStation.List(cmbRadioStation.ListIndex), _
    "-", False) '/Station Name
    
    If sURLS = "Unknown" Then _
    sURLS = RadioStationName
    
Exit Sub
ErrorHandler:
    Err.Clear
End Sub


Private Sub cmdAbout_Click()
    On Local Error Resume Next
    SCR.About
End Sub

Private Sub cmdAddStation_Click()
Dim i As Integer
    On Local Error GoTo ErrorHandler
        For i = 0 To cmbRadioStation.ListCount
            If cmbRadioStation.List(i) = RadioStationURL + "-" + RadioStationName Then
                    MsgBox "The Radio Station link already exists!", vbExclamation, App.Title
                Exit Sub
            End If
        Next i
    cmbRadioStation.AddItem RadioStationURL + "-" + RadioStationName
    MsgBox "Radio Station URL: " & RadioStationURL & vbCrLf & "Radio Station Name: " _
    & RadioStationName & vbCrLf & vbCrLf _
    & "Adding success!", vbInformation, App.Title
Exit Sub
ErrorHandler:
        WriteErrorLogs Err.Number, Err.Description, "FormMain {Sub: cmdAddStation}", True, True
    Err.Clear
End Sub

Private Sub cmdClose_Click()
    Dim sInput As String
    'sInput = InputBox("Enter the correct PassWord for this Option", App.Title, sInput)
    'If sInput = Empty Then Exit Sub
    'If sInput = "hantares" Then:
    
    ' .... Conferma
    If MsgBox("Close Program?" & vbLf & vbLf & "Are you sure?", _
    vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
    
    ' .... EXIT
    sCancel = False: Unload Me
End Sub

Private Sub cmdCloseBackground_Click()
    picBackGround(1).Visible = False
End Sub

Private Sub cmdConnect_Click()
    On Local Error GoTo ErrorHandler
    
    If RadioStationURL = Empty Then
            MsgBox "Select Station URL first!", vbInformation, App.Title
        Exit Sub
    End If
    
    If cmdConnect.Caption = "C&onnect to Url" Then
        If MsgBox("Connect to Radio Station:" & vbLf & RadioStationName _
        & vbCrLf & "Station Link:" & vbCrLf & RadioStationURL, _
        vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
        
        cmdConnect.Caption = "&Disconnect"
        
        If Dir$(App.Path + "\tmp.mp3") <> Empty Then Call Kill(App.Path + "\tmp.mp3")
        
        SCR.Connect RadioStationURL
        
        If SCR.IsConnected = True Then
            scrTime.Enabled = True
            cmdStreaming.Enabled = True
            cmdSaveToMP3.Enabled = True
            
            If SCR.MusicName <> Empty Then cname.Text = SCR.MusicName Else cname.Text = SCR.StreamName
            If SCR.MusicURL <> Empty Then curl.Text = SCR.MusicURL Else curl.Text = SCR.StreamURL
            curl.Text = ReplaceString(curl.Text)
            
            If cmdStreaming.Caption = "&Close Streaming" Then
                cmdStreaming = True
                Call DelayTime(1, True)
                cmdStreaming = True
            End If
            
        End If
        
    ElseIf cmdConnect.Caption = "&Disconnect" Then
        cmdConnect.Caption = "C&onnect to Url"
        
        If SCR.IsConnected Then SCR.Disconnect
        
        If SCR.IsWriting Then
            SCR.StopWriting
            cmdSaveToMP3.Caption = "&Save to MP3"
        End If
        
        cmdSaveToMP3.Enabled = False
        cmdStreaming.Enabled = False
        cmdStop = True
        cmdStop.Enabled = False
        cmdPlay.Enabled = False
        cmdPause.Enabled = False
        scrTime.Enabled = False
        
        If cmdStreaming.Caption = "&Close Streaming" Then cmdStreaming = True
        
    End If
Exit Sub
ErrorHandler:
        WriteErrorLogs Err.Number, Err.Description, "Form Main {Sub: cmdConnect}", True, True
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Delete the selected Station URL?" & vbLf & vbLf _
    & cmbRadioStation.List(cmbRadioStation.ListIndex), _
    vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
    cmbRadioStation.RemoveItem (cmbRadioStation.ListIndex)
    
    On Local Error Resume Next
    If cmbRadioStation.ListCount > 0 Then
        cmbRadioStation.ListIndex = cmbRadioStation.ListCount - 1
    End If
    
End Sub

Private Sub cmdMP3Loop_Click()
    If strLoop = False Then
        strLoop = True
        cmdMP3Loop.ToolTipText = "Loop Yes..."
    Else
        strLoop = False
        cmdMP3Loop.ToolTipText = "Loop No..."
    End If
End Sub

Private Sub cmdMP3Play_Click()
    On Local Error Resume Next
    tMP3.Enabled = False
    MP3.mmStop: MP3.FileName = MP3FileName: MP3.mmPlay
    lblDuration = "Total Time: " & MP3.length
    tMP3.Enabled = True
    cmdMP3Stop.Enabled = True
    cmdMP3Loop.Enabled = True
    cmdMP3Play.Enabled = False
End Sub

Private Sub cmdMP3Stop_Click()
    On Local Error Resume Next
    MP3.mmStop
    tMP3.Enabled = False
    cmdMP3Play.Enabled = True
    cmdMP3Stop.Enabled = False
    cmdMP3Loop.Enabled = False
    lblDuration = "Total Time: 00:00": lblPosition = "Time: 00:00"
End Sub

Private Sub cmdOption_Click()
    frmOption.Show vbModal, Me
End Sub

Private Sub cmdPause_Click()
    On Local Error Resume Next
    VLC.playlist.togglePause: tvcl.Enabled = False
    cmdStop.Enabled = True
    cmdPlay.Enabled = True
    cmdPause.Enabled = False
End Sub

Private Sub cmdPlay_Click()
    On Local Error Resume Next
    VLC.playlist.play: tvcl.Enabled = True
    cmdStop.Enabled = True
    cmdPlay.Enabled = False
    cmdPause.Enabled = True
End Sub

Private Sub cmdSaveToMP3_Click()
    
    If cmdSaveToMP3.Caption = "&Save to MP3" Then
        cmdSaveToMP3.Caption = "&Stop Saving"
        
    ' .... Create Folder of tyhe Song if not exist
    If MakeDir(FormatFileName(SCR.StreamName)) = False Then
        newDir = App.Path + "\Streaming Files List\" + sURLS + "\" + FormatFileName(SCR.MusicName) & "\"
        Call MakeNewDir(newDir)
    End If
    
    ' .... Get the Extra TAG's ....
    Call RetriveExtraTags
    
    ' .... Start Record Song to MP3 File
    SCR.StartWriting newDir + FormatFileName(SCR.MusicName) + ".mp3", AppendSeek
    lblSongPath.Caption = newDir + FormatFileName(SCR.MusicName) + ".mp3"
    lblSonTitle.Caption = FormatFileName(SCR.MusicName)
    
    
    lblHoldFileName.Caption = FormatFileName(SCR.MusicName) + ".mp3"
    lblNewFileName.Caption = FormatFileName(SCR.MusicName) + ".mp3"
    
    Call DelayTime(1, True)
    
    sname.Text = SCR.StreamName
    sgenre.Text = SCR.StreamGenre
    surl.Text = SCR.StreamURL: curl.Text = SCR.StreamURL
    curl.Text = ReplaceString(curl.Text)
    sbr.Text = SCR.StreamBitRate & " Kbps"
    spub.Value = -CInt(SCR.StreamPublic)
    notices.Text = SCR.GetNotices()
    
    ElseIf cmdSaveToMP3.Caption = "&Stop Saving" Then
        cmdSaveToMP3.Caption = "&Save to MP3"
        SCR.StopWriting
    End If
    
End Sub

Private Sub cmdShowListMP3_Click()
    If cmdShowListMP3.Caption = "&Get MP3" Then
        picMP3List.Visible = True
        Dir1.Refresh: File1.Refresh
        cmdShowListMP3.Caption = "&Hide..."
    ElseIf cmdShowListMP3.Caption = "&Hide..." Then
        picMP3List.Visible = False
        cmdShowListMP3.Caption = "&Get MP3"
    End If
End Sub

Private Sub cmdStop_Click()
    On Local Error Resume Next
    VLC.playlist.stop: tvcl.Enabled = False
    
    lblInfoStreaming.Caption = "Time: 00:00"
    cmdStop.Enabled = False
    cmdPlay.Enabled = True
    cmdPause.Enabled = False
End Sub

Private Sub cmdStreaming_Click()
    Dim varOption(0) As String
    On Local Error GoTo ErrorHandler
    
    If SCR.IsConnected = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    If cmdStreaming.Caption = "&Open Streaming" Then
        cmdStreaming.Caption = "&Close Streaming"
    
        ' .... Clear the Playlist
        VLC.playlist.Items.Clear
        ' .... Add radiostation URL
        VLC.playlist.Add RadioStationURL, SCR.StreamName
        ' .... Play the first Item of the Playlist
        VLC.playlist.playItem (0)
    
        ' .... Enable Time Info
        tvcl.Enabled = True
    
        cmdStop.Enabled = True
        cmdPlay.Enabled = False
        cmdPause.Enabled = True
    ElseIf cmdStreaming.Caption = "&Close Streaming" Then
        cmdStreaming.Caption = "&Open Streaming"
        cmdStop = True
        cmdStop.Enabled = False
        cmdPlay.Enabled = False
        cmdPause.Enabled = False
        tvcl.Enabled = False
    End If
    
    Screen.MousePointer = vbDefault
Exit Sub

ErrorHandler:
    Screen.MousePointer = vbDefault
        WriteErrorLogs Err.Number, Err.Description, "FormMain {Sub: cmdStreaming}", True, True
    Err.Clear
End Sub

Private Sub Dir1_Change()
    On Local Error GoTo ErrorHandler
    File1.Path = Dir1.Path: File1.Refresh
Exit Sub
ErrorHandler:
        WriteErrorLogs Err.Number, Err.Description, "FormMain {Sub: Drive1_Change}", True, True
    Err.Clear
End Sub

Private Sub Downloader_DownloadComplete(MaxBytes As Long, SaveFile As String)
    dwComplete = True
    On Local Error Resume Next
    prgBAR.Value = 0
End Sub

Private Sub Downloader_DownloadError(SaveFile As String)
    dwComplete = False
End Sub


Private Sub Downloader_DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)
    On Error Resume Next
    With prgBAR: .Max = MaxBytes: .Value = CurBytes: End With
    lblInfoStreaming.Caption = CurBytes & " of " & MaxBytes
End Sub

Private Sub Drive1_Change()
    On Local Error GoTo ErrorHandler
    If acDrive = Empty Then acDrive = "C:"
    Dir1.Path = Drive1.Drive: Dir1.Refresh
Exit Sub
ErrorHandler:
    If Err.Number = 68 Then ' .... No Drive Selected or Ready...
                Drive1.Drive = acDrive
            Resume Next
        Exit Sub
    End If
        WriteErrorLogs Err.Number, Err.Description, "FormMain {Sub: Drive1_Change}", True, True
    Err.Clear
End Sub

Private Sub File1_DblClick()
    On Local Error Resume Next
    If File1.ListCount > 0 Then
        MP3FileName = File1.Path + "\" + File1.FileName
        cmdMP3Play.Enabled = True
        
        tMP3.Enabled = False
        MP3.mmStop: MP3.FileName = MP3FileName: MP3.mmPlay
        lblDuration = "Total Time: " & MP3.length
        tMP3.Enabled = True
        cmdMP3Stop.Enabled = True
        cmdMP3Loop.Enabled = True
        cmdMP3Play.Enabled = False
        
    Else
        MsgBox "Notthing MP3's file's selected!!!", vbExclamation, App.Title
    End If
End Sub


Private Sub Form_Load()
    Dim ssTop As Long: Dim ssLeft As Long
    Dim hwCurr As Long
    
    On Local Error GoTo ErrorHandler
    
    If Dir$(App.Path + "\Radio Station List.xml") = Empty Then
    ' .... Create file XML Radio Station List
    If CreateRes(101, "XML", "Radio Station List.xml") = False Then
        MsgBox "Error to create Radio Station List XML!" & vbCrLf _
        & "This program will be terminated Now, sorry!", vbExclamation, App.Title
        ' .... EXIT
        sCancel = False: Unload Me
    End If
    End If
    
    ' .... The PC is Connected via NET?
    If GetNetConnectString = False Then
        cmdStreaming.Enabled = False
        cmdConnect.Enabled = False
        cmdSaveToMP3.Enabled = False
        cmdAddStation.Enabled = False
        cmdStop.Enabled = False: cmdPlay.Enabled = False: cmdPause.Enabled = False
        WriteErrorLogs -4125, "Unable to determine the type of network connection!", "ModException {Function: GetNetConnectString}", True, True
        lblInfoStation.Caption = "No Connection Type..."
    Else
        lblInfoStation.Caption = ConnectieType
    End If
    
    ' .... Locate VLC Path
    If VLCLocation = Empty Or Mid$(VLCLocation, 1, 5) = "Error" Then
        If MsgBox("Application VLC v.1.1.8 not found..." & vbLf & vbLf _
            & "Do you want to install only the libraries for the proper functioning of the Program?", _
            vbYesNo + vbQuestion, App.Title) = vbYes Then
            If InstallLibrary(SW_NORMAL) = False Then:
        Else
            lblVLCPlugin.Caption = "Error! VLC Plug-In not found!"
        End If
    Else
        lblVLCPlugin.Caption = VLCLocation
    End If
    
    ' .... Create Folder Radio Streaming
    Call MakeNewDir(App.Path + "\Streaming Files List")
    
    ' .... Init VLC Library
    Set VLC = New VLCPlugin2
    
    SDVolume.Value = 50: VLC.Volume = SDVolume.Value
    
    ' .... No Close Form from [X]
    sCancel = True
    
    ' .... Show Main Form   and Refresh
    frmStreamingRadio.Show: frmStreamingRadio.Refresh
    
    ' .... Center Form
    ssTop = (Screen.Height - frmStreamingRadio.Height) \ 2
    ssLeft = (Screen.Width - frmStreamingRadio.Width) \ 2
    frmStreamingRadio.Move ssLeft, ssTop
    
    ' .... Wait 1 seconds
    Call DelayTime(1, True)
    
    ' .... Populate Radio Station list
    Call FillTVWWithStations(TreeStationURL)
    
    ' .... Populate Combo Radio Stations
    Call AddRadioStation
    
    ' .... Get the First Station
    If cmbRadioStation.ListCount > 0 Then
        RadioStationURL = StripLeft(cmbRadioStation.List(0), _
                "-", True)
        RadioStationName = StripLeft(cmbRadioStation.List(0), _
                "-", False)
        txtStationURL.Text = RadioStationURL
        txtStationTitle.Text = RadioStationName
    End If
    
    ' .... URL Station folder to be NOT empty
    If sURLS = Empty Then sURLS = "Unknown"
    
    ' .... Internal variables program
    IconSet = 1: acDrive = "C:": strLoop = False
    
    ' .... Setting Combo Station List
    If INI.GetKeyValue("STATION", "RADIOSTATION") <> Empty Then _
            cmbRadioStation.ListIndex = INI.GetKeyValue("STATION", "RADIOSTATION") Else _
            cmbRadioStation.ListIndex = 0
    
    ' ...................................................................................
    
    ' .... Write TAGs
    If INI.GetKeyValue("SETTING", "WRITEMP3TAGS") <> Empty Then
        If INI.GetKeyValue("SETTING", "WRITEMP3TAGS") = 1 Then
            sWriteTagOfTrack = True
            SCR.WriteID3TagToFile = False
        ElseIf INI.GetKeyValue("SETTING", "WRITEMP3TAGS") = 0 Then
            sWriteTagOfTrack = False
            SCR.WriteID3TagToFile = False
        End If
    Else
        sWriteTagOfTrack = False
        SCR.WriteID3TagToFile = False
    End If
    
    ' .... Append SONG to File
    If INI.GetKeyValue("SETTING", "APPEND") <> Empty Then
        If INI.GetKeyValue("SETTING", "APPEND") = 1 Then AppendSeek = True
        If INI.GetKeyValue("SETTING", "APPEND") = 0 Then AppendSeek = False
    Else
        AppendSeek = False
    End If
    
    ' .... Delete File?
    If INI.GetKeyValue("SETTING", "DELETE") <> Empty Then
        If INI.GetKeyValue("SETTING", "DELETE") = 1 Then sCheckDelete = True
        If INI.GetKeyValue("SETTING", "DELETE") = 0 Then sCheckDelete = False
    Else
        sCheckDelete = False
    End If
    
    ' .... Create single File?
    If INI.GetKeyValue("SETTING", "CREATESINGLEFILE") <> Empty Then
        If INI.GetKeyValue("SETTING", "CREATESINGLEFILE") = 1 Then sCheckDivide = True
        If INI.GetKeyValue("SETTING", "CREATESINGLEFILE") = 0 Then sCheckDivide = False
    Else
        sCheckDivide = True
    End If
    
    ' .... Save Link Radio Station on Exit?
    If INI.GetKeyValue("SETTING", "SAVELINKONEXIT") <> Empty Then
        If INI.GetKeyValue("SETTING", "SAVELINKONEXIT") = 1 Then schkSaveStation = True
        If INI.GetKeyValue("SETTING", "SAVELINKONEXIT") = 0 Then schkSaveStation = False
    Else
        schkSaveStation = True
    End If

    ' ...................................................................................
    
    ' .... Extra ID3 TAG's
    If INI.GetKeyValue("ID3EXTRATAGINFO", "ENCODEBY") <> Empty Then
        ID3TagEncodedBy = INI.GetKeyValue("ID3EXTRATAGINFO", "ENCODEBY")
    Else
        ID3TagEncodedBy = App.EXEName & " by Salvo Cortesiano"
    End If
    
    If INI.GetKeyValue("ID3EXTRATAGINFO", "COPYRIGHT") <> Empty Then
        ID3TagCopyrightInfo = INI.GetKeyValue("ID3EXTRATAGINFO", "COPYRIGHT")
    Else
        ID3TagCopyrightInfo = "http://www.netshadows.it"
    End If
    
    If INI.GetKeyValue("ID3EXTRATAGINFO", "LANGUAGE") <> Empty Then
        ID3TagLanguages = INI.GetKeyValue("ID3EXTRATAGINFO", "LANGUAGE")
    Else
        ID3TagLanguages = "Italians Language"
    End If
    
    If INI.GetKeyValue("ID3EXTRATAGINFO", "COMMENTS") <> Empty Then
        ID3TagComments = INI.GetKeyValue("ID3EXTRATAGINFO", "COMMENTS")
    Else
        ID3TagComments = "Italians Language"
    End If
    
    ' .... Set default Diver and Path
    Drive1.Drive = "c:"
    Dir1.Path = App.Path + "\Streaming Files List"
    
    ' .... Hide Application to Windows Task ... Meybe ;)
    Call HideAppWinTask(1)
    
    ' .... Retrive the Prodess ID of Application in case of Crash {=PID}
    PID = WindowToProcessId(Me.hWnd)
    INI.DeleteKey "PROCESSID", "PID"
    INI.CreateKeyValue "PROCESSID", "PID", PID
    
    ' .... TaTàaaaaaaaaa ;)
    Call PlaySoundResource(102)
Exit Sub

ErrorHandler:
        WriteErrorLogs Err.Number, Err.Description, "FormMain {Sub: Form_Load}", True, True
    Err.Clear
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' .... Release the VLC Library
    Set VLC = Nothing
    
    ' .... Release the Deugger
    SetUnhandledExceptionFilter ByVal 0&
    
    ' .... Release the Library
    Call FreeLibrary(m_hMod)
    'Call FreeLibrary(m_hMod2)
    
    ' .... UnHook the SO
    If Not InIDE() Then SetErrorMode SEM_NOGPFAULTERRORBOX
    
    ' .... Save the RadioLink List
    If schkSaveStation Then
        If SaveNewRadioLink(, True) = False Then:
    End If
    
    ' .... Save Setting to File *.INI
    Call SaveSettingINI
    
    ' .... Relese Class INI
    Set INI = Nothing
    
    ' .... Ok to exit
    Cancel = sCancel
    
End Sub

Private Sub Form_Resize()
On Local Error GoTo ErrorHandler
If Me.WindowState = vbMinimized Then
        If GetSysTray(True) Then
        If IconSet = 0 Then SetIcon NOTHING_ICON Else SetIcon ICON_PROGRAM
            ShowTip "Radio Streaming v1.0.2b now is Hidden in the Tray-Bar." & vbCrLf & "Double click or Right Click for Menu!", "Radio Streaming v1.0.2b"
            Me.Visible = False
        End If
    ElseIf Me.WindowState = vbNormal Then
        Me.Visible = True
        If GetSysTray(False) Then:
    End If
Exit Sub
ErrorHandler:
    Call WriteErrorLogs(Err.Number, Err.Description, "Form Main {Sub: Form_Resize}", True, True)
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' .... Unload Application
    DestroyWindow Me.hWnd
    
    ' .... Form on Nothing
    Set frmStreamingRadio = Nothing
    
    ' .... End this program
    TerminateProcess GetCurrentProcess, 0
    
    ' .... Close Programs
    End
End Sub



Private Function PositionTime(ByVal stime As String) As String
    Static s As String * 30
    Dim sec, mins: s = stime
    sec = Round(Mid$(s, 1, Len(s)) / 1000)
    If sec < 60 Then PositionTime = "00:" & Format(sec, "00")
    If sec > 59 Then
        mins = Int(sec / 60)
        sec = sec - (mins * 60)
        PositionTime = Format(mins, "00") & ":" & Format(sec, "00")
    End If
End Function

Private Sub FillTVWWithStations(tvw As TreeView)
Dim iPos As Long: Dim i As Long: Dim xj As Long
Dim sEndTag As String: Dim iEndPos As Long
Dim tCount As Integer: Dim p As Long: Dim j As Long
Dim sStation As String: Dim sID As String
Dim sgenre As String: Dim sCurrTrack As String
Dim sBitRate As String: Dim textLine As String: Dim FileHandle As Integer

On Local Error GoTo ErrorHandler

FileHandle = FreeFile: sEndTag = "</genrelist>"

Open App.Path + "\Radio Station List.xml" For Input As #FileHandle
    Do While Not EOF(FileHandle)
        Line Input #FileHandle, textLine: sWholeData = sWholeData & textLine & vbCrLf
        xj = xj + 1
        lblStationUrl.Caption = "Loaded {" & xj & "} Genre Station's..."
    Loop
Close #FileHandle

i = 1: Do: DoEvents
    j = InStr(i, sWholeData, "<genre name=""", vbTextCompare)
    If j > 0 Then
        iPos = InStr(j, sWholeData, "<genre name=""")
        iEndPos = InStr(iPos, sWholeData, """></genre>")
        tvw.Nodes.Add , , "K" & Mid$(sWholeData, iPos + 13, iEndPos - iPos - 13), Mid$(sWholeData, iPos + 13, iEndPos - iPos - 13), 1
        tvw.Nodes.Add tvw.Nodes.Count, tvwChild, , "..."
        i = iEndPos: xj = xj + 1
    End If
    DoEvents: Loop Until j = 0

Exit Sub
ErrorHandler:
    If Err.Number <> 35602 Then _
        WriteErrorLogs Err.Number, Err.Description, "FormMain {Sub: FillTVWWithStations}", True, True
    Err.Clear
End Sub


Private Sub lblAudioSize_Click()
    Dim FSO As New Scripting.FileSystemObject: Dim f As File
    On Local Error Resume Next
    If FSO.FileExists(lblSongPath.Caption) Then
        Set f = FSO.GetFile(lblSongPath.Caption)
        lblAudioSize.Caption = FormatSize(f.SIZE)
    End If
    Set f = Nothing: Set FSO = Nothing
End Sub

Private Sub lblDuration_Click()
'    If sShutDown(WE_REBOOT) = False Then:
    
End Sub

Private Sub lstStationLinks_DblClick()
    On Local Error Resume Next
    RadioStationURL = lstStationLinks.List(lstStationLinks.ListIndex)
    txtStationURL.Text = RadioStationURL
    picBackGround(1).Visible = False
End Sub


Private Sub m_frmSysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
        Dim OpenWeb As Integer: Dim sInput As String

        'sInput = InputBox("Enter the correct PassWord for this Option", App.Title, sInput)
        'If sInput = Empty Then Exit Sub
        'If sInput = "Hantares1966" Then:
        
Select Case sKey
    Case "Open"
        Me.WindowState = vbNormal: Me.Visible = True: Me.Show: Me.ZOrder
    Case "Close"
        cmdClose = True
    Case "About"
        On Local Error Resume Next
        SCR.About
    Case "salvocortesiano"
        If MsgBox("Are you sure to send e-mail to: {salvocortesiano@netshadows.it}?", vbYesNo + vbInformation + _
        vbDefaultButton1, "MailTo: salvocortesiano@hotmail.com") = vbYes Then _
        SendEmail "salvocortesiano@hotmail.com", App.Title, "To-Do " & App.Title, "", ""
    Case "netshadows"
        If MsgBox("Are you sure to visit the web: {www.netshadows.it}?", vbYesNo + vbInformation + _
        vbDefaultButton1, "Open page www.netshadows.it") = vbYes Then _
        OpenWeb = ShellExecute(Me.hWnd, "Open", "http://www.netshadows.it/leombredellarete/forum", "", App.Path, 1)
    Case "Download"
        OpenWeb = ShellExecute(Me.hWnd, "Open", "http://www.netshadows.it/RSRv102.exe", "", App.Path, 1)
    Case "Plg-In"
        OpenWeb = ShellExecute(Me.hWnd, "Open", "http://www.netshadows.it/vlcplugins.exe", "", App.Path, 1)
    Case "Full-binary"
        OpenWeb = ShellExecute(Me.hWnd, "Open", "http://www.netshadows.it/RSRv102.rar", "", App.Path, 1)
    Case "HideTray"
        If IconSet = 1 Then
            SetIcon NOTHING_ICON: IconSet = 0
        ElseIf IconSet = 0 Then
            SetIcon ICON_PROGRAM: IconSet = 1
        End If
    Case "ChangeIcon1"
        If IconSet = 0 Then MsgBox "U must visible Icon Tray before change Task Icon!", vbInformation, App.Title: Exit Sub
        SetIcon ICON_AUDIO
    Case "ChangeIcon2"
        If IconSet = 0 Then MsgBox "U must visible Icon Tray before change Task Icon!", vbInformation, App.Title: Exit Sub
        SetIcon ICON_CDLIST
    Case "ChangeIcon3"
        If IconSet = 0 Then MsgBox "U must visible Icon Tray before change Task Icon!", vbInformation, App.Title: Exit Sub
        SetIcon ICON_LOG
    Case "ChangeIcon4"
        If IconSet = 0 Then MsgBox "U must visible Icon Tray before change Task Icon!", vbInformation, App.Title: Exit Sub
        SetIcon ICON_NOTE
    Case "ChangeIcon5"
        If IconSet = 0 Then MsgBox "U must visible Icon Tray before change Task Icon!", vbInformation, App.Title: Exit Sub
        SetIcon ICON_TODO
    Case "DefaultIconTray"
        If IconSet = 0 Then MsgBox "U must visible Icon Tray before change Task Icon!", vbInformation, App.Title: Exit Sub
        SetIcon ICON_PROGRAM
    End Select
End Sub

Private Sub m_frmSysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
    On Local Error Resume Next
    'Dim sInput As String
        'sInput = InputBox("Enter the correct PassWord for this Option", App.Title, sInput)
        'If sInput = Empty Then Exit Sub
        'If sInput = "Hantares1966" Then:
    Me.WindowState = vbNormal: Me.Visible = True: Me.Show: Me.ZOrder
End Sub


Private Sub m_frmSysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
    On Local Error Resume Next
    'Dim sInput As String
        'sInput = InputBox("Enter the correct PassWord for this Option", App.Title, sInput)
        'If sInput = Empty Then Exit Sub
        'If sInput = "Hantares" Then:
    If (eButton = vbRightButton) Then m_frmSysTray.ShowMenu
End Sub

Private Sub SCR_Error(ErrType As ShouCastRip.EErrType, ErrCode As ShouCastRip.EErrCode, Description As String)
Dim reponse As Long
Select Case ErrType
    Case TErr_WSock
        MsgBox Description, vbExclamation, App.Title
        cmdConnect.Caption = "C&onnect to Url"
            tvcl.Enabled = False
            cmdStreaming.Enabled = True
    Case TErr_Stream
        Select Case ErrCode
            Case Err_Stream_Redirection
                If MsgBox("The source appears to been moved " & Description & "." _
                & vbCrLf & "Do you want to reconnet with this link?", vbYesNo + vbQuestion, _
                App.Title) = vbYes Then
                    SCR.Connect Description
                Else
                    cmdConnect.Caption = "C&onnect to Url"
                End If
            Case Else
                MsgBox ErrCode & "   " & Description, vbExclamation, App.Title
                cmdConnect.Caption = "C&onnect to Url"
                tvcl.Enabled = False
                cmdStreaming.Enabled = True
        End Select
    Case TErr_FileIO
        MsgBox ErrCode & "   " & Description, vbExclamation, App.Title
        cmdConnect.Caption = "C&onnect to Url"
        tvcl.Enabled = False
        cmdStreaming.Enabled = True
End Select

Exit Sub
End Sub

Private Sub SCR_GotHeader(Header As String)
    sname.Text = SCR.StreamName: sgenre.Text = SCR.StreamGenre
    surl.Text = SCR.StreamURL: sbr.Text = SCR.StreamBitRate
    spub.Value = -CInt(SCR.StreamPublic): notices.Text = SCR.GetNotices()
End Sub


Private Sub SCR_MetaChanged(meta As String)
    Dim FSO As New Scripting.FileSystemObject: Dim f As File
    
    On Local Error GoTo ErrorHandler
    
    If SCR.MusicName <> Empty Then cname.Text = SCR.MusicName Else cname.Text = SCR.StreamName
    If SCR.MusicURL <> Empty Then curl.Text = SCR.MusicURL Else curl.Text = SCR.StreamURL
    
    curl.Text = ReplaceString(curl.Text): sname.Text = SCR.StreamName
    sgenre.Text = SCR.StreamGenre: surl.Text = SCR.StreamURL
    sbr.Text = SCR.StreamBitRate & " Kbps": spub.Value = -CInt(SCR.StreamPublic)
    notices.Text = SCR.GetNotices()
    
    If SCR.IsWriting = True And FSO.FileExists(lblSongPath.Caption) Then
        Set f = FSO.GetFile(lblSongPath.Caption)
        lblAudioSize.Caption = FormatSize(f.SIZE)
    End If
    
    If SCR.IsWriting = True And UCase(lblSonTitle.Caption) <> UCase(FormatFileName(SCR.MusicName)) Then
        
        lblNewFileName.Caption = FormatFileName(SCR.MusicName) + ".mp3"
        
        ' .... Create Song Dir
        If MakeDir(FormatFileName(SCR.StreamName)) = False Then
            newDir = App.Path + "\Streaming Files List\" + sURLS + "\" + _
            FormatFileName(SCR.MusicName) & "\": Call MakeNewDir(newDir)
        End If
        
        ' .... Stop Writing
        SCR.StopWriting
        
        ' .... Start Record Song to MP3 File
        SCR.StartWriting newDir + FormatFileName(SCR.MusicName) + ".mp3", AppendSeek
    End If
    
    ' .... Verify the Audio File
    If sCheckDelete Then
            If FSO.FileExists(lblSongPath.Caption) Then
                Set f = FSO.GetFile(newDir + lblHoldFileName.Caption)
            If StripLeft(FormatSizeDUE(f.SIZE), ",", True) <= 1000 Then
                Call Kill(lblSongPath.Caption)
                lblFileDeleted.Caption = lblHoldFileName.Caption & " Deleted!"
            Else
                lblFileDeleted.Caption = "Nothing file Deleted..."
            End If
        End If
    End If
    
    lblSonTitle.Caption = FormatFileName(SCR.MusicName)
    lblSongPath.Caption = newDir + FormatFileName(SCR.MusicName) + ".mp3"
    
    If SCR.IsWriting = True And lblNewFileName.Caption <> lblHoldFileName.Caption Then
        lblHoldFileName.Caption = FormatFileName(SCR.MusicName) + ".mp3"
    End If

    Set f = Nothing: Set FSO = Nothing
Exit Sub
ErrorHandler:
        Err.Clear
    Resume Next
End Sub
Private Sub SCR_StateChanged()
    chkConnected.Value = -CInt(SCR.IsConnected)
    chkScrittura.Value = -CInt(SCR.IsWriting)
    
    If SCR.MusicName <> Empty Then cname.Text = SCR.MusicName Else cname.Text = SCR.StreamName
    If SCR.MusicURL <> Empty Then curl.Text = SCR.MusicURL Else curl.Text = SCR.StreamURL
    curl.Text = ReplaceString(curl.Text)
    
    sname.Text = SCR.StreamName: sgenre.Text = SCR.StreamGenre
    surl.Text = SCR.StreamURL: sbr.Text = SCR.StreamBitRate & " Kbps"
    spub.Value = -CInt(SCR.StreamPublic): notices.Text = SCR.GetNotices()
    cname.Text = SCR.MusicName: curl.Text = SCR.MusicURL
    curl.Text = ReplaceString(curl.Text)
End Sub

Private Sub scrTime_Timer()
    On Local Error Resume Next
    received.Text = FormatSize(SCR.ReceivedBytes)
    written.Text = FormatSize(SCR.WrittenBytes)
    meta.Text = FormatSize(SCR.ReceivedMetaInfo)
    DoEvents
End Sub

Private Sub SDVolume_Change()
    lblVolume.Caption = SDVolume.Value
End Sub

Private Sub SDVolume_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Local Error Resume Next
    VLC.Volume = SDVolume.Value
    lblVolume.Caption = SDVolume.Value
End Sub

Private Sub SDVolume_Scroll()
    On Local Error Resume Next
    VLC.Volume = SDVolume.Value
    lblVolume.Caption = SDVolume.Value
End Sub


Private Sub tMP3_Timer()
    On Local Error Resume Next
    If MP3.isPlaying Then lblPosition = "Time: " & MP3.position
    ' .... Loop
    If MP3.position = MP3.length And strLoop = True Then
        cmdMP3Stop = True: cmdMP3Play = True
        lblDuration.Caption = "Total: " & MP3.length
    ' .... Play Next File
    ElseIf MP3.position = MP3.length And strLoop = False Then
        cmdMP3Stop = True
        lblDuration.Caption = "Total: " & MP3.length
    ' .... Stop Player
    ElseIf MP3.position = MP3.length And strLoop = False Then
        cmdMP3Stop = True: tMP3.Enabled = False
    End If
    If MP3.isPlaying Then lblDuration.Caption = "Total: " & MP3.length
End Sub

Private Sub TreeStationURL_DblClick()
    Dim sItem As String: Dim ojLink(0) As String
    Dim pos1 As Long: Dim pos2 As Long
    Dim textLine As String: Dim FileHandle As Integer
    Dim i As Integer: Dim x As Integer
    
    On Local Error GoTo ErrorHandler
    
    If cmdSaveToMP3.Caption = "&Stop Saving" Then
        
    If MsgBox("You must Stop saving files before selecting another Station!" & vbLf & vbLf & "Stop the Recording file's?", _
        vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
        cmdSaveToMP3 = True
    End If
    
    If cmdConnect.Caption = "&Disconnect" Then
        cmdConnect = True
    End If
    
    If cmdStreaming.Caption = "&Close Streaming" Then
        cmdStreaming = True
    End If
    
    If Me.TreeStationURL.SelectedItem.children > 0 Then
        Exit Sub
    Else
        sItem = Me.TreeStationURL.SelectedItem.Text
        RadioStationURL = StripLeft(sItem, "|", False) ' .... Only URL Station PLS = False
        RadioStationName = StripLeft(sItem, "|", True) ' .... Only the Radio Title = True
        
        txtStationURL.Text = RadioStationURL
        txtStationTitle.Text = RadioStationName
        
        'sURLS = RadioStationName
        
        ' .... Create Folder of Radio Song if not exist
        Call MakeNewDir(App.Path + "\Streaming Files List\" + sURLS)
        
        dwComplete = False
        
        ' .... Download file PLS ...
        Downloader.BeginDownload RadioStationURL, App.Path + "\Streaming Files List\" + sURLS + "\" + RadioStationName + ".pls"
        
        Do While Not dwComplete = True: DoEvents: Loop
        
        FileHandle = FreeFile
        
        Open App.Path + "\Streaming Files List\" + sURLS + "\" + RadioStationName + ".pls" For Input As #FileHandle
            Do While Not EOF(FileHandle)
                    Line Input #FileHandle, textLine
                DoEvents
            Loop
        Close #FileHandle
        
        lstStationLinks.Clear
        
        x = 1
        For i = 0 To 30
        If InStr(textLine, "File" & x & "=") > 0 Then
            pos1 = InStr(pos1 + 1, textLine, "File" & x & "=", vbTextCompare)
            pos2 = InStr(pos1 + 1, textLine, "Title", vbTextCompare)
            ojLink(0) = Mid$(textLine, pos1 + Len("File" & x & "="), pos2 - pos1 - Len("File" & x & "="))
            ojLink(0) = Replace(ojLink(0), " ", "")
            ojLink(0) = Mid$(ojLink(0), 1, Len(ojLink(0)) - 1)
            lstStationLinks.AddItem ojLink(0)
        End If
            x = x + 1
        Next i
        
        ' .... Remuve extra URL's
        If lstStationLinks.ListCount > 0 Then Call ChkLst(lstStationLinks)
    
        If lstStationLinks.ListCount = 1 Then
            RadioStationURL = lstStationLinks.List(0)
            txtStationURL.Text = RadioStationURL
            MsgBox "Radio Streaming URL of " & RadioStationName & ", loading success. To connect this Station click the button {Connect to Url}!", vbInformation, App.Title
        ElseIf lstStationLinks.ListCount > 1 Then
            MsgBox "The Radio Streaming URL " & RadioStationName & ", contains multiple IP's address!" _
            & vbCrLf & "Ok to display the list of IP's...", vbInformation, App.Title
            picBackGround(1).Visible = True
        ElseIf lstStationLinks.ListCount = 0 Then
            MsgBox "Error to retrive the Radio Streaming URL's of " & RadioStationName & "!" _
            & "Ok to continue!", vbExclamation, App.Title
        End If
   End If
   
Exit Sub

ErrorHandler:
        MsgBox "Error: " & Err.Number & "." & vbCrLf & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub

Private Sub TreeStationURL_Expand(ByVal Node As ComctlLib.Node)
Dim pos1 As Long: Dim pos2 As Long
Dim FF As Variant: Dim tmpArray As String
Dim objLink As HTMLLinkElement: Dim objMSHTML As New MSHTML.HTMLDocument
Dim objDocument As MSHTML.HTMLDocument

On Local Error GoTo ErrorHandler

If sURLS = Empty Then sURLS = "Unknown"

If Node.Child.Text = "..." Then
    Screen.MousePointer = 11
    TreeStationURL.Nodes.Remove Node.Child.Index

    If Node.Text = "RB/Urban" Then
        sURLS = "R&B/Urban"
    ElseIf Node.Text = "Classic RB" Then
        sURLS = "Classic R&B"
    ElseIf Node.Text = "Contemporary RB" Then
        sURLS = "Contemporary R&B"
    Else
        sURLS = Node.Text
    End If
    
    Set objDocument = objMSHTML.createDocumentFromUrl("http://www.shoutcast.com/radio/" & sURLS, vbNullString)
    
    While objDocument.readyState <> "complete"
        DoEvents
    Wend
    
    FF = FreeFile
    
     ' .... Create Folder of Radio Song if not exist
    Call MakeNewDir(App.Path + "\Streaming Files List\" + sURLS)
    
    Open App.Path + "\Streaming Files List\" + sURLS + "\" + sURLS + ".txt" For Output As FF
    
    For Each objLink In objDocument.links
        If InStr(objLink, "http://yp.shoutcast.com/sbin/tunein-station.pls?id=") Then
        
            If sHoldLink <> objLink Then
                
                sStation = objLink.Title
                sStation = Replace(sStation, "-", "")
                sStation = Replace(sStation, ".", "")
                sStation = Replace(sStation, "[", "")
                sStation = Replace(sStation, "]", "")
                sStation = Replace(sStation, "(", "")
                sStation = Replace(sStation, ")", "")
                sStation = Replace(sStation, ":", "")
                sStation = Replace(sStation, "/", "")
                sStation = Replace(sStation, "|", "")
                sStation = Replace(sStation, "\", "")
                
                TreeStationURL.Nodes.Add Node.Index, tvwChild, , sStation & "|" & objLink, 2
            
                Print #FF, "Radio title: " & sStation & vbCrLf & "Link: " _
                & objLink & vbCrLf & "====================================" & vbCrLf
            
            End If
        End If
        sHoldLink = objLink
    Next
    
    Close FF
    
    Screen.MousePointer = 0
End If
Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Number & "." & vbCrLf & Err.Description, vbExclamation, App.Title
        Screen.MousePointer = 0
    Err.Clear
End Sub


Private Sub tvcl_Timer()
    On Local Error Resume Next
    lblInfoStreaming.Caption = "Time: " & PositionTime(VLC.input.Time)
    DoEvents
End Sub


Private Sub WS_Connect()
    Dim Cmd, url
    url = sShoutCastURL
    Cmd = "GET " & url & " HTTP/1.0" & vbCrLf & "Accept: */*" & vbCrLf & "Accept: text/html" & vbCrLf & vbCrLf
    WS.SendData Cmd
End Sub

Private Sub WS_DataArrival(ByVal BytesTotal As Long)
    Dim sData As String
    WS.GetData sData, vbString
    sWholeData = sWholeData & sData
    If InStr(1, sWholeData, sEndTag, vbTextCompare) Then
            bDataReceived = True
        WS.Close
    End If
End Sub


Private Sub WS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Screen.MousePointer = vbDefault
    WriteErrorLogs Number, Description, Source & vbCrLf & "    FormMain {Sub: WS_Error}", True, True
End Sub



Private Function FormatSize(SIZE As Long) As String
    On Local Error Resume Next
    If SIZE > 1000000000 Then FormatSize = Format(SIZE / 1073741824, "0.0#") & " GBytes": Exit Function
    If SIZE > 1000000 Then FormatSize = Format(SIZE / 1048576, "0.0#") & " MBytes": Exit Function
    If SIZE > 1000 Then FormatSize = Format(SIZE / 1024, "0.0#") & " KBytes": Exit Function
    FormatSize = Format(SIZE, "0.0#") & " Bytes"
End Function

Private Function FormatFileName(FileName As String)
    Dim FF As String
    On Local Error Resume Next
    FF = FileName
    FF = Replace(FF, "/", "_")
    FF = Replace(FF, "\", "_")
    FF = Replace(FF, "*", "_")
    FF = Replace(FF, "?", "_")
    FF = Replace(FF, ":", "_")
    FF = Replace(FF, """", "_")
    FF = Replace(FF, "<", "_")
    FF = Replace(FF, ">", "_")
    FF = Replace(FF, "|", "_")
    FF = Replace(FF, ".", "_")
    FF = Replace(FF, ";", "_")
    FormatFileName = FF
End Function

Private Function MakeDir(sStation As String) As Boolean
    On Local Error GoTo ErrorHandler
    newDir = IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\") + "Streaming Files List\" + sURLS + "\" + sStation + "\"
    MakeSureDirectoryPathExists newDir
    MakeDir = True
Exit Function
ErrorHandler:
        MakeDir = False
    Err.Clear
End Function

Private Function SaveNewRadioLink(Optional newLink As String = "", Optional saveOnlyList As Boolean = False) As Boolean
    Dim TempText As String
    Dim FF As Integer
    Dim i As Integer
    On Local Error GoTo ErrHandler
    FF = FreeFile
    
    If cmbRadioStation.ListCount = 0 Then Exit Function
    
If saveOnlyList = False Then
    If FileExists(App.Path + "\Stations.txt") Then
        Open App.Path + "\Stations.txt" For Append As #FF
        Print #FF, "|" & newLink
    Else
        Open App.Path + "\Stations.txt" For Output As #FF
        For i = 0 To cmbRadioStation.ListCount
            If cmbRadioStation.List(i) <> Empty Then
                TempText = TempText & cmbRadioStation.List(i) & "|"
            End If
        DoEvents
    Next i
        TempText = TempText & newLink & "|"
        Print #FF, Mid$(TempText, 1, Len(TempText) - 1)
    End If
    Close #FF
    SaveNewRadioLink = True
    cmbRadioStation.AddItem newLink
ElseIf saveOnlyList Then
    Open App.Path + "\Stations.txt" For Output As #FF
        For i = 0 To cmbRadioStation.ListCount
            If cmbRadioStation.List(i) <> "" Then
                TempText = TempText & cmbRadioStation.List(i) & "|"
            End If
        DoEvents
    Next i
        Print #FF, Mid$(TempText, 1, Len(TempText) - 1)
    Close #FF
    SaveNewRadioLink = True
    End If
    Exit Function
ErrHandler:
    SaveNewRadioLink = False
    On Error GoTo 0
        Close #FF
        WriteErrorLogs Err.Number, Err.Description, "FormMain {Sub: SaveNewRadioLink}", True, True
    Err.Clear
End Function

Private Sub AddRadioStation()
    Dim i As Integer
    Dim strLines() As String
    Dim textLine As String
    Dim FF As Integer
    On Local Error GoTo ErrHandler
    FF = FreeFile
    cmbRadioStation.Clear
    If FileExists(App.Path + "\Stations.txt") Then
        Open App.Path + "\Stations.txt" For Input As #FF
            Do While Not EOF(FF)
                Line Input #FF, textLine
                    strLines = Split(textLine, "|")
                For i = 0 To UBound(strLines)
                        cmbRadioStation.AddItem strLines(i) 'StripString(Mid$(strLines(i), 1, Len(strLines(i))), "-", True) & _
                        StripString(Mid$(strLines(i), 1, Len(strLines(i))), "-", False)
                    DoEvents
                Next i
            Loop
        Close #FF
    Else
        strLines = Split(bSTATIONS, "|")
        For i = 0 To UBound(strLines)
                cmbRadioStation.AddItem strLines(i) 'StripString(Mid$(strLines(i), 1, Len(strLines(i))), "-", True) & _
                StripString(Mid$(strLines(i), 1, Len(strLines(i))), "-", False)
            DoEvents
        Next i
    End If
        If cmbRadioStation.ListCount > 0 Then cmbRadioStation.ListIndex = 0
    Exit Sub
ErrHandler:
    WriteErrorLogs Err.Number, Err.Description, "FormMain {Sub: AddRadioStation}", True, True
    Err.Clear
Exit Sub
End Sub

Private Function StripString(ByVal strString As String, Optional sChar As String = "-", Optional RightToLeft As Boolean = False) As String
    Dim i As Integer
    Dim sTmp As String
    On Error Resume Next
    sTmp = Mid$(strString, i + 1, Len(strString))
    For i = 1 To Len(sTmp)
        If Mid(sTmp, i, 1) = sChar And RightToLeft = True Then
                StripString = Mid$(strString, InStrRev(strString, "http", , vbTextCompare))
            Exit Function
        ElseIf Mid(sTmp, i, 1) = sChar And RightToLeft = False Then
            Exit For
        Else
            StripString = Mid$(strString, i + 2, Len(strString))
        End If
    Next
     StripString = Left$(sTmp, i - 1)
End Function

Private Function ChkLst(iList As ListBox)
    On Local Error Resume Next
    Dim i As Integer, x As Integer
        For i = 0 To iList.ListCount - 1
            For x = 0 To iList.ListCount - 1
                If (iList.List(i) = iList.List(x)) And x <> i Then
                    iList.RemoveItem (i)
                End If
            Next x
        Next i
End Function

Private Function VLCLocation() As String
    On Local Error GoTo ErrorHandler
    If Len(Dir$(SpecialFolders.SpecialFolderPath(CSIDL_PROGRAM_FILES) & "VideoLAN\VLC\axvlc.dll")) > 0 Then
        VLCLocation = SpecialFolders.SpecialFolderPath(CSIDL_PROGRAM_FILES) & "VideoLAN\VLC\axvlc.dll"
    End If
Exit Function
ErrorHandler:
        VLCLocation = "Error #" & Err.Number & ". " & Err.Description
    Err.Clear
End Function

Private Function FormatSizeDUE(SIZE As Long) As String
    On Local Error Resume Next
    If SIZE > 1000000000 Then FormatSizeDUE = Format(SIZE / 1073741824, "0.0#"): Exit Function
    If SIZE > 1000000 Then FormatSizeDUE = Format(SIZE / 1048576, "0.0#"): Exit Function
    If SIZE > 1000 Then FormatSizeDUE = Format(SIZE / 1024, "0.0#"): Exit Function
    FormatSizeDUE = Format(SIZE, "0.0#"): Exit Function
End Function

Private Sub SendEmail(Optional Adress As String, Optional Subjet As String, _
            Optional Content As String, Optional CC As String, Optional CCC As String)
    Dim temp As String
    On Local Error Resume Next
    If Len(Subjet) Then temp = "&Subject=" & Subjet
    If Len(Content) Then temp = temp & "&Body=" & Content
    If Len(CC) Then temp = temp & "&CC=" & CC
    If Len(CCC) Then temp = temp & "&BCC=" & CCC
    If Mid(temp, 1, 1) = "&" Then Mid(temp, 1, 1) = "?"
    temp = "mailto:" & Adress & temp
    Call ShellExecute(Me.hWnd, "open", temp, vbNullString, vbNullString, SW_SHOWNORMAL)
End Sub

Private Sub SetIcon(sTypeIcon As sSetIcon)
    On Local Error Resume Next
    Select Case sTypeIcon
    Case 0
        m_frmSysTray.IconHandle = m_frmSysTray.imgIcon(0).Picture.Handle
    Case 1
        m_frmSysTray.IconHandle = m_frmSysTray.imgIcon(1).Picture.Handle
    Case 2
        m_frmSysTray.IconHandle = m_frmSysTray.imgIcon(2).Picture.Handle
    Case 3
        m_frmSysTray.IconHandle = m_frmSysTray.imgIcon(3).Picture.Handle
    Case 4
        m_frmSysTray.IconHandle = m_frmSysTray.imgIcon(4).Picture.Handle
    Case 5
        m_frmSysTray.IconHandle = m_frmSysTray.imgIcon(5).Picture.Handle
    Case 6
        m_frmSysTray.IconHandle = m_frmSysTray.imgIcon(6).Picture.Handle
    End Select
End Sub

Private Sub ShowTip(strMessage As String, strTitle As String)
    On Local Error Resume Next
    m_frmSysTray.ShowBalloonTip strMessage, strTitle, NIIF_INFO
End Sub

Private Function GetSysTray(ByVal sShowOrsHide As Boolean) As Boolean
    On Local Error GoTo ErrorHandler
    If sShowOrsHide = True Then
        Set m_frmSysTray = New frmSysTray
        With m_frmSysTray
            .AddMenuItem "&Restore Radio Streaming v1.0.2b", "Open", True
            .AddMenuItem "-"
            .AddMenuItem "&Netshadows on the Web", "netshadows"
            .AddMenuItem "&Mail to Salvo Cortesiano", "salvocortesiano"
            .AddMenuItem "-"
            .AddMenuItem "&Download the Installation Project", "Download"
            .AddMenuItem "&Download the (Full binary) Project", "Full-binary"
            .AddMenuItem "&Download the VLC Plg-In", "Plg-In"
            .AddMenuItem "-"
            .AddMenuItem "&About...", "About"
            .AddMenuItem "-"
            .AddMenuItem "&Hide Icon Tray", "HideTray"
            .AddMenuItem "&Change Icon Tray (=1)", "ChangeIcon1"
            .AddMenuItem "&Change Icon Tray (=2)", "ChangeIcon2"
            .AddMenuItem "&Change Icon Tray (=3)", "ChangeIcon3"
            .AddMenuItem "&Change Icon Tray (=4)", "ChangeIcon4"
            .AddMenuItem "&Change Icon Tray (=5)", "ChangeIcon5"
            .AddMenuItem "&Default Icon Tray", "DefaultIconTray"
            .AddMenuItem "-"
            .AddMenuItem "&Close Radio Streaming v1.0.2b", "Close"
            .ToolTip = "Radio Streaming v1.0.2b"
        End With
    ElseIf sShowOrsHide = False Then
        Unload m_frmSysTray
        Set m_frmSysTray = Nothing
    End If
    GetSysTray = True
Exit Function
ErrorHandler:
        GetSysTray = False
    Err.Clear
End Function

Private Function SaveSettingINI() As Boolean
    On Local Error GoTo ErrorHandler
    
    INI.DeleteKey "STATION", "RADIOSTATION"
    If cmbRadioStation.ListCount > 0 Then _
    INI.CreateKeyValue "STATION", "RADIOSTATION", cmbRadioStation.ListIndex _
    Else INI.CreateKeyValue "STATION", "RADIOSTATION", 0
    
    SaveSettingINI = True
Exit Function
ErrorHandler:
    WriteErrorLogs Err.Number, Err.Description, "FormMain {Function: SaveSettingINI}", True, True
        SaveSettingINI = False
    Err.Clear
End Function

Private Function ReplaceString(strString As String) As String
    On Error Resume Next
    ReplaceString = Replace(strString, "&amp;quot;", sQuote)
    ReplaceString = Replace(strString, "&amp;#39;", "'")
    ReplaceString = Replace(strString, "&#231;", "ç")
    ReplaceString = Replace(strString, "&#232;", "è")
    ReplaceString = Replace(strString, "&#233;", "é")
    ReplaceString = Replace(strString, "&#224;", "à")
    ReplaceString = Replace(strString, "&#242;", "ò")
    ReplaceString = Replace(strString, "&#249;", "ù")
    ReplaceString = Replace(strString, ":", "-")
    ReplaceString = Replace(strString, ";", "-")
    ReplaceString = Replace(strString, """", "'")
    ReplaceString = Replace(strString, "+", "-")
    ReplaceString = Replace(strString, ".", "_")
    ReplaceString = Replace(strString, "|", "-")
    ReplaceString = Replace(strString, "%", "")
    ReplaceString = Replace(strString, "$", "(dollar)")
    ReplaceString = Replace(strString, "£", "(lit)")
    ReplaceString = Replace(strString, "!", "(esclam)")
    ReplaceString = Replace(strString, "&quot;", "'")
    ReplaceString = Replace(strString, "%20", " ")
    ReplaceString = Replace(strString, "%2D", " ")
    ReplaceString = Replace(strString, "&artist=", "")
    ReplaceString = Replace(strString, "&title=", "-")
    ReplaceString = Replace(strString, "&album=", "-")
    ReplaceString = Replace(strString, "%27n", " and ")
    ReplaceString = Replace(strString, "&duration=", "")
    ReplaceString = Replace(strString, "&songtype=o&overlay=NO&buycd=&website=&picture=icFile", "")
End Function

Private Sub HideAppWinTask(Optional curProcess As Long = 1)
    On Local Error GoTo ErrorHandler
    Call RegisterServiceProcess(GetCurrentProcessId, 1)
Exit Sub
ErrorHandler:
    Err.Clear
End Sub

Private Sub RetriveExtraTags()
    On Local Error Resume Next
    ' .... Extra ID3 TAG's
    If INI.GetKeyValue("ID3EXTRATAGINFO", "ENCODEBY") <> Empty Then
        SCR.ID3EncodedBy = INI.GetKeyValue("ID3EXTRATAGINFO", "ENCODEBY")
    Else
        SCR.ID3EncodedBy = App.EXEName & " by Salvo Cortesiano"
    End If
    
    If INI.GetKeyValue("ID3EXTRATAGINFO", "COPYRIGHT") <> Empty Then
        SCR.ID3Copyright = INI.GetKeyValue("ID3EXTRATAGINFO", "COPYRIGHT")
    Else
        SCR.ID3Copyright = "http://www.netshadows.it"
    End If
    
    If INI.GetKeyValue("ID3EXTRATAGINFO", "LANGUAGE") <> Empty Then
        SCR.ID3Languages = INI.GetKeyValue("ID3EXTRATAGINFO", "LANGUAGE")
    Else
        SCR.ID3Languages = "Italians Language"
    End If
    
    If INI.GetKeyValue("ID3EXTRATAGINFO", "COMMENTS") <> Empty Then
        SCR.ID3Comments = INI.GetKeyValue("ID3EXTRATAGINFO", "COMMENTS")
    Else
        SCR.ID3Comments = "For more music and Softwares go to http://www.netshadows.it/leombredellarete/forum/"
    End If
End Sub
