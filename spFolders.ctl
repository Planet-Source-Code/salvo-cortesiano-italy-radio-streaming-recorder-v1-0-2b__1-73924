VERSION 5.00
Begin VB.UserControl SpecialFolders 
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   255
   InvisibleAtRuntime=   -1  'True
   Picture         =   "spFolders.ctx":0000
   ScaleHeight     =   255
   ScaleWidth      =   255
   ToolboxBitmap   =   "spFolders.ctx":031A
End
Attribute VB_Name = "SpecialFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Author Link: [http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=52229&lngWId=1
' Project revisited by Salvo Cortesiano
' ****************************************************************************************************
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

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetFolderPath Lib "ShFolder" Alias "SHGetFolderPathA" (ByVal hwnd As Long, ByVal CSIDL As Long, ByVal TokenHandle As Long, ByVal Flags As Long, ByVal lpPath As String) As Long

Private Type SHORTITEMID
    cb As Long
    abID As Integer
End Type

Private Type ITEMIDLIST
    mkid As SHORTITEMID
End Type


Public Enum MySpecialFolders
    CSIDL_DESKTOP = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_LOCAL_APPDATA = &H1C
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_WINDOWS = &H24
    CSIDL_SYSTEM = &H25
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_MYPICTURES = &H27
    CSIDL_PROFILE = &H28
    CSIDL_SYSTEMX86 = &H29
    CSIDL_PROGRAM_FILESX86 = &H2A
    CSIDL_PROGRAM_FILES_COMMON = &H2B
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_ADMINTOOLS = &H2F
End Enum
Public Property Get SpecialFolderPath(ByVal Folder As MySpecialFolders) As String
    SpecialFolderPath = GetSpecialFolder(Folder)
End Property

Private Function GetSpecialFolder(ByVal lngCSIDL As Long) As String
    Dim udtIDL As ITEMIDLIST
    Dim lngRtn As Long
    Dim strFolder As String
    Dim Path As String * 260
    lngRtn = SHGetSpecialFolderLocation(hwnd, lngCSIDL, udtIDL)
    If lngRtn = 0 Then
        strFolder = Space$(260)
        lngRtn = SHGetPathFromIDList(ByVal udtIDL.mkid.cb, ByVal strFolder)
        If lngRtn Then
            GetSpecialFolder = Left$(strFolder, InStr(1, strFolder, Chr$(0)) - 1) & "\"
        End If
    Else
        lngRtn = SHGetFolderPath(hwnd, lngCSIDL, 0, 0, Path)
        If lngRtn = 0 Then
            strFolder = Space$(260)
            lngRtn = SHGetPathFromIDList(ByVal udtIDL.mkid.cb, ByVal strFolder)
            If lngRtn Then
                GetSpecialFolder = Left$(strFolder, InStr(1, strFolder, Chr$(0)) - 1) & "\"
            End If
        End If
    End If
End Function

Private Sub UserControl_Resize()
    Width = 255
    Height = 210
End Sub
