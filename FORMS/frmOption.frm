VERSION 5.00
Begin VB.Form frmOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Radio Streaming Option"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOption.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "ID3 Extra TAG's"
      Height          =   1770
      Left            =   75
      TabIndex        =   14
      Top             =   3855
      Width           =   6240
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   1560
         Left            =   30
         ScaleHeight     =   1560
         ScaleWidth      =   6165
         TabIndex        =   15
         Top             =   165
         Width           =   6165
         Begin VB.TextBox txtComments 
            Height          =   315
            Left            =   1155
            TabIndex        =   23
            Top             =   1200
            Width           =   4965
         End
         Begin VB.TextBox txtLanguages 
            Height          =   315
            Left            =   1155
            TabIndex        =   21
            Top             =   840
            Width           =   4965
         End
         Begin VB.TextBox txtCopyrightInfo 
            Height          =   315
            Left            =   1665
            TabIndex        =   19
            Top             =   465
            Width           =   4455
         End
         Begin VB.TextBox txtEncodeBy 
            Height          =   315
            Left            =   1305
            TabIndex        =   17
            Top             =   90
            Width           =   4815
         End
         Begin VB.Label Label4 
            Caption         =   "Comments:"
            Height          =   255
            Left            =   75
            TabIndex        =   22
            Top             =   1245
            Width           =   1050
         End
         Begin VB.Label Label3 
            Caption         =   "Languages:"
            Height          =   255
            Left            =   75
            TabIndex        =   20
            Top             =   855
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Copyright Info:"
            Height          =   255
            Left            =   75
            TabIndex        =   18
            Top             =   480
            Width           =   1620
         End
         Begin VB.Label Label1 
            Caption         =   "Encoded By:"
            Height          =   255
            Left            =   75
            TabIndex        =   16
            Top             =   120
            Width           =   1260
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Oter Option"
      Height          =   735
      Left            =   75
      TabIndex        =   10
      Top             =   3000
      Width           =   6240
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   30
         ScaleHeight     =   465
         ScaleWidth      =   6165
         TabIndex        =   11
         Top             =   225
         Width           =   6165
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "Show &Update"
            Height          =   315
            Left            =   4200
            TabIndex        =   24
            Top             =   60
            Width           =   1905
         End
         Begin VB.CommandButton cmdMailSender 
            Caption         =   "Show &Mail Sender"
            Height          =   315
            Left            =   2145
            TabIndex        =   13
            Top             =   60
            Width           =   1905
         End
         Begin VB.CommandButton cmdInstallaLibrary 
            Caption         =   "Install VLC Lib"
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   1980
         End
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   120
      Picture         =   "frmOption.frx":3D52
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   3
      Top             =   45
      Width           =   510
   End
   Begin VB.Frame Frame1 
      Caption         =   "     Radio Streaming Option"
      Height          =   2670
      Left            =   75
      TabIndex        =   1
      Top             =   240
      Width           =   6225
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2400
         Left            =   45
         ScaleHeight     =   2400
         ScaleWidth      =   6120
         TabIndex        =   2
         Top             =   210
         Width           =   6120
         Begin VB.CheckBox CheckWriteTAG 
            Caption         =   "Write MP3 Info Tags of File..."
            Height          =   285
            Left            =   60
            TabIndex        =   9
            Top             =   2070
            Value           =   1  'Checked
            Width           =   6015
         End
         Begin VB.CheckBox CheckWinStartUp 
            Caption         =   "Launch the Program when Windows starts..."
            Height          =   285
            Left            =   60
            TabIndex        =   8
            Top             =   1695
            Width           =   6015
         End
         Begin VB.CheckBox chkSaveStation 
            Caption         =   "&Save Station Link List on Exit Program..."
            Height          =   255
            Left            =   60
            TabIndex        =   7
            Top             =   1365
            Value           =   1  'Checked
            Width           =   6015
         End
         Begin VB.CheckBox CheckDivide 
            Caption         =   "Create single audio file for each variation of Streaming Title {Artist and Song title}..."
            Height          =   465
            Left            =   60
            TabIndex        =   6
            Top             =   810
            Value           =   1  'Checked
            Width           =   6015
         End
         Begin VB.CheckBox CheckDelete 
            Caption         =   "Delete Audio File < 1.000 Mega bytes (meybe)..."
            Height          =   285
            Left            =   60
            TabIndex        =   5
            Top             =   450
            Value           =   1  'Checked
            Width           =   6015
         End
         Begin VB.CheckBox CheckAppend 
            Caption         =   "Append the Captured Audio to one File only..."
            Height          =   285
            Left            =   60
            TabIndex        =   4
            Top             =   120
            Width           =   6015
         End
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Continue"
      Height          =   315
      Left            =   4980
      TabIndex        =   0
      Top             =   5760
      Width           =   1275
   End
End
Attribute VB_Name = "frmOption"
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

Private sCancel3 As Boolean

Private Sub CheckAppend_Click()
    On Local Error Resume Next
    If CheckAppend.Value = 1 Then
        CheckDelete.Value = 0: CheckDivide.Value = 0
        AppendSeek = True
    Else
        AppendSeek = False
    End If
End Sub

Private Sub CheckDelete_Click()
    On Local Error Resume Next
    If CheckDelete.Value = 1 Then sCheckDelete = True Else sCheckDelete = False
End Sub


Private Sub CheckDivide_Click()
    If CheckDivide.Value = 1 Then
        CheckAppend.Value = 0
        sCheckDivide = True
    Else
        sCheckDivide = False
    End If
End Sub


Private Sub CheckWinStartUp_Click()
    On Local Error Resume Next
    If CheckWinStartUp.Value = 1 Then
        Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Radio Streaming and Recorder v12", App.Path + "\" + App.EXEName + ".exe")
    ElseIf CheckWinStartUp.Value = 0 Then
        Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Radio Streaming and Recorder v12")
    End If
End Sub


Private Sub CheckWriteTAG_Click()
    On Local Error Resume Next
    If CheckWriteTAG.Value = 1 Then
        sWriteTagOfTrack = True
        frmStreamingRadio.SCR.WriteID3TagToFile = True
    Else
        sWriteTagOfTrack = False
        frmStreamingRadio.SCR.WriteID3TagToFile = False
    End If
End Sub

Private Sub chkSaveStation_Click()
    On Local Error Resume Next
    If chkSaveStation.Value = 1 Then schkSaveStation = True Else schkSaveStation = False
End Sub

Private Sub cmdExit_Click()
    sCancel3 = False: Unload Me
End Sub

Private Sub cmdInstallaLibrary_Click()
    On Local Error Resume Next
    If MsgBox("This {Command} install the VLC v1.1.8 Library!" & vbLf & vbLf & "Are you sure to Continue?", _
        vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
    If InstallLibrary(SW_NORMAL) = False Then:
End Sub

Private Sub cmdMailSender_Click()
    On Local Error Resume Next
    frmSender.Show vbModal, Me
End Sub

Private Sub cmdUpdate_Click()
    On Local Error Resume Next
    frmAutoUpdate.Show vbModal, Me
End Sub

Private Sub Form_Load()

    sCancel3 = True
    
    ' .... Write TAGs
    If INI.GetKeyValue("SETTING", "WRITEMP3TAGS") <> Empty Then
        CheckWriteTAG.Value = INI.GetKeyValue("SETTING", "WRITEMP3TAGS")
        If INI.GetKeyValue("SETTING", "WRITEMP3TAGS") = 1 Then _
        frmStreamingRadio.SCR.WriteID3TagToFile = True _
        Else frmStreamingRadio.SCR.WriteID3TagToFile = False
    Else
        CheckWriteTAG.Value = 0
        frmStreamingRadio.SCR.WriteID3TagToFile = False
    End If
    
    ' .... Append SONG to File
    If INI.GetKeyValue("SETTING", "APPEND") <> Empty Then
        CheckAppend.Value = INI.GetKeyValue("SETTING", "APPEND")
    Else
        CheckAppend.Value = 0
    End If
    
    ' .... Delete File?
    If INI.GetKeyValue("SETTING", "DELETE") <> Empty Then _
            CheckDelete.Value = INI.GetKeyValue("SETTING", "DELETE") _
            Else CheckDelete.Value = 1
    
    ' .... Create single File?
    If INI.GetKeyValue("SETTING", "CREATESINGLEFILE") <> Empty Then _
            CheckDivide.Value = INI.GetKeyValue("SETTING", "CREATESINGLEFILE") _
            Else CheckDivide.Value = 1
    
    ' .... Save Link Radio Station on Exit?
    If INI.GetKeyValue("SETTING", "SAVELINKONEXIT") <> Empty Then _
            chkSaveStation.Value = INI.GetKeyValue("SETTING", "SAVELINKONEXIT") _
            Else chkSaveStation.Value = 1
    
    ' .... Open Program on Windows StartUp
    If INI.GetKeyValue("SETTING", "WINSTARTUP") <> Empty Then _
            CheckWinStartUp.Value = INI.GetKeyValue("SETTING", "WINSTARTUP") _
            Else CheckWinStartUp.Value = 0
    
    
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
        ID3TagComments = "For more music and Softwares go to http://www.netshadows.it/leombredellarete/forum/"
    End If
    
    txtEncodeBy.Text = ID3TagEncodedBy: txtCopyrightInfo.Text = ID3TagCopyrightInfo
    txtLanguages.Text = ID3TagLanguages: txtComments.Text = ID3TagComments
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettingINI: Cancel = sCancel3
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmOption = Nothing
End Sub



Private Function SaveSettingINI() As Boolean
    On Local Error GoTo ErrorHandler
    
    INI.DeleteKey "SETTING", "WINSTARTUP"
    INI.CreateKeyValue "SETTING", "WINSTARTUP", CheckWinStartUp.Value
    
    INI.DeleteKey "SETTING", "WRITEMP3TAGS"
    INI.CreateKeyValue "SETTING", "WRITEMP3TAGS", CheckWriteTAG.Value
    
    INI.DeleteKey "SETTING", "APPEND"
    INI.CreateKeyValue "SETTING", "APPEND", CheckAppend.Value
    
    INI.DeleteKey "SETTING", "DELETE"
    INI.CreateKeyValue "SETTING", "DELETE", CheckDelete.Value
    
    INI.DeleteKey "SETTING", "CREATESINGLEFILE"
    INI.CreateKeyValue "SETTING", "CREATESINGLEFILE", CheckDivide.Value
    
    INI.DeleteKey "SETTING", "SAVELINKONEXIT"
    INI.CreateKeyValue "SETTING", "SAVELINKONEXIT", chkSaveStation.Value
    
    ' .... Extra TAG's MP3 Info
    INI.DeleteKey "ID3EXTRATAGINFO", "ENCODEBY"
    INI.CreateKeyValue "ID3EXTRATAGINFO", "ENCODEBY", txtEncodeBy.Text
    
    INI.DeleteKey "ID3EXTRATAGINFO", "COPYRIGHT"
    INI.CreateKeyValue "ID3EXTRATAGINFO", "COPYRIGHT", txtCopyrightInfo.Text
    
    INI.DeleteKey "ID3EXTRATAGINFO", "LANGUAGE"
    INI.CreateKeyValue "ID3EXTRATAGINFO", "LANGUAGE", txtLanguages.Text
    
    INI.DeleteKey "ID3EXTRATAGINFO", "COMMENTS"
    INI.CreateKeyValue "ID3EXTRATAGINFO", "COMMENTS", txtComments.Text
    
    SaveSettingINI = True
Exit Function
ErrorHandler:
    WriteErrorLogs Err.Number, Err.Description, "Form Option {Function: SaveSettingINI}", True, True
    SaveSettingINI = False: Err.Clear
End Function
