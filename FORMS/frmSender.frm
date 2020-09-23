VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSender 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Radio Streaming-Sender Email v1.1.2"
   ClientHeight    =   6870
   ClientLeft      =   1755
   ClientTop       =   1710
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSender.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   7455
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox listAttach 
      Height          =   270
      Left            =   1890
      TabIndex        =   45
      Top             =   3450
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CheckBox CheckHTML 
      Caption         =   "HTML Format"
      Height          =   285
      Left            =   4500
      TabIndex        =   44
      Top             =   1155
      Width           =   1560
   End
   Begin VB.CommandButton cmdRemuve 
      Caption         =   "[X]"
      Enabled         =   0   'False
      Height          =   300
      Left            =   6825
      TabIndex        =   43
      ToolTipText     =   "Delete the Last Attachment File..."
      Top             =   5070
      Width           =   540
   End
   Begin VB.Frame Frame1 
      Caption         =   "Internal SMTP"
      Height          =   1425
      Left            =   45
      TabIndex        =   39
      Top             =   3615
      Width           =   1635
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   30
         ScaleHeight     =   1200
         ScaleWidth      =   1545
         TabIndex        =   40
         Top             =   180
         Width           =   1545
         Begin VB.CheckBox CheckInterlAutentication 
            Caption         =   "Use: Internal Autentication Server"
            Height          =   1095
            Left            =   30
            TabIndex        =   41
            Top             =   60
            Width           =   1470
         End
      End
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   34
      Top             =   765
      Width           =   4200
   End
   Begin VB.CheckBox CheckAutentication 
      Caption         =   "Autentication Required"
      Height          =   285
      Left            =   1755
      TabIndex        =   32
      Top             =   1155
      Width           =   2715
   End
   Begin VB.TextBox txtPassWord 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   30
      Top             =   420
      Width           =   4200
   End
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   420
      Top             =   345
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   2100
      Left            =   6120
      TabIndex        =   23
      Top             =   2550
      Width           =   1275
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   60
         ScaleHeight     =   630
         ScaleWidth      =   1155
         TabIndex        =   27
         Top             =   1380
         Width           =   1155
         Begin VB.OptionButton optSend 
            Caption         =   "Bulk"
            Height          =   255
            Index           =   1
            Left            =   60
            TabIndex        =   29
            ToolTipText     =   "Use the Connect, Send and Disconnect Methods."
            Top             =   330
            Width           =   1075
         End
         Begin VB.OptionButton optSend 
            Caption         =   "Normal"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   28
            ToolTipText     =   "Use the Send Method only for each message."
            Top             =   30
            Value           =   -1  'True
            Width           =   1075
         End
      End
      Begin VB.TextBox txtQty 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   24
         Text            =   "1"
         Top             =   1050
         Width           =   960
      End
      Begin VB.Label lblQty 
         Caption         =   "Messages to send:"
         Height          =   480
         Left            =   120
         TabIndex        =   25
         Top             =   300
         Width           =   1110
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close"
      Height          =   315
      Left            =   6120
      TabIndex        =   22
      Top             =   2010
      Width           =   1275
   End
   Begin VB.ListBox lstStatus 
      BackColor       =   &H8000000F&
      Height          =   1110
      Left            =   1740
      TabIndex        =   20
      Top             =   5430
      Width           =   3660
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   300
      Left            =   6285
      TabIndex        =   8
      ToolTipText     =   "Browser Attachment File's..."
      Top             =   5070
      Width           =   495
   End
   Begin VB.TextBox txtAttach 
      Height          =   285
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5070
      Width           =   3915
   End
   Begin VB.TextBox txtMsg 
      Height          =   1680
      Left            =   1740
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3330
      Width           =   4200
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1740
      TabIndex        =   5
      Top             =   2970
      Width           =   4200
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   1740
      TabIndex        =   2
      Top             =   1890
      Width           =   4200
   End
   Begin VB.TextBox txtFromName 
      Height          =   285
      Left            =   1740
      TabIndex        =   1
      Top             =   1530
      Width           =   4200
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   1740
      TabIndex        =   4
      Top             =   2610
      Width           =   4200
   End
   Begin VB.TextBox txtToName 
      Height          =   285
      Left            =   1740
      TabIndex        =   3
      Top             =   2250
      Width           =   4200
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   60
      Width           =   4200
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   315
      Left            =   6120
      TabIndex        =   10
      Top             =   1590
      Width           =   1275
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   315
      Left            =   6120
      TabIndex        =   9
      Top             =   1200
      Width           =   1275
   End
   Begin VB.Label labellblAttach 
      Alignment       =   2  'Center
      Caption         =   "(0)"
      Height          =   270
      Left            =   5685
      TabIndex        =   46
      Top             =   5100
      Width           =   555
   End
   Begin VB.Label Label4 
      Caption         =   "This Sender is in Beta testing. For any problem, sorry!"
      Height          =   915
      Left            =   5565
      TabIndex        =   42
      Top             =   5580
      Width           =   1740
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   3  'Dot
      Height          =   1050
      Left            =   5490
      Shape           =   4  'Rounded Rectangle
      Top             =   5490
      Width           =   1860
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   930
      Picture         =   "frmSender.frx":3D52
      Top             =   5850
      Width           =   480
   End
   Begin VB.Label lblSNTPUser 
      Caption         =   "Hidden"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   6360
      TabIndex        =   38
      Top             =   735
      Width           =   1035
   End
   Begin VB.Label lblSMTPPassWord 
      Caption         =   "Hidden"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   6360
      TabIndex        =   37
      Top             =   420
      Width           =   1035
   End
   Begin VB.Label lblSMTP 
      Caption         =   "Hidden"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   6360
      TabIndex        =   36
      Top             =   90
      Width           =   1035
   End
   Begin VB.Image imgsMine 
      Height          =   240
      Index           =   2
      Left            =   5970
      Picture         =   "frmSender.frx":7AA4
      Tag             =   "Show"
      ToolTipText     =   "Show"
      Top             =   765
      Width           =   240
   End
   Begin VB.Image imgsMine 
      Height          =   240
      Index           =   1
      Left            =   5970
      Picture         =   "frmSender.frx":802E
      Tag             =   "Show"
      ToolTipText     =   "Show"
      Top             =   435
      Width           =   240
   End
   Begin VB.Image imgsMine 
      Height          =   240
      Index           =   0
      Left            =   5970
      Picture         =   "frmSender.frx":85B8
      Tag             =   "Show"
      ToolTipText     =   "Show"
      Top             =   105
      Width           =   240
   End
   Begin VB.Image imgs 
      Height          =   240
      Index           =   1
      Left            =   15
      Picture         =   "frmSender.frx":8B42
      Top             =   300
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgs 
      Height          =   240
      Index           =   0
      Left            =   30
      Picture         =   "frmSender.frx":90CC
      Top             =   30
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SMTP UserName"
      Height          =   210
      Left            =   120
      TabIndex        =   35
      Top             =   780
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Autentication"
      Height          =   210
      Left            =   135
      TabIndex        =   33
      Top             =   1170
      Width           =   1545
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SMTP PassWord"
      Height          =   210
      Left            =   120
      TabIndex        =   31
      Top             =   450
      Width           =   1545
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3795
      TabIndex        =   26
      Top             =   6585
      Width           =   1590
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status Sender"
      Height          =   210
      Left            =   150
      TabIndex        =   21
      Top             =   5430
      Width           =   1500
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Progress"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   225
      TabIndex        =   19
      Top             =   6600
      Width           =   1740
   End
   Begin VB.Label lblAttach 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attachment"
      Height          =   210
      Left            =   255
      TabIndex        =   18
      Top             =   5130
      Width           =   1455
   End
   Begin VB.Label lblMsg 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      Height          =   210
      Left            =   210
      TabIndex        =   17
      Top             =   3330
      Width           =   1515
   End
   Begin VB.Label lblSubject 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      Height          =   210
      Left            =   180
      TabIndex        =   16
      Top             =   2970
      Width           =   1545
   End
   Begin VB.Label lblFrom 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sender Email"
      Height          =   210
      Left            =   180
      TabIndex        =   15
      Top             =   1965
      Width           =   1545
   End
   Begin VB.Label lblFromName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sender Name"
      Height          =   210
      Left            =   180
      TabIndex        =   14
      Top             =   1590
      Width           =   1545
   End
   Begin VB.Label lblTo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Recipient Email"
      Height          =   210
      Left            =   75
      TabIndex        =   13
      Top             =   2670
      Width           =   1650
   End
   Begin VB.Label lblToName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Recipient Name"
      Height          =   210
      Left            =   165
      TabIndex        =   12
      Top             =   2295
      Width           =   1545
   End
   Begin VB.Label lblServer 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SMTP Server"
      Height          =   210
      Left            =   105
      TabIndex        =   11
      Top             =   105
      Width           =   1545
   End
End
Attribute VB_Name = "frmSender"
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

Option Compare Text

Private WithEvents poSendMail As Sender.clsSendMail
Attribute poSendMail.VB_VarHelpID = -1
Private bSendFailed As Boolean

Private sCancel2 As Boolean
Private sIndex As Long: Private sAutentication As Long: Private InternalAutentication As Long
Private Sub CheckAutentication_Click()
    On Local Error Resume Next
    sAutentication = CheckAutentication.Value
    If CheckAutentication.Value = 1 Then
        txtServer.Enabled = True: txtPassWord.Enabled = True: txtUserName.Enabled = True
        CheckInterlAutentication.Value = 0
    ElseIf CheckAutentication.Value = 0 Then
        txtServer.Enabled = False: txtPassWord.Enabled = False: txtUserName.Enabled = False
    End If
End Sub

Private Sub CheckInterlAutentication_Click()
    If CheckInterlAutentication.Value = 1 Then
        InternalAutentication = 1: CheckAutentication.Value = 0
    ElseIf CheckInterlAutentication.Value = 0 Then
        InternalAutentication = 0: CheckAutentication.Value = 1
    End If
End Sub

Private Sub cmdRemuve_Click()
    Dim i As Integer
    If listAttach.ListCount = 0 Then
        Exit Sub
    ElseIf listAttach.ListCount = 1 Then
        listAttach.RemoveItem (0)
        txtAttach.Text = Empty: cmdRemuve.Enabled = False
    ElseIf listAttach.ListCount > 1 Then
       listAttach.RemoveItem (listAttach.ListCount - 1)
       txtAttach.Text = Empty
       For i = 0 To listAttach.ListCount
            If listAttach.List(i) <> Empty Then _
            txtAttach.Text = txtAttach.Text & ";" & listAttach.List(i)
        Next i
        If txtAttach.Text = Empty Then cmdRemuve.Enabled = False
    End If
    labellblAttach.Caption = "(" & listAttach.ListCount & ")"
End Sub

Private Sub cmdSend_Click()

    Dim lCount As Long: Dim lCtr As Long: Dim t!

    cmdSend.Enabled = False: bSendFailed = False
    lstStatus.Clear: lblTime.Caption = ""
    Screen.MousePointer = vbHourglass
    
    With poSendMail
        
        .UseAuthentication = CheckAutentication.Value
        .AsHTML = CheckHTML.Value
        
        If CheckInterlAutentication.Value = 1 Then .InternalAutentication = True
        
        If CheckAutentication.Value = 1 Then
            ' .... Required for SMTP Server
            .UserName = txtUserName.Text
            .password = txtPassWord.Text
            .SMTPHost = txtServer.Text
        End If
        
        .From = txtFrom.Text
        .FromDisplayName = txtFromName.Text
        .Message = txtMsg.Text
        .Attachment = Trim(txtAttach.Text)

        lCount = Val(txtQty)
        If lCount = 0 Then Exit Sub
        t! = Timer

        If optSend(0).Value = True Then

            For lCtr = 1 To lCount
                .Recipient = txtTo.Text
                .RecipientDisplayName = txtToName.Text
                .Subject = txtSubject & " (Message # " & Str(lCtr) & ")"
                lblTime = "Sending message " & Str(lCtr)
                .send
            Next

        Else
            If .Connect Then
                For lCtr = 1 To lCount
                    lblTime = "Sending message " & Str(lCtr)
                    .Recipient = txtTo.Text
                    .RecipientDisplayName = txtToName.Text
                    .Subject = txtSubject & " (Message # " & Str(lCtr) & ")"
                    .send
                Next
                .Disconnect
            End If
        End If

    End With

    If Not bSendFailed Then lblTime.Caption = Str(lCount) & " Messages sent in " & Format$(Timer - t!, "#,##0.0") & " seconds."
    Screen.MousePointer = vbDefault
    cmdSend.Enabled = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettingToINI: Cancel = sCancel2
End Sub

Private Sub imgsMine_Click(Index As Integer)
    Select Case Index
        Case 0 ' Show SNTP Server
            If imgsMine(Index).Tag = "Show" Then
                txtServer.PasswordChar = Empty
                imgsMine(Index).Tag = "Hide"
                imgsMine(Index).Picture = imgs(1).Picture
                lblSMTP.Caption = "Visible"
                imgsMine(Index).ToolTipText = "Hidden"
            ElseIf imgsMine(Index).Tag = "Hide" Then
                txtServer.PasswordChar = "*"
                imgsMine(Index).Tag = "Show"
                imgsMine(Index).Picture = imgs(0).Picture
                lblSMTP.Caption = "Hidden"
                imgsMine(Index).ToolTipText = "Show"
            End If
        Case 1 ' Show SMTP PassWord
            If imgsMine(Index).Tag = "Show" Then
                txtPassWord.PasswordChar = Empty
                imgsMine(Index).Tag = "Hide"
                imgsMine(Index).Picture = imgs(1).Picture
                lblSMTPPassWord.Caption = "Visible"
                imgsMine(Index).ToolTipText = "Hidden"
            ElseIf imgsMine(Index).Tag = "Hide" Then
                txtPassWord.PasswordChar = "*"
                imgsMine(Index).Tag = "Show"
                imgsMine(Index).Picture = imgs(0).Picture
                lblSMTPPassWord.Caption = "Hidden"
                imgsMine(Index).ToolTipText = "Show"
            End If
        Case 2 ' Show SMTP UserName
            If imgsMine(Index).Tag = "Show" Then
                txtUserName.PasswordChar = Empty
                imgsMine(Index).Tag = "Hide"
                imgsMine(Index).Picture = imgs(1).Picture
                lblSNTPUser.Caption = "Visible"
                imgsMine(Index).ToolTipText = "Hidden"
            ElseIf imgsMine(Index).Tag = "Hide" Then
                txtUserName.PasswordChar = "*"
                imgsMine(Index).Tag = "Show"
                imgsMine(Index).Picture = imgs(0).Picture
                lblSNTPUser.Caption = "Hidden"
                imgsMine(Index).ToolTipText = "Show"
            End If
    End Select
End Sub

Private Sub optSend_Click(Index As Integer)
    sIndex = Index
End Sub

' *****************************************************************************
' The following four Subs capture the Events fired by the vbSendMail component
' *****************************************************************************

Private Sub poSendMail_Progress(lPercentCompete As Long)

    ' vbSendMail 'Progress Event'
    
    lblProgress = lPercentCompete & "% complete"

End Sub

Private Sub poSendMail_SendFailed(Explanation As String)

    ' vbSendMail 'SendFailed Event'

    MsgBox ("Your attempt to send mail failed for the following reason(s): " & vbCrLf & Explanation)
    bSendFailed = True
    lblProgress = ""
    lblTime = ""

End Sub

Private Sub poSendMail_SendSuccesful()

    ' vbSendMail 'SendSuccesful Event'

    lblProgress = "Send Successful!"

End Sub

Private Sub poSendMail_Status(status As String)

    ' vbSendMail 'Status Event'

    lstStatus.AddItem status
    lstStatus.ListIndex = lstStatus.ListCount - 1
    lstStatus.ListIndex = -1

End Sub

Private Sub Form_Load()

    Set poSendMail = New Sender.clsSendMail

    With poSendMail
        .SMTPHostValidation = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX
        .Delimiter = ";"
    End With

    cmDialog.flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    
    sCancel2 = True
    
    ' .... Data Sender
    If INI.GetKeyValue("SENDER", "MESSAGETOSEND") <> Empty Then _
            txtQty.Text = INI.GetKeyValue("SENDER", "MESSAGETOSEND") _
            Else txtQty.Text = "1"
    
    If INI.GetKeyValue("SENDER", "OPTIONSENDER") <> Empty Then
        sIndex = INI.GetKeyValue("SENDER", "OPTIONSENDER")
        optSend(sIndex).Value = True
    Else
        optSend(0).Value = True: sIndex = 0
    End If
    
    If INI.GetKeyValue("SENDER", "AUTENTICATION") <> Empty Then
        sAutentication = INI.GetKeyValue("SENDER", "AUTENTICATION")
        CheckAutentication.Value = sAutentication
        If sAutentication = 1 Then
            txtServer.Enabled = True: txtPassWord.Enabled = True: txtUserName.Enabled = True
        ElseIf sAutentication = 0 Then
            txtServer.Enabled = False: txtPassWord.Enabled = False: txtUserName.Enabled = False
        End If
    Else
        CheckAutentication.Value = 0: sAutentication = 0
        txtServer.Enabled = False: txtPassWord.Enabled = False: txtUserName.Enabled = False
    End If
    
    If INI.GetKeyValue("SENDER", "INTERNALAUTENTICATIONSERVER") <> Empty Then
        InternalAutentication = INI.GetKeyValue("SENDER", "INTERNALAUTENTICATIONSERVER")
        CheckInterlAutentication.Value = InternalAutentication
    Else
        InternalAutentication = 0: CheckInterlAutentication.Value = 0
    End If
    
    If INI.GetKeyValue("SENDER", "SENDASHTML") <> Empty Then _
        CheckHTML.Value = INI.GetKeyValue("SENDER", "SENDASHTML") _
    Else CheckHTML.Value = 0
    
    ' .... Sender Autentication
    If INI.GetKeyValue("SENDER", "SMTP") <> Empty Then _
            txtServer.Text = Decrypt(INI.GetKeyValue("SENDER", "SMTP"), True) _
            Else txtServer.Text = Empty
            
    If INI.GetKeyValue("SENDER", "PASSWORD") <> Empty Then _
            txtPassWord.Text = Decrypt(INI.GetKeyValue("SENDER", "PASSWORD"), True) _
            Else txtPassWord.Text = Empty
            
    If INI.GetKeyValue("SENDER", "USERNAME") <> Empty Then _
            txtUserName.Text = Decrypt(INI.GetKeyValue("SENDER", "USERNAME"), True) _
            Else txtUserName.Text = Empty
    ' ........................................ END

    If INI.GetKeyValue("SENDER", "SENDERNAME") <> Empty Then _
            txtFromName.Text = INI.GetKeyValue("SENDER", "SENDERNAME") _
            Else txtFromName.Text = Empty
            
    If INI.GetKeyValue("SENDER", "SENDEREMAIL") <> Empty Then _
            txtFrom.Text = INI.GetKeyValue("SENDER", "SENDEREMAIL") _
            Else txtFrom.Text = Empty
            
    If INI.GetKeyValue("SENDER", "RECIPIENTNAME") <> Empty Then _
            txtToName.Text = INI.GetKeyValue("SENDER", "RECIPIENTNAME") _
            Else txtToName.Text = Empty
    
    If INI.GetKeyValue("SENDER", "RECIPIENTEMAIL") <> Empty Then _
            txtTo.Text = INI.GetKeyValue("SENDER", "RECIPIENTEMAIL") _
            Else txtTo.Text = Empty
            
    If INI.GetKeyValue("SENDER", "SUBJECT") <> Empty Then _
            txtSubject.Text = INI.GetKeyValue("SENDER", "SUBJECT") _
            Else txtSubject.Text = Empty
            
    If INI.GetKeyValue("SENDER", "BODY") <> Empty Then _
            txtMsg.Text = INI.GetKeyValue("SENDER", "BODY") _
            Else txtMsg.Text = Empty
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set poSendMail = Nothing: Set frmSender = Nothing
End Sub

Private Sub cmdBrowse_Click()
    Dim i As Integer
    On Local Error GoTo ErrorHandler
    
    cmDialog.ShowOpen
    
    If listAttach.ListCount > 0 Then
        For i = 0 To listAttach.ListCount
            If listAttach.List(i) = cmDialog.FileName Then
                    MsgBox "The selected file already exist in the sender list!", vbExclamation, App.Title
                Exit Sub
            End If
        Next i
    End If
    
    If txtAttach.Text = Empty Then
        txtAttach.Text = cmDialog.FileName: cmdRemuve.Enabled = True
        listAttach.AddItem cmDialog.FileName
        labellblAttach.Caption = "(" & listAttach.ListCount & ")"
    Else
        txtAttach.Text = txtAttach.Text & ";" & cmDialog.FileName: cmdRemuve.Enabled = True
        listAttach.AddItem cmDialog.FileName
        labellblAttach.Caption = "(" & listAttach.ListCount & ")"
    End If
Exit Sub
ErrorHandler:
    If Err.Number = 32755 Then
        MsgBox "Abort by User!", vbInformation, App.Title
    Else
        MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation, App.Title
    End If
    Err.Clear
End Sub

Private Sub cmdExit_Click()
    sCancel2 = False: Unload Me
End Sub

Private Sub cmdReset_Click()
    ClearTextBoxesOnForm
    lstStatus.Clear
    lblProgress = Empty: lblTime = Empty
End Sub

Public Sub ClearTextBoxesOnForm()
    On Local Error Resume Next
    Dim Ctl As Control
    For Each Ctl In Me.Controls
        If TypeOf Ctl Is TextBox Then
            Ctl.Text = ""
        End If
    Next
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case 48 To 57                               ' numeric
        Case 8                                      ' backspace
        Case Else: KeyAscii = 0
    End Select

End Sub


Private Sub SaveSettingToINI()
    On Local Error GoTo ErrorHandler
    
    INI.DeleteKey "SENDER", "SMTP"
    INI.CreateKeyValue "SENDER", "SMTP", Encrypt(txtServer.Text, True)
    
    INI.DeleteKey "SENDER", "MESSAGETOSEND"
    INI.CreateKeyValue "SENDER", "MESSAGETOSEND", txtQty.Text
    
    INI.DeleteKey "SENDER", "PASSWORD"
    INI.CreateKeyValue "SENDER", "PASSWORD", Encrypt(txtPassWord.Text, True)
    
    INI.DeleteKey "SENDER", "USERNAME"
    INI.CreateKeyValue "SENDER", "USERNAME", Encrypt(txtUserName.Text, True)
    
    INI.DeleteKey "SENDER", "OPTIONSENDER"
    INI.CreateKeyValue "SENDER", "OPTIONSENDER", sIndex
    
    INI.DeleteKey "SENDER", "AUTENTICATION"
    INI.CreateKeyValue "SENDER", "AUTENTICATION", sAutentication
    
    INI.DeleteKey "SENDER", "SENDERNAME"
    INI.CreateKeyValue "SENDER", "SENDERNAME", txtFromName.Text
    
    INI.DeleteKey "SENDER", "SENDEREMAIL"
    INI.CreateKeyValue "SENDER", "SENDEREMAIL", txtFrom.Text
    
    INI.DeleteKey "SENDER", "RECIPIENTNAME"
    INI.CreateKeyValue "SENDER", "RECIPIENTNAME", txtToName.Text
    
    INI.DeleteKey "SENDER", "RECIPIENTEMAIL"
    INI.CreateKeyValue "SENDER", "RECIPIENTEMAIL", txtTo.Text
    
    INI.DeleteKey "SENDER", "SUBJECT"
    INI.CreateKeyValue "SENDER", "SUBJECT", txtSubject.Text
    
    INI.DeleteKey "SENDER", "BODY"
    INI.CreateKeyValue "SENDER", "BODY", txtMsg.Text
    
    INI.DeleteKey "SENDER", "INTERNALAUTENTICATIONSERVER"
    INI.CreateKeyValue "SENDER", "INTERNALAUTENTICATIONSERVER", CheckInterlAutentication.Value
    
    INI.DeleteKey "SENDER", "SENDASHTML"
    INI.CreateKeyValue "SENDER", "SENDASHTML", CheckHTML.Value

Exit Sub
ErrorHandler:
    WriteErrorLogs Err.Number, Err.Description, "FormSender {Sub: SaveSettingToINI}", True, True
    Err.Clear
End Sub
