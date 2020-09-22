VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Glycerine Server"
   ClientHeight    =   4965
   ClientLeft      =   495
   ClientTop       =   795
   ClientWidth     =   6510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmMain 
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   6015
      Begin VB.CommandButton Command6 
         Caption         =   "kick"
         Height          =   315
         Left            =   5040
         TabIndex        =   41
         Top             =   3000
         Width           =   855
      End
      Begin VB.ListBox lstUsers 
         Height          =   2400
         ItemData        =   "frmMain.frx":08CA
         Left            =   4200
         List            =   "frmMain.frx":08CC
         Sorted          =   -1  'True
         TabIndex        =   37
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "exit"
         Height          =   315
         Left            =   4200
         TabIndex        =   36
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txtText 
         BackColor       =   &H8000000F&
         Height          =   3135
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "online users:"
         Height          =   195
         Left            =   4200
         TabIndex        =   38
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame frmAdmin 
      Caption         =   "admin message"
      Height          =   3495
      Left            =   240
      TabIndex        =   28
      Top             =   480
      Width           =   6015
      Begin VB.CommandButton Command2 
         Caption         =   "send"
         Height          =   315
         Left            =   4440
         TabIndex        =   30
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox txtAdmin 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2685
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Text            =   "frmMain.frx":08CE
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.Frame frmOptions 
      Caption         =   "options"
      Height          =   3495
      Left            =   240
      TabIndex        =   13
      Top             =   480
      Width           =   6015
      Begin VB.CommandButton Command7 
         Caption         =   "reboot"
         Height          =   315
         Left            =   4560
         TabIndex        =   22
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optDebug 
         Caption         =   "Show Debug"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optMOTD 
         Caption         =   "Show MOTD"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optNone 
         Caption         =   "None/Clear"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "hide"
         Height          =   315
         Left            =   4560
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "reload motd"
         Height          =   315
         Left            =   4560
         TabIndex        =   17
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "On"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   1335
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   1155
         Left            =   120
         Picture         =   "frmMain.frx":08DC
         ScaleHeight     =   1095
         ScaleWidth      =   3315
         TabIndex        =   14
         Top             =   2280
         Width           =   3375
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Did you know that SPGC Server will save the value of the view and log error selections? It does."
            Height          =   735
            Left            =   480
            TabIndex        =   15
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   -120
         X2              =   3600
         Y1              =   2175
         Y2              =   2175
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000C&
         X1              =   -120
         X2              =   3600
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   4440
         Y1              =   1455
         Y2              =   1455
      End
      Begin VB.Line Line15 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   4440
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line7 
         BorderColor     =   &H8000000C&
         X1              =   4440
         X2              =   4440
         Y1              =   120
         Y2              =   4080
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         X1              =   4455
         X2              =   4455
         Y1              =   120
         Y2              =   4080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "view:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "log errors:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   750
      End
   End
   Begin VB.Frame frmErrors 
      Caption         =   "errors"
      Height          =   3495
      Left            =   240
      TabIndex        =   25
      Top             =   480
      Width           =   6015
      Begin VB.TextBox txtErrors 
         BackColor       =   &H8000000F&
         Height          =   2775
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   240
         Width           =   5775
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Clear"
         Height          =   315
         Left            =   4560
         TabIndex        =   26
         Top             =   3090
         Width           =   1335
      End
   End
   Begin VB.Frame frmData 
      Caption         =   "data sender"
      Height          =   3495
      Left            =   240
      TabIndex        =   31
      Top             =   480
      Width           =   6015
      Begin VB.TextBox txtProtocalls 
         BackColor       =   &H80000000&
         Height          =   1695
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Text            =   "frmMain.frx":0BE6
         Top             =   1680
         Width           =   5655
      End
      Begin VB.TextBox txtCommand 
         Height          =   1005
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   240
         Width           =   5655
      End
      Begin VB.CommandButton Command5 
         Caption         =   "send"
         Height          =   255
         Left            =   4920
         TabIndex        =   32
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.Frame frmAcc 
      Caption         =   "account manager"
      Height          =   3495
      Left            =   240
      TabIndex        =   7
      Tag             =   "1520"
      Top             =   480
      Width           =   6015
      Begin VB.CommandButton Command8 
         Caption         =   "-"
         Height          =   255
         Left            =   5640
         TabIndex        =   12
         Tag             =   "1520"
         Top             =   1520
         Width           =   255
      End
      Begin VB.CommandButton Command10 
         Caption         =   "ban account"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1520
         Width           =   1455
      End
      Begin VB.ListBox lstUn 
         Height          =   1230
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   5775
      End
      Begin VB.ListBox lstAcc 
         Height          =   1230
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   5775
      End
      Begin VB.CommandButton Command9 
         Caption         =   "unban account"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   1520
         Width           =   1455
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   6000
         Y1              =   1935
         Y2              =   1935
      End
      Begin VB.Line Line13 
         BorderColor     =   &H8000000C&
         X1              =   6000
         X2              =   0
         Y1              =   1920
         Y2              =   1920
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3975
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7011
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   6
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Main"
            Key             =   "Main"
            Object.Tag             =   "Main"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Accounts"
            Key             =   "Accounts"
            Object.Tag             =   "Accounts"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Data sender"
            Key             =   "Data sender"
            Object.Tag             =   "Data sender"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Admin message"
            Key             =   "Admin message"
            Object.Tag             =   "Admin message"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Options"
            Key             =   "Options"
            Object.Tag             =   "Options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Errors"
            Key             =   "Errors"
            Object.Tag             =   "Errors"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   6210
      TabIndex        =   3
      Top             =   4200
      Width           =   6210
      Begin VB.Line Line12 
         BorderColor     =   &H00FFFFFF&
         X1              =   3135
         X2              =   3135
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Line Line11 
         BorderColor     =   &H8000000C&
         X1              =   3120
         X2              =   3120
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Label lblAccPort 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3240
         TabIndex        =   40
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblChatPort 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   2895
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   6240
         X2              =   0
         Y1              =   330
         Y2              =   330
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   6240
         Y1              =   320
         Y2              =   320
      End
      Begin VB.Line Line10 
         BorderColor     =   &H8000000C&
         X1              =   6195
         X2              =   6195
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Line Line9 
         BorderColor     =   &H8000000C&
         X1              =   -240
         X2              =   7200
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   600
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   7200
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "win.console.2.beta"
         Height          =   195
         Left            =   3240
         TabIndex        =   5
         Top             =   75
         Width           =   2820
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "www.ry4.net"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         MouseIcon       =   "frmMain.frx":0D24
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   75
         Width           =   2880
      End
   End
   Begin MSWinsockLib.Winsock winsckACC 
      Index           =   0
      Left            =   8760
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtMOTD 
      Height          =   375
      Left            =   8760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSWinsockLib.Winsock winsck 
      Index           =   0
      Left            =   8760
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lbltoggle 
      Caption         =   "1"
      Height          =   255
      Left            =   8760
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu restore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu quit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vbTray As NOTIFYICONDATA
Private strUsers() As String
Private Sub TrayIt()
    vbTray.cbSize = Len(vbTray)
    vbTray.hwnd = Me.hwnd
    vbTray.uId = vbNull
    vbTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    vbTray.ucallbackMessage = WM_MOUSEMOVE
    vbTray.hIcon = Me.Icon
    vbTray.szTip = Me.Caption & vbNullChar
    Call Shell_NotifyIcon(NIM_ADD, vbTray)
    App.TaskVisible = False
    Me.Hide
End Sub

Private Sub UnTrayIt()
    vbTray.cbSize = Len(vbTray)
    vbTray.hwnd = Me.hwnd
    vbTray.uId = vbNull
    Call Shell_NotifyIcon(NIM_DELETE, vbTray)
End Sub

Private Sub Check1_Click()
Call SaveSetting("Glycerine Server", "Options", "Log Errors", Check1.Value)
End Sub

Private Sub Command10_Click()
On Error Resume Next
Dim strAcc As String
strAcc = lstAcc.Text
lstAcc.RemoveItem (lstAcc.ListIndex)
lstUn.Clear
lstAcc.AddItem (strAcc & "_äœáèÕ")
For n = 0 To lstAcc.ListCount - 1
strUser = Left$(lstAcc.List(n), InStr(lstAcc.List(n), ":") - 1)
lstUn.AddItem (strUser)
Next n
Call SaveListBox(App.Path & "\acc.txt", lstAcc)
End Sub

Private Sub Command11_Click()
txtErrors = ""
End Sub

Private Sub Command4_Click()
If MsgBox("Are you sure you want to kill the server? All connections will be killed!", vbYesNo) = vbYes Then Call UnTrayIt: End
End Sub

Private Sub Command5_Click()
strData = txtCommand
If Mid(strData, 1, 5) = "exit-" Then
strExit$ = Mid(strData, 6)
Call SendToAll("exit-" & strExit$)
For n = 0 To lstUsers.ListCount - 1 Step 1
If lstUsers.List(n) = strExit$ Then
lstUsers.RemoveItem (n)
Exit For
End If
Next n
If Check1.Value = "1" Then txtErrors = txtErrors & vbNewLine & "Error with unload socket. unable to remove: " & strExit$
End If
Call SendToAll(txtCommand)
txtCommand.Text = ""
End Sub

Private Sub Command6_Click()
    strExit$ = lstUsers.Text
    Call SendToAll("exit-" & strExit$)
     For n = 0 To lstUsers.ListCount - 1 Step 1
        If lstUsers.List(n) = strExit$ Then
            lstUsers.RemoveItem (n)
            Exit For
        End If
    Next n
    If Check1.Value = "1" Then txtErrors = txtErrors & vbNewLine & "Error with unload socket. unable to remove: " & strExit$
End Sub

Private Sub Command7_Click()
On Error Resume Next
lbltoggle = "0"
Unload Me
frmReboot.Show
End Sub

Private Sub Command8_Click()
On Error Resume Next
lstAcc.RemoveItem (lstAcc.ListIndex)
lstUn.Clear
For n = 0 To lstAcc.ListCount - 1
strUser = Left$(lstAcc.List(n), InStr(lstAcc.List(n), ":") - 1)
lstUn.AddItem (strUser)
Next n
Call SaveListBox(App.Path & "\acc.txt", lstAcc)
End Sub

Private Sub Command9_Click()
On Error Resume Next
Dim strAcc As String
strAcc = Left$(lstAcc.Text, InStr(lstAcc.Text, "_") - 1)
lstUn.Clear
lstAcc.RemoveItem (lstAcc.ListIndex)
lstAcc.AddItem (strAcc)
For n = 0 To lstAcc.ListCount - 1
strUser = Left$(lstAcc.List(n), InStr(lstAcc.List(n), ":") - 1)
lstUn.AddItem (strUser)
Next n
Call SaveListBox(App.Path & "\acc.txt", lstAcc)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static lngMsg As Long
    Dim blnFlag As Boolean, lngResult As Long
    lngMsg = X / Screen.TwipsPerPixelX
    If blnFlag = False Then
        blnFlag = True
        Select Case lngMsg
            Case WM_LBUTTONDBLCLICK
                Me.WindowState = 0
                Me.Show
            Case WM_RBUTTONUP
                lngResult = SetForegroundWindow(Me.hwnd)
                Me.PopupMenu mnuFile
        End Select
        blnFlag = False
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call UnTrayIt
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then
        Call TrayIt
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lbltoggle = "1" Then
Call TrayIt
Cancel = True
End If
End Sub

Private Sub mnuSysTray_Click()
    Call UnTrayIt
    Me.WindowState = 0
    Me.Show
End Sub

Private Sub SendToAll(ToSend As String)
For lngIndex& = 1 To winsck().Count - 1
If winsck(lngIndex&).State = sckConnected And lngIndex& <> Index Then
Call winsck(lngIndex&).SendData(ToSend)
DoEvents
End If
Next lngIndex&
End Sub

Private Sub Pause(Duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
    DoEvents
    Loop
End Sub

Private Sub LoadText(txtLoad, PathA As String)
    Dim TextString As String
    On Error Resume Next
    Open PathA$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.Text = TextString$
End Sub

Function Search_ListBox(trig$, lst As ListBox) As Long
    Dim items As Long
    Dim n As Long
    items = lst.ListCount - 1
    For n = 0 To items Step 1
        If lst.List(n) = trig$ Then
            Search_ListBox = n
            Exit Function
        End If
    Next n
    Search_ListBox = -1
End Function

Private Sub Command1_Click()
Call TrayIt
End Sub

Private Sub Command2_Click()
SendToAll "adm-" & txtAdmin
txtAdmin = ""
End Sub

Private Sub Command3_Click()
Call LoadText(txtMOTD, App.Path & "\motd.txt")
End Sub

Private Sub Form_Load()
On Error Resume Next
Check1.Value = GetSetting("Glycerine Server", "Options", "Log Errors")
optDebug.Value = GetSetting("Glycerine Server", "Options", "Show Debug")
optMOTD.Value = GetSetting("Glycerine Server", "Options", "Show MOTD")
optNone.Value = GetSetting("Glycerine Server", "Options", "None/Clear")
Call LoadText(txtMOTD, App.Path & "\motd.txt")
Call Loadlistbox(App.Path & "\acc.txt", lstAcc)
lstUn.Clear
For n = 0 To lstAcc.ListCount - 1
strUser = Left$(lstAcc.List(n), InStr(lstAcc.List(n), ":") - 1)
lstUn.AddItem (strUser)
Next n
lblAccPort = "Account port: 5623"
lblChatPort = "Chat port: 5622"
winsckACC(0).LocalPort = "5623"
winsckACC(0).Listen
winsck(0).LocalPort = "5622"
winsck(0).Listen
txtText = "Glycerine: Waiting for connections"
End Sub


Private Sub Label5_Click()
Call Shell("start http://www.ry4.net")
End Sub

Private Sub optDebug_Click()
If optDebug.Value = True Then txtText = ""
Call SaveSetting("Glycerine Server", "Options", "Show Debug", optDebug.Value)
Call SaveSetting("Glycerine Server", "Options", "Show MOTD", optMOTD.Value)
Call SaveSetting("Glycerine Server", "Options", "None/Clear", optNone.Value)
End Sub

Private Sub optMOTD_Click()
If optMOTD.Value = True Then txtText = txtMOTD
Call SaveSetting("Glycerine Server", "Options", "Show Debug", optDebug.Value)
Call SaveSetting("Glycerine Server", "Options", "Show MOTD", optMOTD.Value)
Call SaveSetting("Glycerine Server", "Options", "None/Clear", optNone.Value)
End Sub

Private Sub optNone_Click()
If optNone.Value = True Then txtText = ""
Call SaveSetting("Glycerine Server", "Options", "Show Debug", optDebug.Value)
Call SaveSetting("Glycerine Server", "Options", "Show MOTD", optMOTD.Value)
Call SaveSetting("Glycerine Server", "Options", "None/Clear", optNone.Value)
End Sub

Private Sub quit_Click()
If MsgBox("Are you sure you want to kill the server? All connections will be killed!", vbYesNo) = vbYes Then Call UnTrayIt: End
End Sub

Private Sub restore_Click()
Me.WindowState = 0
Me.Show
End Sub

Private Sub TabStrip1_Click()
Select Case TabStrip1.SelectedItem

Case "Main"
Me.frmMain.Visible = True
Me.frmAcc.Visible = False
Me.frmData.Visible = False
Me.frmAdmin.Visible = False
Me.frmOptions.Visible = False
Me.frmErrors.Visible = False

Case "Accounts"
Me.frmMain.Visible = False
Me.frmAcc.Visible = True
Me.frmData.Visible = False
Me.frmAdmin.Visible = False
Me.frmOptions.Visible = False
Me.frmErrors.Visible = False

Case "Data sender"
Me.frmMain.Visible = False
Me.frmAcc.Visible = False
Me.frmData.Visible = True
Me.frmAdmin.Visible = False
Me.frmOptions.Visible = False
Me.frmErrors.Visible = False

Case "Admin message"
Me.frmMain.Visible = False
Me.frmAcc.Visible = False
Me.frmData.Visible = False
Me.frmAdmin.Visible = True
Me.frmOptions.Visible = False
Me.frmErrors.Visible = False

Case "Options"
Me.frmMain.Visible = False
Me.frmAcc.Visible = False
Me.frmData.Visible = False
Me.frmAdmin.Visible = False
Me.frmOptions.Visible = True
Me.frmErrors.Visible = False

Case "Errors"
Me.frmMain.Visible = False
Me.frmAcc.Visible = False
Me.frmData.Visible = False
Me.frmAdmin.Visible = False
Me.frmOptions.Visible = False
Me.frmErrors.Visible = True

End Select
End Sub

Private Sub winsck_Connect(Index As Integer)
On Error GoTo Err:
If optDebug.Value = True Then
txtText = txtText & vbNewLine & "Connected! - " & winsck(Index).RemoteHostIP & ":" & winsck(Index).RemotePort
txtText.SelLength = Len(txtText)
End If
Exit Sub
Err:
If Check1.Value = "1" Then txtErrors = txtErrors & vbNewLine & "Error: " & Err.Description & "  -  " & Err.Number
End Sub

Private Sub winsck_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error GoTo Err:
   Dim lngIndex As Long, blnFlag As Boolean
      For lngIndex& = 1 To winsck().UBound
         If winsck(lngIndex&).State = sckClosed Then
             blnFlag = True
             Exit For
         End If
      Next lngIndex&
      If blnFlag = False Then
         lngIndex& = winsck().UBound + 1
         Load winsck(lngIndex&)
      End If
   Call winsck(lngIndex&).Accept(requestID&)
   Call winsck_Connect(Index)

Exit Sub
Err:
If Check1.Value = "1" Then txtErrors = txtErrors & vbNewLine & "Error: " & Err.Description & "  -  " & Err.Number
End Sub

Private Sub winsck_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Call winsck(Index).GetData(Data$, vbString)
Call StepThrough(Data$, Index)
End Sub

Private Sub winsck_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If Check1.Value = "1" Then txtErrors = txtErrors & vbNewLine & "Error: " & Err.Description & "  -  " & Err.Number
End Sub

Private Sub winsckACC_Connect(Index As Integer)
On Error GoTo Err:
If optDebug.Value = True Then
txtText = txtText & vbNewLine & "Connected! - " & winsck(Index).RemoteHostIP & ":" & winsck(Index).RemotePort
txtText.SelLength = Len(txtText)
End If
Exit Sub
Err:
If Check1.Value = "1" Then txtErrors = txtErrors & vbNewLine & "Error: " & Err.Description & "  -  " & Err.Number
End Sub

Private Sub winsckACC_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error GoTo Err:
   Dim lngIndex As Long, blnFlag As Boolean
      For lngIndex& = 1 To winsckACC().UBound
         If winsckACC(lngIndex&).State = sckClosed Then
      On Error GoTo Err:
       blnFlag = True
             Exit For
         End If
      Next lngIndex&
      If blnFlag = False Then
         lngIndex& = winsckACC().UBound + 1
         Load winsckACC(lngIndex&)
      End If
   Call winsckACC(lngIndex&).Accept(requestID&)
   Call winsckACC_Connect(Index)
Exit Sub
Err:
If Check1.Value = "1" Then txtErrors = txtErrors & vbNewLine & "Error: " & Err.Description & "  -  " & Err.Number
End Sub

Private Sub winsckACC_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Data As String
Call winsckACC(Index).GetData(Data$, vbString)
Call StepThrough2(Data$, Index)
End Sub
Sub StepThrough(Data As String, Index As Integer)
If optDebug.Value = True Then
txtText = txtText & vbNewLine & "debug: " & Data$
txtText.SelLength = Len(txtText)
End If

    If Mid(Data$, 1, 4) = "req-" Then
    strName$ = Mid(Data, 5)
    strName$ = DecryptA(strName$)
    If optDebug.Value = True Then txtText = txtText & vbNewLine & "debug: DECRYPTED DATA " & strName$
    strJoin$ = Left$(strName$, InStr(strName$, ":") - 1)
    If Not Search_ListBox(strName$, lstAcc) = "-1" Then
    If Search_ListBox(strJoin$, lstUsers) = "-1" Then
    Call winsck(Index).SendData("ok-" & strJoin$)
    Call SendToAll("join-" & strJoin$)
    lstUsers.AddItem (strJoin$)
    DoEvents
    Call winsck(Index).SendData("motd-" & txtMOTD)
    Pause 1
    For n = 0 To lstUsers.ListCount - 1
    strN = strN & "-" & lstUsers.List(n)
    Next n
    winsck(Index).SendData ("users-" & strN & "-")
    Pause 0.7
    Else
    Call winsck(Index).SendData("inuse-" & strJoin$)
    Exit Sub
    End If
    Else
    Call winsck(Index).SendData("inuse-" & strJoin$)
    Exit Sub
    End If
    End If
    
    If Mid(Data$, 1, 5) = "exit-" Then
    strExit$ = Mid(Data, 6)
    Call SendToAll("exit-" & strExit$)
     For n = 0 To lstUsers.ListCount - 1 Step 1
        If lstUsers.List(n) = strExit$ Then
            lstUsers.RemoveItem (n)
            Exit For
        End If
    Next n
    If Check1.Value = "1" Then txtErrors = txtErrors & vbNewLine & "Error with unload socket. unable to remove: " & strExit$
    End If
    
    If Mid(Data$, 1, 3) = "pm-" Then
    strString$ = Mid(Data, 4)
    Call SendToAll("pm-" & strString$)
    End If
    
    If Mid(Data$, 1, 3) = "me-" Then
    strString$ = Mid(Data, 4)
    Call SendToAll("me-" & strString$)
    End If
    
    If Mid(Data$, 1, 4) = "msg-" Then
    strString$ = Mid(Data, 5)
    Call SendToAll("msg-" & strString$)
    End If
    
    
Exit Sub
Err:
If Check1.Value = "1" Then txtErrors = txtErrors & vbNewLine & "Error: " & Err.Description & "  -  " & Err.Number
End Sub
Sub StepThrough2(Data As String, Index As Integer)
On Error Resume Next
If Mid(Data$, 1, 4) = "jrm-" Then
Data$ = Mid(Data$, 5)
strName$ = Left$(Data$, InStr(Data$, ":") - 1)
strPW$ = Right(Data$, Len(Data$) - InStr(Data$, ":"))
strName$ = ReplaceString(strName$, "-", "_")
strName$ = ReplaceString(strName$, ":", "_")
strName$ = ReplaceString(strName$, ";", "_")
strName$ = ReplaceString(strName$, " ", "_")
If Search_ListBox(strName$, lstUn) = "-1" Then
Call winsckACC(Index).SendData("accok-" & strName$)
lstAcc.AddItem (strName$ & ":" & strPW$)
lstUn.Clear
For n = 0 To lstAcc.ListCount - 1
strUser = Left$(lstAcc.List(n), InStr(lstAcc.List(n), ":") - 1)
lstUn.AddItem (strUser)
Next n
Call SaveListBox(App.Path & "\acc.txt", lstAcc)
Else
Call winsckACC(Index).SendData("accno-" & strName$)
Exit Sub
End If
End If
End Sub

