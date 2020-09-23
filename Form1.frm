VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{0F5979DA-E576-11D3-8086-0010A4FA0BE6}#35.0#0"; "MobileFBUS.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{17ED04B9-6C71-11D4-87A3-DAA6B6B40E8F}#6.0#0"; "LongTimer.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New mail Inbox SMS notifier"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9585
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin LngTmr.LongTimer LongTimer1 
      Left            =   6720
      Top             =   1380
      _ExtentX        =   873
      _ExtentY        =   873
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Default         =   -1  'True
      Height          =   465
      Left            =   7920
      TabIndex        =   2
      Top             =   2850
      Width           =   1605
   End
   Begin MobileFBUS.MobileFBUSControl Mobile 
      Height          =   795
      Left            =   8580
      TabIndex        =   1
      Top             =   1290
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2685
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   4736
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "From:"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Subject:"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   7290
      Top             =   1380
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   7920
      Top             =   1380
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Label pword 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5490
      TabIndex        =   11
      Top             =   3570
      Width           =   1395
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Password:"
      Height          =   195
      Left            =   5520
      TabIndex        =   10
      Top             =   3330
      Width           =   1365
   End
   Begin VB.Label aname 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3930
      TabIndex        =   9
      Top             =   3570
      Width           =   1395
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Account name:"
      Height          =   195
      Left            =   3930
      TabIndex        =   8
      Top             =   3330
      Width           =   1395
   End
   Begin VB.Label cport 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2730
      TabIndex        =   7
      Top             =   3570
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "COM Port:"
      Height          =   195
      Left            =   2760
      TabIndex        =   6
      Top             =   3330
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Mobile number:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3330
      Width           =   2505
   End
   Begin VB.Label mnumber 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   3570
      Width           =   2535
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   90
      TabIndex        =   3
      Top             =   2850
      Width           =   7695
   End
   Begin VB.Menu mnuMailCheck 
      Caption         =   "Mail Checker"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit Mail checker"
      End
      Begin VB.Menu Spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore Mail checker"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Private mobilenumber
Private z As Integer

Private Sub Command1_Click()
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub

Private Sub Form_Resize()
If WindowState = vbMinimized Then
Me.Hide
Me.Refresh
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
.szTip = Me.Caption & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
Else
Shell_NotifyIcon NIM_DELETE, nid
End If
End Sub
Private Sub Form_Load()

Open "MailChek.ini" For Input As #1

Do While Not EOF(1)

Line Input #1, a$
    Dim i As Long
    i = InStr(a$, "=")
        If i Then
        Field = UCase(Left$(a$, i - 1))
    Else
        Field = UCase(Trim$(a$))
    End If

Select Case Field

    Case "MOBILEPORT": mobileport = Mid$(a$, i + 1)

    Case "CHECKINTERVAL"
        x = Mid(a, i + 1)
        LongTimer1.Interval = (x * 60000)

    Case "MOBILENUMBER": mobilenumber = Mid$(a$, i + 1)
    
    Case "ACCOUNTNAME": MAPISession1.UserName = Mid$(a$, i + 1)
    
    Case "ACCOUNTPASSWORD": MAPISession1.Password = Mid$(a$, i + 1)

End Select

Loop

Close #1

mnumber.Caption = mobilenumber
cport.Caption = mobileport
aname.Caption = MAPISession1.UserName
pword.Caption = MAPISession1.Password

'Don't download any new messages - just list the ones already in the inbox.
'That way we can tell if any new ones arrive! Note that with Outlook Express,
'you have to "sign on" to download any new mail. The Fetch lists all messages
'in the inbox. It does not "fetch" them from your ISP's mail server.
'Essentially, all this part does is to get the initial inbox list.

MAPISession1.DownLoadMail = False
MAPISession1.SignOn
MAPIMessages1.SessionID = MAPISession1.SessionID

MAPIMessages1.Fetch

z = MAPIMessages1.MsgCount

With ListView1
      .ListItems.Clear
      .ColumnHeaders.Clear
      .ColumnHeaders.Add , , "From:"
      .ColumnHeaders.Add , , "Subject:"
      .ColumnHeaders.Add , , "Date:"
       .View = lvwReport
      .Sorted = False
End With

For i = 1 To z

MAPIMessages1.MsgIndex = i - 1

Set itmx = ListView1.ListItems.Add(, , MAPIMessages1.MsgOrigDisplayName)
      
itmx.SubItems(1) = MAPIMessages1.MsgSubject

itmx.SubItems(2) = MAPIMessages1.MsgDateReceived

Next i

MAPISession1.SignOff

Call lvAutosizeControl(ListView1)

Status.Caption = "Waiting " & (LongTimer1.Interval / 60000) & " minutes to check..."

Mobile.Connect mobileport

End Sub

Private Sub Form_Terminate()
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub

Private Sub LongTimer1_Timer()

'This is the part that gets any new email. It compares the inbox message count
'with the new count about to be made. It does this by signing on to MAPI,
'downloading any new messages, doing another count, and checking the difference.
'If there is a difference, it will send an SMS message for each new email received.
'Long timer is used because the standard Windows timer only does 65 seconds.
'We need something a bit longer. Don't want to hammer the ISP mail server!

Status.Caption = "Logging on to mail..."

MAPISession1.DownLoadMail = True
MAPISession1.SignOn
MAPIMessages1.SessionID = MAPISession1.SessionID

MAPIMessages1.Fetch

If MAPIMessages1.MsgCount > z Then

Status.Caption = "New mail received! Sending SMS now!"

For i = z To (MAPIMessages1.MsgCount - 1)

MAPIMessages1.MsgIndex = i

Set itmx = ListView1.ListItems.Add(, , MAPIMessages1.MsgOrigDisplayName)
      
itmx.SubItems(1) = MAPIMessages1.MsgSubject

itmx.SubItems(2) = MAPIMessages1.MsgDateReceived

'Send the new mail notification

Mobile.SendSMSMessage mobilenumber, "A new mail from " & Chr(34) & MAPIMessages1.MsgOrigDisplayName & Chr(34) & " with the subject " & Chr(34) & MAPIMessages1.MsgSubject & Chr(34) & " has been received."

Next i

End If

z = MAPIMessages1.MsgCount

MAPISession1.SignOff

Call lvAutosizeControl(ListView1)

Status.Caption = "Waiting " & (LongTimer1.Interval / 60000) & " minutes to check again..."

End Sub


Private Sub lvAutosizeControl(lv As ListView)

   Dim col2adjust As Long

   For col2adjust = 0 To lv.ColumnHeaders.Count - 1
   
      Call SendMessage(lv.hwnd, _
                       LVM_SETCOLUMNWIDTH, _
                       col2adjust, _
                       ByVal LVSCW_AUTOSIZE_USEHEADER)

   Next

End Sub
Private Sub mnuRestore_Click()
WindowState = vbNormal
Me.Show
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Sys As Long
Sys = x / Screen.TwipsPerPixelX
Select Case Sys
Case WM_LBUTTONDOWN:
Me.PopupMenu mnuMailCheck
End Select
End Sub
