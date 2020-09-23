VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Connection1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "Game - Form4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin BattleStar.TSPanel TSPanel1 
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   7858
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Frame PlayerName1 
         Caption         =   "Player Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   3255
         Begin VB.TextBox PlayerNText1 
            Height          =   285
            Left            =   240
            MaxLength       =   18
            TabIndex        =   18
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.Frame StatusMenu1 
         Caption         =   "Connection Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   3255
         Begin VB.CheckBox SMCheck1 
            Caption         =   "Host This Game"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1575
         End
         Begin VB.Frame Status1 
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   15
            Top             =   1800
            Width           =   3015
            Begin VB.Label StatusText1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Not Connected."
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   240
               TabIndex        =   16
               Top             =   360
               Width           =   1155
            End
         End
         Begin VB.PictureBox ClientHost1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            ForeColor       =   &H80000008&
            Height          =   975
            Index           =   2
            Left            =   120
            ScaleHeight     =   945
            ScaleWidth      =   2985
            TabIndex        =   10
            Top             =   720
            Visible         =   0   'False
            Width           =   3015
            Begin VB.TextBox HostText1 
               Height          =   285
               Index           =   0
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   12
               Text            =   "0.0.0.0"
               Top             =   120
               Width           =   1575
            End
            Begin VB.TextBox HostText1 
               Height          =   285
               Index           =   1
               Left            =   1320
               TabIndex        =   11
               Text            =   "1295"
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label HostLabel1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Local IP :"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   480
               TabIndex        =   14
               Top             =   120
               Width           =   675
            End
            Begin VB.Label HostLabel1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Local Port :"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   360
               TabIndex        =   13
               Top             =   480
               Width           =   810
            End
         End
         Begin VB.PictureBox ClientHost1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            ForeColor       =   &H80000008&
            Height          =   975
            Index           =   1
            Left            =   120
            ScaleHeight     =   945
            ScaleWidth      =   2985
            TabIndex        =   5
            Top             =   720
            Width           =   3015
            Begin VB.TextBox ClientText1 
               Height          =   285
               Index           =   0
               Left            =   1320
               MaxLength       =   15
               TabIndex        =   7
               Top             =   120
               Width           =   1575
            End
            Begin VB.TextBox ClientText1 
               Height          =   285
               Index           =   1
               Left            =   1320
               MaxLength       =   4
               TabIndex        =   6
               Text            =   "1295"
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label ClientLabel1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Remote IP :"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   360
               TabIndex        =   9
               Top             =   120
               Width           =   840
            End
            Begin VB.Label ClientLabel1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Remote Port :"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   1
               Left            =   240
               TabIndex        =   8
               Top             =   480
               Width           =   975
            End
         End
      End
      Begin VB.CommandButton StartStopIt1 
         Caption         =   "&Start"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   3960
         Width           =   1215
      End
   End
   Begin VB.Timer Watcher1 
      Left            =   1680
      Top             =   6360
   End
   Begin VB.ListBox Outbox1 
      Height          =   1425
      Left            =   2160
      TabIndex        =   1
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox Inbox1 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Con1 
      Left            =   1680
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Connection1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Con1_ConnectionRequest(ByVal requestID As Long)
    Call Connection1.Con1.Close
    Call Connection1.Con1.Accept(requestID)
''===============================================================================''
    Call Connection1.Outbox1.Clear
    Call Connection1.Outbox1.AddItem("@", 0)
    Call Connection1.Outbox1.AddItem("@", 0)
    Call Connection1.Outbox1.AddItem("@", 0)
End Sub

Private Sub Con1_DataArrival(ByVal bytesTotal As Long)
Dim Input1 As String
''===============================================================================''
    Let SocketSendData1 = True
    Call Connection1.Con1.GetData(Input1, , bytesTotal)
    If Not (Trim$(Input1) = "@") Then
        Call Connection1.Inbox1.AddItem(Input1)
    End If
''===============================================================================''
    If Not (Connection1.Outbox1.ListCount = 0) Then
        Call Connection1.Outbox1.RemoveItem(0)
    End If
End Sub

Private Sub Form_Load()
    Let Connection1.Caption = ("Battle Star v" & App.Major & "." & App.Minor & "." & App.Revision)
    Call Startup_Procedure1
    Call Load_Connection1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call Connection1.Con1.Close
    End
End Sub

Private Sub PlayerNText1_Change()
    If (Trim$(Connection1.PlayerNText1.Text) = "") Then
        Let Connection1.StartStopIt1.Enabled = False
        Let Connection1.PlayerNText1.Text = ""
    Else
        Let Connection1.StartStopIt1.Enabled = True
    End If
End Sub

Private Sub SMCheck1_Click()
    If (Connection1.SMCheck1.Value = 1) Then
        Let Connection1.ClientHost1(1).Visible = False
        Let Connection1.ClientHost1(2).Visible = True
    ElseIf (Connection1.SMCheck1.Value = 0) Then
        Let Connection1.ClientHost1(1).Visible = True
        Let Connection1.ClientHost1(2).Visible = False
    End If
End Sub

Private Sub StartStopIt1_Click()
    Let ConnectionType1 = (Connection1.SMCheck1.Value - -1)
    Call StartStopClient_Host1
End Sub

Private Sub Watcher1_Timer()
    Call Connection1Watcher1Loop1
End Sub

'Thraddash Software - Trevor Lewis

