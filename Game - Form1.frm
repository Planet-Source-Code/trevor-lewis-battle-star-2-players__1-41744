VERSION 5.00
Begin VB.Form Main1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Thraddash Software"
   ClientHeight    =   9150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10305
   Icon            =   "Game - Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Game - Form1.frx":08CA
   ScaleHeight     =   9150
   ScaleWidth      =   10305
   Begin VB.PictureBox Frame1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8475
      Left            =   0
      Picture         =   "Game - Form1.frx":A7FC
      ScaleHeight     =   565
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   565
      TabIndex        =   0
      Top             =   0
      Width           =   8475
      Begin VB.CommandButton MessageButton1 
         Caption         =   ".."
         Default         =   -1  'True
         Height          =   255
         Left            =   7680
         TabIndex        =   14
         ToolTipText     =   "Send Message"
         Top             =   6240
         Width           =   255
      End
      Begin VB.TextBox SendMessage1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   6000
         TabIndex        =   1
         Top             =   6240
         Width           =   1575
      End
      Begin VB.PictureBox ShipHit1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   720
         ScaleHeight     =   105
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   233
         TabIndex        =   11
         Top             =   120
         Width           =   3495
         Begin VB.Label Destroyed1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Destroyed"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label ShipsName1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Ships Name_"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   165
            Left            =   120
            TabIndex        =   12
            Top             =   1320
            Width           =   795
         End
      End
      Begin VB.PictureBox MsgBox1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   4200
         ScaleHeight     =   89
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   217
         TabIndex        =   8
         Top             =   6720
         Width           =   3255
         Begin VB.PictureBox MsgBox2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   120
            ScaleHeight     =   81
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   209
            TabIndex        =   9
            Top             =   120
            Width           =   3135
            Begin VB.Label MsgText1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Thraddash Software"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Index           =   0
               Left            =   480
               TabIndex        =   10
               Top             =   120
               Width           =   1440
            End
            Begin VB.Image MsgPic1 
               Appearance      =   0  'Flat
               Height          =   225
               Index           =   0
               Left            =   120
               Picture         =   "Game - Form1.frx":31B4A
               Top             =   120
               Width           =   225
            End
         End
      End
      Begin VB.VScrollBar VMsgScroll1 
         Enabled         =   0   'False
         Height          =   1335
         Left            =   7560
         Max             =   0
         TabIndex        =   7
         Top             =   6720
         Width           =   255
      End
      Begin VB.HScrollBar HMsgScroll1 
         Enabled         =   0   'False
         Height          =   255
         LargeChange     =   50
         Left            =   4200
         Max             =   0
         SmallChange     =   10
         TabIndex        =   6
         Top             =   8160
         Width           =   3255
      End
      Begin VB.PictureBox RGrid1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4065
         Left            =   4440
         MouseIcon       =   "Game - Form1.frx":320C0
         MousePointer    =   99  'Custom
         ScaleHeight     =   4065
         ScaleWidth      =   4065
         TabIndex        =   3
         Top             =   0
         Width           =   4065
         Begin VB.Image Target1 
            Enabled         =   0   'False
            Height          =   420
            Left            =   0
            Picture         =   "Game - Form1.frx":323CA
            Top             =   0
            Visible         =   0   'False
            Width           =   420
         End
      End
      Begin VB.PictureBox LGrid1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4065
         Left            =   0
         MouseIcon       =   "Game - Form1.frx":32B94
         ScaleHeight     =   4065
         ScaleWidth      =   4065
         TabIndex        =   2
         Top             =   4440
         Width           =   4065
         Begin VB.Line PosLineXY1 
            BorderColor     =   &H0000FF00&
            Index           =   1
            Visible         =   0   'False
            X1              =   240
            X2              =   240
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Line PosLineXY1 
            BorderColor     =   &H0000FF00&
            Index           =   0
            Visible         =   0   'False
            X1              =   120
            X2              =   120
            Y1              =   600
            Y2              =   2280
         End
         Begin VB.Image Edit1 
            Enabled         =   0   'False
            Height          =   495
            Left            =   600
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Image Explosion1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   495
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.Image RemoteVictory1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   7320
         Top             =   4560
         Width           =   135
      End
      Begin VB.Image RemoteVictory1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   7200
         Top             =   4560
         Width           =   135
      End
      Begin VB.Image LocalVictory1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1080
         Top             =   3720
         Width           =   135
      End
      Begin VB.Image LocalVictory1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   960
         Top             =   3720
         Width           =   135
      End
      Begin VB.Image RemoteAmount1 
         Appearance      =   0  'Flat
         Height          =   135
         Left            =   5400
         Top             =   4200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image LocalAmount1 
         Appearance      =   0  'Flat
         Height          =   135
         Left            =   2040
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label RemoteName1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Thraddash Software"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6840
         TabIndex        =   5
         Top             =   4200
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label LocalName1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Thraddash Software"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   4080
         Visible         =   0   'False
         Width           =   1440
      End
   End
End
Attribute VB_Name = "Main1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call Load_Main1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Let Connection1.Visible = True
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button And 1) Then
        Call AllowFormToMove1(Main1)
    ElseIf (Button And 2) Then
        Call PopupMenu(Data1.Menu1)
    End If
End Sub

Private Sub HMsgScroll1_Change()
    Let Main1.MsgBox2.Left = -(Main1.HMsgScroll1.Value)
End Sub

Private Sub HMsgScroll1_Scroll()
    Let Main1.MsgBox2.Left = -(Main1.HMsgScroll1.Value)
End Sub

Private Sub LGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (EditModeOnOff1) Then
        If (Button And 1) Then
            Call PlaceShipOnGrid1
        ElseIf (Button And 2) Then
            Call RotatePlayersShips1
        End If
    End If
End Sub

Private Sub LGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let LocalGridRefX = Int((GridBlocksX / 100) * Int((X / LGrid1.Width) * 100))
    Let LocalGridRefY = Int((GridBlocksY / 100) * Int((Y / LGrid1.Height) * 100))
''===============================================================================''
    If (LocalGridRefX < 0) Then Let LocalGridRefX = 0
    If (LocalGridRefY < 0) Then Let LocalGridRefY = 0
''===============================================================================''
    If (LocalGridRefX > (GridBlocksX - 1)) Then Let LocalGridRefX = (GridBlocksX - 1)
    If (LocalGridRefY > (GridBlocksY - 1)) Then Let LocalGridRefY = (GridBlocksY - 1)
''===============================================================================''
    If (EditModeOnOff1) Then Call MovePlayersShips1
End Sub

Private Sub MessageButton1_Click()
    Call SendDataToRemoteUser1("##RUMessage##:" & Main1.SendMessage1.Text)
    Call SendDataToLocalUser1("##LUMessage##:" & Main1.SendMessage1.Text)
''===============================================================================''
    Let Main1.SendMessage1.Text = ""
End Sub

Private Sub RGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim GridList1 As Long
''===============================================================================''
    If (RemoteGridEnabled1) Then
        GridList1 = ((GridBlocksY * RemoteGridRefY) - -(RemoteGridRefX))
        If (Data1.RemoteList1.List(GridList1) = 0) Then
            Let Data1.RemoteList1.List(GridList1) = 1
            Call SendDataToRemoteUser1("##FireAt##:" & FillSpaces1((RemoteGridRefX), DataSpacing1) & FillSpaces1((RemoteGridRefY), DataSpacing1))
            Call DisEnableRemoteGrid1(False)
        Else
            Call AddAMessage1((0), ("You Have Already Fired There !"))
        End If
  End If
End Sub

Private Sub RGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Let RemoteGridRefX = Int((GridBlocksX / 100) * Int((X / RGrid1.Width) * 100))
    Let RemoteGridRefY = Int((GridBlocksY / 100) * Int((Y / RGrid1.Height) * 100))
''===============================================================================''
    If (RemoteGridRefX < 0) Then Let RemoteGridRefX = 0
    If (RemoteGridRefY < 0) Then Let RemoteGridRefY = 0
''===============================================================================''
    If (RemoteGridRefX > (GridBlocksX - 1)) Then Let RemoteGridRefX = (GridBlocksX - 1)
    If (RemoteGridRefY > (GridBlocksY - 1)) Then Let RemoteGridRefY = (GridBlocksY - 1)
''===============================================================================''
    If (RemoteGridEnabled1) Then
        If (Main1.Target1.Visible = False) Then Let Main1.Target1.Visible = True
        Call Main1.Target1.Move((GridSizeX * RemoteGridRefX), (GridSizeY * RemoteGridRefY))
    End If
End Sub

Private Sub SendMessage1_Change()
    If (Len(Trim(Main1.SendMessage1.Text)) = 0) Then
        Let Main1.MessageButton1.Enabled = False
        Let Main1.SendMessage1.Text = ""
    Else
        Let Main1.MessageButton1.Enabled = True
    End If
End Sub

Private Sub VMsgScroll1_Change()
    Call RefreshMsgDisplay1
End Sub

'Thraddash Software - Trevor Lewis

