VERSION 5.00
Begin VB.Form Data1 
   Caption         =   "Grid Data And Animations"
   ClientHeight    =   7860
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8910
   Icon            =   "Game - Form2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.PictureBox RemoteHitMiss1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   5
      Left            =   3240
      Picture         =   "Game - Form2.frx":000C
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   22
      Top             =   5400
      Width           =   420
   End
   Begin VB.PictureBox RemoteHitMiss1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   4
      Left            =   2760
      Picture         =   "Game - Form2.frx":097E
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   21
      Top             =   5400
      Width           =   420
   End
   Begin VB.HScrollBar RVictory1 
      Height          =   255
      Left            =   120
      Max             =   100
      TabIndex        =   20
      Top             =   6840
      Width           =   2535
   End
   Begin VB.HScrollBar LVictory1 
      Height          =   255
      Left            =   120
      Max             =   100
      TabIndex        =   19
      Top             =   6600
      Width           =   2535
   End
   Begin VB.ListBox Messages1 
      Height          =   2010
      Index           =   1
      Left            =   5760
      TabIndex        =   18
      Top             =   4440
      Width           =   1455
   End
   Begin VB.ListBox Messages1 
      Height          =   2010
      Index           =   0
      Left            =   4320
      TabIndex        =   17
      Top             =   4440
      Width           =   1455
   End
   Begin VB.PictureBox RemoteHitMiss1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   3
      Left            =   3240
      Picture         =   "Game - Form2.frx":12F0
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   16
      Top             =   4920
      Width           =   420
   End
   Begin VB.PictureBox RemoteHitMiss1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   2
      Left            =   2760
      Picture         =   "Game - Form2.frx":1C62
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   15
      Top             =   4920
      Width           =   420
   End
   Begin VB.PictureBox RemoteHitMiss1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   1
      Left            =   3240
      Picture         =   "Game - Form2.frx":25D4
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   14
      Top             =   4440
      Width           =   420
   End
   Begin VB.PictureBox RemoteHitMiss1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   0
      Left            =   2760
      Picture         =   "Game - Form2.frx":2F46
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   13
      Top             =   4440
      Width           =   420
   End
   Begin VB.PictureBox HitPicture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4680
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   9
      Top             =   7200
      Width           =   495
   End
   Begin VB.ListBox ShipData2 
      Height          =   2010
      Index           =   0
      Left            =   2040
      TabIndex        =   8
      Top             =   4440
      Width           =   615
   End
   Begin VB.ListBox LocalList2 
      Height          =   2400
      Left            =   5760
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox HitPicture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4080
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   6
      Top             =   7200
      Width           =   495
   End
   Begin VB.ListBox ShipData1 
      Height          =   2010
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   4440
      Width           =   615
   End
   Begin VB.ListBox ShipData1 
      Height          =   2010
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   1335
   End
   Begin VB.ListBox RemoteList1 
      Height          =   2400
      Left            =   7320
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox LocalList1 
      Height          =   2400
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox RemoteGrid1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4065
      Left            =   120
      Picture         =   "Game - Form2.frx":38B8
      ScaleHeight     =   4065
      ScaleWidth      =   4065
      TabIndex        =   1
      Top             =   120
      Width           =   4065
   End
   Begin VB.PictureBox LocalGrid1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4065
      Left            =   120
      Picture         =   "Game - Form2.frx":15CEA
      ScaleHeight     =   4065
      ScaleWidth      =   4065
      TabIndex        =   0
      Top             =   120
      Width           =   4065
   End
   Begin VB.Image MsgPictures1 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   8
      Left            =   7320
      Picture         =   "Game - Form2.frx":2811C
      Top             =   5160
      Width           =   225
   End
   Begin VB.Image MsgPictures1 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   7
      Left            =   8400
      Picture         =   "Game - Form2.frx":28692
      Top             =   4800
      Width           =   225
   End
   Begin VB.Image Numbers1 
      Height          =   225
      Index           =   9
      Left            =   8280
      Picture         =   "Game - Form2.frx":28C08
      Top             =   6120
      Width           =   120
   End
   Begin VB.Image Numbers1 
      Height          =   225
      Index           =   8
      Left            =   8040
      Picture         =   "Game - Form2.frx":28D0A
      Top             =   6120
      Width           =   120
   End
   Begin VB.Image Numbers1 
      Height          =   225
      Index           =   7
      Left            =   7800
      Picture         =   "Game - Form2.frx":28E0C
      Top             =   6120
      Width           =   120
   End
   Begin VB.Image Numbers1 
      Height          =   225
      Index           =   6
      Left            =   7560
      Picture         =   "Game - Form2.frx":28F0E
      Top             =   6120
      Width           =   120
   End
   Begin VB.Image Numbers1 
      Height          =   225
      Index           =   5
      Left            =   7320
      Picture         =   "Game - Form2.frx":29010
      Top             =   6120
      Width           =   120
   End
   Begin VB.Image Numbers1 
      Height          =   225
      Index           =   4
      Left            =   8280
      Picture         =   "Game - Form2.frx":29112
      Top             =   5880
      Width           =   120
   End
   Begin VB.Image Numbers1 
      Height          =   225
      Index           =   3
      Left            =   8040
      Picture         =   "Game - Form2.frx":29214
      Top             =   5880
      Width           =   120
   End
   Begin VB.Image Numbers1 
      Height          =   225
      Index           =   2
      Left            =   7800
      Picture         =   "Game - Form2.frx":29316
      Top             =   5880
      Width           =   120
   End
   Begin VB.Image Numbers1 
      Height          =   225
      Index           =   1
      Left            =   7560
      Picture         =   "Game - Form2.frx":29418
      Top             =   5880
      Width           =   120
   End
   Begin VB.Image Numbers1 
      Height          =   225
      Index           =   0
      Left            =   7320
      Picture         =   "Game - Form2.frx":2951A
      Top             =   5880
      Width           =   120
   End
   Begin VB.Image MsgPictures1 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   6
      Left            =   8040
      Picture         =   "Game - Form2.frx":2961C
      Top             =   4800
      Width           =   225
   End
   Begin VB.Image MsgPictures1 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   5
      Left            =   7680
      Picture         =   "Game - Form2.frx":29B92
      Top             =   4800
      Width           =   225
   End
   Begin VB.Image MsgPictures1 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   4
      Left            =   7320
      Picture         =   "Game - Form2.frx":2A108
      Top             =   4800
      Width           =   225
   End
   Begin VB.Image MsgPictures1 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   3
      Left            =   8400
      Picture         =   "Game - Form2.frx":2A67E
      Top             =   4440
      Width           =   225
   End
   Begin VB.Image MsgPictures1 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   2
      Left            =   8040
      Picture         =   "Game - Form2.frx":2ABF4
      Top             =   4440
      Width           =   225
   End
   Begin VB.Image MsgPictures1 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   1
      Left            =   7680
      Picture         =   "Game - Form2.frx":2B16A
      Top             =   4440
      Width           =   225
   End
   Begin VB.Image MsgPictures1 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   0
      Left            =   7320
      Picture         =   "Game - Form2.frx":2B6E0
      Top             =   4440
      Width           =   225
   End
   Begin VB.Label PlayerReady1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   6600
      TabIndex        =   12
      Top             =   3000
      Width           =   90
   End
   Begin VB.Label PlayerReady1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   6480
      TabIndex        =   11
      Top             =   3000
      Width           =   90
   End
   Begin VB.Label HostClient1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "-1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6480
      TabIndex        =   10
      Top             =   2760
      Width           =   135
   End
   Begin VB.Image RemoteAmount1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   195
      Index           =   5
      Left            =   5400
      Picture         =   "Game - Form2.frx":2BC56
      Top             =   3960
      Width           =   1005
   End
   Begin VB.Image RemoteAmount1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   195
      Index           =   4
      Left            =   5400
      Picture         =   "Game - Form2.frx":2BEF0
      Top             =   3720
      Width           =   1005
   End
   Begin VB.Image RemoteAmount1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   195
      Index           =   3
      Left            =   5400
      Picture         =   "Game - Form2.frx":2C18A
      Top             =   3480
      Width           =   1005
   End
   Begin VB.Image RemoteAmount1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   195
      Index           =   2
      Left            =   5400
      Picture         =   "Game - Form2.frx":2C424
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Image RemoteAmount1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   195
      Index           =   1
      Left            =   5400
      Picture         =   "Game - Form2.frx":2C6BE
      Top             =   3000
      Width           =   1005
   End
   Begin VB.Image RemoteAmount1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   195
      Index           =   0
      Left            =   5400
      Picture         =   "Game - Form2.frx":2C958
      Top             =   2760
      Width           =   1005
   End
   Begin VB.Image LocalAmount1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   195
      Index           =   5
      Left            =   4320
      Picture         =   "Game - Form2.frx":2CBF2
      Top             =   3960
      Width           =   1005
   End
   Begin VB.Image LocalAmount1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   195
      Index           =   4
      Left            =   4320
      Picture         =   "Game - Form2.frx":2CE8C
      Top             =   3720
      Width           =   1005
   End
   Begin VB.Image LocalAmount1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   195
      Index           =   3
      Left            =   4320
      Picture         =   "Game - Form2.frx":2D126
      Top             =   3480
      Width           =   1005
   End
   Begin VB.Image LocalAmount1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   195
      Index           =   2
      Left            =   4320
      Picture         =   "Game - Form2.frx":2D3C0
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Image LocalAmount1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   195
      Index           =   1
      Left            =   4320
      Picture         =   "Game - Form2.frx":2D65A
      Top             =   3000
      Width           =   1005
   End
   Begin VB.Image LocalAmount1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   195
      Index           =   0
      Left            =   4320
      Picture         =   "Game - Form2.frx":2D8F4
      Top             =   2760
      Width           =   1005
   End
   Begin VB.Image Explosion1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   450
      Index           =   6
      Left            =   3000
      Picture         =   "Game - Form2.frx":2DB8E
      Top             =   7320
      Width           =   450
   End
   Begin VB.Image Explosion1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   450
      Index           =   5
      Left            =   2520
      Picture         =   "Game - Form2.frx":2E358
      Top             =   7320
      Width           =   450
   End
   Begin VB.Image Explosion1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   450
      Index           =   4
      Left            =   2040
      Picture         =   "Game - Form2.frx":2EB22
      Top             =   7320
      Width           =   450
   End
   Begin VB.Image Explosion1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   450
      Index           =   3
      Left            =   1560
      Picture         =   "Game - Form2.frx":2F2EC
      Top             =   7320
      Width           =   450
   End
   Begin VB.Image Explosion1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   450
      Index           =   2
      Left            =   1080
      Picture         =   "Game - Form2.frx":2FAB6
      Top             =   7320
      Width           =   450
   End
   Begin VB.Image Explosion1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   450
      Index           =   1
      Left            =   600
      Picture         =   "Game - Form2.frx":30280
      Top             =   7320
      Width           =   450
   End
   Begin VB.Image Explosion1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   450
      Index           =   0
      Left            =   120
      Picture         =   "Game - Form2.frx":30A4A
      Top             =   7320
      Width           =   450
   End
   Begin VB.Image Explosion1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   450
      Index           =   7
      Left            =   3480
      Picture         =   "Game - Form2.frx":31214
      Top             =   7320
      Width           =   450
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Begin VB.Menu Menu1_Sub2 
         Caption         =   "&Start A New Game"
      End
      Begin VB.Menu Menu1_Line1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_Sub1 
         Caption         =   "E&xit Game"
      End
   End
End
Attribute VB_Name = "Data1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LVictory1_Change()
Dim tmpString1 As String
''===============================================================================''
    If (Data1.LVictory1.Value = Data1.LVictory1.Max) Then
        Let Data1.LVictory1.Value = Data1.LVictory1.Min
    End If
''===============================================================================''
    Let tmpString1 = Trim(Data1.LVictory1.Value)
    If (Len(tmpString1) = 1) Then
        Let tmpString1 = ("0" & tmpString1)
    End If
''===============================================================================''
    Let Main1.LocalVictory1(0).Picture = Data1.Numbers1(Val(Left$(tmpString1, 1))).Picture
    Let Main1.LocalVictory1(1).Picture = Data1.Numbers1(Val(Right$(tmpString1, 1))).Picture
End Sub

Private Sub Menu1_Sub1_Click()
    End
End Sub

Private Sub Menu1_Sub2_Click()
    Call SendDataToRemoteUser1("##StartANewGame##")
    Call SendDataToLocalUser1("##StartANewGame##")
End Sub

Private Sub RVictory1_Change()
Dim tmpString1 As String
''===============================================================================''
    If (Data1.RVictory1.Value = Data1.RVictory1.Max) Then
        Let Data1.RVictory1.Value = Data1.RVictory1.Min
    End If
''===============================================================================''
    Let tmpString1 = Trim(Data1.RVictory1.Value)
    If (Len(tmpString1) = 1) Then
        Let tmpString1 = ("0" & tmpString1)
    End If
''===============================================================================''
    Let Main1.RemoteVictory1(0).Picture = Data1.Numbers1(Val(Left$(tmpString1, 1))).Picture
    Let Main1.RemoteVictory1(1).Picture = Data1.Numbers1(Val(Right$(tmpString1, 1))).Picture
End Sub

'Thraddash Software - Trevor Lewis

