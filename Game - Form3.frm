VERSION 5.00
Begin VB.Form Ships1 
   Caption         =   "Pictures Of Space Ships And Masks"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13740
   Icon            =   "Game - Form3.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   13740
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.PictureBox Portrait1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1680
      Index           =   4
      Left            =   6600
      Picture         =   "Game - Form3.frx":000C
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   236
      TabIndex        =   39
      Top             =   3480
      Width           =   3540
   End
   Begin VB.PictureBox Portrait1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1680
      Index           =   3
      Left            =   10080
      Picture         =   "Game - Form3.frx":6B8E
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   236
      TabIndex        =   38
      Top             =   1800
      Width           =   3540
   End
   Begin VB.PictureBox Portrait1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1680
      Index           =   2
      Left            =   10080
      Picture         =   "Game - Form3.frx":D710
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   236
      TabIndex        =   37
      Top             =   120
      Width           =   3540
   End
   Begin VB.PictureBox Portrait1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1680
      Index           =   1
      Left            =   6600
      Picture         =   "Game - Form3.frx":14292
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   236
      TabIndex        =   36
      Top             =   1800
      Width           =   3540
   End
   Begin VB.PictureBox Portrait1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1680
      Index           =   0
      Left            =   6600
      Picture         =   "Game - Form3.frx":1AE14
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   236
      TabIndex        =   35
      Top             =   120
      Width           =   3540
   End
   Begin VB.PictureBox ShipDamaged2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   16
      Left            =   5040
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   34
      Top             =   6360
      Width           =   300
   End
   Begin VB.PictureBox ShipDamaged2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   12
      Left            =   6120
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   33
      Top             =   4440
      Width           =   300
   End
   Begin VB.PictureBox ShipDamaged2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   8
      Left            =   3720
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   32
      Top             =   2880
      Width           =   300
   End
   Begin VB.PictureBox ShipDamaged2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   3720
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   31
      Top             =   1680
      Width           =   300
   End
   Begin VB.PictureBox ShipDamaged2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   6120
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   30
      Top             =   120
      Width           =   300
   End
   Begin VB.PictureBox ShipFixed2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   16
      Left            =   4560
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   29
      Top             =   6360
      Width           =   300
   End
   Begin VB.PictureBox ShipFixed2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   12
      Left            =   5640
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   28
      Top             =   4440
      Width           =   300
   End
   Begin VB.PictureBox ShipFixed2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   8
      Left            =   3240
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   27
      Top             =   2880
      Width           =   300
   End
   Begin VB.PictureBox ShipFixed2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   4
      Left            =   3240
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   26
      Top             =   1680
      Width           =   300
   End
   Begin VB.PictureBox ShipFixed2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   5640
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   25
      Top             =   120
      Width           =   300
   End
   Begin VB.PictureBox ShipDamaged1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Index           =   16
      Left            =   3600
      Picture         =   "Game - Form3.frx":21996
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   55
      TabIndex        =   24
      Top             =   6360
      Width           =   825
   End
   Begin VB.PictureBox ShipFixed1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Index           =   16
      Left            =   2640
      Picture         =   "Game - Form3.frx":21B98
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   55
      TabIndex        =   23
      Top             =   6360
      Width           =   825
   End
   Begin VB.PictureBox ShipMask1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Index           =   16
      Left            =   1680
      Picture         =   "Game - Form3.frx":22BE2
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   55
      TabIndex        =   22
      Top             =   6360
      Width           =   825
   End
   Begin VB.PictureBox ShipDamaged1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1635
      Index           =   12
      Left            =   4320
      Picture         =   "Game - Form3.frx":22DE4
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   82
      TabIndex        =   19
      Top             =   4440
      Width           =   1230
   End
   Begin VB.PictureBox ShipFixed1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1635
      Index           =   12
      Left            =   3000
      Picture         =   "Game - Form3.frx":2334A
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   82
      TabIndex        =   18
      Top             =   4440
      Width           =   1230
   End
   Begin VB.PictureBox ShipMask1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1635
      Index           =   12
      Left            =   1680
      Picture         =   "Game - Form3.frx":25B50
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   82
      TabIndex        =   17
      Top             =   4440
      Width           =   1230
   End
   Begin VB.PictureBox ShipDamaged1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1230
      Index           =   8
      Left            =   2640
      Picture         =   "Game - Form3.frx":260B6
      ScaleHeight     =   82
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   14
      Top             =   2880
      Width           =   420
   End
   Begin VB.PictureBox ShipFixed1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1230
      Index           =   8
      Left            =   2160
      Picture         =   "Game - Form3.frx":26248
      ScaleHeight     =   82
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   13
      Top             =   2880
      Width           =   420
   End
   Begin VB.PictureBox ShipMask1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1230
      Index           =   8
      Left            =   1680
      Picture         =   "Game - Form3.frx":26F82
      ScaleHeight     =   82
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   12
      Top             =   2880
      Width           =   420
   End
   Begin VB.PictureBox ShipDamaged1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Index           =   4
      Left            =   2640
      Picture         =   "Game - Form3.frx":27114
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   9
      Top             =   1680
      Width           =   420
   End
   Begin VB.PictureBox ShipDamaged1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1230
      Index           =   0
      Left            =   4320
      Picture         =   "Game - Form3.frx":2723A
      ScaleHeight     =   82
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   82
      TabIndex        =   8
      Top             =   120
      Width           =   1230
   End
   Begin VB.PictureBox ShipFixed1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Index           =   4
      Left            =   2160
      Picture         =   "Game - Form3.frx":2765C
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   7
      Top             =   1680
      Width           =   420
   End
   Begin VB.PictureBox ShipMask1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Index           =   4
      Left            =   1680
      Picture         =   "Game - Form3.frx":280A2
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   6
      Top             =   1680
      Width           =   420
   End
   Begin VB.PictureBox ShipFixed1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1230
      Index           =   0
      Left            =   3000
      Picture         =   "Game - Form3.frx":281C8
      ScaleHeight     =   82
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   82
      TabIndex        =   3
      Top             =   120
      Width           =   1230
   End
   Begin VB.PictureBox ShipMask1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1230
      Index           =   0
      Left            =   1680
      Picture         =   "Game - Form3.frx":2A0F2
      ScaleHeight     =   82
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   82
      TabIndex        =   2
      Top             =   120
      Width           =   1230
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   19
      Left            =   1200
      Picture         =   "Game - Form3.frx":2A514
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   375
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   18
      Left            =   840
      Picture         =   "Game - Form3.frx":2AD5A
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   375
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   17
      Left            =   480
      Picture         =   "Game - Form3.frx":2B5A0
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   375
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   16
      Left            =   120
      Picture         =   "Game - Form3.frx":2BDE6
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label ShipName1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Interceptor"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   21
      Top             =   6840
      Width           =   765
   End
   Begin VB.Label Direction1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   7080
      Width           =   90
   End
   Begin VB.Label Direction1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   5040
      Width           =   90
   End
   Begin VB.Label ShipName1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "War Ship"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   4800
      Width           =   660
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   15
      Left            =   1200
      Picture         =   "Game - Form3.frx":2C62C
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   14
      Left            =   840
      Picture         =   "Game - Form3.frx":2DDC6
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   13
      Left            =   480
      Picture         =   "Game - Form3.frx":2F628
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   12
      Left            =   120
      Picture         =   "Game - Form3.frx":30DC2
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   11
      Left            =   1200
      Picture         =   "Game - Form3.frx":32624
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   10
      Left            =   840
      Picture         =   "Game - Form3.frx":32CCE
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   9
      Left            =   480
      Picture         =   "Game - Form3.frx":333C0
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   8
      Left            =   120
      Picture         =   "Game - Form3.frx":33A6A
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label ShipName1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Defender"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   660
   End
   Begin VB.Label Direction1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   90
   End
   Begin VB.Label Direction1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   90
   End
   Begin VB.Label ShipName1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Scout Ship"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   780
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   7
      Left            =   1200
      Picture         =   "Game - Form3.frx":3415C
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   375
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   6
      Left            =   840
      Picture         =   "Game - Form3.frx":345D6
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   375
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   5
      Left            =   480
      Picture         =   "Game - Form3.frx":34AAC
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   375
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   4
      Left            =   120
      Picture         =   "Game - Form3.frx":34F26
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   375
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   3
      Left            =   1200
      Picture         =   "Game - Form3.frx":353FC
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   2
      Left            =   840
      Picture         =   "Game - Form3.frx":36676
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   1
      Left            =   480
      Picture         =   "Game - Form3.frx":378F0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Image ShipOutline1 
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "Game - Form3.frx":38B6A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Label ShipName1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Battle Cruiser"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   930
   End
   Begin VB.Label Direction1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   90
   End
End
Attribute VB_Name = "Ships1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Thraddash Software - Trevor Lewis

