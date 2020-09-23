VERSION 5.00
Begin VB.UserControl TSPanel 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   ControlContainer=   -1  'True
   ScaleHeight     =   3000
   ScaleWidth      =   3000
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.PictureBox CaptionPanel1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2000
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   120
      Width           =   2000
      Begin VB.Label Caption1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TSPanel"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   645
      End
      Begin VB.Label Caption1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TSPanel"
         ForeColor       =   &H80000010&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Caption1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TSPanel"
         ForeColor       =   &H80000014&
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   645
      End
   End
End
Attribute VB_Name = "TSPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Exx As Long, Why As Long

Private var_Font3D As Byte
Private var_Alignment As Byte
Private var_AutoSize As Byte
Private var_BevelInner As Byte
Private var_BevelOuter As Byte
Private var_BevelWidth As Long
Private var_BorderWidth As Long
Public Enum enum_BevelInnerOuterOpitons
    tsNoneBevel = 0
    tsInsetBevel = 1
    tsRaisedBevel = 2
End Enum
Public Enum enum_AlignmentOptions
    tsLeftTop = 0
    tsLeftMiddle = 1
    tsLeftBottom = 2
    tsRightTop = 3
    tsRightMiddle = 4
    tsRightBottom = 5
    tsCenterTop = 6
    tsCenterMiddle = 7
    tsCenterBottom = 8
End Enum
Public Enum enum_AutoSizeOptions
    tsNoneAutoSize = 0
    tsWidthToCaption = 1
    tsHeightToCaption = 2
    tsBothToCaption = 3
End Enum
Public Enum enum_Font3DOptions
    tsNoneFont3D = 0
    tsRaisedLight = 1
    tsRaisedHeavy = 2
    tsInsetLight = 3
    tsInsetHeavy = 4
End Enum

Public Property Let Font3D(ByVal Input1 As enum_Font3DOptions)
    Let var_Font3D = Input1
    Call fnc_DesignTheControl
End Property
Public Property Get Font3D() As enum_Font3DOptions
    Let Font3D = var_Font3D
End Property
Public Property Let ForeColor(ByVal Input1 As OLE_COLOR)
    Let UserControl.Caption1(0).ForeColor = Input1
    Call fnc_DesignTheControl
End Property
Public Property Get ForeColor() As OLE_COLOR
    Let ForeColor = UserControl.Caption1(0).ForeColor
End Property
Public Property Let AutoSize(ByVal Input1 As enum_AutoSizeOptions)
    Let var_AutoSize = Input1
    Call fnc_DesignTheControl
End Property
Public Property Get AutoSize() As enum_AutoSizeOptions
    Let AutoSize = var_AutoSize
End Property
Public Property Let Alignment(ByVal Input1 As enum_AlignmentOptions)
    Let var_Alignment = Input1
    Call fnc_DesignTheControl
End Property
Public Property Get Alignment() As enum_AlignmentOptions
    Let Alignment = var_Alignment
End Property
Public Property Set Font(ByVal Input1 As Font)
    Set UserControl.Caption1(0).Font = Input1
    Call fnc_DesignTheControl
End Property
Public Property Get Font() As Font
    Set Font = UserControl.Caption1(0).Font
End Property
Public Property Let Caption(ByVal Input1 As String)
    Let UserControl.Caption1(0).Caption = Input1
    Call fnc_DesignTheControl
End Property
Public Property Get Caption() As String
    Let Caption = UserControl.Caption1(0).Caption
End Property
Public Property Let BackColor(ByVal Input1 As OLE_COLOR)
    Let UserControl.BackColor = Input1
    Let UserControl.CaptionPanel1.BackColor = Input1
    Call fnc_DesignTheControl
End Property
Public Property Get BackColor() As OLE_COLOR
    Let BackColor = UserControl.BackColor
End Property
Public Property Let BevelInner(Input1 As enum_BevelInnerOuterOpitons)
    Let var_BevelInner = Input1
    Call fnc_DesignTheControl
End Property
Public Property Get BevelInner() As enum_BevelInnerOuterOpitons
    Let BevelInner = var_BevelInner
End Property
Public Property Let BevelOuter(Input1 As enum_BevelInnerOuterOpitons)
    Let var_BevelOuter = Input1
    Call fnc_DesignTheControl
End Property
Public Property Get BevelOuter() As enum_BevelInnerOuterOpitons
    Let BevelOuter = var_BevelOuter
End Property
Public Property Let BevelWidth(Input1 As Long)
    If (Input1 < 0) Then Let Input1 = 0
    If (Input1 > 30) Then Let Input1 = 30
    Let var_BevelWidth = Input1
    Call fnc_DesignTheControl
End Property
Public Property Get BevelWidth() As Long
    Let BevelWidth = var_BevelWidth
End Property
Public Property Let BorderWidth(Input1 As Long)
    If (Input1 < 0) Then Let Input1 = 0
    If (Input1 > 30) Then Let Input1 = 30
    Let var_BorderWidth = Input1
    Call fnc_DesignTheControl
End Property
Public Property Get BorderWidth() As Long
    Let BorderWidth = var_BorderWidth
End Property

Private Sub fnc_DesignTheControl()
Dim BorderColor(1) As Long
Dim tmpLng1 As Long, tmpLng2 As Long, BorderCounter As Long
Dim tmpByte1 As Byte, BevelBorders(1) As Byte
Dim ShowControlContents1 As Boolean
    
    Call UserControl.Cls
    Let BorderCounter = 0
    Let BevelBorders(0) = var_BevelOuter
    Let BevelBorders(1) = var_BevelInner
    Let ShowControlContents1 = False
    
    'Draw The Controls Borders
    For tmpByte1 = 0 To UBound(BevelBorders)
        Select Case BevelBorders(tmpByte1)
            Case 1  'Border Inset
                Let BorderColor(0) = &H80000010     'Button Shadow
                Let BorderColor(1) = &H80000014     'Button Highlight
            Case 2  'Border Raised
                Let BorderColor(0) = &H80000014     'Button Highlight
                Let BorderColor(1) = &H80000010     'Button Shadow
        End Select
        If Not (BevelBorders(tmpByte1) = 0) Then
            For tmpLng1 = 0 To (var_BevelWidth - 1)
                UserControl.Line ((BorderCounter - -tmpLng1) * Exx, (BorderCounter - -tmpLng1) * Why)-((((UserControl.Width / Exx) - (BorderCounter - -tmpLng1)) * Exx), (BorderCounter - -tmpLng1) * Why), BorderColor(0)
                UserControl.Line ((BorderCounter - -tmpLng1) * Exx, (BorderCounter - -tmpLng1) * Why)-(((BorderCounter - -tmpLng1) * Exx), (((UserControl.Height / Why) - (BorderCounter - -tmpLng1)) * Why)), BorderColor(0)
                UserControl.Line ((BorderCounter - -tmpLng1) * Exx, (((UserControl.Height / Why) - (BorderCounter - -tmpLng1)) * Why) - (1 * Why))-((((UserControl.Width / Exx) - (BorderCounter - -tmpLng1)) * Exx), (((UserControl.Height / Why) - (BorderCounter - -tmpLng1)) * Why) - (1 * Why)), BorderColor(1)
                UserControl.Line ((((UserControl.Width / Exx) - (BorderCounter - -tmpLng1)) * Exx) - (1 * Exx), (BorderCounter - -tmpLng1) * Why)-((((UserControl.Width / Exx) - (BorderCounter - -tmpLng1)) * Exx) - (1 * Exx), (((UserControl.Height / Why) - (BorderCounter - -tmpLng1)) * Why)), BorderColor(1)
            Next tmpLng1
            Let BorderCounter = (BorderCounter - -var_BevelWidth)
        End If
        If (tmpByte1 = 0) Then
            If Not (BevelBorders(1) = 0) Then
                Let BorderCounter = (BorderCounter - -var_BorderWidth)
            End If
        End If
    Next tmpByte1
        
    'Position The Caption Area
    With UserControl.CaptionPanel1
        Let .Visible = False
        Call .Move((BorderCounter * Exx), (BorderCounter * Why))
        Let tmpLng1 = ((UserControl.Width - (BorderCounter * Exx)) - .Left)
        Let tmpLng2 = ((UserControl.Height - (BorderCounter * Why)) - .Top)
        If (tmpLng1 >= 0) And (tmpLng2 >= 0) Then
            Let .Width = tmpLng1
            Let .Height = tmpLng2
            Let ShowControlContents1 = True
        Else
            If Not (tmpLng1 >= .Left) Then Let .Width = 0
            If Not (tmpLng2 >= .Top) Then Let .Height = 0
        End If
        'Position The Text
        Select Case var_Alignment   'Left, Center & Right Text (X)
            Case 0, 1, 2: UserControl.Caption1(0).Left = (1 * Exx)
            Case 3, 4, 5: UserControl.Caption1(0).Left = (.Width - (UserControl.Caption1(0).Width - -(1 * Exx)))
            Case 6, 7, 8: UserControl.Caption1(0).Left = ((.Width / 2) - (UserControl.Caption1(0).Width / 2))
        End Select
        Select Case var_Alignment   'Top, Middle & Bottom Text (Y)
            Case 0, 3, 6: UserControl.Caption1(0).Top = (1 * Why)
            Case 1, 4, 7: UserControl.Caption1(0).Top = ((.Height / 2) - (UserControl.Caption1(0).Height / 2))
            Case 2, 5, 8: UserControl.Caption1(0).Top = (.Height - (UserControl.Caption1(0).Height - -(1 * Why)))
        End Select
    End With
    
    'AutoSize The Control To The Caption
    With UserControl.CaptionPanel1
        Select Case var_AutoSize        'AutoSize Captions Width
            Case 1, 3
                Let tmpLng1 = (UserControl.Caption1(0).Width - -(2 * Exx))
                If Not (.Width = tmpLng1) Then
                    Let UserControl.Width = (((BorderCounter * 2) * Exx) - -tmpLng1)
                End If
        End Select
        Select Case var_AutoSize        'AutoSize Captions Height
            Case 2, 3
                Let tmpLng1 = (UserControl.Caption1(0).Height - -(2 * Why))
                If Not (.Height = tmpLng1) Then
                    Let UserControl.Height = (((BorderCounter * 2) * Why) - -tmpLng1)
                End If
        End Select
    End With
    
    'Apply 3D Effects To The Font
    If ShowControlContents1 Then
        Select Case var_Font3D
            Case 0
                If Not (UserControl.Caption1(0).Visible) Then Let UserControl.Caption1(0).Visible = True
                If Not (UserControl.Caption1(1).Visible = False) Then Let UserControl.Caption1(1).Visible = False
                If Not (UserControl.Caption1(2).Visible = False) Then Let UserControl.Caption1(2).Visible = False
            Case 1
                Set UserControl.Caption1(2).Font = UserControl.Caption1(0).Font
                Call UserControl.Caption1(2).Move(UserControl.Caption1(0).Left - (1 * Exx), UserControl.Caption1(0).Top - (1 * Why))
                Call UserControl.Caption1(0).ZOrder(0)
                If Not (UserControl.Caption1(0).Visible) Then Let UserControl.Caption1(0).Visible = True
                If Not (UserControl.Caption1(1).Visible = False) Then Let UserControl.Caption1(1).Visible = False
                If Not (UserControl.Caption1(2).Visible) Then Let UserControl.Caption1(2).Visible = True
            Case 2
                Set UserControl.Caption1(1).Font = UserControl.Caption1(0).Font
                Set UserControl.Caption1(2).Font = UserControl.Caption1(0).Font
                Call UserControl.Caption1(1).Move(UserControl.Caption1(0).Left - -(1 * Exx), UserControl.Caption1(0).Top - -(1 * Why))
                Call UserControl.Caption1(2).Move(UserControl.Caption1(0).Left - (1 * Exx), UserControl.Caption1(0).Top - (1 * Why))
                Call UserControl.Caption1(1).ZOrder(0)
                Call UserControl.Caption1(0).ZOrder(0)
                If Not (UserControl.Caption1(0).Visible) Then Let UserControl.Caption1(0).Visible = True
                If Not (UserControl.Caption1(1).Visible) Then Let UserControl.Caption1(1).Visible = True
                If Not (UserControl.Caption1(2).Visible) Then Let UserControl.Caption1(2).Visible = True
            Case 3
                Set UserControl.Caption1(2).Font = UserControl.Caption1(0).Font
                Call UserControl.Caption1(2).Move(UserControl.Caption1(0).Left - -(1 * Exx), UserControl.Caption1(0).Top - -(1 * Why))
                Call UserControl.Caption1(0).ZOrder(0)
                If Not (UserControl.Caption1(0).Visible) Then Let UserControl.Caption1(0).Visible = True
                If Not (UserControl.Caption1(1).Visible = False) Then Let UserControl.Caption1(1).Visible = False
                If Not (UserControl.Caption1(2).Visible) Then Let UserControl.Caption1(2).Visible = True
            Case 4
                Set UserControl.Caption1(1).Font = UserControl.Caption1(0).Font
                Set UserControl.Caption1(2).Font = UserControl.Caption1(0).Font
                Call UserControl.Caption1(1).Move(UserControl.Caption1(0).Left - (1 * Exx), UserControl.Caption1(0).Top - (1 * Why))
                Call UserControl.Caption1(2).Move(UserControl.Caption1(0).Left - -(1 * Exx), UserControl.Caption1(0).Top - -(1 * Why))
                Call UserControl.Caption1(2).ZOrder(0)
                Call UserControl.Caption1(0).ZOrder(0)
                If Not (UserControl.Caption1(0).Visible) Then Let UserControl.Caption1(0).Visible = True
                If Not (UserControl.Caption1(1).Visible) Then Let UserControl.Caption1(1).Visible = True
                If Not (UserControl.Caption1(2).Visible) Then Let UserControl.Caption1(2).Visible = True
        End Select
    End If
    If (Trim$(UserControl.Caption1(0).Caption) = "") Then Let ShowControlContents1 = False
    If Not (UserControl.CaptionPanel1.Visible = ShowControlContents1) Then Let UserControl.CaptionPanel1.Visible = ShowControlContents1
    Call UserControl.CaptionPanel1.Refresh
    
End Sub

Private Sub Caption1_Change(Index As Integer)
    If (Index = 0) Then
        Let UserControl.Caption1(1).Caption = UserControl.Caption1(0).Caption
        Let UserControl.Caption1(2).Caption = UserControl.Caption1(0).Caption
    End If
End Sub

Private Sub UserControl_Initialize()
    Let Exx = Screen.TwipsPerPixelX: Let Why = Screen.TwipsPerPixelY
    'Control Defaults
    UserControl.Circle (0, 0), 50
    UserControl.Line ((100), (100))-((200), (200))
    Let UserControl.BackColor = &H8000000F
    Let UserControl.CaptionPanel1.BackColor = &H8000000F
    Let UserControl.Caption1(0).ForeColor = &H80000008
    Let UserControl.Caption1(0).Caption = ""
    Let var_Alignment = 7                   'Align The Caption Center Middle
    Let var_AutoSize = 0                    'No AutoSizing The Control
    Let var_BevelInner = 0                  'No Inner Bevel
    Let var_BevelOuter = 2                  'Raise The Outer Bevel
    Let var_BevelWidth = 1                  'Have 1 Level Of Outer Border
    Let var_BorderWidth = 3                 'Have 3 Levels Of Border Width
    'Design The Control
    Call fnc_DesignTheControl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Let UserControl.Caption1(0).Caption = PropBag.ReadProperty("Caption", "")
    Let UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Let UserControl.CaptionPanel1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Let UserControl.Caption1(0).ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Set UserControl.Caption1(0).Font = PropBag.ReadProperty("Font")
    Let var_Alignment = PropBag.ReadProperty("Alignment", 7)
    Let var_AutoSize = PropBag.ReadProperty("AutoSize", 0)
    Let var_BevelInner = PropBag.ReadProperty("BevelInner", 0)
    Let var_BevelOuter = PropBag.ReadProperty("BevelOuter", 2)
    Let var_BevelWidth = PropBag.ReadProperty("BevelWidth", 1)
    Let var_BorderWidth = PropBag.ReadProperty("BorderWidth", 3)
    Let var_Font3D = PropBag.ReadProperty("Font3D", 0)
    Call fnc_DesignTheControl
End Sub

Private Sub UserControl_Resize()
    Call fnc_DesignTheControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", UserControl.Caption1(0).Caption, "")
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.Caption1(0).ForeColor, &H80000008)
    Call PropBag.WriteProperty("Font", UserControl.Caption1(0).Font)
    Call PropBag.WriteProperty("Alignment", var_Alignment, 7)
    Call PropBag.WriteProperty("AutoSize", var_AutoSize, 0)
    Call PropBag.WriteProperty("BevelInner", var_BevelInner, 0)
    Call PropBag.WriteProperty("BevelOuter", var_BevelOuter, 2)
    Call PropBag.WriteProperty("BevelWidth", var_BevelWidth, 1)
    Call PropBag.WriteProperty("BorderWidth", var_BorderWidth, 3)
    Call PropBag.WriteProperty("Font3D", var_Font3D, 0)
End Sub

'Thraddash Software - Trevor Lewis

