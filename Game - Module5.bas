Attribute VB_Name = "Conversation1"
Option Explicit

Public Sub AddAMessage1(Input1 As Long, Input2 As String)
    Call Data1.Messages1(0).AddItem(Trim$(Input1))
    Call Data1.Messages1(1).AddItem(Trim$(Input2))
''===============================================================================''
    If (Data1.Messages1(0).ListCount > 500) Then
        Call Data1.Messages1(0).RemoveItem(0)
        Call Data1.Messages1(1).RemoveItem(0)
    End If
''===============================================================================''
    Call RebuildMsgDisplay1
End Sub

Public Sub RebuildMsgDisplay1()
Dim T As Long
''===============================================================================''
    If ((Data1.Messages1(0).ListCount - 1) <= Main1.MsgPic1.UBound) Then
        Let NumVisibleMsgs1 = (Data1.Messages1(0).ListCount - 1)
        If (Main1.VMsgScroll1.Enabled) Then
            Let Main1.VMsgScroll1.Enabled = False
            Let Main1.VMsgScroll1.Max = 0
            Let Main1.VMsgScroll1.Value = 0
        End If
    Else
        Let NumVisibleMsgs1 = (Main1.MsgPic1.UBound)
        Let Main1.VMsgScroll1.Max = ((Data1.Messages1(0).ListCount - 1) - NumVisibleMsgs1)
        If Not (Main1.VMsgScroll1.Enabled) Then Let Main1.VMsgScroll1.Enabled = True
    End If
''===============================================================================''
    For T = (NumVisibleMsgs1 - -1) To Main1.MsgPic1.UBound
        Let Main1.MsgPic1(T).Visible = False
        Let Main1.MsgText1(T).Visible = False
    Next T
''===============================================================================''
    Let Main1.VMsgScroll1.Value = Main1.VMsgScroll1.Max
''===============================================================================''
    Call RefreshMsgDisplay1
End Sub

Public Sub RefreshMsgDisplay1()
Dim T As Long
Dim Add1 As Long
Dim MoveIt1 As Long
''===============================================================================''
    Let MoveIt1 = 0
    Let Add1 = Main1.VMsgScroll1.Value
''===============================================================================''
    For T = 0 To NumVisibleMsgs1
        Let Main1.MsgPic1(T).Visible = False
        Let Main1.MsgText1(T).Visible = False
        Let Main1.MsgPic1(T).Picture = Data1.MsgPictures1(Data1.Messages1(0).List(T - -Add1)).Picture
        Select Case (Data1.Messages1(0).List(T - -Add1))
            Case 0: Let Main1.MsgText1(T).ForeColor = QBColor(13)
            Case 1: Let Main1.MsgText1(T).ForeColor = QBColor(12)
            Case 2: Let Main1.MsgText1(T).ForeColor = QBColor(10)
            Case 3: Let Main1.MsgText1(T).ForeColor = QBColor(14)
            Case 4: Let Main1.MsgText1(T).ForeColor = QBColor(7)
            Case 5: Let Main1.MsgText1(T).ForeColor = QBColor(15)
            Case 6: Let Main1.MsgText1(T).ForeColor = QBColor(12)
            Case 7: Let Main1.MsgText1(T).ForeColor = QBColor(14)
            Case 8: Let Main1.MsgText1(T).ForeColor = QBColor(7)
        End Select
        Let Main1.MsgText1(T).Caption = Data1.Messages1(1).List(T - -Add1)
        If (MoveIt1 < (Main1.MsgText1(T).Left - -Main1.MsgText1(T).Width)) Then Let MoveIt1 = (Main1.MsgText1(T).Left - -Main1.MsgText1(T).Width - -(5))
        Let Main1.MsgPic1(T).Visible = True
        Let Main1.MsgText1(T).Visible = True
    Next T
''===============================================================================''
    If (MoveIt1 > Main1.MsgBox1.ScaleWidth) Then
        Let Main1.MsgBox2.Width = MoveIt1
        Let Main1.HMsgScroll1.Max = (Main1.MsgBox2.ScaleWidth - Main1.MsgBox1.ScaleWidth)
        If (Main1.HMsgScroll1.Enabled = False) Then Let Main1.HMsgScroll1.Enabled = True
    Else
        Let Main1.MsgBox2.Width = Main1.MsgBox1.Width
        If Not (Main1.HMsgScroll1.Max = 0) Then Let Main1.HMsgScroll1.Max = 0
        If (Main1.HMsgScroll1.Enabled = True) Then Let Main1.HMsgScroll1.Enabled = False
    End If
End Sub

'Thraddash Software - Trevor Lewis

