Attribute VB_Name = "General1"
Option Explicit

''===============================================================================''
Public Exx As Long, Why As Long
''===============================================================================''
Public NumVisibleMsgs1 As Long
''===============================================================================''
Public LocalGridRefX As Long, LocalGridRefY As Long
Public RemoteGridRefX As Long, RemoteGridRefY As Long
Public EditorGridRefX As Long, EditorGridRefY As Long
''===============================================================================''
Public Const GridBlocksX = 10
Public Const GridBlocksY = 10
Public Const GridSizeX = 27
Public Const GridSizeY = 27
''===============================================================================''
Public Const DataSpacing1 = 5
''===============================================================================''
Public RemoteGridEnabled1 As Boolean
Public EditModeOnOff1 As Boolean
Public EndOfGame1 As Boolean
Public EditModeCurrentShip1 As Long
''===============================================================================''
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
''===============================================================================''

Public Sub Startup_Procedure1()
    Let Exx = Screen.TwipsPerPixelX
    Let Why = Screen.TwipsPerPixelY
''===============================================================================''
    Call Randomize
''===============================================================================''
    Call DrawShipRotations1
End Sub

Public Sub Load_Main1()
Dim T As Long
''===============================================================================''
    Call ChangePlayerNames1("", "")
''===============================================================================''
    Call ShapeForm1(Main1, Main1.Point(0, 0))
''===============================================================================''
    Let Main1.ScaleMode = vbTwips
    Let Main1.Frame1.ScaleMode = vbPixels
''===============================================================================''
    Call Main1.Frame1.Move((0), (0))
''===============================================================================''
    Let Main1.Width = Main1.Frame1.Width
    Let Main1.Height = Main1.Frame1.Height
    Call Main1.Move(((Screen.Width / 2) - (Main1.Width / 2)), ((Screen.Height / 2) - (Main1.Height / 2)))
''===============================================================================''
    With Main1.LGrid1
        Let .ScaleMode = vbPixels
        Let .BorderStyle = 0
        Let .AutoRedraw = True
        Let .AutoSize = True
        Let .Picture = Data1.LocalGrid1.Image
        Call .Move((0), (294))
    End With
    With Main1.RGrid1
        Let .ScaleMode = vbPixels
        Let .BorderStyle = 0
        Let .AutoRedraw = True
        Let .AutoSize = True
        Let .Picture = Data1.RemoteGrid1.Image
        Call .Move((294), (0))
    End With
''===============================================================================''
    Call Main1.HMsgScroll1.Move((278), (541), (219), (16))
    Call Main1.VMsgScroll1.Move((500), (443), (16), (95))
    Call Main1.MsgBox1.Move((279), (444), (217), (93))
    Call Main1.MsgBox2.Move((0), (0), (Main1.MsgBox1.Width), (Main1.MsgBox1.Height))
    For T = 0 To 4
        If (T > 0) Then
            Call Load(Main1.MsgPic1(T))
            Call Load(Main1.MsgText1(T))
        End If
        Let Main1.MsgPic1(T).Stretch = True
        Let Main1.MsgPic1(T).Width = (15)
        Let Main1.MsgPic1(T).Height = (15)
        Let Main1.MsgText1(T).AutoSize = False
        Let Main1.MsgText1(T).AutoSize = True
        If (T = 0) Then
            Call Main1.MsgPic1(T).Move((3), (3))
        Else
            Call Main1.MsgPic1(T).Move((Main1.MsgPic1(T - 1).Left), (Main1.MsgPic1(T - 1).Top - -Main1.MsgPic1(T - 1).Height - -(3)))
        End If
        Call Main1.MsgText1(T).Move((Main1.MsgPic1(T).Left - -Main1.MsgPic1(T).Width - -(4)), (Main1.MsgPic1(T).Top - -(Main1.MsgPic1(T).Height / 2) - (Main1.MsgText1(T).Height / 2)))
        Let Main1.MsgPic1(T).Visible = False
        Let Main1.MsgText1(T).Visible = Main1.MsgPic1(T).Visible
    Next T
''===============================================================================''
    Call Main1.ShipHit1.Move((50), (9), (236), (112))
    Call Main1.ShipHit1.Cls
''===============================================================================''
    Let Main1.ShipsName1.AutoSize = False: Let Main1.ShipsName1.AutoSize = True
    Call Main1.ShipsName1.Move((3), (Main1.ShipHit1.Height - (Main1.ShipsName1.Height - -(3))))
    Let Main1.ShipsName1.Caption = ""
''===============================================================================''
    Let Main1.Destroyed1.AutoSize = False: Let Main1.Destroyed1.AutoSize = True
    Call Main1.Destroyed1.Move((3), (3))
    Let Main1.Destroyed1.Caption = "Destroyed"
    Let Main1.Destroyed1.Visible = False
''===============================================================================''
    Call Main1.SendMessage1.Move((398), (411), (109), (20))
    Let Main1.SendMessage1.Text = ""
    Call Main1.MessageButton1.Move((509), (411), (20), (20))
    Let Main1.MessageButton1.Enabled = False
    Let Main1.MessageButton1.Default = True
''===============================================================================''
    Let Main1.LocalVictory1(0).Stretch = True
    Let Main1.LocalVictory1(1).Stretch = True
    Call Main1.LocalVictory1(0).Move((65), (248), (8), (15))
    Call Main1.LocalVictory1(1).Move((74), (248), (8), (15))
    Let Main1.LocalVictory1(0).Picture = Data1.Numbers1(0).Picture
    Let Main1.LocalVictory1(1).Picture = Data1.Numbers1(0).Picture
    Let Data1.LVictory1.Value = 0
    Let Main1.RemoteVictory1(0).Stretch = True
    Let Main1.RemoteVictory1(1).Stretch = True
    Call Main1.RemoteVictory1(0).Move((483), (302), (8), (15))
    Call Main1.RemoteVictory1(1).Move((492), (302), (8), (15))
    Let Main1.RemoteVictory1(0).Picture = Data1.Numbers1(0).Picture
    Let Main1.RemoteVictory1(1).Picture = Data1.Numbers1(0).Picture
    Let Data1.RVictory1.Value = 0
''===============================================================================''
    Call Data1.Messages1(0).Clear
    Call Data1.Messages1(1).Clear
    Call RebuildMsgDisplay1
''===============================================================================''
    Let Main1.Explosion1.Visible = False
    Call Main1.Explosion1.ZOrder(0)
    Let Main1.Edit1.Visible = False
    Let Main1.Target1.Visible = False
''===============================================================================''
    Call DisEnableRemoteGrid1(False)
    Call EditModeOnOrOff1(False)
    Let EndOfGame1 = True
''===============================================================================''
    Let Main1.PosLineXY1(0).X1 = (0)
    Let Main1.PosLineXY1(1).Y1 = (0)
    Let Main1.PosLineXY1(0).Visible = False
    Let Main1.PosLineXY1(1).Visible = False
''===============================================================================''
    Let Main1.LocalAmount1.Picture = Data1.LocalAmount1(0).Picture
    Let Main1.RemoteAmount1.Picture = Data1.RemoteAmount1(0).Picture
''===============================================================================''
    Let Main1.LocalAmount1.Stretch = True: Let Main1.LocalAmount1.Stretch = False
    Let Main1.RemoteAmount1.Stretch = True: Let Main1.RemoteAmount1.Stretch = False
''===============================================================================''
    Call Main1.LocalAmount1.Move((140), (271))
    Call Main1.RemoteAmount1.Move((354), (283))
    Let Main1.LocalAmount1.Visible = False
    Let Main1.RemoteAmount1.Visible = False
End Sub

Public Sub Project_Pause1(Input1 As Double)
Dim tmpDbl1 As Double, tmpDbl2(3) As Double
''===============================================================================''
    Let tmpDbl1 = Val(Input1)
    Let tmpDbl2(0) = (Timer - -tmpDbl1)
    Let tmpDbl2(1) = Timer: Let tmpDbl2(3) = Timer
    Do Until (Timer >= tmpDbl2(0))
        If (Timer < tmpDbl2(1)) Then
            Let tmpDbl2(0) = (Timer - -(tmpDbl1 - tmpDbl2(2)))
            Let tmpDbl2(1) = Timer: Let tmpDbl2(3) = Timer
        Else
            Let tmpDbl2(1) = Timer
            Let tmpDbl2(2) = (Timer - tmpDbl2(3))
        End If
        Call Sleep(1)
        DoEvents
    Loop
End Sub

Public Sub SendDataToRemoteUser1(Input1 As String)
    Connection1.Outbox1.AddItem (Input1)
End Sub

Public Sub SendDataToLocalUser1(Input1 As String)
    Connection1.Inbox1.AddItem (Input1)
End Sub

Public Sub ProcessReceivedData1()
Dim Input1 As String
''===============================================================================''
    If Not (Connection1.Inbox1.ListCount = 0) Then
        Let Input1 = Connection1.Inbox1.List(0)
        Call Connection1.Inbox1.RemoveItem(0)
''===============================================================================''
        If Trim$(UCase$(Left$(Input1, 18)) = "##SENDMEYOURNAME##") Then
            Call SendDataToRemoteUser1("##PlayerName##:" & Trim$(Connection1.PlayerNText1.Text))
        ElseIf Trim$(UCase$(Left$(Input1, 15)) = "##PLAYERNAME##:") Then
            Call ChangePlayerNames1(Trim$(Connection1.PlayerNText1.Text), Trim$(Right$(Input1, (Len(Input1) - 15))))
        ElseIf Trim$(UCase$(Left$(Input1, 17)) = "##STARTANEWGAME##") Then
            Call StartNewGame1
        ElseIf Trim$(UCase$(Left$(Input1, 11)) = "##FIREAT##:") Then
            Call DestroyBlock1(Trim$(Left$((Right$(Input1, (DataSpacing1 * 2))), DataSpacing1)), Trim$(Right$(Input1, DataSpacing1)))
            If Not (EndOfGame1) Then
                Call DisEnableRemoteGrid1(True)
                Call AddAMessage1((3), ("Your Turn !"))
                Call Beep
            Else
                Call DisEnableRemoteGrid1(False)
            End If
        ElseIf Trim$(UCase$(Left$(Input1, 15)) = "##READYTOPLAY##") Then
            Let Data1.PlayerReady1(1).Caption = 1
        ElseIf Trim$(UCase$(Left$(Input1, 14)) = "##YOUGOFIRST##") Then
            Call DisEnableRemoteGrid1(True)
            Call AddAMessage1((3), ("You Get To Go First !"))
            Call Beep
        ElseIf Trim$(UCase$(Left$(Input1, 15)) = "##YOUGOSECOND##") Then
            Call AddAMessage1((0), (Trim$(Main1.RemoteName1.Caption) & " Will Go First !"))
        ElseIf Trim$(UCase$(Left$(Input1, 10)) = "#HITSHIP#:") Then
            Call ProcessHitMiss1(Right$(Input1, (Len(Input1) - 10)), (0))
        ElseIf Trim$(UCase$(Left$(Input1, 11)) = "#MISSSHIP#:") Then
            Call ProcessHitMiss1(Right$(Input1, (Len(Input1) - 11)), (1))
        ElseIf Trim$(UCase$(Left$(Input1, 14)) = "##RUMESSAGE##:") Then
            Call AddAMessage1((7), (Trim$(Main1.RemoteName1.Caption) & ": " & Right$(Input1, (Len(Input1) - 14))))
        ElseIf Trim$(UCase$(Left$(Input1, 14)) = "##LUMESSAGE##:") Then
            Call AddAMessage1((8), (Trim$(Main1.LocalName1.Caption) & ": " & Right$(Input1, (Len(Input1) - 14))))
        ElseIf Trim$(UCase$(Left$(Input1, 12)) = "#SHIPPIECE#:") Then
            If (EndOfGame1) Then DrawHitMissBlocks1 Trim$(Left$((Right$(Input1, (DataSpacing1 * 2))), DataSpacing1)), Trim$(Right$(Input1, DataSpacing1)), (2)
        End If
    End If
End Sub

'Thraddash Software - Trevor Lewis

