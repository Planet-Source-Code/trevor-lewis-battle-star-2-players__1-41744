Attribute VB_Name = "GameCode1"
Option Explicit

Public Sub DisEnableRemoteGrid1(Input1 As Boolean)
    If (Input1) Then
        Let RemoteGridEnabled1 = True
        Let Main1.RGrid1.MousePointer = 99
    ElseIf Not (Input1) Then
        Let Main1.RGrid1.MousePointer = 0
        Let RemoteGridEnabled1 = False
        Let Main1.Target1.Visible = False
    End If
End Sub

Public Sub EditModeOnOrOff1(Input1 As Boolean)
    If (Input1) Then
        Let Main1.LGrid1.MousePointer = 99
        Let EditModeOnOff1 = True
    ElseIf Not (Input1) Then
        Let Main1.LGrid1.MousePointer = 0
        Let EditModeOnOff1 = False
        Let Main1.Edit1.Visible = False
    End If
End Sub

Public Sub StartNewGame1()
Dim T As Long
''===============================================================================''
    Call AddAMessage1((0), ("A New Game Has Been Started."))
    Call AddAMessage1((0), ("Right-Click To Rotate The Ship."))
''===============================================================================''
    Let Data1.Menu1_Sub2.Enabled = False
''===============================================================================''
    Let Data1.PlayerReady1(0).Caption = 0
    Let Data1.PlayerReady1(1).Caption = 0
''===============================================================================''
    Let Main1.PosLineXY1(0).X1 = (0)
    Let Main1.PosLineXY1(1).Y1 = (0)
    Let Main1.PosLineXY1(0).Visible = False
    Let Main1.PosLineXY1(1).Visible = False
''===============================================================================''
    Call DisEnableRemoteGrid1(False)
''===============================================================================''
    Let Main1.LGrid1.Picture = Data1.LocalGrid1.Image
    Let Main1.RGrid1.Picture = Data1.RemoteGrid1.Image
    Call Main1.ShipHit1.Cls
    Let Main1.ShipsName1.Caption = ""
    Let Main1.Destroyed1.Visible = False
    Let Main1.LocalAmount1.Visible = False
    Let Main1.RemoteAmount1.Visible = False
''===============================================================================''
    Call Data1.LocalList1.Clear
    Call Data1.LocalList2.Clear
    Call Data1.RemoteList1.Clear
    For T = 1 To (GridBlocksX * GridBlocksY)
        Call Data1.LocalList1.AddItem("")
        Call Data1.LocalList2.AddItem(Trim$(0))
        Data1.RemoteList1.AddItem (Trim$(0))
    Next T
''===============================================================================''
    For T = Data1.ShipData1.LBound To Data1.ShipData1.UBound
        Call Data1.ShipData1(T).Clear
    Next T
    Call Data1.ShipData2(0).Clear
    For T = Ships1.ShipName1.LBound To Ships1.ShipName1.UBound
        Let Ships1.Direction1(T).Caption = Int(Rnd * 4)
        Call Data1.ShipData1(0).AddItem(Ships1.ShipName1(T).Caption)
        Call Data1.ShipData1(1).AddItem(0)
        Call Data1.ShipData2(0).AddItem(1)
    Next T
''===============================================================================''
    Call EditModeOnOrOff1(True)
    Let EditModeCurrentShip1 = -1
    Let EndOfGame1 = False
''===============================================================================''
    Call NextShipToPosition1
End Sub

Public Sub MovePlayersShips1()
Dim tmpInt1 As Long, tmpInt2 As Long
''===============================================================================''
    Let Main1.Edit1.Picture = Ships1.ShipOutline1(4 * EditModeCurrentShip1 - -(Ships1.Direction1(EditModeCurrentShip1)))
    If Not (Main1.Edit1.Visible) Then Let Main1.Edit1.Visible = True
''===============================================================================''
    If (Main1.Edit1.Width > GridSizeX) Then Let tmpInt1 = Int(((Main1.Edit1.Width - 1) / GridSizeX)) Else Let tmpInt1 = 0
    If (Main1.Edit1.Height > GridSizeY) Then Let tmpInt2 = Int(((Main1.Edit1.Height - 1) / GridSizeY)) Else Let tmpInt2 = 0
''===============================================================================''
    If (LocalGridRefX > (GridBlocksX - tmpInt1)) Then Let tmpInt1 = (GridBlocksX - tmpInt1) Else Let tmpInt1 = LocalGridRefX
    If (LocalGridRefY > (GridBlocksY - tmpInt2)) Then Let tmpInt2 = (GridBlocksY - tmpInt2) Else Let tmpInt2 = LocalGridRefY
''===============================================================================''
    Let EditorGridRefX = tmpInt1
    Let EditorGridRefY = tmpInt2
''===============================================================================''
    Call Main1.Edit1.Move((GridSizeX * EditorGridRefX), (GridSizeY * EditorGridRefY))
End Sub

Public Sub RotatePlayersShips1()
    Let Ships1.Direction1(EditModeCurrentShip1).Caption = (Ships1.Direction1(EditModeCurrentShip1).Caption - -1)
    If (Ships1.Direction1(EditModeCurrentShip1).Caption >= 4) Then Let Ships1.Direction1(EditModeCurrentShip1).Caption = 0
    Let Main1.Edit1.Picture = Ships1.ShipOutline1((4 * EditModeCurrentShip1) - -(Ships1.Direction1(EditModeCurrentShip1)))
''===============================================================================''
    Call MovePlayersShips1
End Sub

Public Sub PlaceShipOnGrid1()
Dim CurPosX As Long, CurPosY As Long
Dim CurShip1 As Long, CurMask1 As Long
Dim tmpInt1 As Long, tmpInt2 As Long
Dim Getter1 As Long, Getter2 As Long
Dim Mode1 As Long
Dim HitCount1 As Long
Dim Drawn1 As Boolean
Dim PicWidth1 As Long, PicHeight1 As Long
Dim ListAdd1 As String
''===============================================================================''
    Let CurPosX = EditorGridRefX
    Let CurPosY = EditorGridRefY
    Let CurShip1 = EditModeCurrentShip1
    Let CurMask1 = ((4 * EditModeCurrentShip1) - -(Ships1.Direction1(CurShip1).Caption))
    Let PicWidth1 = (((Ships1.ShipMask1(CurMask1).ScaleWidth - 1) / GridSizeX) - 1)
    Let PicHeight1 = (((Ships1.ShipMask1(CurMask1).ScaleHeight - 1) / GridSizeY) - 1)
    Let Drawn1 = False
    Let HitCount1 = 0
''===============================================================================''
    For Mode1 = 1 To 2
        For tmpInt2 = 0 To PicHeight1
            For tmpInt1 = 0 To PicWidth1
                Let Getter1 = Ships1.ShipMask1(CurMask1).Point(((GridSizeX / 2) - -(GridSizeX * tmpInt1)), ((GridSizeY / 2) - -(GridSizeX * tmpInt2)))
                If ((Getter1 >= 0) And (Getter1 <= 10)) Then
                    Let Getter2 = ((GridBlocksY * (CurPosY - -tmpInt2)) - -(CurPosX - -tmpInt1))
''===============================================================================''
                    If (Mode1 = 1) Then
                        If Not (Trim$(Data1.LocalList1.List(Getter2)) = "") Then
                            Call AddAMessage1((0), ("You Cannot Place Over Another Ship !"))
                            Exit Sub
                        Else
                            Let HitCount1 = (HitCount1 - -1)
                        End If
                    ElseIf (Mode1 = 2) Then
                        Let ListAdd1 = (FillSpaces1((EditModeCurrentShip1), (DataSpacing1))) & (FillSpaces1((CurMask1), (DataSpacing1))) & (FillSpaces1((PicWidth1), (DataSpacing1))) & (FillSpaces1((PicHeight1), (DataSpacing1))) & (FillSpaces1((tmpInt1), (DataSpacing1))) & (FillSpaces1((tmpInt2), (DataSpacing1))) & (FillSpaces1((Getter2), (DataSpacing1)))
                        Let Data1.LocalList1.List(Getter2) = (ListAdd1)
                        Let Data1.LocalList2.List(Getter2) = (0)
                        If Not (Drawn1) Then
                            Let Data1.ShipData1(1).List(CurShip1) = (HitCount1)
                            Let Data1.ShipData2(0).List(CurShip1) = (HitCount1)
              
                            Call BitBlt(Main1.LGrid1.hDC, (CurPosX * GridSizeX), (CurPosY * GridSizeY), Ships1.ShipFixed2(CurMask1).ScaleWidth, Ships1.ShipFixed2(CurMask1).ScaleHeight, Ships1.ShipFixed2(CurMask1).hDC, 0, 0, SRCPAINT)
                            Call BitBlt(Main1.LGrid1.hDC, (CurPosX * GridSizeX), (CurPosY * GridSizeY), Ships1.ShipFixed1(CurMask1).ScaleWidth, Ships1.ShipFixed1(CurMask1).ScaleHeight, Ships1.ShipFixed1(CurMask1).hDC, 0, 0, SRCAND)

                            Call Main1.LGrid1.Refresh
                            Let Drawn1 = True
                        End If
                    End If
''===============================================================================''
                End If
            Next tmpInt1
        Next tmpInt2
    Next Mode1
''===============================================================================''
  If (Drawn1) Then Call NextShipToPosition1
End Sub

Public Sub NextShipToPosition1()
    Let EditModeCurrentShip1 = (EditModeCurrentShip1 - -1)
    Let Main1.Edit1.Visible = False
''===============================================================================''
    If (EditModeCurrentShip1 > (Data1.ShipData1(0).ListCount - 1)) Then
        Call EditModeOnOrOff1(False)
        Let Data1.PlayerReady1(0).Caption = 1
        Call SendDataToRemoteUser1("##ReadyToPlay##")
        Call ShowAmountOfPlayerShips1
        Call AddAMessage1((0), ("Good Luck !"))
        If Not (Data1.PlayerReady1(1).Caption = 1) Then Call AddAMessage1((0), ("Waiting For " & Trim(Main1.RemoteName1.Caption) & " To Finish..."))
    Else
        Call AddAMessage1((2), ("Position Your : " & Ships1.ShipName1(EditModeCurrentShip1).Caption & "..."))
    End If
End Sub

Public Function FillSpaces1(Input1 As String, Input2 As Long) As String
Dim tmpString1 As String
Dim tmpInt1 As Long
''===============================================================================''
    Let tmpString1 = Input1
''===============================================================================''
    For tmpInt1 = 1 To (Input2 - (Len(Input1)))
        Let tmpString1 = (" " & tmpString1)
    Next tmpInt1
''===============================================================================''
    Let FillSpaces1 = tmpString1
End Function

Public Sub ChangePlayerNames1(Player1 As String, Player2 As String)
Dim TempPos1 As Long, TempPos2 As Long
''===============================================================================''
    Let Main1.LocalName1.Caption = Trim(Player1)
    Let Main1.RemoteName1.Caption = Trim(Player2)
''===============================================================================''
    Let Main1.LocalName1.AutoSize = False
    Let Main1.RemoteName1.AutoSize = False
    Let Main1.LocalName1.AutoSize = True
    Let Main1.RemoteName1.AutoSize = True
''===============================================================================''
    Let TempPos1 = (((289 - 271) / 2) - (Main1.LocalName1.Height / 2))
    Let TempPos2 = (((293 - 275) / 2) - (Main1.RemoteName1.Height / 2))
''===============================================================================''
    Call Main1.LocalName1.Move((11 - -TempPos1), (271 - -TempPos1))
    Call Main1.RemoteName1.Move((554 - (Main1.RemoteName1.Width - -TempPos2)), (275 - -TempPos2))
''===============================================================================''
    Let Main1.LocalName1.Visible = (True)
    Let Main1.RemoteName1.Visible = (True)
End Sub

'Thraddash Software - Trevor Lewis

