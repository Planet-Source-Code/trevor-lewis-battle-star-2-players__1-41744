Attribute VB_Name = "Graphics1"
Option Explicit

Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Sub AniGridLines1(GridX As Long, GridY As Long)
Dim T As Long
Dim PosX1 As Long, PosY1 As Long
Dim PosX2 As Long, PosY2 As Long
Dim PosX3 As Long, PosY3 As Long
Dim PosX4 As Long, PosY4 As Long
Const Mover1 = 15
''===============================================================================''
    Let PosX1 = ((GridSizeX / 2) - -(GridSizeX * GridX))
    Let PosY1 = ((GridSizeY / 2) - -(GridSizeY * GridY))
    Let PosX2 = (Main1.PosLineXY1(0).X1)
    Let PosY2 = (Main1.PosLineXY1(1).Y1)
''===============================================================================''
    If (PosX1 >= PosX2) Then
        Let PosX3 = (PosX1 - PosX2): Let PosX4 = 0
    Else
        Let PosX3 = (PosX2 - PosX1): Let PosX4 = 1
    End If
''===============================================================================''
    If (PosY1 >= PosY2) Then
        Let PosY3 = (PosY1 - PosY2): Let PosY4 = 0
    Else
        Let PosY3 = (PosY2 - PosY1): Let PosY4 = 1
    End If
''===============================================================================''
    For T = 0 To Mover1
''===============================================================================''
        If (PosX4 = 0) Then
            With Main1.PosLineXY1(0)
                Let .X1 = (PosX2 - -((PosX3 / Mover1) * T))
                Let .X2 = (.X1)
                Let .Y1 = (1)
                Let .Y2 = (Main1.LGrid1.Height - 1)
            End With
        ElseIf (PosX4 = 1) Then
            With Main1.PosLineXY1(0)
                Let .X1 = (PosX2 - ((PosX3 / Mover1) * T))
                Let .X2 = (.X1)
                Let .Y1 = (1)
                Let .Y2 = (Main1.LGrid1.Height - 1)
            End With
        End If
''===============================================================================''
        If (PosY4 = 0) Then
            With Main1.PosLineXY1(1)
                Let .X1 = (1)
                Let .X2 = (Main1.LGrid1.Width - 1)
                Let .Y1 = (PosY2 - -((PosY3 / Mover1) * T))
                Let .Y2 = (.Y1)
            End With
        ElseIf (PosY4 = 1) Then
            With Main1.PosLineXY1(1)
                Let .X1 = (1)
                Let .X2 = (Main1.LGrid1.Width - 1)
                Let .Y1 = (PosY2 - ((PosY3 / Mover1) * T))
                Let .Y2 = (.Y1)
            End With
        End If
''===============================================================================''
        If Not (Main1.PosLineXY1(0).Visible) Then
            Main1.PosLineXY1(0).Visible = True
            Main1.PosLineXY1(1).Visible = True
        End If
        Call Project_Pause1(0.01)
    Next T
End Sub

Public Sub AniExplosion1(GridX As Long, GridY As Long)
Dim T As Long
''===============================================================================''
    Let GridX = (GridSizeX * GridX)
    Let GridY = (GridSizeY * GridY)
''===============================================================================''
    Call Main1.Explosion1.Move((GridX), (GridY))
    For T = Data1.Explosion1.LBound To Data1.Explosion1.UBound
        Let Main1.Explosion1.Picture = Data1.Explosion1(T).Picture
        If (T = 0) Then
            With Main1.Explosion1
                Let .Stretch = False
                Let .Stretch = True
                Let .Width = (GridSizeX - -1)
                Let .Height = (GridSizeY - -1)
                Let .Visible = True
            End With
        End If
        Call Project_Pause1(0.1)
        DoEvents
    Next T
    Let Main1.Explosion1.Visible = False
End Sub

Public Function DestroyBlock1(GridX As Long, GridY As Long) As String
Dim tmpInt1 As Long
Dim GridList1 As Long
Dim ListInfo1 As String
Dim PicData1(6) As Long
Dim MissHit1 As Boolean
Dim ExColor1 As Long
Dim SendMsg1 As String
''===============================================================================''
    Call AniGridLines1((GridX), (GridY))
    Call Project_Pause1(0.2)
''===============================================================================''
    Let GridList1 = ((GridBlocksY * GridY) - -(GridX))
    Let ListInfo1 = Data1.LocalList1.List(GridList1)
''===============================================================================''
    If (Trim$(ListInfo1) = "") Then
        Let MissHit1 = False
    Else
        Let MissHit1 = True
        For tmpInt1 = 0 To UBound(PicData1)
            Let PicData1(tmpInt1) = -1
            Let PicData1(tmpInt1) = CLng(Trim(Right(Left((ListInfo1), (DataSpacing1 * (tmpInt1 - -1))), DataSpacing1)))
        Next tmpInt1
    End If
''===============================================================================''
    Call Main1.LGrid1.PaintPicture(Data1.LocalGrid1.Image, (GridSizeX * GridX), (GridSizeY * GridY), (GridSizeX - -1), (GridSizeY - -1), (GridSizeX * GridX), (GridSizeY * GridY), (GridSizeX - -1), (GridSizeY - -1))
''===============================================================================''
    Let Data1.HitPicture1.Width = ((GridSizeX - -1) * Exx)
    Let Data1.HitPicture1.Height = ((GridSizeY - -1) * Why)
    Let Data1.HitPicture2.Width = ((GridSizeX - -1) * Exx)
    Let Data1.HitPicture2.Height = ((GridSizeY - -1) * Why)
    If (MissHit1) Then
        Call Data1.HitPicture1.Cls
        Let ExColor1 = Ships1.ShipDamaged1(PicData1(1)).Point(0, 0)
        Call Data1.HitPicture1.PaintPicture(Ships1.ShipDamaged1(PicData1(1)).Image, 0, 0, (GridSizeX - -1), (GridSizeY - -1), (GridSizeX * PicData1(4)), (GridSizeX * PicData1(5)), (GridSizeX - -1), (GridSizeY - -1))
        Call Data1.HitPicture2.PaintPicture(Ships1.ShipDamaged2(PicData1(1)).Image, 0, 0, (GridSizeX - -1), (GridSizeY - -1), (GridSizeX * PicData1(4)), (GridSizeX * PicData1(5)), (GridSizeX - -1), (GridSizeY - -1))
        Call BitBlt(Main1.LGrid1.hDC, (GridX * GridSizeX), (GridY * GridSizeY), Data1.HitPicture2.ScaleWidth, Data1.HitPicture2.ScaleHeight, Data1.HitPicture2.hDC, 0, 0, SRCPAINT)
        Call BitBlt(Main1.LGrid1.hDC, (GridX * GridSizeX), (GridY * GridSizeY), Data1.HitPicture1.ScaleWidth, Data1.HitPicture1.ScaleHeight, Data1.HitPicture1.hDC, 0, 0, SRCAND)
        Call Main1.LGrid1.Refresh
        If (Data1.LocalList2.List(PicData1(6)) = 0) Then
            Let Data1.ShipData1(1).List(PicData1(0)) = (Data1.ShipData1(1).List(PicData1(0)) - 1)
            Let Data1.LocalList2.List(PicData1(6)) = (1)
        End If
        If (Data1.ShipData1(1).List(PicData1(0)) = 0) Then
            Call AddAMessage1((6), ("Your " & Data1.ShipData1(0).List(PicData1(0)) & " Was Destroyed !"))
        ElseIf (Data1.ShipData1(1).List(PicData1(0)) > 0) Then
            Call AddAMessage1((6), ("Your " & Data1.ShipData1(0).List(PicData1(0)) & " Was Hit !"))
        End If
        Let SendMsg1 = "#HitShip#:" & (FillSpaces1((GridX), (DataSpacing1))) & (FillSpaces1((GridY), (DataSpacing1))) & (FillSpaces1((PicData1(0)), (DataSpacing1))) & (FillSpaces1(Data1.ShipData1(1).List(PicData1(0)), (DataSpacing1)))
        Call ShowAmountOfPlayerShips1
    ElseIf Not (MissHit1) Then
        Call Data1.HitPicture1.Cls
        Let ExColor1 = Data1.RemoteHitMiss1(2).Point(0, 0)
        Call Data1.HitPicture1.PaintPicture(Data1.RemoteHitMiss1(2).Image, 0, 0, (GridSizeX - -1), (GridSizeY - -1), 0, 0, (GridSizeX - -1), (GridSizeY - -1))
        Call Data1.HitPicture2.PaintPicture(Data1.RemoteHitMiss1(3).Image, 0, 0, (GridSizeX - -1), (GridSizeY - -1), 0, 0, (GridSizeX - -1), (GridSizeY - -1))
        Call BitBlt(Main1.LGrid1.hDC, (GridX * GridSizeX), (GridY * GridSizeY), Data1.HitPicture2.ScaleWidth, Data1.HitPicture2.ScaleHeight, Data1.HitPicture2.hDC, 0, 0, SRCPAINT)
        Call BitBlt(Main1.LGrid1.hDC, (GridX * GridSizeX), (GridY * GridSizeY), Data1.HitPicture1.ScaleWidth, Data1.HitPicture1.ScaleHeight, Data1.HitPicture1.hDC, 0, 0, SRCAND)
        Call Main1.LGrid1.Refresh
        Let SendMsg1 = "#MissShip#:" & (FillSpaces1((GridX), (DataSpacing1))) & (FillSpaces1((GridY), (DataSpacing1))) & (FillSpaces1((0), (DataSpacing1))) & (FillSpaces1((0), (DataSpacing1)))
    End If
''===============================================================================''
    Call SendDataToRemoteUser1(SendMsg1)
    Call AniExplosion1((GridX), (GridY))
End Function

Public Sub ShowAmountOfPlayerShips1()
Dim T As Long
Dim Amount1 As Long, Amount2 As Long
Dim ShipNames1 As String, ShipNames2 As String
''===============================================================================''
    Let Amount1 = 0
    Let Amount2 = 0
    Let ShipNames1 = ("")
    Let ShipNames2 = ("")
''===============================================================================''
    For T = 0 To (Data1.ShipData1(1).ListCount - 1)
        If Not (CInt(Data1.ShipData1(1).List(T)) = 0) Then
            If (Amount1 < 5) Then
                If Not (Len(Trim(ShipNames1)) = 0) Then Let ShipNames1 = (ShipNames1 & ", ")
                Let ShipNames1 = (ShipNames1 & Trim$(Data1.ShipData1(0).List(T)))
                Let Amount1 = (Amount1 - -1)
            End If
        End If
        If Not (CInt(Data1.ShipData2(0).List(T)) = 0) Then
            If (Amount2 < 5) Then
                If Not (Len(Trim(ShipNames2)) = 0) Then Let ShipNames2 = (ShipNames2 & ", ")
                Let ShipNames2 = (ShipNames2 & Trim$(Data1.ShipData1(0).List(T)))
                Let Amount2 = (Amount2 - -1)
            End If
        End If
    Next T
''===============================================================================''
    Let Main1.LocalAmount1.Picture = Data1.LocalAmount1(Amount1).Picture
    Let Main1.LocalAmount1.ToolTipText = Trim$(ShipNames1)
    Let Main1.RemoteAmount1.Picture = Data1.RemoteAmount1(Amount2).Picture
    Let Main1.RemoteAmount1.ToolTipText = Trim$(ShipNames2)
''===============================================================================''
    If (Amount1 = 0) Then
        Call AddAMessage1((5), ("You Have Lost This Game !"))
        Let Data1.Menu1_Sub2.Enabled = True
        Let EndOfGame1 = True
        Call DisEnableRemoteGrid1(False)
        Let Data1.RVictory1.Value = (Data1.RVictory1.Value - -1)
    ElseIf (Amount2 = 0) Then
        Call AddAMessage1((5), ("You Win !"))
        Let Data1.Menu1_Sub2.Enabled = True
        Let EndOfGame1 = True
        Call DisEnableRemoteGrid1(False)
        Let Data1.LVictory1.Value = (Data1.LVictory1.Value - -1)
        Call ShowRemotePlayerShipPieces1
    End If
''===============================================================================''
    If Not (Main1.LocalAmount1.Visible) Then
        Let Main1.LocalAmount1.Visible = True
        Let Main1.RemoteAmount1.Visible = True
    End If
End Sub

Public Sub DrawShipRotations1()
Dim T As Long, S As Long
Dim Width1 As Long, Height1 As Long
Dim PosX1 As Long, PosY1 As Long
Dim Color0(1) As Long
Dim Color1(2) As Long
Dim Color2(1) As Long
''===============================================================================''
    For T = Ships1.ShipFixed1.LBound To Ships1.ShipFixed1.UBound Step 4
        Let Width1 = Ships1.ShipFixed1(T).ScaleWidth
        Let Height1 = Ships1.ShipFixed1(T).ScaleHeight
        For S = 1 To 3
            Call Load(Ships1.ShipMask1(T - -S))
            Call Load(Ships1.ShipFixed1(T - -S))
            Call Load(Ships1.ShipFixed2(T - -S))
            Call Load(Ships1.ShipDamaged1(T - -S))
            Call Load(Ships1.ShipDamaged2(T - -S))
        Next S
''===============================================================================''
        For S = 0 To 3
            If ((S Mod 2) = 1) Then
                Let Ships1.ShipFixed1(T - -S).Width = (Height1 * Exx)
                Let Ships1.ShipFixed1(T - -S).Height = (Width1 * Why)
                Let Ships1.ShipFixed2(T - -S).Width = (Height1 * Exx)
                Let Ships1.ShipFixed2(T - -S).Height = (Width1 * Why)
                Let Ships1.ShipDamaged1(T - -S).Width = (Height1 * Exx)
                Let Ships1.ShipDamaged1(T - -S).Height = (Width1 * Why)
                Let Ships1.ShipDamaged2(T - -S).Width = (Height1 * Exx)
                Let Ships1.ShipDamaged2(T - -S).Height = (Width1 * Why)
                Let Ships1.ShipMask1(T - -S).Width = (Height1 * Exx)
                Let Ships1.ShipMask1(T - -S).Height = (Width1 * Why)
            Else
                Let Ships1.ShipFixed1(T - -S).Width = (Width1 * Exx)
                Let Ships1.ShipFixed1(T - -S).Height = (Height1 * Why)
                Let Ships1.ShipFixed2(T - -S).Width = (Width1 * Exx)
                Let Ships1.ShipFixed2(T - -S).Height = (Height1 * Why)
                Let Ships1.ShipDamaged1(T - -S).Width = (Width1 * Exx)
                Let Ships1.ShipDamaged1(T - -S).Height = (Height1 * Why)
                Let Ships1.ShipDamaged2(T - -S).Width = (Width1 * Exx)
                Let Ships1.ShipDamaged2(T - -S).Height = (Height1 * Why)
                Let Ships1.ShipMask1(T - -S).Width = (Width1 * Exx)
                Let Ships1.ShipMask1(T - -S).Height = (Height1 * Why)
            End If
        Next S
''===============================================================================''
        Let Color0(0) = GetPixel((Ships1.ShipFixed1(T).hDC), (0), (0))
        Let Color0(1) = GetPixel((Ships1.ShipDamaged1(T).hDC), (0), (0))
        For PosX1 = 0 To Width1
            For PosY1 = 0 To Height1
        
                Let Color1(0) = GetPixel((Ships1.ShipFixed1(T).hDC), (PosX1), (PosY1))
                If (Color0(0) = Color1(0)) Then
                    Let Color1(0) = QBColor(15)
                    Let Color2(0) = QBColor(0)
                Else
                    Let Color2(0) = QBColor(15)
                End If
                Call SetPixel((Ships1.ShipFixed1(T).hDC), (PosX1), (PosY1), (Color1(0)))
                Call SetPixel((Ships1.ShipFixed1(T - -1).hDC), (PosY1), ((Width1 - 1) - PosX1), (Color1(0)))
                Call SetPixel((Ships1.ShipFixed1(T - -2).hDC), ((Width1 - 1) - PosX1), ((Height1 - 1) - PosY1), (Color1(0)))
                Call SetPixel((Ships1.ShipFixed1(T - -3).hDC), ((Height1 - 1) - PosY1), (PosX1), (Color1(0)))
                Call SetPixel((Ships1.ShipFixed2(T).hDC), (PosX1), (PosY1), (Color2(0)))
                Call SetPixel((Ships1.ShipFixed2(T - -1).hDC), (PosY1), ((Width1 - 1) - PosX1), (Color2(0)))
                Call SetPixel((Ships1.ShipFixed2(T - -2).hDC), ((Width1 - 1) - PosX1), ((Height1 - 1) - PosY1), (Color2(0)))
                Call SetPixel((Ships1.ShipFixed2(T - -3).hDC), ((Height1 - 1) - PosY1), (PosX1), (Color2(0)))
        
                Let Color1(1) = GetPixel((Ships1.ShipDamaged1(T).hDC), (PosX1), (PosY1))
                If (Color0(1) = Color1(1)) Then
                    Let Color1(1) = QBColor(15)
                    Let Color2(1) = QBColor(0)
                Else
                    Let Color2(1) = QBColor(15)
                End If
                Call SetPixel((Ships1.ShipDamaged1(T).hDC), (PosX1), (PosY1), (Color1(1)))
                Call SetPixel((Ships1.ShipDamaged1(T - -1).hDC), (PosY1), ((Width1 - 1) - PosX1), (Color1(1)))
                Call SetPixel((Ships1.ShipDamaged1(T - -2).hDC), ((Width1 - 1) - PosX1), ((Height1 - 1) - PosY1), (Color1(1)))
                Call SetPixel((Ships1.ShipDamaged1(T - -3).hDC), ((Height1 - 1) - PosY1), (PosX1), (Color1(1)))
                Call SetPixel((Ships1.ShipDamaged2(T).hDC), (PosX1), (PosY1), (Color2(1)))
                Call SetPixel((Ships1.ShipDamaged2(T - -1).hDC), (PosY1), ((Width1 - 1) - PosX1), (Color2(1)))
                Call SetPixel((Ships1.ShipDamaged2(T - -2).hDC), ((Width1 - 1) - PosX1), ((Height1 - 1) - PosY1), (Color2(1)))
                Call SetPixel((Ships1.ShipDamaged2(T - -3).hDC), ((Height1 - 1) - PosY1), (PosX1), (Color2(1)))
      
                Let Color1(2) = GetPixel((Ships1.ShipMask1(T).hDC), (PosX1), (PosY1))
                Call SetPixel((Ships1.ShipMask1(T - -1).hDC), (PosY1), ((Width1 - 1) - PosX1), (Color1(2)))
                Call SetPixel((Ships1.ShipMask1(T - -2).hDC), ((Width1 - 1) - PosX1), ((Height1 - 1) - PosY1), (Color1(2)))
                Call SetPixel((Ships1.ShipMask1(T - -3).hDC), ((Height1 - 1) - PosY1), (PosX1), (Color1(2)))
      
            Next PosY1
        Next PosX1
''===============================================================================''
        DoEvents
    Next T
''===============================================================================''
End Sub

Public Sub ProcessHitMiss1(Input1 As String, Input2 As Long)
Dim tmpInt1 As Long
Dim PicData1(3) As Long
''===============================================================================''
    For tmpInt1 = 0 To UBound(PicData1)
        Let PicData1(tmpInt1) = -1
        Let PicData1(tmpInt1) = CLng(Trim$(Right$(Left$((Input1), (DataSpacing1 * (tmpInt1 - -1))), DataSpacing1)))
    Next tmpInt1
''===============================================================================''
    If (Input2 = 0) Then
        Call DrawHitMissBlocks1((PicData1(0)), (PicData1(1)), (0))
    ElseIf (Input2 = 1) Then
        Call DrawHitMissBlocks1((PicData1(0)), (PicData1(1)), (1))
    End If
''===============================================================================''
    If (Input2 = 0) Then
        Let Data1.ShipData2(0).List(PicData1(2)) = (PicData1(3))
        If (PicData1(3) = 0) Then
            Call Main1.ShipHit1.PaintPicture(Ships1.Portrait1(PicData1(2)).Image, (0), (0), Ships1.Portrait1(PicData1(2)).ScaleWidth, Ships1.Portrait1(PicData1(2)).ScaleHeight, 0, 0, Main1.ShipHit1.ScaleWidth, Main1.ShipHit1.ScaleHeight)
            Let Main1.Destroyed1.Visible = True
            Let Main1.ShipsName1.Caption = (Data1.ShipData1(0).List(PicData1(2)) & "_")
            Call AddAMessage1((1), ("You Destroyed " & Trim(Main1.RemoteName1.Caption) & "'s " & Data1.ShipData1(0).List(PicData1(2)) & " !"))
        ElseIf Not (PicData1(3) = 0) Then
            Call Main1.ShipHit1.PaintPicture(Ships1.Portrait1(PicData1(2)).Image, (0), (0), Ships1.Portrait1(PicData1(2)).ScaleWidth, Ships1.Portrait1(PicData1(2)).ScaleHeight, 0, 0, Main1.ShipHit1.ScaleWidth, Main1.ShipHit1.ScaleHeight)
            Let Main1.Destroyed1.Visible = False
            Let Main1.ShipsName1.Caption = (Data1.ShipData1(0).List(PicData1(2)) & "_")
            Call AddAMessage1((1), ("You Hit " & Trim$(Main1.RemoteName1.Caption) & "'s " & Data1.ShipData1(0).List(PicData1(2)) & " !"))
        End If
        Call ShowAmountOfPlayerShips1
    ElseIf (Input2 = 1) Then
        Call AddAMessage1((4), ("You Missed !"))
    End If
End Sub

Public Sub DrawHitMissBlocks1(GridX As Long, GridY As Long, HitMiss1 As Long)
    Let Data1.HitPicture1.Width = ((GridSizeX - -1) * Exx)
    Let Data1.HitPicture1.Height = ((GridSizeY - -1) * Why)
    Let Data1.HitPicture2.Width = ((GridSizeX - -1) * Exx)
    Let Data1.HitPicture2.Height = ((GridSizeY - -1) * Why)
''===============================================================================''
    Call Data1.HitPicture1.PaintPicture(Data1.RemoteHitMiss1(HitMiss1 * 2).Image, 0, 0, (GridSizeX - -1), (GridSizeY - -1), 0, 0, (GridSizeX - -1), (GridSizeY - -1))
    Call Data1.HitPicture2.PaintPicture(Data1.RemoteHitMiss1((HitMiss1 * 2) - -1).Image, 0, 0, (GridSizeX - -1), (GridSizeY - -1), 0, 0, (GridSizeX - -1), (GridSizeY - -1))
''===============================================================================''
    Call BitBlt(Main1.RGrid1.hDC, (GridX * GridSizeX), (GridY * GridSizeY), Data1.HitPicture2.ScaleWidth, Data1.HitPicture2.ScaleHeight, Data1.HitPicture2.hDC, 0, 0, SRCPAINT)
    Call BitBlt(Main1.RGrid1.hDC, (GridX * GridSizeX), (GridY * GridSizeY), Data1.HitPicture1.ScaleWidth, Data1.HitPicture1.ScaleHeight, Data1.HitPicture1.hDC, 0, 0, SRCAND)
''===============================================================================''
    Call Main1.RGrid1.Refresh
End Sub

Public Sub ShowRemotePlayerShipPieces1()
Dim GridX As Long, GridY As Long
Dim GridList1 As Long
Dim ListInfo1 As String
''===============================================================================''
    For GridX = 0 To (GridBlocksX - 1)
        For GridY = 0 To (GridBlocksY - 1)
            If (EndOfGame1) Then
                Let GridList1 = ((GridBlocksY * GridY) - -(GridX))
                Let ListInfo1 = Data1.LocalList1.List(GridList1)
                If (Not (Len(Trim(ListInfo1)) = 0)) Then
                    If (Not (Data1.LocalList2.List(GridList1) = 1)) Then
                        Call SendDataToRemoteUser1("#ShipPiece#:" & (FillSpaces1((GridX), (DataSpacing1))) & (FillSpaces1((GridY), (DataSpacing1))))
                    End If
                End If
            End If
        Next GridY
    Next GridX
End Sub

'Thraddash Software - Trevor Lewis

