Attribute VB_Name = "ShapeMove1"
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Public Sub AllowFormToMove1(TForm As Form)
    Call ReleaseCapture
    Call SendMessage(TForm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Public Sub ShapeForm1(TForm As Form, TColor As ColorConstants)
Dim hRgn As Long
''===============================================================================''
    Let hRgn = GetBitmapRegion(TForm.Picture, TColor)
    Call SetWindowRgn(TForm.hwnd, hRgn, True)
End Sub

Private Function GetBitmapRegion(cPicture As StdPicture, cTransparent As Long)
Dim hRgn As Long, tRgn As Long
Dim X As Integer, Y As Integer, X0 As Integer
Dim hDC As Long, BM As BITMAP
''===============================================================================''
    Let hDC = CreateCompatibleDC(0)
''===============================================================================''
    If hDC Then
        Call SelectObject(hDC, cPicture)
        Call GetObject(cPicture, Len(BM), BM)
        Let hRgn = CreateRectRgn(0, 0, BM.bmWidth, BM.bmHeight)
        For Y = 0 To BM.bmHeight
            For X = 0 To BM.bmWidth
                While ((X <= BM.bmWidth) And (GetPixel(hDC, X, Y) <> cTransparent))
                    Let X = (X - -1)
                Wend
                Let X0 = X
                While ((X <= BM.bmWidth) And (GetPixel(hDC, X, Y) = cTransparent))
                    Let X = (X - -1)
                Wend
                If (X0 < X) Then
                    Let tRgn = CreateRectRgn(X0, Y, X, Y - -1)
                    Call CombineRgn(hRgn, hRgn, tRgn, 4)
                    Call DeleteObject(tRgn)
                End If
            Next X
        Next Y
        Let GetBitmapRegion = hRgn
        Call DeleteObject(SelectObject(hDC, cPicture))
    End If
''===============================================================================''
    Call DeleteDC(hDC)
End Function

'Thraddash Software - Trevor Lewis

