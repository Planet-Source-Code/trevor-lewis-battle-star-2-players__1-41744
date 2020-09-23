Attribute VB_Name = "Connect1"
Option Explicit

Public SocketSendData1 As Boolean
Public DataResend1  As Long
Public StartStop1 As Boolean
Public ConnectionType1 As Long

Public Sub StartStopClient_Host1()
    Let StartStop1 = Not (StartStop1)
''===============================================================================''
    If (StartStop1) Then
        Let Connection1.StartStopIt1.Caption = "&Stop"
        Let Connection1.PlayerNText1.Enabled = False
        Call Connection1.Outbox1.Clear
''===============================================================================''
        If (ConnectionType1 = 1) Then
            Let Data1.HostClient1.Caption = 1
            Let Connection1.SMCheck1.Enabled = False
            Let Connection1.ClientText1(0).Enabled = False
            Let Connection1.ClientText1(1).Enabled = False
            Call Connection1.Con1.Close
            Let Connection1.Con1.LocalPort = 0
            Let Connection1.Con1.RemotePort = Connection1.ClientText1(1).Text
            Let Connection1.Con1.RemoteHost = Connection1.ClientText1(0).Text
            Let Connection1.StatusText1.Caption = "Connecting..."
            Call Connection1.Con1.Connect
        ElseIf (ConnectionType1 = 2) Then
            Let Data1.HostClient1.Caption = 0
            Let Connection1.SMCheck1.Enabled = False
            Let Connection1.HostText1(0).Enabled = False
            Let Connection1.HostText1(1).Enabled = False
            Call Connection1.Con1.Close
            Let Connection1.Con1.LocalPort = Connection1.HostText1(1).Text
            Let Connection1.Con1.RemotePort = 0
            Let Connection1.Con1.RemoteHost = ""
            Let Connection1.StatusText1.Caption = "Waiting For Connection..."
            Call Connection1.Con1.Listen
        End If
''===============================================================================''
    ElseIf Not (StartStop1) Then
        Let Data1.HostClient1.Caption = -1
        Let Connection1.StartStopIt1.Caption = "&Start"
        Let Connection1.StartStopIt1.Enabled = True
        Let Connection1.PlayerNText1.Enabled = True
        Call Connection1.Con1.Close
''===============================================================================''
        If (ConnectionType1 = 1) Then
            Let Connection1.SMCheck1.Enabled = True
            Let Connection1.ClientText1(0).Enabled = True
            Let Connection1.ClientText1(1).Enabled = True
        ElseIf (ConnectionType1 = 2) Then
            Let Connection1.SMCheck1.Enabled = True
            Let Connection1.HostText1(0).Enabled = True
            Let Connection1.HostText1(1).Enabled = True
            Let Connection1.StatusText1.Caption = "Not Connected."
        End If
  
        If Not (Connection1.Visible) Then
            Let Connection1.StatusText1.Caption = "Remote User Disconnected."
            Call Unload(Main1)
        End If
''===============================================================================''
    End If
End Sub

Public Sub Connection1Watcher1Loop1()
    Call ProcessReceivedData1
    If (Connection1.Con1.State = 7) Then
    
        If (Connection1.StartStopIt1.Enabled) Then
            Let Connection1.StatusText1.Caption = "Connected."
            Let Connection1.StartStopIt1.Enabled = False
''===============================================================================''
            Call Load(Main1)
            Let Connection1.Visible = False
            Call SendDataToRemoteUser1("##SendMeYourName##")
            Call Main1.Show
            If (Data1.HostClient1.Caption = 0) Then
                Call SendDataToRemoteUser1("##StartANewGame##")
                Call SendDataToLocalUser1("##StartANewGame##")
            End If
''===============================================================================''
        End If
''===============================================================================''
        If (Data1.HostClient1.Caption = 0) Then
            If ((Data1.PlayerReady1(0).Caption = 1) And (Data1.PlayerReady1(1).Caption = 1)) Then
                Let Data1.PlayerReady1(0).Caption = 2
                Let Data1.PlayerReady1(1).Caption = 2
      
                If (Int(Rnd * 2) = 1) Then
                    Call SendDataToRemoteUser1("##YouGoFirst##")
                    Call SendDataToLocalUser1("##YouGoSecond##")
                Else
                    Call SendDataToLocalUser1("##YouGoFirst##")
                    Call SendDataToRemoteUser1("##YouGoSecond##")
                End If
            End If
        End If
''===============================================================================''
        If (SocketSendData1) Then
            Let SocketSendData1 = False
            If (Connection1.Outbox1.ListCount = 0) Then Connection1.Outbox1.AddItem ("@")
            If (Trim$(Connection1.Outbox1.List(0)) = "") Then
                Let Connection1.Outbox1.List(0) = "@"
            End If
            Call Connection1.Con1.SendData(Connection1.Outbox1.List(0))
            Let DataResend1 = (Timer - -10)
        ElseIf (Not (SocketSendData1)) And (Timer >= DataResend1) Then
            Call Connection1.Outbox1.AddItem("@", 0)
            Let SocketSendData1 = True
        End If
    ElseIf (Connection1.Con1.State = 8) Then
        Let Connection1.StatusText1.Caption = "Disconnected."
        Call Connection1.Con1.Close
        Call StartStopClient_Host1
    ElseIf (Connection1.Con1.State = 9) Then
        Let Connection1.StatusText1.Caption = "Could Not Connect."
        Call Connection1.Con1.Close
        Call StartStopClient_Host1
    End If
End Sub

Public Sub Load_Connection1()
Dim T As Long
Dim Size1 As Long
    Let Exx = Screen.TwipsPerPixelX
    Let Why = Screen.TwipsPerPixelY
''===============================================================================''
    Let Connection1.Watcher1.Interval = 1
    Let StartStop1 = False
    Let Connection1.SMCheck1.Value = 1
    Let Connection1.SMCheck1.Value = 0
    Let Connection1.HostText1(0).Text = Trim(Connection1.Con1.LocalIP)
''===============================================================================''
    Let Connection1.TSPanel1.Width = (220 * Exx)
''===============================================================================''
    Call Connection1.PlayerName1.Move((5 * Exx), (8 * Why))
    Let Connection1.PlayerName1.Width = ((Connection1.TSPanel1.Width) - (Connection1.PlayerName1.Left * 2))
    Let Connection1.PlayerNText1.Height = 0
    Call Connection1.PlayerNText1.Move((14 * Exx), (23 * Why), (Connection1.PlayerName1.Width - (28 * Exx)))
    Let Connection1.PlayerName1.Height = (Connection1.PlayerNText1.Top - -Connection1.PlayerNText1.Height - -Connection1.PlayerNText1.Left)
''===============================================================================''
    Call Connection1.StatusMenu1.Move((Connection1.PlayerName1.Left), (Connection1.PlayerName1.Top - -Connection1.PlayerName1.Height - -(7 * Why)), (Connection1.PlayerName1.Width))
    Call Connection1.SMCheck1.Move((10 * Exx), (20 * Why), (Connection1.StatusMenu1.Width - (20 * Exx)), 0)
    Call Connection1.ClientHost1(1).Move((10 * Exx), (Connection1.SMCheck1.Top - -Connection1.SMCheck1.Height - -(5 * Why)))
    Let Connection1.ClientHost1(1).Width = (Connection1.StatusMenu1.Width - (Connection1.ClientHost1(1).Left * 2))
''===============================================================================''
    Let Size1 = 0
    For T = 0 To 1
        Let Connection1.ClientLabel1(T).AutoSize = False
        Let Connection1.ClientLabel1(T).AutoSize = True
        Let Connection1.ClientText1(T).Height = 0
        Let Connection1.HostLabel1(T).AutoSize = False
        Let Connection1.HostLabel1(T).AutoSize = True
        Let Connection1.HostText1(T).Height = 0
    If (T = 0) Then
        Let Connection1.ClientText1(T).Top = (9 * Why)
        Let Connection1.HostText1(T).Top = (9 * Why)
    Else
        Let Connection1.ClientText1(T).Top = (Connection1.ClientText1(T - 1).Top - -Connection1.ClientText1(T - 1).Height - -(6 * Why))
        Let Connection1.HostText1(T).Top = (Connection1.HostText1(T - 1).Top - -Connection1.HostText1(T - 1).Height - -(6 * Why))
    End If
        Call Connection1.ClientLabel1(T).Move((9 * Exx), (Connection1.ClientText1(T).Top - -(Connection1.ClientText1(T).Height / 2) - (Connection1.ClientLabel1(T).Height / 2)))
        Call Connection1.HostLabel1(T).Move((9 * Exx), (Connection1.HostText1(T).Top - -(Connection1.HostText1(T).Height / 2) - (Connection1.HostLabel1(T).Height / 2)))
  
        If (Size1 < (Connection1.ClientLabel1(T).Left - -Connection1.ClientLabel1(T).Width)) Then Let Size1 = (Connection1.ClientLabel1(T).Left - -Connection1.ClientLabel1(T).Width)
        If (Size1 < (Connection1.HostLabel1(T).Left - -Connection1.HostLabel1(T).Width)) Then Let Size1 = (Connection1.HostLabel1(T).Left - -Connection1.HostLabel1(T).Width)
    Next T
''===============================================================================''
    For T = 0 To 1
        Let Connection1.ClientLabel1(T).Left = (Size1 - Connection1.ClientLabel1(T).Width)
        Let Connection1.HostLabel1(T).Left = (Size1 - Connection1.HostLabel1(T).Width)
        Let Connection1.ClientText1(T).Left = (Size1 - -(7 * Exx))
        Let Connection1.HostText1(T).Left = (Size1 - -(7 * Exx))
        Let Connection1.ClientText1(T).Width = Connection1.ClientHost1(1).Width - ((Connection1.ClientText1(0).Top) - -Connection1.ClientText1(T).Left - -(1 * Exx))
        Let Connection1.HostText1(T).Width = Connection1.ClientHost1(1).Width - ((Connection1.HostText1(0).Top) - -Connection1.HostText1(T).Left - -(1 * Exx))
    Next T
''===============================================================================''
    Let Connection1.ClientHost1(1).Height = (Connection1.ClientText1(1).Top - -Connection1.ClientText1(1).Height - -Connection1.ClientText1(0).Top - -(2 * Why))
    Call Connection1.ClientHost1(2).Move((Connection1.ClientHost1(1).Left), (Connection1.ClientHost1(1).Top), (Connection1.ClientHost1(1).Width), (Connection1.ClientHost1(1).Height))
''===============================================================================''
    Call Connection1.Status1.Move((Connection1.ClientHost1(1).Left), (Connection1.ClientHost1(1).Top - -Connection1.ClientHost1(1).Height - -(6 * Why)), (Connection1.ClientHost1(1).Width), (Connection1.PlayerName1.Height))
    Let Connection1.StatusText1.AutoSize = False
    Let Connection1.StatusText1.AutoSize = True
    Call Connection1.StatusText1.Move((Connection1.PlayerNText1.Left), (Connection1.PlayerNText1.Top - -(Connection1.PlayerNText1.Height / 2) - (Connection1.StatusText1.Height / 2)))
''===============================================================================''
    Let Connection1.StatusMenu1.Height = (Connection1.Status1.Top - -Connection1.Status1.Height - -Connection1.Status1.Left)
''===============================================================================''
    Let Connection1.StartStopIt1.Width = (Connection1.Status1.Width / 2.5)
    Let Connection1.StartStopIt1.Height = (27 * Why)
    Call Connection1.StartStopIt1.Move(((Connection1.StatusMenu1.Left - -Connection1.StatusMenu1.Width) - Connection1.StartStopIt1.Width), (Connection1.StatusMenu1.Top - -Connection1.StatusMenu1.Height - -(6 * Why)))
''===============================================================================''
    Let Connection1.TSPanel1.Height = (Connection1.StartStopIt1.Top - -Connection1.StartStopIt1.Height - -Connection1.PlayerName1.Left - -(2 * Why))
    Call Connection1.TSPanel1.Move((5 * Exx), (5 * Why))
''===============================================================================''
    Let Connection1.Width = (((Connection1.TSPanel1.Left * 2) - -Connection1.TSPanel1.Width) - -(Connection1.Width - Connection1.ScaleWidth))
    Let Connection1.Height = (((Connection1.TSPanel1.Top * 2) - -Connection1.TSPanel1.Height) - -(Connection1.Height - Connection1.ScaleHeight))
End Sub

'Thraddash Software - Trevor Lewis

