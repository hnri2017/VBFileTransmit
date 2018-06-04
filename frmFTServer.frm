VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Begin VB.Form frmFTServer 
   Caption         =   "FTServer"
   ClientHeight    =   3210
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7035
   Icon            =   "frmFTServer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   7035
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   2400
      Width           =   1095
   End
   Begin FTServer.LabelProgressBar LabelProgressBar1 
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   1440
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox Check1 
      Caption         =   "最小化隐藏"
      Height          =   180
      Left            =   5640
      TabIndex        =   8
      Top             =   120
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开启服务"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   40
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1080
      TabIndex        =   4
      Top             =   80
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2580
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Left            =   4560
      Top             =   600
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   5040
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   5520
      Top             =   600
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label Label1 
      Height          =   180
      Index           =   4
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "服务状态："
      Height          =   180
      Index           =   3
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "侦听端口："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   180
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "当前连接列表："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1260
   End
End
Attribute VB_Name = "frmFTServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const mconstrBar As String = "--"


Private Function mfCloseAllConnect() As Boolean
    Dim ctlSck As MSWinsockLib.Winsock
    
    For Each ctlSck In Winsock1
        If ctlSck.State <> 0 Then
            ctlSck.Close
            gArr(ctlSck.Index) = gArr(0)
            If ctlSck.Index <> 0 Then Unload ctlSck
        End If
    Next
    
    List1.Clear
    Label1.Item(1).Caption = 0
    
End Function

Private Function mfConnect() As Boolean
    With Winsock1.Item(0)
        If Command1.Caption = gVar.ServerStart Then
            Dim strPort As String
            
            strPort = Trim(Text1.Text)
            If Len(strPort) = 0 Then
                strPort = GetSetting(App.Title, gVar.RegTcpSection, gVar.RegTcpKeyPort, gVar.TCPPort)
            End If
            strPort = CStr(CLng(Val(strPort)))
            If Val(strPort) > 65535 Or Val(strPort) < 0 Then strPort = gVar.TCPPort
            If strPort <> Text1.Text Then Text1.Text = strPort
            SaveSetting App.Title, gVar.RegTcpSection, gVar.RegTcpKeyPort, strPort
            
            If .State <> 0 Then .Close
            .LocalPort = strPort
            .Listen
            Command1.Caption = gVar.ServerClose
            Label1.Item(4).Caption = gVar.ServerStarted
            Label1.Item(4).ForeColor = vbBlue
        Else
            .Close
            Command1.Caption = gVar.ServerStart
            Label1.Item(4).Caption = gVar.ServerNotStarted
            Label1.Item(4).ForeColor = vbRed
            Call mfCloseAllConnect
        End If
    End With
End Function


Private Sub Command1_Click()
    Const conInterval As Long = 2
    Static sngLastTime As Single
    Dim sngCurTime As Single
    
    sngCurTime = Timer
    If sngCurTime - sngLastTime < conInterval Then
        MsgBox "两次点击时间间隔小于" & conInterval & "秒！", vbExclamation
        Exit Sub
    End If
    sngLastTime = sngCurTime
    
    Call mfConnect
    
End Sub

Private Sub Command2_Click()
    Call gfLoadSkin(Me, SkinFramework1, sMS07)
End Sub

Private Sub Form_Load()
        
    If App.PrevInstance Then
        MsgBox "服务端已打开！", vbExclamation
        Unload Me
        Exit Sub
    End If
    
    Timer1.Interval = 1000
    Check1.Value = 1
    
    Call gsInitialize
    Call gfNotifyIconAdd(Me)
    Call gfLoadSkin(Me, SkinFramework1, , True)
    Call mfConnect
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '处理托盘图标上鼠标事件
    Dim sngMsg As Single
    
    sngMsg = X / Screen.TwipsPerPixelX
    
    Select Case sngMsg
        Case WM_RBUTTONUP
            
        Case WM_LBUTTONDBLCLK
            With Me
                If .WindowState = vbMinimized Then
                    .WindowState = vbNormal
                    .Show
                    .SetFocus
                Else
                    .WindowState = vbMinimized
                End If
            End With
        Case Else
    End Select
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        If Check1.Value = 1 Then
            Me.Hide
            Call gfNotifyIconBalloon(Me, "最小化到系统托盘图标啦", "提示")
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call gfNotifyIconDelete(Me)
End Sub

Private Sub Timer1_Timer()
    If Winsock1.Item(0).State = 2 Then
        If Command1.Caption <> gVar.ServerClose Then
            Command1.Caption = gVar.ServerClose
            Label1.Item(4).Caption = gVar.ServerStarted
            Label1.Item(4).ForeColor = vbBlue
        End If
    ElseIf Winsock1.Item(0).State = 9 Then
        If Label1.Item(4).Caption <> gVar.ServerError Then
            Command1.Caption = gVar.ServerStart
            Label1.Item(4).Caption = gVar.ServerError
            Label1.Item(4).ForeColor = vbRed
            Call mfCloseAllConnect
        End If
    Else
        If Label1.Item(4).Caption <> gVar.ServerNotStarted Then
            Command1.Caption = gVar.ServerStart
            Label1.Item(4).Caption = gVar.ServerNotStarted
            Label1.Item(4).ForeColor = vbRed
            Call mfCloseAllConnect
        End If
    End If
End Sub

Private Sub Winsock1_Close(Index As Integer)
    Dim K As Long
    
    If Index = 0 Then Exit Sub
    If List1.ListCount = 0 Then Exit Sub
    
    For K = 0 To List1.ListCount - 1
        If (InStr(List1.List(K), Winsock1.Item(Index).RemoteHostIP) > 0) _
            And (InStr(List1.List(K), mconstrBar & Winsock1.Item(Index).Tag & mconstrBar) > 0) Then
            List1.RemoveItem K
            Unload Winsock1.Item(Index)
            gArr(Index) = gArr(0)
            Label1.Item(1).Caption = List1.ListCount
            Exit For
        End If
    Next
    
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim ctlSck As Winsock
    Dim K As Long
    
    If Index <> 0 Then Exit Sub
    
    For Each ctlSck In Winsock1
        If ctlSck.Index = K Then
            K = K + 1
        Else
            Exit For
        End If
    Next
    
    With Winsock1
        If K = .Count Then ReDim Preserve gArr(K)
        gArr(K) = gArr(0)
        
        Load .Item(K)
        .Item(K).Accept requestID
        .Item(K).Tag = requestID
        
        List1.AddItem .Item(K).RemoteHostIP & mconstrBar & CStr(requestID) & mconstrBar & K
        Label1.Item(1).Caption = List1.ListCount
    End With
    
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strGet As String
    Dim byteGet() As Byte
    
    With gArr(Index)
        If Not .FileTransmitState Then
            Winsock1.Item(Index).GetData strGet
            If Not gfRestoreInfo(strGet, Winsock1.Item(Index)) Then
                
            End If
Debug.Print "Server GetInfo:" & strGet, bytesTotal
            
        Else
            If .FileNumber = 0 Then
                .FileNumber = FreeFile
                Open .FilePath For Binary As #.FileNumber
            End If
            
            ReDim byteGet(bytesTotal - 1)
            Winsock1.Item(Index).GetData byteGet, vbArray + vbByte
            Put #.FileNumber, , byteGet
            .FileSizeCompleted = .FileSizeCompleted + bytesTotal
            
            If .FileSizeCompleted >= .FileSizeTotal Then
                Close #.FileNumber
                Call gfSendInfo(gVar.PTFileEnd, Winsock1.Item(Index))
                gArr(Index) = gArr(0)
Debug.Print "Received Over"
            End If
            
        End If
    End With
End Sub
