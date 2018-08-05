VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "ftcskfm.ocx"
Begin VB.Form frmFTClient 
   Caption         =   "FTClient"
   ClientHeight    =   2865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6750
   Icon            =   "frmFTClient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   6750
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command4 
      Caption         =   "接收"
      Height          =   400
      Left            =   2280
      TabIndex        =   12
      Top             =   2280
      Width           =   800
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Index           =   1
      Left            =   3480
      TabIndex        =   11
      Top             =   40
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Index           =   0
      Left            =   960
      TabIndex        =   9
      Top             =   40
      Width           =   1600
   End
   Begin VB.CommandButton Command3 
      Caption         =   "发送"
      Height          =   400
      Left            =   2280
      TabIndex        =   7
      Top             =   1200
      Width           =   800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打开…"
      Height          =   400
      Left            =   5400
      TabIndex        =   6
      Top             =   750
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   600
      TabIndex        =   4
      Top             =   800
      Width           =   4815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   1
      Left            =   4920
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   6120
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "建立连接"
      Height          =   400
      Left            =   2640
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin FTClient.LabelProgressBar LabelProgressBar1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   5055
      _ExtentX        =   8916
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
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   4440
      Top             =   360
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "连接端口："
      Height          =   180
      Index           =   4
      Left            =   2640
      TabIndex        =   10
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "服务器IP："
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "文件："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "未连接"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "服务器连接状态："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1440
   End
End
Attribute VB_Name = "frmFTClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Function mfConnect() As Boolean
    With Winsock1.Item(1)
        If Command1.Caption = gVar.ConnectToServer Then
            Dim strIP As String, strPort As String
            
            strIP = Trim(Text2.Item(0).Text)
            If Len(strIP) = 0 Then
                strIP = GetSetting(gVar.RegAppName, gVar.RegTcpSection, gVar.RegTcpKeyIP, gVar.TCPIP)
            End If
            strIP = gfCheckIP(strIP)
            If strIP <> Text2.Item(0).Text Then Text2.Item(0).Text = strIP
            SaveSetting gVar.RegAppName, gVar.RegTcpSection, gVar.RegTcpKeyIP, strIP
            
            strPort = Trim(Text2.Item(1).Text)
            If Len(strPort) = 0 Then
                strPort = GetSetting(gVar.RegAppName, gVar.RegTcpSection, gVar.RegTcpKeyPort, gVar.TCPPort)
            End If
            strPort = CStr(CLng(Val(strPort)))
            If Val(strPort) > 65535 Or Val(strPort) < 0 Then strPort = gVar.TCPPort
            If strPort <> Text2.Item(1).Text Then Text2.Item(1).Text = strPort
            SaveSetting gVar.RegAppName, gVar.RegTcpSection, gVar.RegTcpKeyPort, strPort
            
            If .State <> 0 Then .Close
            .RemoteHost = strIP
            .RemotePort = strPort
            .Connect
            If .State = 7 Then gVar.TCPConnected = True
        ElseIf Command1.Caption = gVar.DisConnectFromServer Then
            .Close
            gVar.TCPConnected = False
        End If
    End With
End Function


Private Sub Command1_Click()
    Const conInterval As Long = 1
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
    With CommonDialog1
        .DialogTitle = "选择传输文件"
        .Filter = "All(*.*)|*.*|Word(*.doc;*.docx)|*.doc;*.docx|Excel(*.xls;*.xlsx)|*.xls;*.xlsx" & _
                  "|Picture(*.jpg;*.bmp;*.png)|*.jpg;*.bmp;*.png|CAD(*.dwg;*.dxf)|*.dwg;*.dxf" & _
                  "|PDF(*.pdf)|*.pdf|Text(*.txt)|*.txt"
        .Flags = cdlOFNFileMustExist
        .ShowOpen
        Text1.Text = .FileName
    End With
End Sub

Private Sub Command3_Click()
    '发送
    
    Const conInterval As Long = 2
    Const conMaxFile As Long = 500
    
    Static sngLastTime As Single
    
    Dim sngCurTime As Single
    Dim strFile As String
    
    sngCurTime = Timer
    If sngCurTime - sngLastTime < conInterval Then
        MsgBox "两次点击时间间隔小于" & conInterval & "秒！", vbExclamation
        Exit Sub
    End If
    sngLastTime = sngCurTime
    
    strFile = Trim(Text1.Text)
    If Len(strFile) = 0 Then
        Command2.SetFocus
        MsgBox "请先选择一个文件！", vbExclamation
        Exit Sub
    End If
    If Not gfDirFile(strFile) Then
        MsgBox strFile & vbCrLf & vbCrLf & "该文件不存在，无法传输！", vbExclamation
        Exit Sub
    End If
    
    gArr(1) = gArr(0)
    With gArr(1)
        .FileFolder = gVar.FolderNameTemp
        .FileName = Mid(strFile, InStrRev(strFile, "\") + 1)
        .FileSizeTotal = FileLen(strFile)
        .FilePath = strFile
        If .FileSizeTotal > (conMaxFile * 1024 * 1024) Or .FileSizeTotal < 0 Then
            MsgBox "传输的单个文件大小不能超过" & conMaxFile & "M！", vbExclamation
            Exit Sub
        End If
    End With
    
    If Winsock1.Item(1).State <> 7 Then
        MsgBox "请先建立连接！", vbExclamation
        Exit Sub
    End If
    
    If gfSendInfo(gfFileInfoJoin(1), Winsock1.Item(1)) Then
        With LabelProgressBar1
            .Value = 0
            .Max = gArr(1).FileSizeTotal
            .Min = 0
        End With
    Else
        MsgBox "文件信息传输失败！", vbExclamation
        Exit Sub
    End If
        
End Sub

Private Sub Command4_Click()
    '接收
    
End Sub

Private Sub Form_Load()
    Dim strUP As String
    
    ReDim gArr(1)
    Timer1.Interval = 1000
    
    Call gsInitialize
    Call gfStartUpSet
    Call gfLoadSkin(Me, SkinFramework1, , True)
    Call mfConnect
    
    strUP = gVar.AppPath & gVar.UpdateExeName & " " & gVar.ClientExeName & _
            gVar.CmdSeparator & gVar.CmdLineHide    '隐式打开更新检测程序
'    strUP = gVar.AppPath & gVar.UpdateExeName & " " & gVar.ClientExeName   '显示打开更新检测程序窗口
    If Not gfShell(strUP) Then
        MsgBox "更新程序启动异常！", vbExclamation, "警告"
    End If
    
    If LCase(App.EXEName & ".exe") <> LCase(gVar.ClientExeName) Then
        MsgBox "不可擅自修改可执行的应用程序文件名！", vbCritical, "严重警报"
        Unload Me   '防止exe文件名被改
    End If
End Sub

Private Sub Timer1_Timer()
    Const conConn As Byte = 1       '连接状态检测间隔conConn秒
    Const conState As Byte = 5      '连接服务器检测间隔conState秒
    
    Static byteConn As Byte
    Static byteState As Byte
    
    byteConn = byteConn + 1
    byteState = byteState + 1
    
    If byteConn >= conConn Then
        If Winsock1.Item(1).State = 7 Then
            If Command1.Caption <> gVar.DisConnectFromServer Then
                Command1.Caption = gVar.DisConnectFromServer
                Label1.Item(1).Caption = gVar.Connected
                Label1.Item(1).ForeColor = vbBlue
                gVar.TCPConnected = False
            End If
        ElseIf Winsock1.Item(1).State = 9 Then
            If Label1.Item(1).Caption <> gVar.ConnectError Then
                Command1.Caption = gVar.ConnectToServer
                Label1.Item(1).Caption = gVar.ConnectError
                Label1.Item(1).ForeColor = vbRed
                gVar.TCPConnected = False
            End If
        Else
            If Command1.Caption <> gVar.ConnectToServer Then
                Command1.Caption = gVar.ConnectToServer
                Label1.Item(1).Caption = gVar.DisConnected
                Label1.Item(1).ForeColor = vbRed
                gVar.TCPConnected = False
            End If
        End If
        byteConn = 0    '复位静态变量
    End If
    
    If byteState >= conState Then
        If Winsock1.Item(1).State <> 7 Then
            Call mfConnect
        End If
        byteState = 0   '复位静态变量
    End If
    
End Sub

Private Sub Winsock1_Close(Index As Integer)
    '连接关闭时清空传输信息
    If UBound(gArr) = 1 Then gArr(1) = gArr(0)
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strGet As String
    
    With gArr(Index)
        If Not .FileTransmitState Then
            Winsock1.Item(Index).GetData strGet
            If InStr(strGet, gVar.PTFileStart) > 0 Then
                Call gfSendFile(.FilePath, Winsock1.Item(Index))
                Call gsFormEnable(Me)
            End If
Debug.Print "Client GetInfo:" & strGet, bytesTotal
        Else
            
        End If
    End With
    
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    If Index <> 0 Then
        If gArr(Index).FileTransmitState Then
Debug.Print "ClientWinsockError:" & Index & "--" & Err.Number & "  " & Err.Description
            Close #gArr(Index).FileNumber
            gArr(Index) = gArr(0)
            Call mfConnect
            Call gsFormEnable(Me, True)
        End If
    End If
    
End Sub

Private Sub Winsock1_SendComplete(Index As Integer)

    If Index = 0 Then Exit Sub
    With gArr(Index)
        If .FileTransmitState Then
            LabelProgressBar1.Value = .FileSizeCompleted
            If .FileSizeCompleted < .FileSizeTotal Then
                Call gfSendFile(.FilePath, Winsock1.Item(Index))
            Else
                gArr(Index) = gArr(0)
                Call gsFormEnable(Me, True)
Debug.Print "Send Over"
            End If
        End If
    End With
    
End Sub
