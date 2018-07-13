VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "更新程序"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5370
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   4335
   End
   Begin FTCUpdate.LabelProgressBar LabelProgressBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   4335
      _ExtentX        =   7646
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
   Begin VB.Timer Timer1 
      Left            =   1920
      Top             =   2160
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   1
      Left            =   1320
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   225
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrCmd As String   'EXEname
Dim mblnCheckStart As Boolean   '已开始检查标识
Dim mblnUpdateFinish As Boolean     '更新完成标识

Private Function mfCheckUpdate() As Boolean
    '检查更新
    Dim strFileLoc As String, strFileNet As String, strVerLoc As String, strVerNet As String
    
    strFileLoc = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & mstrCmd
    If Not gfDirFile(strFileLoc) Then Exit Function
    strVerLoc = gfBackVersion(strFileLoc)
    If Not mfVersionCompare(strVerLoc, strVerNet) Then Exit Function
    
    '有新版需要更新
    
    mfCheckUpdate = True
    
End Function

Private Function mfConnect() As Boolean
    Dim strIP As String, strPort As String
    Static lngCount As Long
            
    lngCount = lngCount + 1
    If lngCount = 2 Then Exit Function    '尝试百次后不再连接了
    
    With Winsock1.Item(1)
        If Label1(1).Caption = gVar.DisConnected Then
            strIP = GetSetting(gVar.RegAppName, gVar.RegTcpSection, gVar.RegTcpKeyIP, gVar.TCPIP)
            strIP = gfCheckIP(strIP)

            strPort = GetSetting(gVar.RegAppName, gVar.RegTcpSection, gVar.RegTcpKeyPort, gVar.TCPPort)
            strPort = CStr(CLng(Val(strPort)))
            If Val(strPort) > 65535 Or Val(strPort) < 0 Then strPort = gVar.TCPPort

            If .State <> 0 Then .Close
            .RemoteHost = strIP
            .RemotePort = strPort
            .Connect
            If .State = 7 Then gVar.TCPConnected = True
        End If
    End With
End Function

Private Function mfVersionCompare(ByVal strVerOld As String, ByVal strVerNew As String) As Boolean
    '新旧版本号比较
    Dim ArrOld() As String, ArrNew() As String
    Dim K As Long, C As Long
    
    ArrOld = Split(strVerOld, ".")
    ArrNew = Split(strVerNew, ".")
    K = UBound(ArrOld)
    C = UBound(ArrNew)
    If K = C And K = 4 Then
        For K = 0 To C
            If Val(ArrNew(K)) > Val(ArrOld(K)) Then
                mfVersionCompare = True '说明有新版本
                Exit For
            End If
        Next
    End If
    
End Function

Private Sub Form_Load()
    Dim strIP As String, strPort As String
    
    Call gsInitialize
    
    '检测是否传入命令行参数进来，没有则退出程序
    mstrCmd = Command()
    If UCase(mstrCmd) <> gVar.CmdLineStr Then
'        GoTo LineUnload
'        mstrCmd = gVar.CmdLineStr
    End If
    
    Text1.BackColor = Me.BackColor
    Label1(1).Caption = gVar.DisConnected
    Label1(1).BackColor = vbRed
    Timer1.Interval = 1000
    Timer1.Enabled = True
      
    Exit Sub
    
LineUnload:
    Unload Me   '此行以下除End Sub不可再跟任何有效代码
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
            Label1.Item(1).Caption = gVar.Connected
            Label1.Item(1).BackColor = vbGreen
            gVar.TCPConnected = True
            If Not mblnCheckStart Then
                mblnCheckStart = True
                Call mfCheckUpdate
            End If
        Else
            Label1.Item(1).Caption = gVar.DisConnected
            Label1.Item(1).BackColor = vbRed
            gVar.TCPConnected = False
        End If
        byteConn = 0    '复位静态变量
    End If
    
    If byteState >= conState Then
        If Winsock1.Item(1).State <> 7 Then
            If Not mblnUpdateFinish Then Call mfConnect
        End If
        byteState = 0   '复位静态变量
    End If
End Sub
