Attribute VB_Name = "basCommon"
'要求Winsock控件在客户端与服务端都必须建成数组，且其Index值与对应的数组变量的下标要相同

Option Explicit


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'以下API函数Shell_NotifyIcon与一堆常量、枚举、结构体都有关托盘
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
    lpData As NOTIFYICONDATA) As Long

Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const IMAGE_ICON = 1

Public Const NIF_ICON = &H2     'hIcon成员起作用
Public Const NIF_INFO = &H10    '使用气球提示 代替普通的提示框
Public Const NIF_MESSAGE = &H1  'uCallbackMessage成员起作用
Public Const NIF_STATE = &H8    'dwState和dwStateMask成员起作用
Public Const NIF_TIP = &H4      'szTip成员起作用

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIM_SETFOCUS = &H3
Public Const NIM_SETVERSION = &H4
Public Const NIM_VERSION = &H5

Public Const WM_USER As Long = &H400
Public Const NIN_BALLOONSHOW = (WM_USER + 2)
Public Const NIN_BALLOONHIDE = (WM_USER + 3)
Public Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Public Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

Public Const NOTIFYICON_VERSION = 3 '使用Windows2000风格，另一个常量值0表示使用Windows95风格

Public Const NIS_HIDDEN = &H1       '图标隐藏
Public Const NIS_SHAREDICON = &H2   '图标共享

Public Const WM_NOTIFY As Long = &H4E
Public Const WM_COMMAND As Long = &H111
Public Const WM_CLOSE As Long = &H10

Public Const WM_MOUSEMOVE As Long = &H200
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
Public Const WM_LBUTTONDBLCLK As Long = &H203
Public Const WM_RBUTTONDOWN As Long = &H204
Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_RBUTTONDBLCLK As Long = &H206
Public Const WM_MBUTTONDOWN As Long = &H207
Public Const WM_MBUTTONUP As Long = &H208
Public Const WM_MBUTTONDBLCLK As Long = &H209

Public Const SW_RESTORE = 9
Public Const SW_HIDE = 0

Public Type NOTIFYICONDATA
    cbSize As Long  '结构大小（字节）
    hWnd As Long    '处理消息的窗口句柄
    uID As Long     '托盘图标的标识符
    uFlags As Long  '此成员表明哪些其他成员起作用
    uCallbackMessage As Long        '应用程序定义的消息标示
    hIcon As Long                   '托盘图标句柄
    szTip As String * 128   '提示信息,不知为何，长度不设128弹不出气泡

    dwState As Long         '图标状态
    dwStateMask As Long     '指明dwState成员的哪些位可以被设置或访问
    szInfo As String * 256      '气球提示信息
    uTimeoutOrVersion As Long  '气球提示消失时间或版本
    szInfoTitle As String * 64  '气球提示标题
    dwInfoFlags As Long         '给气球提示框增加一个图标
End Type

Public Enum enmNotifyIconFlag
    NIIF_NONE = &H0     '没有图标
    NIIF_INFO = &H1     '信息图标
    NIIF_WARNING = &H2  '警告图标
    NIIF_ERROR = &H3    '错误图标
    NIIF_GUID = &H5     'Version6.0保留
    NIIF_ICON_MASK = &HF    'Version6.0保留
    NIIF_NOSOUND = &H10     'Version6.0禁止播放相应声音
End Enum

Public Enum enmNotifyIconMouseEvent  '鼠标事件
    MouseMove = &H200
    LeftUp = &H202
    LeftDown = &H201
    LeftDbClick = &H203
    RightUp = &H205
    RightDown = &H204
    RightDbClick = &H206
    MiddleUp = &H208
    MiddleDown = &H207
    MiddleDbClick = &H209
    BalloonClick = (WM_USER + 5)
End Enum

Public gNotifyIconData As NOTIFYICONDATA


Public Enum enmFileTransimitType    '文件传输类型枚举
    ftSend = 1      '发送
    ftReceive = 2   '接收
End Enum

Public Enum enmSkinResChoose
    sNone = 0
    sMSVst = 1
    sMS07 = 2
End Enum

Public Type gtypeCommonVariant  '自定义公用常量
    TCPIP As String     '服务器IP地址
    TCPPort As Long     '服务器端口
    TCPConnMax As Long  '最大连接数
    ChunkSize As Long   '文件传输时的分块大小
    WaitTime As Long    '每段文件传输时的等待时间，单位秒
    
    RegTcpSection As String     'section值
    RegTcpKeyIP As String       'key_IP值
    RegTcpKeyPort As String     'key_port值
    RegSkinSection As String    '
    RegSkinKeyFile As String
        
    ServerStart As String       '启动服务
    ServerClose As String       '关闭服务
    ServerError As String       '异常
    ServerStarted As String     '已启动
    ServerNotStarted As String  '未启动
    
    Connected As String             '已连接
    DisConnected As String          '未连接
    ConnectError As String          '连接异常
    ConnectToServer As String       '建立连接
    DisConnectFromServer As String  '断开连接
    
    FolderNameTemp As String    '文件夹名称：Temp
    FolderNameData As String    '文件夹名称：Data
    FolderNameBin As String     '文件夹名称：Bin
    
    PTFileName As String    '协议：文件名标识
    PTFileSize As String    '协议：文件大小标识
    PTFileFolder As String  '协议：文件要保存的文件夹名标识
    PTFileStart As String   '协议：文件开始传输标识
    PTFileEnd As String     '协议：文件结束传输标识
    PTFileSend As String    '协议：文件发送标识
    PTFileReceive As String '协议：文件接收标识
End Type

Public Type gtypeFileTransmitVariant    '自定义文件传输变量
    FileNumber As Integer       '文件传输时打开的文件号
    FilePath As String          '文件名，含全路径
    FileName As String          '仅文件名，不含路径
    FileFolder As String        '文件存储位置的文件夹名称，暂不支持其它路径，默认定在App.Path下
    FileSizeTotal As Long       '文件总大小
    FileSizeCompleted As Long   '文件已传输大小

    FileTransmitState As Boolean    '是否在传输文件
End Type


Public gVar As gtypeCommonVariant
Public gArr() As gtypeFileTransmitVariant



Public Function gfCheckIP(ByVal strIP As String) As String
    Dim K As Long
    Dim arrIP() As String
    
    arrIP = Split(strIP, ".")
    If UBound(arrIP) <> 3 Then GoTo LineOver
    For K = 0 To 3
        If Not IsNumeric(arrIP(K)) Then GoTo LineOver
        If Val(arrIP(K)) < 0 Or Val(arrIP(K)) > 255 Then GoTo LineOver
        arrIP(K) = CStr(Val(arrIP(K)))
    Next
    gfCheckIP = arrIP(0) & "." & arrIP(1) & "." & arrIP(2) & "." & arrIP(3)
    Exit Function
    
LineOver:
    gfCheckIP = "127.0.0.1"
End Function

Public Function gfDirFile(ByVal strFile As String) As Boolean
    Dim strDir As String
    
    strFile = Trim(strFile)
    If Len(strFile) = 0 Then Exit Function
    
    On Error GoTo LineErr
    
    strDir = Dir(strFile, vbHidden + vbReadOnly + vbSystem)
    If Len(strDir) > 0 Then
        SetAttr strFile, vbNormal
        gfDirFile = True
    End If
    
    Exit Function
LineErr:
    Debug.Print "Error:gfDirFile--" & Err.Number & "  " & Err.Description
End Function

Public Function gfDirFolder(ByVal strFolder As String) As Boolean
    Dim strDir As String
    
    strFolder = Trim(strFolder)
    If Len(strFolder) = 0 Then Exit Function
    
    On Error GoTo LineErr
    
    strDir = Dir(strFolder, vbHidden + vbReadOnly + vbSystem + vbDirectory)
    If Len(strDir) = 0 Then
        MkDir strFolder
    Else
        SetAttr strFolder, vbNormal
    End If
    gfDirFolder = True
    
    Exit Function
LineErr:
    Debug.Print "Error:gfDirFolder--" & Err.Number & "  " & Err.Description
End Function

Public Function gfFileInfoJoin(ByVal intIndex As Integer, Optional ByVal enmType As enmFileTransimitType = ftSend) As String
    '文件信息拼接
    Dim strType As String
    
    strType = IIf(enmType = ftReceive, gVar.PTFileReceive, gVar.PTFileSend) '确定文件传输类型
    With gArr(intIndex)
        gfFileInfoJoin = gVar.PTFileFolder & .FileFolder & gVar.PTFileName & .FileName & gVar.PTFileSize & .FileSizeTotal & strType
    End With
    
End Function

Public Function gfLoadSkin(ByRef frmCur As Form, ByRef skFRM As XtremeSkinFramework.SkinFramework, _
    Optional ByVal lngResource As enmSkinResChoose, Optional ByVal blnFromReg As Boolean) As Boolean
    '加载主题
    Dim lngReg As Long, strRes As String, strIni As String
    
    lngReg = GetSetting(App.Title, gVar.RegSkinSection, gVar.RegSkinKeyFile, 0)
    If blnFromReg Then  '如果从注册表中获取资源文件，则按注册表中值修改lngResource的值
        If lngReg > 2 Then lngReg = 0
        lngResource = lngReg
    End If
    
    Select Case lngResource '选择窗口风格资源文件
        Case 1
            strRes = App.Path & "\bin\ftsmsvst.dll"
            strIni = "NormalBlue.ini"   'NormalBlue NormalBlack NormalSilver
        Case 2
            strRes = App.Path & "\bin\ftsms07.dll"
            strIni = "NormalBlue.ini"
        Case Else
    End Select
    
    With skFRM
        .LoadSkin strRes, strIni
        .ApplyOptions = .ApplyOptions Or xtpSkinApplyMetrics Or xtpSkinApplyMenus
        .ApplyWindow frmCur.hWnd
    End With
    
    If lngReg <> lngResource Then Call SaveSetting(App.Title, gVar.RegSkinSection, gVar.RegSkinKeyFile, lngResource)
    
End Function

Public Function gfNotifyIconAdd(ByRef frmCur As Form) As Boolean
    '生成托盘图标
    With gNotifyIconData
        .hWnd = frmCur.hWnd
        .uID = frmCur.Icon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP Or NIF_INFO
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = frmCur.Icon.Handle
        .szTip = App.Title & " " & App.Major & "." & App.Minor & _
            "." & App.Revision & vbNullChar   '鼠标移动托盘图标时显示的Tip信息
        .cbSize = Len(gNotifyIconData)
    End With
    Call Shell_NotifyIcon(NIM_ADD, gNotifyIconData)
End Function

Public Function gfNotifyIconBalloon(ByRef frmCur As Form, ByVal BalloonInfo As String, _
    ByVal BalloonTitle As String, Optional IconFlag As enmNotifyIconFlag = NIIF_INFO) As Boolean
    '托盘图标弹出气泡信息
    With gNotifyIconData
        .dwInfoFlags = IconFlag
        .szInfoTitle = BalloonTitle & vbNullChar
        .szInfo = BalloonInfo & vbNullChar
        .cbSize = Len(gNotifyIconData)
    End With
    Call gfNotifyIconModify(gNotifyIconData)
End Function

Public Function gfNotifyIconDelete(ByRef frmCur As Form) As Boolean
    '删除托盘图标
    Call Shell_NotifyIcon(NIM_DELETE, gNotifyIconData)
End Function

Public Function gfNotifyIconModify(nfIconData As NOTIFYICONDATA) As Boolean
    '修改托盘图标信息
    gNotifyIconData = nfIconData
    Call Shell_NotifyIcon(NIM_MODIFY, gNotifyIconData)
End Function

Public Function gfRestoreInfo(ByVal strInfo As String, sckGet As MSWinsockLib.Winsock) As Boolean
    '还原接收到的文件信息
    
    With gArr(sckGet.Index)
        If InStr(strInfo, gVar.PTFileFolder) > 0 Then
            Dim lngFod As Long, lngFile As Long, lngSize As Long
            Dim lngSend As Long, lngReceive As Long, lngType As Long
            Dim strFod As String, strSize As String, strType As String
            
            lngFod = InStr(strInfo, gVar.PTFileFolder)
            lngFile = InStr(strInfo, gVar.PTFileName)
            lngSize = InStr(strInfo, gVar.PTFileSize)
            lngSend = InStr(strInfo, gVar.PTFileSend)
            lngReceive = InStr(strInfo, gVar.PTFileReceive)
            
            If lngFile > 0 Then
                gArr(sckGet.Index) = gArr(0)
                
                If (lngSend > 0 And lngReceive > 0) Or (lngSend = 0 And lngReceive = 0) Then Exit Function
                strType = IIf(lngSend > 0, gVar.PTFileSend, gVar.PTFileReceive)
                lngType = IIf(lngSend > 0, lngSend, lngReceive)
                
                .FileFolder = Mid(strInfo, lngFod + Len(gVar.PTFileFolder), lngFile - (lngFod + Len(gVar.PTFileFolder)))
                strFod = App.Path & "\" & .FileFolder
                If Not gfDirFolder(strFod) Then Exit Function
                
                .FileName = Mid(strInfo, lngFile + Len(gVar.PTFileName), lngSize - (lngFile + Len(gVar.PTFileName)))
                
                strSize = Mid(strInfo, lngSize + Len(gVar.PTFileSize), lngType - (lngSize + Len(gVar.PTFileSize)))
                If Not IsNumeric(strSize) Then Exit Function
                
                If strType <> Mid(strInfo, lngType) Then Exit Function
                
                If strType = gVar.PTFileSend Then
                    .FileSizeTotal = CLng(strSize)
                    .FilePath = strFod & "\" & .FileName
                    .FileTransmitState = True
                    Call gfSendInfo(gVar.PTFileStart, sckGet)
                ElseIf strType = gVar.PTFileReceive Then
                    
                End If
                gfRestoreInfo = True
            End If
        End If
    End With

End Function

Public Function gfSendFile(ByVal strFile As String, sckSend As MSWinsockLib.Winsock) As Boolean
    Dim lngSendSize As Long, lngRemain As Long
    Dim byteSend() As Byte
    
    With gArr(sckSend.Index)
        If Not .FileTransmitState Then
            .FileNumber = FreeFile
            Open strFile For Binary As #.FileNumber
            .FileTransmitState = True
        End If
        
        lngSendSize = gVar.ChunkSize
        lngRemain = .FileSizeTotal - Loc(.FileNumber)
        If lngSendSize > lngRemain Then lngSendSize = lngRemain
        
        ReDim byteSend(lngSendSize - 1)
        Get #.FileNumber, , byteSend
        sckSend.SendData byteSend
        
        .FileSizeCompleted = .FileSizeCompleted + lngSendSize
        If .FileSizeCompleted = .FileSizeTotal Then Close #.FileNumber
        
    End With
    
End Function

Public Function gfSendInfo(ByVal strInfo As String, sckSend As MSWinsockLib.Winsock) As Boolean
    If sckSend.State = 7 Then
        sckSend.SendData strInfo
        DoEvents
'''        Call Sleep(200)
        gfSendInfo = True
    End If
End Function

Public Sub gsFormEnable(frmCur As Form, Optional ByVal blnState As Boolean)
    With frmCur
        If blnState Then
            .Enabled = True
            .MousePointer = 0
        Else
            .Enabled = False
            .MousePointer = 13
        End If
    End With
End Sub

Public Sub gsInitialize()
    With gVar
        .TCPIP = "127.0.0.1"
        .TCPPort = 9898
        .TCPConnMax = 20
        .ChunkSize = 8000
        .WaitTime = 5
        
        .RegTcpKeyIP = "IP"
        .RegTcpKeyPort = "Port"
        .RegTcpSection = "TCP"
        
        .RegSkinKeyFile = "File"
        .RegSkinSection = "Res"
        
        .ServerClose = "关闭服务"
        .ServerError = "异常"
        .ServerNotStarted = "未启动"
        .ServerStart = "开启服务"
        .ServerStarted = "已启动"
        
        .Connected = "已连接"
        .DisConnected = "未连接"
        .ConnectError = "连接异常"
        .ConnectToServer = "建立连接"
        .DisConnectFromServer = "断开连接"
        
        .FolderNameBin = "Bin"
        .FolderNameData = "Data"
        .FolderNameTemp = "Temp"
        
        .PTFileName = "<FileName>"
        .PTFileSize = "<FileSize>"
        .PTFileFolder = "<FileFolder>"
        
        .PTFileStart = "<FileStart>"
        .PTFileEnd = "<FileEnd>"
        
        .PTFileSend = "<FileSend>"
        .PTFileReceive = "<FileReceive>"
    End With
End Sub
