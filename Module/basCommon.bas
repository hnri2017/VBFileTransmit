Attribute VB_Name = "basCommon"
'要求Winsock控件在客户端与服务端都必须建成数组，且其Index值与对应的数组变量的下标要相同

Option Explicit


'查找窗口，发送信息
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'使用 ShellExecute 打开文件或执行程序
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'hWnd：用于指定父窗口句柄。当函数调用过程出现错误时，它将作为Windows消息窗口的父窗口
'Operation：用于指定要进行的操作。其中:
'''edit 用编辑器打开 lpFile 指定的文档，如果 lpFile 不是文档，则会失败;
'''explore 浏览 lpFile 指定的文件夹
'''find 搜索 lpDirectory 指定的目录
'''open 打开 lpFile 文件，lpFile 可以是文件或文件夹
'''print 打印 lpFile，如果 lpFile 不是文档，则函数失败
'''properties 显示属性
'''runas 请求以管理员权限运行，比如以管理员权限运行某个exe
'''NULL 执行默认”open”动作
'FileName：用于指定要打开的文件名、要执行的程序文件名或要浏览的文件夹名
'Parameters：若FileName参数是一个可执行程序，则此参数指定命令行参数，否则此参数应为nil或PChar(0)
'Directory：用于指定默认目录
'ShowCmd：若FileName参数是一个可执行程序，则此参数指定程序窗口的初始显示方式，否则此参数应设置为0

'若ShellExecute函数调用成功，则返回值为被执行程序的实例句柄。若返回值小于32，则表示出现错误,错误如下:
Public Const NO_ERROR = 0   '系统内存或资源不足
Public Const ERROR_FILE_NOT_FOUND = 2&  '找不到指定的文件
Public Const ERROR_PATH_NOT_FOUND = 3&  '找不到指定路径
Public Const ERROR_BAD_FORMAT = 11&     '.exe文件无效
Public Const SE_ERR_ACCESSDENIED = 5    '拒绝访问指定文件
Public Const SE_ERR_ASSOCINCOMPLETE = 27    '文件名关联无效或不完整
Public Const SE_ERR_DDEBUSY = 30    'DDE事务正在处理，DDE事务无法完成
Public Const SE_ERR_DDEFAIL = 29    'DDE事务失败
Public Const SE_ERR_DDETIMEOUT = 28 '请求超时，无法完成DDE事务请求
Public Const SE_ERR_DLLNOTFOUND = 32    '未找到指定dll
Public Const SE_ERR_FNF = 2         '未找到指定文件
Public Const SE_ERR_NOASSOC = 31    '未找到与给的文件拓展名关联的应用程序，比如打印不可打印的文件等
Public Const SE_ERR_OOM = 8         '内存不足，无法完成操作
Public Const SE_ERR_PNF = 3         '未找到指定路径
Public Const SE_ERR_SHARE = 26      '发生共享冲突

'ShellExecute参数nShowCmd所用的常量ShowWindow() Commands
Public Const SW_HIDE = 0        '隐藏窗口，活动状态给令一个窗口
Public Const SW_SHOWNORMAL = 1  '与SW_RESTORE相同
Public Const SW_NORMAL = 1      '
Public Const SW_SHOWMINIMIZED = 2   '最小化窗口，并将其激活
Public Const SW_SHOWMAXIMIZED = 3   'SHOWMAXIMIZED 最大化窗口，并将其激活
Public Const SW_MAXIMIZE = 3        '
Public Const SW_SHOWNOACTIVATE = 4  '用最近的大小和位置显示一个窗口，同时不改变活动窗口
Public Const SW_SHOW = 5            '用当前的大小和位置显示一个窗口，同时令其进入活动状态
Public Const SW_MINIMIZE = 6        '最小化窗口，活动状态给令一个窗口
Public Const SW_SHOWMINNOACTIVE = 7 '最小化一个窗口，同时不改变活动窗口
Public Const SW_SHOWNA = 8          '用当前的大小和位置显示一个窗口，不改变活动窗口
Public Const SW_RESTORE = 9         '用原来的大小和位置显示一个窗口，同时令其进入活动状态
Public Const SW_SHOWDEFAULT = 10    '
Public Const SW_MAX = 10            '


'注册表操作API与类型
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Public Const HKEY_USER_RUN As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"  '软件开机自动启动注册表子键位置

Public Enum genumRegRootDirectory   '注册表根键
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
End Enum

Public Enum genumRegDataType    '注册表值类型
    REG_SZ = 1          ' Unicode nul terminated string
    REG_EXPAND_SZ = 2   ' Unicode nul terminated string
    REG_BINARY = 3      ' Free form binary
    REG_DWORD = 4       ' 32-bit number
End Enum

Public Enum genumRegOperateType '注册表操作类型
    RegRead = 1
    RegWrite = 2
    RegDelete = 3
End Enum


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)  '程序暂停运行（毫秒）


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

Public Type NOTIFYICONDATA
    cbSize As Long  '结构大小（字节）
    hwnd As Long    '处理消息的窗口句柄
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
    sNone = 0   '无
    sMSVst = 1  'MicrosoftVista风格
    sMS07 = 2   'MicrosoftOffice2007风格
End Enum

Public Type gtypeCommonVariant  '自定义公用常量
    TCPIP As String     '服务器IP地址
    TCPPort As Long     '服务器端口
    TCPConnMax As Long  '最大连接数
    TCPConnected As Boolean     '连接成功标识
    TCPServerStarted As Boolean '服务器启动标识
    ChunkSize As Long   '文件传输时的分块大小
    WaitTime As Long    '每段文件传输时的等待时间，单位秒
    
    AppPath As String           'App路径，确保最后字符为"\"
    ClientExeName As String     '客户端Exe文件名 / 命令行参数值
    CmdSeparator As String      '命令行间隔符
    CmdLineHide As String       '命令行参数之隐藏
    NewSetupFileName As String  '更新安装包的文件名
    
    RegAppName As String
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
    
    PTVersionOfClient As String     '协议：客户端版本号
    PTVersionNotUpdate As String    '协议：不需要更新
    PTVersionNeedUpdate As String   '协议：需要更新
    
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



Public Function gfBackVersion(ByVal strFile As String) As String
    '返回文件的版本号
    Dim objFile As Scripting.FileSystemObject
    
    If Not gfDirFile(strFile) Then Exit Function
    Set objFile = New FileSystemObject
    gfBackVersion = objFile.GetFileVersion(strFile)

    Set objFile = Nothing
End Function


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

Public Function gfCloseApp(ByVal strName As String) As Boolean
    '关闭指定应用程序进程
    
    Dim winHwnd As Long
    Dim RetVal As Long
    Dim objWMIService As Object
    Dim colProcessList As Object
    Dim objProcess As Object
    
    On Error GoTo LineErr
    
''    winHwnd = FindWindow(vbNullString, strName) '查找窗口，strName内容即任务栏上看到的窗口标题
''    If winHwnd <> 0 Then    '不为0表示找到窗口
''        RetVal = PostMessage(winHwnd, WM_CLOSE, 0&, 0&) '发送关闭窗口信息,返回值为0表示关闭失败
''    End If
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("select * from Win32_Process where Name='" & strName & "' ")
    For Each objProcess In colProcessList
        RetVal = objProcess.Terminate
        If RetVal <> 0 Then Exit Function   '经观察=0时关闭进程成功，不成功时返回值不为零
    Next
    
    gfCloseApp = True   '全部关闭成功或不存在该进程名时
    
LineErr:
    Set objWMIService = Nothing
    Set colProcessList = Nothing
    Set objProcess = Nothing
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
    
    lngReg = GetSetting(gVar.RegAppName, gVar.RegSkinSection, gVar.RegSkinKeyFile, 0)
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
        .ApplyWindow frmCur.hwnd
    End With
    
    If lngReg <> lngResource Then Call SaveSetting(gVar.RegAppName, gVar.RegSkinSection, gVar.RegSkinKeyFile, lngResource)
    
End Function

Public Function gfNotifyIconAdd(ByRef frmCur As Form) As Boolean
    '生成托盘图标
    With gNotifyIconData
        .hwnd = frmCur.hwnd
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

Public Function gfRegOperate(ByVal RegHKEY As genumRegRootDirectory, ByVal lpSubKey As String, _
    ByVal lpValueName As String, Optional ByVal lpType As genumRegDataType = REG_SZ, _
    Optional ByRef lpValue As String, Optional ByVal lpOp As genumRegOperateType = RegRead) As Boolean
    '
    Dim Ret As Long, hKey As Long, lngLength As Long
    Dim Buff() As Byte
    
    
    Ret = RegOpenKey(RegHKEY, lpSubKey, hKey)
    If Ret = 0 Then
        Select Case lpOp
            Case RegDelete
                Ret = RegDeleteValue(hKey, lpValueName)
                If Ret = 0 Then
                    gfRegOperate = True
                End If
                
            Case RegWrite
                lngLength = LenB(StrConv(lpValue, vbFromUnicode))   '不用LenB与StrConv的话lpValue字符串长度对不上
                Ret = RegSetValueEx(hKey, lpValueName, 0, lpType, ByVal lpValue, lngLength)
                If Ret = 0 Then
                    gfRegOperate = True
'Debug.Print "W", lpValue, lngLength
                End If
                
            Case Else
                Ret = RegQueryValueEx(hKey, lpValueName, 0, lpType, ByVal 0, lngLength) '获取值的长度
                If Ret = 0 And lngLength > 0 Then
                    ReDim Buff(lngLength - 1)   '重定义缓冲大小
                    Ret = RegQueryValueEx(hKey, lpValueName, 0, lpType, Buff(0), lngLength) '取值
                    If Ret = 0 And lngLength > 1 Then
                        ReDim Preserve Buff(lngLength - 2)
                        lpValue = StrConv(Buff, vbUnicode)
                        gfRegOperate = True
'Debug.Print "R", lpValue, lngLength - 1
                    End If
                End If
                
        End Select
    End If
    
    Call RegCloseKey(hKey)
    
End Function


Public Function gfRestoreInfo(ByVal strInfo As String, sckGet As MSWinsockLib.Winsock) As Boolean
    '还原接收到的文件信息
    
    With gArr(sckGet.Index)
        If InStr(strInfo, gVar.PTFileFolder) > 0 Then
            '此判断似乎仅适应于客户端向服务端上传文件时，其它情形有待确认
            
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
                    Call gfSendInfo(gVar.PTFileStart, sckGet)
                    .FileTransmitState = True
                ElseIf strType = gVar.PTFileReceive Then
                    
                End If
                gfRestoreInfo = True
            End If
        ElseIf InStr(strInfo, gVar.PTVersionNotUpdate) > 0 Then
            
        End If
    End With

End Function

Public Function gfSendFile(ByVal strFile As String, sckSend As MSWinsockLib.Winsock) As Boolean
    Dim lngSendSize As Long, lngRemain As Long
    Dim byteSend() As Byte
    
    With gArr(sckSend.Index)
        If .FileNumber = 0 Then
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

Public Function gfShellExecute(ByVal strFile As String) As Boolean
    '执行程序或打开文件或文件夹
    '''Call ShellExecute(Me.hwnd, "open", strFile, vbNullString, vbNullString, 1)

    Dim lngRet As Long
    Dim strDir As String
    
    lngRet = ShellExecute(GetDesktopWindow, "open", strFile, vbNullString, vbNullString, vbNormalFocus)

    ' 没有关联的程序
    If lngRet = SE_ERR_NOASSOC Then
         strDir = Space$(260)
         lngRet = GetSystemDirectory(strDir, Len(strDir))
         strDir = Left$(strDir, lngRet)
       ' 显示打开方式窗口
         lngRet = ShellExecute(GetDesktopWindow, vbNullString, "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & strFile, strDir, vbNormalFocus)
    End If
    
    If lngRet > 32 Then gfShellExecute = True
    
End Function

Public Function gfStartUpSet() As Boolean
    
    '开机自启动设置
    Dim strReg As String, strCur As String
    Dim blnReg As Boolean
    
    strCur = Chr(34) & App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & App.EXEName & ".exe" & Chr(34) & "-s"
    blnReg = gfRegOperate(HKEY_LOCAL_MACHINE, HKEY_USER_RUN, App.EXEName, REG_SZ, strReg, RegRead)
    If blnReg Then
        If LCase(strCur) <> LCase(strReg) Then
            blnReg = False
'''Debug.Print LCase(strCur)
'''Debug.Print LCase(strReg)
        End If
    End If
    If Not blnReg Then
        blnReg = gfRegOperate(HKEY_LOCAL_MACHINE, HKEY_USER_RUN, App.EXEName, REG_SZ, strCur, RegWrite)
        If Not blnReg Then
            '记录设置开机自动启动失败
            
        End If
    End If
    
End Function

Public Function gfVersionCompare(ByVal strVerCL As String, ByVal strVerSV As String) As String
    '新旧版本号比较
    Dim ArrCL() As String, ArrSV() As String
    Dim K As Long, C As Long
    
    ArrCL = Split(strVerCL, ".")
    ArrSV = Split(strVerSV, ".")
    K = UBound(ArrCL)
    C = UBound(ArrSV)
    If K = C And K = 3 Then
        For K = 0 To C
            If Not IsNumeric(ArrCL(K)) Then
                gfVersionCompare = "客户端版本异常"
                Exit For
            End If
            If Not IsNumeric(ArrSV(K)) Then
                gfVersionCompare = "服务端版本异常"
                Exit For
            End If
            
            If Val(ArrSV(K)) > Val(ArrCL(K)) Then
                gfVersionCompare = "1" '说明有新版本
                Exit For
            End If
        Next
        If K = C + 1 Then gfVersionCompare = "0" '说明没有新版，不用更新
    Else
        If K = 3 And C <> 3 Then
            gfVersionCompare = "服务端版本获取异常"
        ElseIf C = 3 And K <> 3 Then
            gfVersionCompare = "客户端版本获取异常"
        Else
            gfVersionCompare = "版本获取异常"
        End If
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
        .ChunkSize = 5734
        .WaitTime = 5
        
        .AppPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
        .ClientExeName = "exeFTClient.exe"
        .NewSetupFileName = "FTClientSetup.exe"
        .CmdLineHide = "Hide"
        .CmdSeparator = " / "
        
        .RegAppName = "FT"
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
        
        .PTVersionNeedUpdate = "<VersionNeedUpdate>"
        .PTVersionNotUpdate = "<VersionNotUpdate>"
        .PTVersionOfClient = "<VersionOfClient>"
        
    End With
    
    
End Sub
