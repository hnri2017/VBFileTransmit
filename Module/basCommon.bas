Attribute VB_Name = "basCommon"
'Ҫ��Winsock�ؼ��ڿͻ��������˶����뽨�����飬����Indexֵ���Ӧ������������±�Ҫ��ͬ

Option Explicit


'���Ҵ��ڣ�������Ϣ
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'ʹ�� ShellExecute ���ļ���ִ�г���
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'hWnd������ָ�������ھ�������������ù��̳��ִ���ʱ��������ΪWindows��Ϣ���ڵĸ�����
'Operation������ָ��Ҫ���еĲ���������:
'''edit �ñ༭���� lpFile ָ�����ĵ������ lpFile �����ĵ������ʧ��;
'''explore ��� lpFile ָ�����ļ���
'''find ���� lpDirectory ָ����Ŀ¼
'''open �� lpFile �ļ���lpFile �������ļ����ļ���
'''print ��ӡ lpFile����� lpFile �����ĵ�������ʧ��
'''properties ��ʾ����
'''runas �����Թ���ԱȨ�����У������Թ���ԱȨ������ĳ��exe
'''NULL ִ��Ĭ�ϡ�open������
'FileName������ָ��Ҫ�򿪵��ļ�����Ҫִ�еĳ����ļ�����Ҫ������ļ�����
'Parameters����FileName������һ����ִ�г�����˲���ָ�������в���������˲���ӦΪnil��PChar(0)
'Directory������ָ��Ĭ��Ŀ¼
'ShowCmd����FileName������һ����ִ�г�����˲���ָ�����򴰿ڵĳ�ʼ��ʾ��ʽ������˲���Ӧ����Ϊ0

'��ShellExecute�������óɹ����򷵻�ֵΪ��ִ�г����ʵ�������������ֵС��32�����ʾ���ִ���,��������:
Public Const NO_ERROR = 0   'ϵͳ�ڴ����Դ����
Public Const ERROR_FILE_NOT_FOUND = 2&  '�Ҳ���ָ�����ļ�
Public Const ERROR_PATH_NOT_FOUND = 3&  '�Ҳ���ָ��·��
Public Const ERROR_BAD_FORMAT = 11&     '.exe�ļ���Ч
Public Const SE_ERR_ACCESSDENIED = 5    '�ܾ�����ָ���ļ�
Public Const SE_ERR_ASSOCINCOMPLETE = 27    '�ļ���������Ч������
Public Const SE_ERR_DDEBUSY = 30    'DDE�������ڴ���DDE�����޷����
Public Const SE_ERR_DDEFAIL = 29    'DDE����ʧ��
Public Const SE_ERR_DDETIMEOUT = 28 '����ʱ���޷����DDE��������
Public Const SE_ERR_DLLNOTFOUND = 32    'δ�ҵ�ָ��dll
Public Const SE_ERR_FNF = 2         'δ�ҵ�ָ���ļ�
Public Const SE_ERR_NOASSOC = 31    'δ�ҵ�������ļ���չ��������Ӧ�ó��򣬱����ӡ���ɴ�ӡ���ļ���
Public Const SE_ERR_OOM = 8         '�ڴ治�㣬�޷���ɲ���
Public Const SE_ERR_PNF = 3         'δ�ҵ�ָ��·��
Public Const SE_ERR_SHARE = 26      '���������ͻ

'ShellExecute����nShowCmd���õĳ���ShowWindow() Commands
Public Const SW_HIDE = 0        '���ش��ڣ��״̬����һ������
Public Const SW_SHOWNORMAL = 1  '��SW_RESTORE��ͬ
Public Const SW_NORMAL = 1      '
Public Const SW_SHOWMINIMIZED = 2   '��С�����ڣ������伤��
Public Const SW_SHOWMAXIMIZED = 3   'SHOWMAXIMIZED ��󻯴��ڣ������伤��
Public Const SW_MAXIMIZE = 3        '
Public Const SW_SHOWNOACTIVATE = 4  '������Ĵ�С��λ����ʾһ�����ڣ�ͬʱ���ı�����
Public Const SW_SHOW = 5            '�õ�ǰ�Ĵ�С��λ����ʾһ�����ڣ�ͬʱ�������״̬
Public Const SW_MINIMIZE = 6        '��С�����ڣ��״̬����һ������
Public Const SW_SHOWMINNOACTIVE = 7 '��С��һ�����ڣ�ͬʱ���ı�����
Public Const SW_SHOWNA = 8          '�õ�ǰ�Ĵ�С��λ����ʾһ�����ڣ����ı�����
Public Const SW_RESTORE = 9         '��ԭ���Ĵ�С��λ����ʾһ�����ڣ�ͬʱ�������״̬
Public Const SW_SHOWDEFAULT = 10    '
Public Const SW_MAX = 10            '


'ע������API������
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Public Const HKEY_USER_RUN As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"  '��������Զ�����ע����Ӽ�λ��

Public Enum genumRegRootDirectory   'ע������
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
End Enum

Public Enum genumRegDataType    'ע���ֵ����
    REG_SZ = 1          ' Unicode nul terminated string
    REG_EXPAND_SZ = 2   ' Unicode nul terminated string
    REG_BINARY = 3      ' Free form binary
    REG_DWORD = 4       ' 32-bit number
End Enum

Public Enum genumRegOperateType 'ע����������
    RegRead = 1
    RegWrite = 2
    RegDelete = 3
End Enum


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)  '������ͣ���У����룩


'����API����Shell_NotifyIcon��һ�ѳ�����ö�١��ṹ�嶼�й�����
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
    lpData As NOTIFYICONDATA) As Long

Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const IMAGE_ICON = 1

Public Const NIF_ICON = &H2     'hIcon��Ա������
Public Const NIF_INFO = &H10    'ʹ��������ʾ ������ͨ����ʾ��
Public Const NIF_MESSAGE = &H1  'uCallbackMessage��Ա������
Public Const NIF_STATE = &H8    'dwState��dwStateMask��Ա������
Public Const NIF_TIP = &H4      'szTip��Ա������

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

Public Const NOTIFYICON_VERSION = 3 'ʹ��Windows2000�����һ������ֵ0��ʾʹ��Windows95���

Public Const NIS_HIDDEN = &H1       'ͼ������
Public Const NIS_SHAREDICON = &H2   'ͼ�깲��

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
    cbSize As Long  '�ṹ��С���ֽڣ�
    hwnd As Long    '������Ϣ�Ĵ��ھ��
    uID As Long     '����ͼ��ı�ʶ��
    uFlags As Long  '�˳�Ա������Щ������Ա������
    uCallbackMessage As Long        'Ӧ�ó��������Ϣ��ʾ
    hIcon As Long                   '����ͼ����
    szTip As String * 128   '��ʾ��Ϣ,��֪Ϊ�Σ����Ȳ���128����������

    dwState As Long         'ͼ��״̬
    dwStateMask As Long     'ָ��dwState��Ա����Щλ���Ա����û����
    szInfo As String * 256      '������ʾ��Ϣ
    uTimeoutOrVersion As Long  '������ʾ��ʧʱ���汾
    szInfoTitle As String * 64  '������ʾ����
    dwInfoFlags As Long         '��������ʾ������һ��ͼ��
End Type

Public Enum enmNotifyIconFlag
    NIIF_NONE = &H0     'û��ͼ��
    NIIF_INFO = &H1     '��Ϣͼ��
    NIIF_WARNING = &H2  '����ͼ��
    NIIF_ERROR = &H3    '����ͼ��
    NIIF_GUID = &H5     'Version6.0����
    NIIF_ICON_MASK = &HF    'Version6.0����
    NIIF_NOSOUND = &H10     'Version6.0��ֹ������Ӧ����
End Enum

Public Enum enmNotifyIconMouseEvent  '����¼�
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


Public Enum enmFileTransimitType    '�ļ���������ö��
    ftSend = 1      '����
    ftReceive = 2   '����
End Enum

Public Enum enmSkinResChoose
    sNone = 0   '��
    sMSVst = 1  'MicrosoftVista���
    sMS07 = 2   'MicrosoftOffice2007���
End Enum

Public Type gtypeCommonVariant  '�Զ��幫�ó���
    TCPIP As String     '������IP��ַ
    TCPPort As Long     '�������˿�
    TCPConnMax As Long  '���������
    TCPConnected As Boolean     '���ӳɹ���ʶ
    TCPServerStarted As Boolean '������������ʶ
    ChunkSize As Long   '�ļ�����ʱ�ķֿ��С
    WaitTime As Long    'ÿ���ļ�����ʱ�ĵȴ�ʱ�䣬��λ��
    
    AppPath As String           'App·����ȷ������ַ�Ϊ"\"
    ClientExeName As String     '�ͻ���Exe�ļ��� / �����в���ֵ
    CmdSeparator As String      '�����м����
    CmdLineHide As String       '�����в���֮����
    NewSetupFileName As String  '���°�װ�����ļ���
    
    RegAppName As String
    RegTcpSection As String     'sectionֵ
    RegTcpKeyIP As String       'key_IPֵ
    RegTcpKeyPort As String     'key_portֵ
    RegSkinSection As String    '
    RegSkinKeyFile As String
        
    ServerStart As String       '��������
    ServerClose As String       '�رշ���
    ServerError As String       '�쳣
    ServerStarted As String     '������
    ServerNotStarted As String  'δ����
    
    Connected As String             '������
    DisConnected As String          'δ����
    ConnectError As String          '�����쳣
    ConnectToServer As String       '��������
    DisConnectFromServer As String  '�Ͽ�����
    
    FolderNameTemp As String    '�ļ������ƣ�Temp
    FolderNameData As String    '�ļ������ƣ�Data
    FolderNameBin As String     '�ļ������ƣ�Bin
    
    PTFileName As String    'Э�飺�ļ�����ʶ
    PTFileSize As String    'Э�飺�ļ���С��ʶ
    PTFileFolder As String  'Э�飺�ļ�Ҫ������ļ�������ʶ
    PTFileStart As String   'Э�飺�ļ���ʼ�����ʶ
    PTFileEnd As String     'Э�飺�ļ����������ʶ
    PTFileSend As String    'Э�飺�ļ����ͱ�ʶ
    PTFileReceive As String 'Э�飺�ļ����ձ�ʶ
    
    PTVersionOfClient As String     'Э�飺�ͻ��˰汾��
    PTVersionNotUpdate As String    'Э�飺����Ҫ����
    PTVersionNeedUpdate As String   'Э�飺��Ҫ����
    
End Type

Public Type gtypeFileTransmitVariant    '�Զ����ļ��������
    FileNumber As Integer       '�ļ�����ʱ�򿪵��ļ���
    FilePath As String          '�ļ�������ȫ·��
    FileName As String          '���ļ���������·��
    FileFolder As String        '�ļ��洢λ�õ��ļ������ƣ��ݲ�֧������·����Ĭ�϶���App.Path��
    FileSizeTotal As Long       '�ļ��ܴ�С
    FileSizeCompleted As Long   '�ļ��Ѵ����С

    FileTransmitState As Boolean    '�Ƿ��ڴ����ļ�
End Type


Public gVar As gtypeCommonVariant
Public gArr() As gtypeFileTransmitVariant



Public Function gfBackVersion(ByVal strFile As String) As String
    '�����ļ��İ汾��
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
    '�ر�ָ��Ӧ�ó������
    
    Dim winHwnd As Long
    Dim RetVal As Long
    Dim objWMIService As Object
    Dim colProcessList As Object
    Dim objProcess As Object
    
    On Error GoTo LineErr
    
''    winHwnd = FindWindow(vbNullString, strName) '���Ҵ��ڣ�strName���ݼ��������Ͽ����Ĵ��ڱ���
''    If winHwnd <> 0 Then    '��Ϊ0��ʾ�ҵ�����
''        RetVal = PostMessage(winHwnd, WM_CLOSE, 0&, 0&) '���͹رմ�����Ϣ,����ֵΪ0��ʾ�ر�ʧ��
''    End If
    
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("select * from Win32_Process where Name='" & strName & "' ")
    For Each objProcess In colProcessList
        RetVal = objProcess.Terminate
        If RetVal <> 0 Then Exit Function   '���۲�=0ʱ�رս��̳ɹ������ɹ�ʱ����ֵ��Ϊ��
    Next
    
    gfCloseApp = True   'ȫ���رճɹ��򲻴��ڸý�����ʱ
    
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
    '�ļ���Ϣƴ��
    Dim strType As String
    
    strType = IIf(enmType = ftReceive, gVar.PTFileReceive, gVar.PTFileSend) 'ȷ���ļ���������
    With gArr(intIndex)
        gfFileInfoJoin = gVar.PTFileFolder & .FileFolder & gVar.PTFileName & .FileName & gVar.PTFileSize & .FileSizeTotal & strType
    End With
    
End Function

Public Function gfLoadSkin(ByRef frmCur As Form, ByRef skFRM As XtremeSkinFramework.SkinFramework, _
    Optional ByVal lngResource As enmSkinResChoose, Optional ByVal blnFromReg As Boolean) As Boolean
    '��������
    Dim lngReg As Long, strRes As String, strIni As String
    
    lngReg = GetSetting(gVar.RegAppName, gVar.RegSkinSection, gVar.RegSkinKeyFile, 0)
    If blnFromReg Then  '�����ע����л�ȡ��Դ�ļ�����ע�����ֵ�޸�lngResource��ֵ
        If lngReg > 2 Then lngReg = 0
        lngResource = lngReg
    End If
    
    Select Case lngResource 'ѡ�񴰿ڷ����Դ�ļ�
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
    '��������ͼ��
    With gNotifyIconData
        .hwnd = frmCur.hwnd
        .uID = frmCur.Icon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP Or NIF_INFO
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = frmCur.Icon.Handle
        .szTip = App.Title & " " & App.Major & "." & App.Minor & _
            "." & App.Revision & vbNullChar   '����ƶ�����ͼ��ʱ��ʾ��Tip��Ϣ
        .cbSize = Len(gNotifyIconData)
    End With
    Call Shell_NotifyIcon(NIM_ADD, gNotifyIconData)
End Function

Public Function gfNotifyIconBalloon(ByRef frmCur As Form, ByVal BalloonInfo As String, _
    ByVal BalloonTitle As String, Optional IconFlag As enmNotifyIconFlag = NIIF_INFO) As Boolean
    '����ͼ�굯��������Ϣ
    With gNotifyIconData
        .dwInfoFlags = IconFlag
        .szInfoTitle = BalloonTitle & vbNullChar
        .szInfo = BalloonInfo & vbNullChar
        .cbSize = Len(gNotifyIconData)
    End With
    Call gfNotifyIconModify(gNotifyIconData)
End Function

Public Function gfNotifyIconDelete(ByRef frmCur As Form) As Boolean
    'ɾ������ͼ��
    Call Shell_NotifyIcon(NIM_DELETE, gNotifyIconData)
End Function

Public Function gfNotifyIconModify(nfIconData As NOTIFYICONDATA) As Boolean
    '�޸�����ͼ����Ϣ
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
                lngLength = LenB(StrConv(lpValue, vbFromUnicode))   '����LenB��StrConv�Ļ�lpValue�ַ������ȶԲ���
                Ret = RegSetValueEx(hKey, lpValueName, 0, lpType, ByVal lpValue, lngLength)
                If Ret = 0 Then
                    gfRegOperate = True
'Debug.Print "W", lpValue, lngLength
                End If
                
            Case Else
                Ret = RegQueryValueEx(hKey, lpValueName, 0, lpType, ByVal 0, lngLength) '��ȡֵ�ĳ���
                If Ret = 0 And lngLength > 0 Then
                    ReDim Buff(lngLength - 1)   '�ض��建���С
                    Ret = RegQueryValueEx(hKey, lpValueName, 0, lpType, Buff(0), lngLength) 'ȡֵ
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
    '��ԭ���յ����ļ���Ϣ
    
    With gArr(sckGet.Index)
        If InStr(strInfo, gVar.PTFileFolder) > 0 Then
            '���ж��ƺ�����Ӧ�ڿͻ����������ϴ��ļ�ʱ�����������д�ȷ��
            
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
    'ִ�г������ļ����ļ���
    '''Call ShellExecute(Me.hwnd, "open", strFile, vbNullString, vbNullString, 1)

    Dim lngRet As Long
    Dim strDir As String
    
    lngRet = ShellExecute(GetDesktopWindow, "open", strFile, vbNullString, vbNullString, vbNormalFocus)

    ' û�й����ĳ���
    If lngRet = SE_ERR_NOASSOC Then
         strDir = Space$(260)
         lngRet = GetSystemDirectory(strDir, Len(strDir))
         strDir = Left$(strDir, lngRet)
       ' ��ʾ�򿪷�ʽ����
         lngRet = ShellExecute(GetDesktopWindow, vbNullString, "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & strFile, strDir, vbNormalFocus)
    End If
    
    If lngRet > 32 Then gfShellExecute = True
    
End Function

Public Function gfStartUpSet() As Boolean
    
    '��������������
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
            '��¼���ÿ����Զ�����ʧ��
            
        End If
    End If
    
End Function

Public Function gfVersionCompare(ByVal strVerCL As String, ByVal strVerSV As String) As String
    '�¾ɰ汾�űȽ�
    Dim ArrCL() As String, ArrSV() As String
    Dim K As Long, C As Long
    
    ArrCL = Split(strVerCL, ".")
    ArrSV = Split(strVerSV, ".")
    K = UBound(ArrCL)
    C = UBound(ArrSV)
    If K = C And K = 3 Then
        For K = 0 To C
            If Not IsNumeric(ArrCL(K)) Then
                gfVersionCompare = "�ͻ��˰汾�쳣"
                Exit For
            End If
            If Not IsNumeric(ArrSV(K)) Then
                gfVersionCompare = "����˰汾�쳣"
                Exit For
            End If
            
            If Val(ArrSV(K)) > Val(ArrCL(K)) Then
                gfVersionCompare = "1" '˵�����°汾
                Exit For
            End If
        Next
        If K = C + 1 Then gfVersionCompare = "0" '˵��û���°棬���ø���
    Else
        If K = 3 And C <> 3 Then
            gfVersionCompare = "����˰汾��ȡ�쳣"
        ElseIf C = 3 And K <> 3 Then
            gfVersionCompare = "�ͻ��˰汾��ȡ�쳣"
        Else
            gfVersionCompare = "�汾��ȡ�쳣"
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
        
        .ServerClose = "�رշ���"
        .ServerError = "�쳣"
        .ServerNotStarted = "δ����"
        .ServerStart = "��������"
        .ServerStarted = "������"
        
        .Connected = "������"
        .DisConnected = "δ����"
        .ConnectError = "�����쳣"
        .ConnectToServer = "��������"
        .DisConnectFromServer = "�Ͽ�����"
        
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
