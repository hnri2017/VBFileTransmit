Attribute VB_Name = "basCommon"
'Ҫ��Winsock�ؼ��ڿͻ��������˶����뽨�����飬����Indexֵ���Ӧ������������±�Ҫ��ͬ

Option Explicit


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

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

Public Const SW_RESTORE = 9
Public Const SW_HIDE = 0

Public Type NOTIFYICONDATA
    cbSize As Long  '�ṹ��С���ֽڣ�
    hWnd As Long    '������Ϣ�Ĵ��ھ��
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
    sNone = 0
    sMSVst = 1
    sMS07 = 2
End Enum

Public Type gtypeCommonVariant  '�Զ��幫�ó���
    TCPIP As String     '������IP��ַ
    TCPPort As Long     '�������˿�
    TCPConnMax As Long  '���������
    ChunkSize As Long   '�ļ�����ʱ�ķֿ��С
    WaitTime As Long    'ÿ���ļ�����ʱ�ĵȴ�ʱ�䣬��λ��
    
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
    
    lngReg = GetSetting(App.Title, gVar.RegSkinSection, gVar.RegSkinKeyFile, 0)
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
        .ApplyWindow frmCur.hWnd
    End With
    
    If lngReg <> lngResource Then Call SaveSetting(App.Title, gVar.RegSkinSection, gVar.RegSkinKeyFile, lngResource)
    
End Function

Public Function gfNotifyIconAdd(ByRef frmCur As Form) As Boolean
    '��������ͼ��
    With gNotifyIconData
        .hWnd = frmCur.hWnd
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

Public Function gfRestoreInfo(ByVal strInfo As String, sckGet As MSWinsockLib.Winsock) As Boolean
    '��ԭ���յ����ļ���Ϣ
    
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
    End With
End Sub
