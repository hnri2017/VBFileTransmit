VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���³���"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   4335
   End
   Begin FTCUpdate.LabelProgressBar LabelProgressBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1920
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Top             =   2520
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   1
      Left            =   1320
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   180
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   2000
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnHide As Boolean     '���´��������ش�ģʽ����ʾ��ģʽ
Dim mblnCheckStart As Boolean   '�ѿ�ʼ����ʶ
Dim mblnUpdateFinish As Boolean     '������ɱ�ʶ


Private Function mfCheckUpdate() As Boolean
    '������
    Dim strFileLoc As String, strFileNet As String, strVerLoc As String, strVerNet As String
    
    strFileLoc = gVar.AppPath & gVar.ClientExeName
    If Not gfDirFile(strFileLoc) Then Exit Function
    strVerLoc = Trim(gfBackVersion(strFileLoc))
    If Len(strVerLoc) = 0 Then Exit Function
    
    If Winsock1.Item(1).State <> 7 Then Exit Function
    Call mfSetText("����������֤�汾�С���", vbBlue)
    Call gfSendInfo(gVar.PTVersionOfClient & strVerLoc, Winsock1.Item(1))
    
End Function

Private Function mfConnect() As Boolean
    Dim strIP As String, strPort As String
    Static lngCount As Long
            
    lngCount = lngCount + 1
    If lngCount = 2 Then
        Call mfSetText("�汾���ʧ�ܣ��޷����ӷ�������" & vbCrLf & _
                       "��ȷ�Ϸ����������������������и��³���", vbRed)
        Exit Function    '���԰ٴκ���������
    End If
    
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

Private Sub mfSetLabel(ByVal strCaption As String, ByVal backColor As Long)
    Label1.Item(1).Caption = strCaption
    Label1.Item(1).backColor = backColor
End Sub

Private Sub mfSetText(ByVal strTxt As String, ByVal ForeColor As Long)
    Text1.Text = strTxt
    Text1.ForeColor = ForeColor
End Sub

Private Function mfShellSetup(ByVal strFile As String) As Boolean
    '�رտͻ��˳���ִ�и��°�װ��
    
    Dim strClient As String
    
    If MsgBox("�Ƿ�����ִ�и��³���", vbQuestion + vbYesNo, "��װѯ��") = vbYes Then
        If gfCloseApp(gVar.ClientExeName) Then  '�رտͻ���exe
            If gfShellExecute(strFile) Then     '���а�װ��
                Unload Me
            End If
        Else
            MsgBox "��ȷ���ѹرտͻ��˳��򣬲��������и��³���", vbInformation, "����"
        End If
    Else
        Unload Me
    End If
End Function


Private Sub Form_Load()
    
    Dim strCmd As String, arrCmd() As String
    
    Label1.Item(0).Caption = ""
    ReDim gArr(0 To 1)
    Call gsInitialize
    
    '����Ƿ��������в���������û�����˳�����
    strCmd = Command
    If Len(strCmd) = 0 Then
        GoTo LineUnload '��ֱֹ���������³��򣬱�����������
    Else
        arrCmd = Split(strCmd, gVar.CmdSeparator)
        
        If UCase(arrCmd(0)) <> UCase(gVar.ClientExeName) Then
            GoTo LineUnload    '��������е�һ���ַ��̶�Ϊexe�ļ�������������Ϊ�Ƿ��������³��򣬲�׼ִ��
        End If
        
        If UBound(arrCmd) > 0 Then  '�ж�����������Ƿ�������ش�������
            If LCase(arrCmd(1)) = LCase(gVar.CmdLineHide) Then
                mblnHide = True
                Me.Hide
            End If
        End If
    End If
    
    
    Text1.backColor = Me.backColor
    Call mfSetLabel(gVar.DisConnected, vbRed)
    Call mfConnect
    Timer1.Interval = 1000
    Timer1.Enabled = True

    Exit Sub
    
LineUnload:
    Unload Me   '�������³�End Sub�����ٸ��κ���Ч����
End Sub

Private Sub Timer1_Timer()
    Const conConn As Byte = 1       '����״̬�����conConn��
    Const conState As Byte = 5      '���ӷ����������conState��
    
    Static byteConn As Byte
    Static byteState As Byte
    Static byteDotCount As Byte
    
    byteConn = byteConn + 1
    byteState = byteState + 1
    
    If byteConn >= conConn Then
        If Winsock1.Item(1).State = 7 Then
            Call mfSetLabel(gVar.Connected, vbGreen)
            gVar.TCPConnected = True
            If Not mblnCheckStart Then
                mblnCheckStart = True
                Call mfCheckUpdate
            End If
        Else
            Call mfSetLabel(gVar.DisConnected, vbRed)
            gVar.TCPConnected = False
        End If
        byteConn = 0    '��λ��̬����
    End If
    
    If byteState >= conState Then
        If Winsock1.Item(1).State <> 7 Then
            If Not mblnUpdateFinish Then Call mfConnect
        End If
        byteState = 0   '��λ��̬����
    End If
    
    If gArr(1).FileTransmitState Then
        byteDotCount = byteDotCount + 1
        If byteDotCount > 6 Then byteDotCount = 1
        Label1.Item(0).Caption = "����������" & String(byteDotCount, "��")
    End If
    
End Sub

Private Sub Winsock1_Close(Index As Integer)
    '���䱻�ر�
    If UBound(gArr) = 1 Then
        gArr(1) = gArr(0)
'Debug.Print "Winsock1_Close trigger all time ?"
    End If
    
    If mblnCheckStart Then
        Call mfSetText("�����������жϣ��汾���¼��ʧ�ܣ�", vbRed)
        mblnCheckStart = False
    End If
    Label1.Item(0).Caption = ""
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    '���շ������˴�����Ϣ���ļ�
    
    Dim strGet As String    '�����ַ���Ϣ
    Dim byteGet() As Byte   '�����ļ�
    
    With gArr(Index)
        If Not .FileTransmitState Then
            '�ַ���Ϣ����״̬��
            
            Winsock1.Item(Index).GetData strGet
            If Not gfRestoreInfo(strGet, Winsock1.Item(Index)) Then
                
            End If
            
            If InStr(strGet, gVar.PTVersionNeedUpdate) > 0 Then
                Dim strVer As String
                
                strVer = Mid(strGet, Len(gVar.PTVersionNeedUpdate) + 1)
                Call mfSetText("�����°棺" & strVer, vbBlue)
            End If
            
            If InStr(strGet, gVar.PTVersionNotUpdate) > 0 Then
                Dim strNot As String
                
                If Len(strGet) = Len(gVar.PTVersionNotUpdate) Then
                    strNot = "����ǰ�İ汾�������°汾������Ҫ���¡�"
                    Call mfSetText(strNot, vbGreen)
                    If mblnHide Then Unload Me  '����ģʽ�򿪸��´���ʱ���޸�����ֱ���˳�
                Else
                    strNot = Mid(strGet, Len(gVar.PTVersionNotUpdate) + 1)
                    strNot = "�汾����쳣��" & strNot
                    Call mfSetText(strNot, vbMagenta)
                End If
                
                mblnUpdateFinish = True
            End If
Debug.Print "Get Server Info:" & strGet, bytesTotal
            '�ַ���Ϣ����״̬��
            
        Else
            '�ļ�����״̬��
            
            If .FileNumber = 0 Then
                .FileNumber = FreeFile
                Open .FilePath For Binary As #.FileNumber
                
                LabelProgressBar1.Min = 0
                LabelProgressBar1.Max = .FileSizeTotal
                LabelProgressBar1.Value = 0
            End If
            
            ReDim byteGet(bytesTotal - 1)
            Winsock1.Item(Index).GetData byteGet, vbArray + vbByte
            Put #.FileNumber, , byteGet
            .FileSizeCompleted = .FileSizeCompleted + bytesTotal
            LabelProgressBar1.Value = .FileSizeCompleted
            
            If .FileSizeCompleted >= .FileSizeTotal Then
                Dim strSetupFile As String
                
                strSetupFile = .FilePath
                Close #.FileNumber
                Call gfSendInfo(gVar.PTFileEnd, Winsock1.Item(Index))
                gArr(Index) = gArr(0)
                Label1.Item(0).Caption = "������ɣ�"
                
                Call mfShellSetup(strSetupFile)
                
Debug.Print "Received Over"
            End If
            
            '�ļ�����״̬��
        End If
    End With
    
End Sub


Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Index <> 0 Then
        If gArr(Index).FileTransmitState Then   '
            Close #gArr(Index).FileNumber
            gArr(Index) = gArr(0)
        End If
    End If
End Sub
