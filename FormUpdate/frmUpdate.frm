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
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   360
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin FTCUpdate.LabelProgressBar LabelProgressBar1 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   3855
      _ExtentX        =   6800
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
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    
    Call gsInitialize
    
    '检测是否传入命令行参数进来，没有则退出程序
    Dim strCmd As String
    
    strCmd = Command()
    If UCase(strCmd) <> gVar.CmdLineStr Then
        GoTo LineUnload
    End If
    
    Exit Sub
    
LineUnload:
    Unload Me   '此行以下除End Sub不可再跟任何有效代码
End Sub
