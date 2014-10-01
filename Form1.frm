VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2040
      Top             =   120
   End
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = &HFFFFFFFF
 
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
 
Dim lngPId As Long
Dim lngPHandle As Long
 

Private Sub Command1_Click() '录音
  MMControl1.DeviceType = "WaveAudio"   '打开设备的类型
  MMControl1.FileName = App.Path & "\record2.wav" '零时文件以及位置，'【"d:\record"可以随意确定，看你自己的爱好了】
  MMControl1.Command = "open"     '打开
  MMControl1.Command = "record"  '录音命令---开始录音
  Label1.Caption = "正在录音"
  Timer1.Interval = 1000
  Timer1.Enabled = True

  Command2.Enabled = True
  Command1.Enabled = False
End Sub
Private Sub Command2_Click() '停止
  MMControl1.Command = "stop"
  Timer1.Enabled = False
  miaojishi = 0
  fengjishi = 0
  Label1.Caption = "录音已经停止"
  Command1.Enabled = False
  Command2.Enabled = False
  MMControl1.Command = "save"
  lngPId = Shell("pscp " & Chr(34) & App.Path & "\record2.wav" & Chr(34) & " root@192.168.1.5:/", vbNormalFocus)
  lngPHandle = OpenProcess(SYNCHRONIZE, 0, lngPId)
  If lngPHandle <> 0 Then
    Call WaitForSingleObject(lngPHandle, INFINITE) '无限等待，直到程式结束
    Call CloseHandle(lngPHandle)
  End If
  Shell "putty -pw gerenyinsi -m " & Chr(34) & App.Path & "\play.sh" & Chr(34) & " root@192.168.1.5", vbNormalFocus
End Sub

