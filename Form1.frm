VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   7605
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   4320
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3450
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim socket As SocketClass
' **************start******************
'网络相关

'发送数据
Function getSendMsg() As String
    
    getSendMsg = "00000402<root><charset>00</charset><version>1.0</version><signType>HEX</signType><clientNo>100120</clientNo><clientType>00</clientType><clientMac>3C-97-0E-67-40-4D</clientMac><clientIP>192.168.2.186</clientIP><service>offlineOrderSearch</service><requestId>3935</requestId><merchantId>888021050450001</merchantId><payCode>111111</payCode><barCode>1</barCode><sign>cfea956de6ae57c3d1445a4905a23bad</sign></root>"
    
End Function
'解析数据
Function parseData(param As String)
    
    MsgBox param
    
    Set socket = Nothing
    
End Function
'更新网络状态
Function updateStatus(param As String)
    
    StatusBar1.SimpleText = param
    
End Function
'更新超时状态
Function upDateTimeOutInfo(param As String)
    
    MsgBox param
    
    Set socket = Nothing
    
End Function
'发送网络异常
Function throwException(param As ErrObject)
    
    Debug.Print param.Description
    
    Set socket = Nothing
    
End Function

'***************end*********************

'测试
Private Sub Command1_Click()
    
    Set socket = New SocketClass
    
    Set socket.clsObject = Me
    
    socket.ServerIp = "192.168.2.239"
    
    socket.ServerPort = "32011"
    
    socket.ConnectTimeout = 5
    
    socket.ReceiveTimeout = 60
    
    socket.SendMessage
    
End Sub




'业务相关



