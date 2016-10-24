VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SocketClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents LhSocket_ As LhSocket
Attribute LhSocket_.VB_VarHelpID = -1
Private WithEvents LhTimeout_ As LhTimeout
Attribute LhTimeout_.VB_VarHelpID = -1

Private mObject As Object '接收数据类
Private mConnectTimeout As Integer '连接超时时间
Private mReceiveTimeout As Integer  '接收数据超时时间
Private mServerIP As String '主机IP
Private mServerPort As String '连接端口
Private mTimeType As Integer '计时类型


Public Property Set clsObject(object As Object)
    Set mObject = object
End Property

Public Property Let ConnectTimeout(param As Integer)
    mConnectTimeout = param
End Property

Public Property Let ReceiveTimeout(param As Integer)
    mReceiveTimeout = param
End Property

Public Property Let ServerIp(param As String)
    mServerIP = param
End Property

Public Property Let ServerPort(param As Integer)
    mServerPort = param
End Property

Public Sub SendMessage()
    On Error GoTo ErrHandle
    Set LhSocket_ = New LhSocket
    Set LhTimeout_ = New LhTimeout
    LhSocket_.RemoteHost = mServerIP
    LhSocket_.RemotePort = mServerPort
    LhSocket_.Connect
    mTimeType = 1
    LhTimeout_.TimeOutValue = mConnectTimeout
    LhTimeout_.StartTimer
    mObject.updateStatus "正在连接服务器。。。。"
    Exit Sub
ErrHandle:
    Call SetNothing
    mObject.throwException Err
End Sub
 
Private Sub LhSocket__OnConnect()
    On Error GoTo ErrHandle
    LhTimeout_.StopTimer
    LhSocket_.SendData mObject.getSendMsg
    
    mObject.updateStatus "连接服务器成功，正在发送报文。。。"
    
    Exit Sub
ErrHandle:
    Call SetNothing
    mObject.throwException Err
End Sub

Private Sub LhSocket__OnDataArrival(ByVal bytesTotal As Long)
    
    On Error GoTo ErrHandle
    
    Dim RetChar As String
    
    LhSocket_.GetData RetChar, vbString
    LhTimeout_.StopTimer
    LhSocket_.CloseSocket
    mObject.updateStatus "接收成功，正在解析数据"
    mObject.parseData RetChar
    Call SetNothing
    Exit Sub
ErrHandle:
    Call SetNothing
    mObject.throwException Err
End Sub

Private Sub LhSocket__OnSendComplete()
    On Error GoTo ErrHandle
    mTimeType = 2
    LhTimeout_.TimeOutValue = mReceiveTimeout
    LhTimeout_.StartTimer
    
    mObject.updateStatus "发送成功，正在接收数据中。。。。"
    
    Exit Sub
ErrHandle:
    Call SetNothing
    mObject.throwException Err
End Sub

Private Sub LhTimeout__TimeOut()
    On Error GoTo ErrHandle
    LhTimeout_.StopTimer
    Dim timeOutinfo As String
    If mTimeType = 1 Then
        timeOutinfo = "连接服务器超时"
    End If
    
    If mTimeType = 2 Then
        timeOutinfo = "接收数据超时"
    End If
    Call SetNothing
    mObject.upDateTimeOutInfo timeOutinfo
    Exit Sub
ErrHandle:
    Call SetNothing
    
    mObject.throwException Err
    
    
End Sub
Private Sub SetNothing()
    Set LhTimeout_ = Nothing
    Set LhSocket_ = Nothing
End Sub
