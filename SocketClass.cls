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

Private mObject As Object '����������
Private mConnectTimeout As Integer '���ӳ�ʱʱ��
Private mReceiveTimeout As Integer  '�������ݳ�ʱʱ��
Private mServerIP As String '����IP
Private mServerPort As String '���Ӷ˿�
Private mTimeType As Integer '��ʱ����


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
    mObject.updateStatus "�������ӷ�������������"
    Exit Sub
ErrHandle:
    Call SetNothing
    mObject.throwException Err
End Sub
 
Private Sub LhSocket__OnConnect()
    On Error GoTo ErrHandle
    LhTimeout_.StopTimer
    LhSocket_.SendData mObject.getSendMsg
    
    mObject.updateStatus "���ӷ������ɹ������ڷ��ͱ��ġ�����"
    
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
    mObject.updateStatus "���ճɹ������ڽ�������"
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
    
    mObject.updateStatus "���ͳɹ������ڽ��������С�������"
    
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
        timeOutinfo = "���ӷ�������ʱ"
    End If
    
    If mTimeType = 2 Then
        timeOutinfo = "�������ݳ�ʱ"
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