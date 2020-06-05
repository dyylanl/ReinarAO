Attribute VB_Name = "mdlPing"
Option Explicit

Private Const ICMP_SUCCESS As Long = 0
Private Const ICMP_STATUS_BUFFER_TO_SMALL = 11001                   'Buffer Too Small
Private Const ICMP_STATUS_DESTINATION_NET_UNREACH = 11002           'Destination Net Unreachable
Private Const ICMP_STATUS_DESTINATION_HOST_UNREACH = 11003          'Destination Host Unreachable
Private Const ICMP_STATUS_DESTINATION_PROTOCOL_UNREACH = 11004      'Destination Protocol Unreachable
Private Const ICMP_STATUS_DESTINATION_PORT_UNREACH = 11005          'Destination Port Unreachable
Private Const ICMP_STATUS_NO_RESOURCE = 11006                       'No Resources
Private Const ICMP_STATUS_BAD_OPTION = 11007                        'Bad Option
Private Const ICMP_STATUS_HARDWARE_ERROR = 11008                    'Hardware Error
Private Const ICMP_STATUS_LARGE_PACKET = 11009                      'Packet Too Big
Private Const ICMP_STATUS_REQUEST_TIMED_OUT = 11010                 'Request Timed Out
Private Const ICMP_STATUS_BAD_REQUEST = 11011                       'Bad Request
Private Const ICMP_STATUS_BAD_ROUTE = 11012                         'Bad Route
Private Const ICMP_STATUS_TTL_EXPIRED_TRANSIT = 11013               'TimeToLive Expired Transit
Private Const ICMP_STATUS_TTL_EXPIRED_REASSEMBLY = 11014            'TimeToLive Expired Reassembly
Private Const ICMP_STATUS_PARAMETER = 11015                         'Parameter Problem
Private Const ICMP_STATUS_SOURCE_QUENCH = 11016                     'Source Quench
Private Const ICMP_STATUS_OPTION_TOO_BIG = 11017                    'Option Too Big
Private Const ICMP_STATUS_BAD_DESTINATION = 11018                   'Bad Destination
Private Const ICMP_STATUS_NEGOTIATING_IPSEC = 11032                 'Negotiating IPSEC
Private Const ICMP_STATUS_GENERAL_FAILURE = 11050                   'General Failure

Public Const WINSOCK_ERROR = "Windows Sockets not responding correctly."
Public Const INADDR_NONE As Long = &HFFFFFFFF
Public Const WSA_SUCCESS = 0
Public Const WS_VERSION_REQD As Long = &H101



Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long

Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal CP As String) As Long


Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long

Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 256) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Long
    iMaxUDPDG As Long
    lpVendorInfo As Long
End Type

Private Declare Function IcmpSendEcho Lib "icmp.dll" _
                                      (ByVal IcmpHandle As Long, _
                                       ByVal DestinationAddress As Long, _
                                       ByVal RequestData As String, _
                                       ByVal RequestSize As Long, _
                                       ByVal RequestOptions As Long, _
                                       ReplyBuffer As ICMP_ECHO_REPLY, _
                                       ByVal ReplySize As Long, _
                                       ByVal Timeout As Long) As Long

Private Type IP_OPTION_INFORMATION
    Ttl As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Byte
    OptionsData As Long
End Type

Public Type ICMP_ECHO_REPLY
    address As Long
    Status As Long
    RoundTripTime As Long
    DataSize As Long
    Reserved As Integer
    ptrData As Long
    Options As IP_OPTION_INFORMATION
    Data As String * 250
End Type
Public Function ping(sAddress As String, Reply As ICMP_ECHO_REPLY) As Long

    Dim hIcmp As Long
    Dim lAddress As Long
    Dim lTimeOut As Long
    Dim StringToSend As String


    StringToSend = "hello"


    lTimeOut = 1000    'ms


    lAddress = inet_addr(sAddress)

    If (lAddress <> -1) And (lAddress <> 0) Then


        hIcmp = IcmpCreateFile()

        If hIcmp Then

            Call IcmpSendEcho(hIcmp, lAddress, StringToSend, Len(StringToSend), 0, Reply, Len(Reply), lTimeOut)


            ping = Reply.Status


            IcmpCloseHandle hIcmp
        Else
            Debug.Print "failure opening icmp handle."
            ping = -1
        End If
    Else
        ping = -1
    End If

End Function
Public Sub SocketsCleanup()

    WSACleanup

End Sub

Public Function SocketsInitialize() As Boolean

    Dim WSAD As WSADATA

    SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = ICMP_SUCCESS

End Function

Public Function EvaluatePingResponse(PingResponse As Long) As String

    Select Case PingResponse


        Case ICMP_SUCCESS: EvaluatePingResponse = "Success!"


        Case ICMP_STATUS_BUFFER_TO_SMALL: EvaluatePingResponse = "Buffer Too Small"
        Case ICMP_STATUS_DESTINATION_NET_UNREACH: EvaluatePingResponse = "Destination Net Unreachable"
        Case ICMP_STATUS_DESTINATION_HOST_UNREACH: EvaluatePingResponse = "Destination Host Unreachable"
        Case ICMP_STATUS_DESTINATION_PROTOCOL_UNREACH: EvaluatePingResponse = "Destination Protocol Unreachable"
        Case ICMP_STATUS_DESTINATION_PORT_UNREACH: EvaluatePingResponse = "Destination Port Unreachable"
        Case ICMP_STATUS_NO_RESOURCE: EvaluatePingResponse = "No Resources"
        Case ICMP_STATUS_BAD_OPTION: EvaluatePingResponse = "Bad Option"
        Case ICMP_STATUS_HARDWARE_ERROR: EvaluatePingResponse = "Hardware Error"
        Case ICMP_STATUS_LARGE_PACKET: EvaluatePingResponse = "Packet Too Big"
        Case ICMP_STATUS_REQUEST_TIMED_OUT: EvaluatePingResponse = "Request Timed Out"
        Case ICMP_STATUS_BAD_REQUEST: EvaluatePingResponse = "Bad Request"
        Case ICMP_STATUS_BAD_ROUTE: EvaluatePingResponse = "Bad Route"
        Case ICMP_STATUS_TTL_EXPIRED_TRANSIT: EvaluatePingResponse = "TimeToLive Expired Transit"
        Case ICMP_STATUS_TTL_EXPIRED_REASSEMBLY: EvaluatePingResponse = "TimeToLive Expired Reassembly"
        Case ICMP_STATUS_PARAMETER: EvaluatePingResponse = "Parameter Problem"
        Case ICMP_STATUS_SOURCE_QUENCH: EvaluatePingResponse = "Source Quench"
        Case ICMP_STATUS_OPTION_TOO_BIG: EvaluatePingResponse = "Option Too Big"
        Case ICMP_STATUS_BAD_DESTINATION: EvaluatePingResponse = "Bad Destination"
        Case ICMP_STATUS_NEGOTIATING_IPSEC: EvaluatePingResponse = "Negotiating IPSEC"
        Case ICMP_STATUS_GENERAL_FAILURE: EvaluatePingResponse = "General Failure"


        Case Else: EvaluatePingResponse = "Unknown Response"

    End Select

End Function

Sub AveriguarTiempoLlegada()
    Dim Reply As ICMP_ECHO_REPLY
    Dim lngSuccess As Long
    Dim strIPAddress As String


    If SocketsInitialize() Then


        strIPAddress = "64.233.161.104"


        lngSuccess = ping(strIPAddress, Reply)

        AddtoRichTextBox frmMain.rectxt, "Ping : " & Reply.RoundTripTime & " ms", 2, 51, 223, 1, 1

        SocketsCleanup

    Else


        Debug.Print WINSOCK_ERROR

    End If
End Sub

