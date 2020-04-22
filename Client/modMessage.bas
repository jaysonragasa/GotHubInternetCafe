Attribute VB_Name = "modMessage"
' my un arranged constant variables

Option Explicit

Public LogedIn                          As Boolean

Public Const MAX_CHUNK                  As Long = 4196

Public Const CONN_MESSAGE_PASSWORD      As String = "DEADMOUSE"
Public Const CONN_CONNECT               As String = "1"
Public Const CONN_DISCONNECT            As String = "2"
Public Const CONN_CONNECTED             As String = "3"
Public Const CONN_DISCONNECTED          As String = "4"
Public Const CONN_LOGIN                 As String = "5"
Public Const CONN_LOGOUT                As String = "6"
Public Const CONN_HI                    As String = "7"
Public Const CONN_HELLO                 As String = "8"
Public Const CONN_CANCEL                As String = "9"
Public Const CONN_CHATMSG               As String = "10"
Public Const CONN_ENUMWIN               As String = "11"
Public Const CONN_SENDENUMWIN           As String = "12"
Public Const CONN_CLOSEAPP              As String = "13"
Public Const CONN_REQUESTSTATUS         As String = "14"
Public Const CONN_CAPTURESCREEN         As String = "15"
Public Const CONN_CAPTUREFORM           As String = "16"
Public Const CONN_FILESTAT              As String = "17"
Public Const CONN_LOCKWS                As String = "18"
Public Const CONN_MONITOROFF            As String = "19"
Public Const CONN_SHUTDOWN              As String = "20"
Public Const CONN_MOUSECLICK            As String = "21"
Public Const CONN_TIMESUP               As String = "22"
Public Const CONN_LOGEDIN               As String = "23"
