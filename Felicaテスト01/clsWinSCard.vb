'*******************************************************************************************************
'       クラス名：clsWinSCard
' コンストラクタ：clsWinSCard()
'       メソッド：getCardID() As Boolean
'     プロパティ：Timeout_MilliSecond As Integer  タイムアウトする時間(ミリ秒)
'                 CardType            As String   カードの種類      (ReadOnly)
'                 IsFelica            As Boolean  Felicaフラグ      (ReadOnly)
'                 IsMifare            As Boolean  Mifeaフラグ       (ReadOnly)
'                 IDm                 As String   FelicaのIDm       (ReadOnly)
'                 PMm                 As String   FelicaのPMm       (ReadOnly)
'                 UID                 As String   MifareのUID       (ReadOnly)
'                 ErrorMsg            As String   エラーメッセージ  (ReadOnly)
'-------------------------------------------------------------------------------------------------------
' winscard.dllの機能を使ってFelicaのIDmやPMm、MifareのUIDを取得する
' 実行前に設定する項目   Timeout_MilliSecond ※設定しない場合、PaSoRiにカードをかざすまで無限に待機する
' 実行後に設定される項目 CardType, IsFelica, IsMifare, IDm, PMm, UID, ErrorMsg
'*******************************************************************************************************

'*********
' 参考URL
'*********
' ソフテックだより 第２０３号（2014年2月5日発行） 技術レポート「非接触ICカード技術"FeliCa(フェリカ)"のIDm読み取り方法」
' http://www.softech.co.jp/mm_140205_pc.htm
' VB.netでPC/SC(NFC通信)してFelicaやMifareを読み取るサンプル
' http://log.windows78.net/2015/03/1295/
' EternalWindows セキュリティ / スマートカード
' http://eternalwindows.jp/security/scard/scard00.html

'******
' MSDN
'******
' プラットフォーム呼び出しのデータ型
' https://msdn.microsoft.com/ja-jp/library/ac7ay120(v=vs.110).aspx
' SCardEstablishContext
' https://msdn.microsoft.com/ja-jp/library/windows/desktop/aa379479(v=vs.85).aspx
' SCardListReaders
' https://msdn.microsoft.com/ja-jp/library/windows/desktop/aa379793(v=vs.85).aspx
' SCardGetStatusChange
' https://msdn.microsoft.com/ja-jp/library/windows/desktop/aa379773(v=vs.85).aspx
' SCardConnect
' https://msdn.microsoft.com/ja-jp/library/windows/desktop/aa379473(v=vs.85).aspx
' SCardStatus
' https://msdn.microsoft.com/ja-jp/library/windows/desktop/aa379803(v=vs.85).aspx
' SCardTransmit
' https://msdn.microsoft.com/ja-jp/library/windows/desktop/aa379804(v=vs.85).aspx
' SCardDisconnect
' https://msdn.microsoft.com/ja-jp/library/windows/desktop/aa379475(v=vs.85).aspx
' SCardFreeMemory
' https://msdn.microsoft.com/ja-jp/library/windows/desktop/aa379488(v=vs.85).aspx
' SCardReleaseContext
' https://msdn.microsoft.com/ja-jp/library/windows/desktop/aa379798(v=vs.85).aspx

Imports System
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Security.Cryptography


Public Class clsWinSCard
    '==========
    ' 定数定義
    '==========
    Private Const SCARD_S_SUCCESS                 As Integer  = 0
    Private Const SCARD_F_INTERNAL_ERROR          As Integer  = &H80100001
    Private Const SCARD_E_CANCELLED               As Integer  = &H80100002
    Private Const SCARD_E_INVALID_HANDLE          As Integer  = &H80100003
    Private Const SCARD_E_INVALID_PARAMETER       As Integer  = &H80100004
    Private Const SCARD_E_INVALID_TARGET          As Integer  = &H80100005
    Private Const SCARD_E_NO_MEMORY               As Integer  = &H80100006
    Private Const SCARD_F_WAITED_TOO_LONG         As Integer  = &H80100007
    Private Const SCARD_E_INSUFFICIENT_BUFFER     As Integer  = &H80100008
    Private Const SCARD_E_UNKNOWN_READER          As Integer  = &H80100009
    Private Const SCARD_E_TIMEOUT                 As Integer  = &H8010000A
    Private Const SCARD_E_SHARING_VIOLATION       As Integer  = &H8010000B
    Private Const SCARD_E_NO_SMARTCARD            As Integer  = &H8010000C
    Private Const SCARD_E_UNKNOWN_CARD            As Integer  = &H8010000D
    Private Const SCARD_E_CANT_DISPOSE            As Integer  = &H8010000E
    Private Const SCARD_E_PROTO_MISMATCH          As Integer  = &H8010000F
    Private Const SCARD_E_NOT_READY               As Integer  = &H80100010
    Private Const SCARD_E_INVALID_VALUE           As Integer  = &H80100011
    Private Const SCARD_E_SYSTEM_CANCELLED        As Integer  = &H80100012
    Private Const SCARD_E_COMM_ERROR              As Integer  = &H80100013
    Private Const SCARD_F_UNKNOWN_ERROR           As Integer  = &H80100014
    Private Const SCARD_E_INVALID_ATR             As Integer  = &H80100015
    Private Const SCARD_E_NOT_TRANSACTED          As Integer  = &H80100016
    Private Const SCARD_E_READER_UNAVAILABLE      As Integer  = &H80100017
    Private Const SCARD_P_SHUTDOWN                As Integer  = &H80100018
    Private Const SCARD_E_PCI_TOO_SMALL           As Integer  = &H80100019
    Private Const SCARD_E_READER_UNSUPPORTED      As Integer  = &H8010001A
    Private Const SCARD_E_DUPLICATE_READER        As Integer  = &H8010001B
    Private Const SCARD_E_CARD_UNSUPPORTED        As Integer  = &H8010001C
    Private Const SCARD_E_NO_SERVICE              As Integer  = &H8010001D
    Private Const SCARD_E_SERVICE_STOPPED         As Integer  = &H8010001E
    Private Const SCARD_E_UNEXPECTED              As Integer  = &H8010001F
    Private Const SCARD_E_ICC_INSTALLATION        As Integer  = &H80100020
    Private Const SCARD_E_ICC_CREATEORDER         As Integer  = &H80100021
    Private Const SCARD_E_UNSUPPORTED_FEATURE     As Integer  = &H80100022
    Private Const SCARD_E_DIR_NOT_FOUND           As Integer  = &H80100023
    Private Const SCARD_E_FILE_NOT_FOUND          As Integer  = &H80100024
    Private Const SCARD_E_NO_DIR                  As Integer  = &H80100025
    Private Const SCARD_E_NO_FILE                 As Integer  = &H80100026
    Private Const SCARD_E_NO_ACCESS               As Integer  = &H80100027
    Private Const SCARD_E_WRITE_TOO_MANY          As Integer  = &H80100028
    Private Const SCARD_E_BAD_SEEK                As Integer  = &H80100029
    Private Const SCARD_E_INVALID_CHV             As Integer  = &H8010002A
    Private Const SCARD_E_UNKNOWN_RES_MNG         As Integer  = &H8010002B
    Private Const SCARD_E_NO_SUCH_CERTIFICATE     As Integer  = &H8010002C
    Private Const SCARD_E_CERTIFICATE_UNAVAILABLE As Integer  = &H8010002D
    Private Const SCARD_E_NO_READERS_AVAILABLE    As Integer  = &H8010002E
    Private Const SCARD_E_COMM_DATA_LOST          As Integer  = &H8010002F
    Private Const SCARD_E_NO_KEY_CONTAINER        As Integer  = &H80100030
    Private Const SCARD_E_SERVER_TOO_BUSY         As Integer  = &H80100031
    Private Const SCARD_E_PIN_CACHE_EXPIRED       As Integer  = &H80100032
    Private Const SCARD_E_NO_PIN_CACHE            As Integer  = &H80100033
    Private Const SCARD_E_READ_ONLY_CARD          As Integer  = &H80100034
    Private Const SCARD_W_UNSUPPORTED_CARD        As Integer  = &H80100065
    Private Const SCARD_W_UNRESPONSIVE_CARD       As Integer  = &H80100066
    Private Const SCARD_W_UNPOWERED_CARD          As Integer  = &H80100067
    Private Const SCARD_W_RESET_CARD              As Integer  = &H80100068
    Private Const SCARD_W_REMOVED_CARD            As Integer  = &H80100069
    Private Const SCARD_W_SECURITY_VIOLATION      As Integer  = &H8010006A
    Private Const SCARD_W_WRONG_CHV               As Integer  = &H8010006B
    Private Const SCARD_W_CHV_BLOCKED             As Integer  = &H8010006C
    Private Const SCARD_W_EOF                     As Integer  = &H8010006D
    Private Const SCARD_W_CANCELLED_BY_USER       As Integer  = &H8010006E
    Private Const SCARD_W_CARD_NOT_AUTHENTICATED  As Integer  = &H8010006F
    Private Const SCARD_W_CACHE_ITEM_NOT_FOUND    As Integer  = &H80100070
    Private Const SCARD_W_CACHE_ITEM_STALE        As Integer  = &H80100071
    Private Const SCARD_W_CACHE_ITEM_TOO_BIG      As Integer  = &H80100072
    Private Const SCARD_PROTOCOL_T0               As Integer  = 1
    Private Const SCARD_PROTOCOL_T1               As Integer  = 2
    Private Const SCARD_PROTOCOL_RAW              As Integer  = 4
    Private Const SCARD_SCOPE_USER                As UInteger = 0
    Private Const SCARD_SCOPE_TERMINAL            As UInteger = 1
    Private Const SCARD_SCOPE_SYSTEM              As UInteger = 2
    Private Const SCARD_STATE_UNAWARE             As Integer  = &H0
    Private Const SCARD_STATE_IGNORE              As Integer  = &H1
    Private Const SCARD_STATE_CHANGED             As Integer  = &H2
    Private Const SCARD_STATE_UNKNOWN             As Integer  = &H4
    Private Const SCARD_STATE_UNAVAILABLE         As Integer  = &H8
    Private Const SCARD_STATE_EMPTY               As Integer  = &H10
    Private Const SCARD_STATE_PRESENT             As Integer  = &H20
    Private Const SCARD_STATE_AIRMATCH            As Integer  = &H40
    Private Const SCARD_STATE_EXCLUSIVE           As Integer  = &H80
    Private Const SCARD_STATE_INUSE               As Integer  = &H100
    Private Const SCARD_STATE_MUTE                As Integer  = &H200
    Private Const SCARD_STATE_UNPOWERED           As Integer  = &H400
    Private Const SCARD_SHARE_EXCLUSIVE           As Integer  = &H1
    Private Const SCARD_SHARE_SHARED              As Integer  = &H2
    Private Const SCARD_SHARE_DIRECT              As Integer  = &H3
    Private Const SCARD_LEAVE_CARD                As Integer  = 0
    Private Const SCARD_RESET_CARD                As Integer  = 1
    Private Const SCARD_UNPOWER_CARD              As Integer  = 2
    Private Const SCARD_EJECT_CARD                As Integer  = 3

    Private Const SCARD_UNKNOWN                   As Integer  = 0
    Private Const SCARD_ABSENT                    As Integer  = 1
    Private Const SCARD_PRESENT                   As Integer  = 2
    Private Const SCARD_SWALLOWED                 As Integer  = 3
    Private Const SCARD_POWERED                   As Integer  = 4
    Private Const SCARD_NEGOTIABLE                As Integer  = 5
    Private Const SCARD_SPECIFIC                  As Integer  = 6

    '========================================
    ' SCardGetStatusChange()で使用する構造体
    '========================================
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Private Structure SCARD_READERSTATE
        <MarshalAs(UnmanagedType.LPTStr)> _
        Public szReader       As String
        Public pvUserData     As IntPtr
        Public dwCurrentState As UInteger
        Public dwEventState   As UInteger
        Public cbAtr          As UInteger
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=36)> _
        Public rgbAtr()       As Byte

        Sub Initialize()
            szReader       = ""
            pvUserData     = 0
            dwCurrentState = 0
            dwEventState   = 0
            cbAtr          = 0
            ReDim rgbAtr(35)
        End Sub
    End Structure

    '=================================
    ' SCardTransmit()で使用する構造体
    '=================================
    <StructLayout(LayoutKind.Sequential)> _
    Private Structure SCARD_IO_REQUEST
        Public dwProtocol  As UInteger
        Public cbPciLength As UInteger
    End Structure

    '==============
    ' WinSCard API
    '==============
    <DllImport("winscard.dll", EntryPoint:="SCardEstablishContext", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function SCardEstablishContext( _
        ByVal dwScope     As UInteger, _
        ByVal pvReserved1 As IntPtr, _
        ByVal pvReserved2 As IntPtr, _
        ByRef phContext   As IntPtr _
    ) As UInteger
    End Function

    <DllImport("winscard.dll", EntryPoint:="SCardListReaders", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function SCardListReaders( _
        ByVal hContext    As IntPtr, _
        ByVal mszGroups   As Byte(), _
        ByVal mszReaders  As Byte(), _
        ByRef pcchReaders As UInt32 _
    ) As UInteger
    End Function

    <DllImport("winscard.dll", EntryPoint:="SCardGetStatusChange", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function SCardGetStatusChange( _
        ByVal hContext       As IntPtr, _
        ByVal dwTimeout      As Integer, _
        ByRef rgReaderStates As SCARD_READERSTATE, _
        ByVal cReaders       As Integer _
    ) As UInteger
    End Function

    <DllImport("winscard.dll", EntryPoint:="SCardConnect", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function SCardConnect( _
        ByVal hContext             As IntPtr, _
        ByVal szReader             As String, _
        ByVal dwShareMode          As UInteger, _
        ByVal dwPreferredProtocols As UInteger, _
        ByRef phCard               As IntPtr, _
        ByRef pdwActiveProtocol    As IntPtr _
    ) As UInteger
    End Function

    <DllImport("winscard.dll", EntryPoint:="SCardStatus", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function SCardStatus( _
        ByVal hCard         As IntPtr, _
        ByVal szReaderName  As StringBuilder, _
        ByRef pcchReaderLen As UInteger, _
        ByRef pdwState      As UInteger, _
        ByRef pdwProtocol   As UInteger, _
        ByVal pbAtr         As Byte(), _
        ByRef pcbAtrLen     As Integer _
    ) As UInteger
    End Function

    <DllImport("winscard.dll", EntryPoint:="SCardTransmit", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function SCardTransmit( _
        ByVal hCard          As IntPtr, _
        ByVal pioSendRequest As IntPtr, _
        ByVal SendBuff       As Byte(), _
        ByVal SendBuffLen    As Integer, _
        ByRef pioRecvRequest As SCARD_IO_REQUEST, _
        ByVal RecvBuff       As Byte(), _
        ByRef RecvBuffLen    As Integer _
    ) As UInteger
    End Function

    <DllImport("winscard.dll", EntryPoint:="SCardDisconnect", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function SCardDisconnect( _
        ByVal hCard       As IntPtr, _
        ByVal Disposition As Integer _
    ) As UInteger
    End Function

    '<DllImport("winscard.dll", EntryPoint:="SCardFreeMemory", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)> _
    'Private Shared Function SCardFreeMemory( _
    '    ByVal hContext As IntPtr, _
    '    ByVal pvMem    As IntPtr _
    ') As UInteger
    'End Function


    <DllImport("winscard.dll", EntryPoint:="SCardControl", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function SCardControl(ByVal hCard As IntPtr,
        ByVal controlCode As Integer,
        ByVal inBuffer As Byte(),
        ByVal inBufferLen As Integer,
        ByVal outBuffer As Byte(),
        ByVal outBufferLen As Integer,
        ByRef bytesReturned As Integer) As UInteger
    End Function






    <DllImport("winscard.dll", EntryPoint:="SCardReleaseContext", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function SCardReleaseContext(ByVal phContext As IntPtr) As UInteger
    End Function

    '===========
    ' Win32 API
    '===========
    <DllImport("kernel32.dll", EntryPoint:="LoadLibrary", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function LoadLibrary(ByVal lpFileName As String) As IntPtr
    End Function

    <DllImport("kernel32.dll", EntryPoint:="FreeLibrary", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Sub FreeLibrary(ByVal handle As IntPtr)
    End Sub

    <DllImport("kernel32.dll", EntryPoint:="GetProcAddress", CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function GetProcAddress(ByVal handle As IntPtr, ByVal procName As String) As IntPtr
    End Function

    '=========================
    ' エラーメッセージ取得(1)
    '=========================
    Private Function GetErrorMessage(ByVal errNo As UInteger) As String
        Dim strMessage As String = vbNullString

        Select Case errNo
            Case GetUInteger(SCARD_F_INTERNAL_ERROR)
                strMessage = "[Error] " + errNo.ToString + " INTERNAL ERROR"
            Case GetUInteger(SCARD_E_CANCELLED)
                strMessage = "[Error] " + errNo.ToString + " CANCELLED"
            Case GetUInteger(SCARD_E_INVALID_HANDLE)
                strMessage = "[Error] " + errNo.ToString + " INVALID HANDLE"
            Case GetUInteger(SCARD_E_INVALID_PARAMETER)
                strMessage = "[Error] " + errNo.ToString + " INVALID PARAMETER"
            Case GetUInteger(SCARD_E_INVALID_TARGET)
                strMessage = "[Error] " + errNo.ToString + " INVALID TARGET"
            Case GetUInteger(SCARD_E_NO_MEMORY)
                strMessage = "[Error] " + errNo.ToString + " NO MEMORY"
            Case GetUInteger(SCARD_F_WAITED_TOO_LONG)
                strMessage = "[Error] " + errNo.ToString + " WAITED TOO LONG"
            Case GetUInteger(SCARD_E_INSUFFICIENT_BUFFER)
                strMessage = "[Error] " + errNo.ToString + " INSUFFICIENT BUFFER"
            Case GetUInteger(SCARD_E_UNKNOWN_READER)
                strMessage = "[Error] " + errNo.ToString + " UNKNOWN READER"
            Case GetUInteger(SCARD_E_TIMEOUT)
                strMessage = "[Error] " + errNo.ToString + " TIMEOUT"
            Case GetUInteger(SCARD_E_SHARING_VIOLATION)
                strMessage = "[Error] " + errNo.ToString + " SHARING VIOLATION"
            Case GetUInteger(SCARD_E_NO_SMARTCARD)
                strMessage = "[Error] " + errNo.ToString + " NO SMARTCARD"
            Case GetUInteger(SCARD_E_UNKNOWN_CARD)
                strMessage = "[Error] " + errNo.ToString + " UNKNOWN CARD"
            Case GetUInteger(SCARD_E_CANT_DISPOSE)
                strMessage = "[Error] " + errNo.ToString + " CANT DISPOSE"
            Case GetUInteger(SCARD_E_PROTO_MISMATCH)
                strMessage = "[Error] " + errNo.ToString + " PROTO MISMATCH"
            Case GetUInteger(SCARD_E_NOT_READY)
                strMessage = "[Error] " + errNo.ToString + " NOT READY"
            Case GetUInteger(SCARD_E_INVALID_VALUE)
                strMessage = "[Error] " + errNo.ToString + " INVALID VALUE"
            Case GetUInteger(SCARD_E_SYSTEM_CANCELLED)
                strMessage = "[Error] " + errNo.ToString + " SYSTEM CANCELLED"
            Case GetUInteger(SCARD_E_COMM_ERROR)
                strMessage = "[Error] " + errNo.ToString + " COMM ERROR"
            Case GetUInteger(SCARD_F_UNKNOWN_ERROR)
                strMessage = "[Error] " + errNo.ToString + " UNKNOWN ERROR"
            Case GetUInteger(SCARD_E_INVALID_ATR)
                strMessage = "[Error] " + errNo.ToString + " INVALID ATR"
            Case GetUInteger(SCARD_E_NOT_TRANSACTED)
                strMessage = "[Error] " + errNo.ToString + " NOT TRANSACTED"
            Case GetUInteger(SCARD_E_READER_UNAVAILABLE)
                strMessage = "[Error] " + errNo.ToString + " READER UNAVAILABLE"
            Case GetUInteger(SCARD_P_SHUTDOWN)
                strMessage = "[Error] " + errNo.ToString + " SHUTDOWN"
            Case GetUInteger(SCARD_E_PCI_TOO_SMALL)
                strMessage = "[Error] " + errNo.ToString + " PCI TOO SMALL"
            Case GetUInteger(SCARD_E_READER_UNSUPPORTED)
                strMessage = "[Error] " + errNo.ToString + " READER UNSUPPORTED"
            Case GetUInteger(SCARD_E_DUPLICATE_READER)
                strMessage = "[Error] " + errNo.ToString + " DUPLICATE READER"
            Case GetUInteger(SCARD_E_CARD_UNSUPPORTED)
                strMessage = "[Error] " + errNo.ToString + " CARD UNSUPPORTED"
            Case GetUInteger(SCARD_E_NO_SERVICE)
                strMessage = "[Error] " + errNo.ToString + " NO SERVICE"
            Case GetUInteger(SCARD_E_SERVICE_STOPPED)
                strMessage = "[Error] " + errNo.ToString + " SERVICE STOPPED"
            Case GetUInteger(SCARD_E_UNEXPECTED)
                strMessage = "[Error] " + errNo.ToString + " UNEXPECTED"
            Case GetUInteger(SCARD_E_ICC_INSTALLATION)
                strMessage = "[Error] " + errNo.ToString + " ICC INSTALLATION"
            Case GetUInteger(SCARD_E_ICC_CREATEORDER)
                strMessage = "[Error] " + errNo.ToString + " ICC CREATEORDER"
            Case GetUInteger(SCARD_E_UNSUPPORTED_FEATURE)
                strMessage = "[Error] " + errNo.ToString + " UNSUPPORTED FEATURE"
            Case GetUInteger(SCARD_E_DIR_NOT_FOUND)
                strMessage = "[Error] " + errNo.ToString + " DIR NOT FOUND"
            Case GetUInteger(SCARD_E_FILE_NOT_FOUND)
                strMessage = "[Error] " + errNo.ToString + " FILE NOT FOUND"
            Case GetUInteger(SCARD_E_NO_DIR)
                strMessage = "[Error] " + errNo.ToString + " NO DIR"
            Case GetUInteger(SCARD_E_NO_FILE)
                strMessage = "[Error] " + errNo.ToString + " NO FILE"
            Case GetUInteger(SCARD_E_NO_ACCESS)
                strMessage = "[Error] " + errNo.ToString + " NO ACCESS"
            Case GetUInteger(SCARD_E_WRITE_TOO_MANY)
                strMessage = "[Error] " + errNo.ToString + " WRITE TOO MANY"
            Case GetUInteger(SCARD_E_BAD_SEEK)
                strMessage = "[Error] " + errNo.ToString + " BAD SEEK"
            Case GetUInteger(SCARD_E_INVALID_CHV)
                strMessage = "[Error] " + errNo.ToString + " INVALID CHV"
            Case GetUInteger(SCARD_E_UNKNOWN_RES_MNG)
                strMessage = "[Error] " + errNo.ToString + " UNKNOWN RES MNG"
            Case GetUInteger(SCARD_E_NO_SUCH_CERTIFICATE)
                strMessage = "[Error] " + errNo.ToString + " NO SUCH CERTIFICATE"
            Case GetUInteger(SCARD_E_CERTIFICATE_UNAVAILABLE)
                strMessage = "[Error] " + errNo.ToString + " CERTIFICATE UNAVAILABLE"
            Case GetUInteger(SCARD_E_NO_READERS_AVAILABLE)
                strMessage = "[Error] " + errNo.ToString + " NO READERS AVAILABLE"
            Case GetUInteger(SCARD_E_COMM_DATA_LOST)
                strMessage = "[Error] " + errNo.ToString + " COMM DATA LOST"
            Case GetUInteger(SCARD_E_NO_KEY_CONTAINER)
                strMessage = "[Error] " + errNo.ToString + " NO KEY CONTAINER"
            Case GetUInteger(SCARD_E_SERVER_TOO_BUSY)
                strMessage = "[Error] " + errNo.ToString + " SERVER TOO BUSY"
            Case GetUInteger(SCARD_E_PIN_CACHE_EXPIRED)
                strMessage = "[Error] " + errNo.ToString + " PIN CACHE EXPIRED"
            Case GetUInteger(SCARD_E_NO_PIN_CACHE)
                strMessage = "[Error] " + errNo.ToString + " NO PIN CACHE"
            Case GetUInteger(SCARD_E_READ_ONLY_CARD)
                strMessage = "[Error] " + errNo.ToString + " READ ONLY CARD"
            Case GetUInteger(SCARD_W_UNSUPPORTED_CARD)
                strMessage = "[Error] " + errNo.ToString + " UNSUPPORTED CARD"
            Case GetUInteger(SCARD_W_UNRESPONSIVE_CARD)
                strMessage = "[Error] " + errNo.ToString + " UNRESPONSIVE CARD"
            Case GetUInteger(SCARD_W_UNPOWERED_CARD)
                strMessage = "[Error] " + errNo.ToString + " UNPOWERED CARD"
            Case GetUInteger(SCARD_W_RESET_CARD)
                strMessage = "[Error] " + errNo.ToString + " RESET CARD"
            Case GetUInteger(SCARD_W_REMOVED_CARD)
                strMessage = "[Error] " + errNo.ToString + " REMOVED CARD"
            Case GetUInteger(SCARD_W_SECURITY_VIOLATION)
                strMessage = "[Error] " + errNo.ToString + " SECURITY VIOLATION"
            Case GetUInteger(SCARD_W_WRONG_CHV)
                strMessage = "[Error] " + errNo.ToString + " WRONG CHV"
            Case GetUInteger(SCARD_W_CHV_BLOCKED)
                strMessage = "[Error] " + errNo.ToString + " CHV BLOCKED"
            Case GetUInteger(SCARD_W_EOF)
                strMessage = "[Error] " + errNo.ToString + " EOF"
            Case GetUInteger(SCARD_W_CANCELLED_BY_USER)
                strMessage = "[Error] " + errNo.ToString + " CANCELLED BY USER"
            Case GetUInteger(SCARD_W_CARD_NOT_AUTHENTICATED)
                strMessage = "[Error] " + errNo.ToString + " CARD NOT AUTHENTICATED"
            Case GetUInteger(SCARD_W_CACHE_ITEM_NOT_FOUND)
                strMessage = "[Error] " + errNo.ToString + " CACHE ITEM NOT FOUND"
            Case GetUInteger(SCARD_W_CACHE_ITEM_STALE)
                strMessage = "[Error] " + errNo.ToString + " CACHE ITEM STALE"
            Case GetUInteger(SCARD_W_CACHE_ITEM_TOO_BIG)
                strMessage = "[Error] " + errNo.ToString + " CACHE ITEM TOO BIG"
            Case Else
                strMessage = "[Error] " + errNo.ToString + " OTHER ERROR"
        End Select

        Return strMessage
    End Function

    '=========================
    ' エラーメッセージ取得(2)
    '=========================
    Private Function GetErrorMessage_SCardStatus(ByVal errNo As UInteger) As String
        Dim strMessage As String = vbNullString

        Select Case errNo
            Case GetUInteger(SCARD_UNKNOWN)
                strMessage = "[Error] " + errNo.ToString + " The card state is unknown or unexpected."
            Case GetUInteger(SCARD_ABSENT)
                strMessage = "[Error] " + errNo.ToString + " There is no card in the reader."
            Case GetUInteger(SCARD_PRESENT)
                strMessage = "[Error] " + errNo.ToString + " There is a card in the reader, but it has not been moved into position for use."
            Case GetUInteger(SCARD_SWALLOWED)
                strMessage = "[Error] " + errNo.ToString + " There is a card in the reader in position for use. The card is not powered."
            Case GetUInteger(SCARD_POWERED)
                strMessage = "[Error] " + errNo.ToString + " Power is being provided to the card, but the reader driver is unaware of the mode of the card."
            Case GetUInteger(SCARD_NEGOTIABLE)
                strMessage = "[Error] " + errNo.ToString + " The card has been reset and is awaiting PTS negotiation."
            Case Else
                strMessage = "[Error] " + errNo.ToString + " OTHER ERROR"
        End Select

        Return strMessage
    End Function

    '==========================================
    ' 符号あり整数型から、符号なし整数型に変換
    '==========================================
    Private Function GetUInteger(ByVal p_Value As Integer) As UInteger
        Return p_Value + 2 ^ 32
    End Function

    '================
    ' プロパティ宣言
    '================
    Private _Timeout_MilliSecond As Integer  'タイムアウトする時間(ミリ秒)
    Private _CardType As String   'カードの種類
    Private _IsFelica As Boolean  'Felicaフラグ
    Private _IsMifare As Boolean  'Mifareフラグ
    Private _IDm As String   'FelicaのIDm
    Private bIDm(7) As Byte
    Private _PMm As String   'FelicaのPMm
    Private bPMm(7) As Byte
    Private _UID As String   'MifareのUID
    Private _S_PAD0 As String
    Private _ErrorMsg As String   'エラーメッセージ
    Private Mac(7) As Byte
    Private Mac_A(7) As Byte
    Private Ck(15) As Byte, Ck1(7) As Byte, Ck2(7) As Byte
    Private Rc(15) As Byte, Rc1(7) As Byte, Rc2(7) As Byte
    Private Sk1(7) As Byte, Sk2(7) As Byte
    Private Iv(7) As Byte, Key(23) As Byte
    Private CardMasterKey(23) As Byte
    Private CardMasterKeyString As String

    Public Property Timeout_MilliSecond() As Integer
        Get
            Return _Timeout_MilliSecond
        End Get
        Set(value As Integer)
            _Timeout_MilliSecond = value
        End Set
    End Property

    Public ReadOnly Property CardType() As String
        Get
            Return _CardType
        End Get
    End Property

    Public ReadOnly Property IsFelica() As Boolean
        Get
            Return _IsFelica
        End Get
    End Property

    Public ReadOnly Property IsMifare() As Boolean
        Get
            Return _IsMifare
        End Get
    End Property

    Public ReadOnly Property IDm() As String
        Get
            Return _IDm
        End Get
    End Property

    Public ReadOnly Property PMm() As String
        Get
            Return _PMm
        End Get
    End Property

    Public ReadOnly Property UID() As String
        Get
            Return _UID
        End Get
    End Property

    Public Property S_PAD0() As String
        Get
            Return _S_PAD0
        End Get
        Set(value As String)
            _S_PAD0 = value
        End Set
    End Property

    Public ReadOnly Property ErrorMsg() As String
        Get
            Return _ErrorMsg
        End Get
    End Property

    '============
    ' メンバ変数
    '============
    Private hContext As IntPtr
    Private mReader As String
    Private hCard As IntPtr

    'Private SCARD_PCI_T0  As IntPtr
    Private SCARD_PCI_T1 As IntPtr
    Private SCARD_PCI_RAW As IntPtr

    Private sendBuffer_IDm As Byte() = New Byte() {&HFF, &HCA, &H0, &H0, &H0}  'Felica IDm取得コマンド
    Private sendBuffer_PMm As Byte() = New Byte() {&HFF, &HCA, &H1, &H0, &H0}  'Felica PMm取得コマンド
    Private sendBuffer_UID As Byte() = New Byte() {&HFF, &HCA, &H0, &H0, &H4}  'Mifare UID取得コマンド
    Private sendBuffer_S_PAD0 As Byte() = New Byte() {&HFF, &HB0, &H0, &H0, &H3, &H0, &H0, &H10}
    Private writeBuffer_SPAD0 As Byte() = New Byte() {&HFF, &HD6, &H0, &H0, 16, &H30, &H31, &H32, &H33, _
                                                      &H34, &H36, &H37, &H38, &H39, &H2A, &H2A, &H2A, &H2A, &H2A, &H2A, &H2A, &H0} 'UPDATE BINARY
    Private clearBuffer_SPAD0 As Byte() = New Byte() {&HFF, &HD6, &H0, &H0, 16, &H0, &H0, &H0, &H0, _
                                                      &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0} 'UPDATE BINARY



    Private URtn As UInteger

    '================
    ' コンストラクタ
    '================
    Public Sub New(ByVal CardMasterKeyString As String)
        '--------
        ' 初期化
        '--------
        Dim hLoader As IntPtr = LoadLibrary("winscard.dll")
        'SCARD_PCI_T0  = GetProcAddress(hLoader, "g_rgSCardT0Pci")
        SCARD_PCI_T1 = GetProcAddress(hLoader, "g_rgSCardT1Pci")
        SCARD_PCI_RAW = GetProcAddress(hLoader, "g_rgSCardRawPci")
        FreeLibrary(hLoader)

        Me._Timeout_MilliSecond = Timeout.Infinite
        Me._CardType = String.Empty
        Me._IsFelica = False
        Me._IsMifare = False
        Me._IDm = String.Empty
        Me._PMm = String.Empty
        Me._UID = String.Empty
        Me._S_PAD0 = String.Empty
        Me._ErrorMsg = String.Empty
        Me.Iv = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        'Me.CardMasterKeyString = "GBRC kanyama 2020"
        Me.CardMasterKey = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, _
                                 &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, _
                                 &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0 _
                                 }

        Dim chr() As Byte
        If CardMasterKeyString = "" Then
            CardMasterKeyString = "GBRC 2020"
        End If
        If CardMasterKeyString.Length > 24 Then
            CardMasterKeyString = Left(CardMasterKeyString, 24)
        End If
        chr = System.Text.Encoding.UTF8.GetBytes(CardMasterKeyString)
        Array.Copy(chr, 0, Me.CardMasterKey, 0, chr.Length)


    End Sub

    '======================================
    ' (1)SCardEstablishContext
    ' ⇒リソースマネージャのハンドルを取得
    '======================================
    Private Function EstablishContext() As Boolean
        URtn = SCardEstablishContext(SCARD_SCOPE_USER, Nothing, Nothing, Me.hContext)
        If URtn <> SCARD_S_SUCCESS Then
            Me._ErrorMsg = GetErrorMessage(URtn)
            Return False
        End If

        Return True
    End Function

    '==================================
    ' (2)SCardListReaders
    ' ⇒リーダー／ライターの名称を取得
    '==================================
    Private Function ListReaders() As Boolean
        Dim pcchReaders As UInteger = 0
        Dim mszReaders  As Byte()

        URtn = SCardListReaders(Me.hContext, Nothing, Nothing, pcchReaders)
        If URtn <> SCARD_S_SUCCESS Then
            Me._ErrorMsg = GetErrorMessage(URtn)
            Return False
        End If

        mszReaders = New Byte(Convert.ToInt32(pcchReaders) * 2 - 1) {}
        URtn = SCardListReaders(Me.hContext, Nothing, mszReaders, pcchReaders)
        If URtn <> SCARD_S_SUCCESS Then
            Me._ErrorMsg = GetErrorMessage(URtn)
            Return False
        End If

        '取得したカードリーダーの名前をNULLで分割
        Me.mReader = ""
        Dim reader As String() = (New UnicodeEncoding).GetString(mszReaders).Split(vbNullChar)
        For I As Integer = 0 To reader.Length() - 1
            'PaSoRiの文字があるか確認
            If reader(I).IndexOf("pasori", StringComparison.InvariantCultureIgnoreCase) >= 0 Then
                Me.mReader = reader(I)
                Exit For
            End If
        Next

        If Me.mReader = "" Then
            Me._ErrorMsg = "[Error] PaSoRi NOT FOUND."
            Return False
        End If

        Return True
    End Function

    '==================================
    ' (3)SCardGetStatusChange
    ' ⇒リーダー／ライターの状態を取得
    '==================================
    Private Function GetStatusChange() As Boolean
        Dim srReaderState As New SCARD_READERSTATE
        srReaderState.Initialize()
        srReaderState.szReader = Me.mReader
        srReaderState.dwCurrentState = SCARD_STATE_UNAWARE

        URtn = SCardGetStatusChange(Me.hContext, 100, srReaderState, 1)
        If URtn <> SCARD_S_SUCCESS Then
            Me._ErrorMsg = GetErrorMessage(URtn)
            Return False
        End If

        If (srReaderState.dwEventState And SCARD_STATE_EMPTY) <> 0 Then
            srReaderState.dwCurrentState = srReaderState.dwEventState

            'URtn = SCardGetStatusChange(hContext, Timeout.Infinite, srReaderState, 1)    '待ち時間、無限
            URtn = SCardGetStatusChange(hContext, _Timeout_MilliSecond, srReaderState, 1) '待ち時間を指定して実行
            If URtn <> SCARD_S_SUCCESS Then
                Me._ErrorMsg = GetErrorMessage(URtn)
                Return False
            End If

            If (srReaderState.dwEventState And SCARD_STATE_PRESENT) <> 0 Then
                'カードがセットされました
                Return True
            ElseIf (srReaderState.dwEventState And SCARD_STATE_UNAVAILABLE) <> 0 Then
                'カードリーダが外されました
                Me._ErrorMsg = "[Error] Card is Unavailable."
                Return False
            Else
            End If
        ElseIf (srReaderState.dwEventState And SCARD_STATE_PRESENT) <> 0 Then
            'カードは既にセットされています
            Return True
        Else
        End If

        Me._ErrorMsg = "[Error] Unknown error at GetStatusChange()."
        Return False
    End Function

    '=================
    ' (4)SCardConnect
    ' ⇒カードなしで接続
    '=================
    Private Function ConnectDirect() As Boolean
        Dim pdwActiveProtocol As IntPtr = IntPtr.Zero
        Dim time As UInteger = 0

        'カードに接続してみる
        URtn = SCardConnect(Me.hContext, _
                            Me.mReader, _
                            SCARD_SHARE_DIRECT, _
                            SCARD_SCOPE_USER, _
                            Me.hCard, _
                            pdwActiveProtocol)
        'エラーの場合
        If URtn <> SCARD_S_SUCCESS Then
            Me._ErrorMsg = GetErrorMessage(URtn)
            Return False
        End If

        Return True
    End Function





    '=================
    ' (4)SCardConnect
    ' ⇒カードに接続
    '=================
    Private Function Connect() As Boolean
        Dim pdwActiveProtocol As IntPtr = IntPtr.Zero
        Dim time As UInteger = 0

        'カードに接続してみる
        URtn = SCardConnect(Me.hContext, _
                            Me.mReader, _
                            SCARD_SHARE_SHARED, _
                            SCARD_PROTOCOL_T1, _
                            Me.hCard, _
                            pdwActiveProtocol)
        'エラーの場合
        If URtn <> SCARD_S_SUCCESS Then
            Me._ErrorMsg = GetErrorMessage(URtn)
            Return False
        End If

        Return True
    End Function

    '================================================
    ' (5)SCardStatus
    ' ⇒ATR(Answer To Reset)を取得しカード種別を判定
    '================================================
    Private Function Status() As Boolean
        Dim szReaderName As StringBuilder
        Dim pcchReaderLen As UInteger = 0
        Dim dwState As UInteger = 0
        Dim pdwProtocol As IntPtr = IntPtr.Zero
        Dim pbAtr As Byte()
        Dim dwAtrLen As UInteger = 0

        pcchReaderLen = 0
        dwAtrLen = 0
        URtn = SCardStatus(
                   Me.hCard, _
                   Nothing, _
                   pcchReaderLen, _
                   dwState, _
                   pdwProtocol, _
                   Nothing, _
                   dwAtrLen)
        If URtn = SCARD_S_SUCCESS Then
            If dwState <> SCARD_SPECIFIC Then
                Me._ErrorMsg = GetErrorMessage_SCardStatus(dwState)
                Return False
            End If
        Else
            Me._ErrorMsg = GetErrorMessage(URtn)
            Return False
        End If

        szReaderName = New StringBuilder(Convert.ToInt32(pcchReaderLen))
        pbAtr = New Byte(dwAtrLen - 1) {}
        URtn = SCardStatus(
                   Me.hCard, _
                   szReaderName, _
                   pcchReaderLen, _
                   dwState, _
                   pdwProtocol, _
                   pbAtr, _
                   dwAtrLen)
        If URtn = SCARD_S_SUCCESS Then
            If dwState <> SCARD_SPECIFIC Then
                Me._ErrorMsg = GetErrorMessage_SCardStatus(dwState)
                Return False
            End If
        Else
            Me._ErrorMsg = GetErrorMessage(URtn)
            Return False
        End If

        ' ATRのデータ長をチェック
        If pbAtr.Length() < 15 Then
            Me._ErrorMsg = "[Error] ATR data is too short."
            Return False
        End If

        ' カードの種類をチェック
        '--------------------------
        ' 00 01: MIFARE Classic 1K
        ' 00 02: MIFARE Classic 4K
        ' 00 03: MIFARE Ultralight
        ' 00 26: MIFARE Mini
        ' 00 3B: FeliCa
        '--------------------------
        If pbAtr(13) = &H0 AndAlso pbAtr(14) = &H1 Then
            Me._CardType = "Mifare Classic 1K"
            Me._IsMifare = True
        ElseIf pbAtr(13) = &H0 AndAlso pbAtr(14) = &H2 Then
            Me._CardType = "Mifare Classic 4K"
            Me._IsMifare = True
        ElseIf pbAtr(13) = &H0 AndAlso pbAtr(14) = &H3 Then
            Me._CardType = "Mifare Ultralight"
            Me._IsMifare = True
        ElseIf pbAtr(13) = &H0 AndAlso pbAtr(14) = &H26 Then
            Me._CardType = "Mifare Mini"
            Me._IsMifare = True
        ElseIf pbAtr(13) = &H0 AndAlso pbAtr(14) = &H3B Then
            Me._CardType = "Felica"
            Me._IsFelica = True
        Else
            Me._ErrorMsg = "[Error] Card is not Felica or Mifare."
            Return False
        End If

        Return True
    End Function

    '======================================
    ' (6)SCardTransmit
    ' ⇒カードにコマンド送信・データを受信
    '--------------------------------------
    '【第１引数】[in]  送信するコマンド
    '【第２引数】[out] 受信用のバッファ
    '【第３引数】[out] 受信データの長さ
    '======================================
    Private Function Transmit(ByVal sendBuffer As Byte(), ByRef recvBuffer As Byte(), ByRef recvBufferLen As Integer) As Boolean
        ReDim recvBuffer(511)
        recvBufferLen = recvBuffer.Length
        Dim pioRecvRequest As SCARD_IO_REQUEST = New SCARD_IO_REQUEST()
        pioRecvRequest.cbPciLength = 255

        URtn = SCardTransmit(Me.hCard, _
                             Me.SCARD_PCI_T1, _
                             sendBuffer, _
                             sendBuffer.Length, _
                             pioRecvRequest, _
                             recvBuffer, _
                             recvBufferLen)
        If URtn <> SCARD_S_SUCCESS Then
            Me._ErrorMsg = GetErrorMessage(URtn)
            Return False
        End If

        Return True
    End Function

    '======================================
    ' (6)SCardTransmit
    ' ⇒カードにコマンド送信・データを受信
    '--------------------------------------
    '【第１引数】[in]  送信するコマンド
    '【第２引数】[out] 受信用のバッファ
    '【第３引数】[out] 受信データの長さ
    '======================================
    Private Function TransmitDirect(ByVal sendBuffer As Byte(), ByRef recvBuffer As Byte(), ByRef recvBufferLen As Integer) As Boolean
        ReDim recvBuffer(511)
        recvBufferLen = recvBuffer.Length
        Dim pioRecvRequest As SCARD_IO_REQUEST = New SCARD_IO_REQUEST()
        pioRecvRequest.cbPciLength = 255

        URtn = SCardTransmit(Me.hCard, _
                             Me.SCARD_PCI_RAW, _
                             sendBuffer, _
                             sendBuffer.Length, _
                             pioRecvRequest, _
                             recvBuffer, _
                             recvBufferLen)
        If URtn <> SCARD_S_SUCCESS Then
            Me._ErrorMsg = GetErrorMessage(URtn)
            Return False
        End If

        Return True
    End Function

    '======================================
    ' (6)SCardControl
    ' ⇒カードまたはＲ／Ｗにコマンド送信・データを受信
    '--------------------------------------
    '【第１引数】[in]  送信するコマンド
    '【第２引数】[out] 受信用のバッファ
    '【第３引数】[out] 受信データの長さ
    '======================================
    Private Function Control(ByVal sendBuffer As Byte(), ByRef recvBuffer As Byte(), ByRef recvBufferLen As Integer) As Boolean
        ReDim recvBuffer(511)
        Dim controlCode As Integer = &H3136B0   ' 固定
        Dim bytesReturned As Integer
        recvBufferLen = recvBuffer.Length
        'Dim pioRecvRequest As SCARD_IO_REQUEST = New SCARD_IO_REQUEST()
        'pioRecvRequest.cbPciLength = 255

        URtn = SCardControl(Me.hCard, controlCode, sendBuffer, sendBuffer.Length, recvBuffer, recvBufferLen, bytesReturned)

        If URtn <> SCARD_S_SUCCESS Then
            Me._ErrorMsg = GetErrorMessage(URtn)
            Return False
        End If

        Return True
    End Function



    '========================
    ' (7)SCardDisconnect
    ' ⇒カードとの通信を切断
    '========================
    Private Sub Disconnect()
        If Me.hCard <> IntPtr.Zero Then
            SCardDisconnect(Me.hCard, SCARD_LEAVE_CARD)
        End If
    End Sub

    '======================================
    ' (8)SCardReleaseContext
    ' ⇒リソースマネージャのハンドルを解放
    '======================================
    Private Sub ReleaseContext()
        If Me.hContext <> IntPtr.Zero Then
            SCardReleaseContext(Me.hContext)
        End If
    End Sub

    '======================
    ' カードのIDを取得する
    '======================
    Public Function getCardID() As Boolean
        '------------------
        ' プロパティ初期化
        '------------------
        Me._CardType = String.Empty
        Me._IsFelica = False
        Me._IsMifare = False
        Me._IDm = String.Empty
        Me.bIDm = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Me._PMm = String.Empty
        Me.bPMm = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Me._UID = String.Empty
        Me._S_PAD0 = String.Empty
        Me._ErrorMsg = String.Empty
        Me.Mac = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Me.Mac_A = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Me.Ck1 = New Byte() {&H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7}
        Me.Ck2 = New Byte() {&H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF}
        'Me.Rc1 = New Byte() {&H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7}
        'Me.Rc2 = New Byte() {&H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF}

        '------------------------------------
        ' FelicaのIDm,PMm・MifareのUIDを取得
        '------------------------------------
        Try
            ' (1)SCardEstablishContext
            ' ⇒リソースマネージャのハンドルを取得
            If Not Me.EstablishContext() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (2)SCardListReaders
            ' ⇒リーダー／ライターの名称を取得
            If Not Me.ListReaders() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If


            ' (3)SCardConnect
            ' ⇒カードなしでＲ／Ｗに接続
            If Not Me.ConnectDirect() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (4)SessionStart
            ' ⇒コマンドモードのセッション開始
            If Not Me.SessionStart() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (5)ChangeFelicaProtocol
            ' ⇒Felicaのプロトコルに変更
            If Not Me.ChangeFelicaProtocol() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (6)FelicaPolling
            ' ⇒Felicaカードのポーリングの開始（IDmとPMmの読み取り）
            Dim SysCode As Byte() = New Byte() {&H88, &HB4}
            If Not Me.FelicaPolling(SysCode) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (7)FelicaDumpBinary
            ' ⇒Felicaカードの内容をダンプ
            Dim A As String = ""
            If Not Me.FelicaDumpBinary(A) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If
            Me._S_PAD0 = A

            ' (8)SessionEnd
            ' ⇒コマンドモードのセッション終了
            If Not Me.SessionEnd() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (9)SCardDisconnect
            ' ⇒カードとの通信を切断
            Me.Disconnect()

            ' (10)SCardReleaseContext
            ' ⇒リソースマネージャのハンドルを解放
            Me.ReleaseContext()

            Return True
        Catch ex As Exception
            Me._ErrorMsg = "[Error] " + ex.Message
            Return False
        End Try

    End Function



    '======================
    ' カードのIDを取得する
    '======================
    Public Function getDataWithMac_A() As Boolean
        Dim chr1 As Byte() = New Byte(0) {}
        Dim chrlen As Integer
        '------------------
        ' プロパティ初期化
        '------------------
        Me._CardType = String.Empty
        Me._IsFelica = False
        Me._IsMifare = False
        Me._IDm = String.Empty
        Me.bIDm = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Me._PMm = String.Empty
        Me.bPMm = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Me._UID = String.Empty
        Me._S_PAD0 = String.Empty
        Me._ErrorMsg = String.Empty
        Me.Mac = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Me.Mac_A = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        'Me.Ck1 = New Byte() {&H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7}
        'Me.Ck2 = New Byte() {&H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF}
        'Me.Rc1 = New Byte() {&H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7}
        'Me.Rc2 = New Byte() {&H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF}

        '------------------------------------
        ' FelicaのIDm,PMm・MifareのUIDを取得
        '------------------------------------
        Try
            ' (1)SCardEstablishContext
            ' ⇒リソースマネージャのハンドルを取得
            If Not Me.EstablishContext() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (2)SCardListReaders
            ' ⇒リーダー／ライターの名称を取得
            If Not Me.ListReaders() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If


            ' (3)SCardConnect
            ' ⇒カードなしでＲ／Ｗに接続
            If Not Me.ConnectDirect() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (4)SessionStart
            ' ⇒コマンドモードのセッション開始
            If Not Me.SessionStart() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (5)ChangeFelicaProtocol
            ' ⇒Felicaのプロトコルに変更
            If Not Me.ChangeFelicaProtocol() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (6)FelicaPolling
            ' ⇒Felicaカードのポーリングの開始（IDmとPMmの読み取り）
            Dim SysCode As Byte() = New Byte() {&H88, &HB4}
            If Not Me.FelicaPolling(SysCode) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' カードに書かれたIDと暗号鍵から個別暗号鍵を作成し、Me.Ck()変数にコピーする。
            ' この個別暗号鍵と同じものが事前にカードに書き込まれている。
            If Not Me.CalcCKBlock() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' 乱数Me.Rc1とMe.Rc2を作成し、カードのRcブロックに書き込む。
            ' 上記のMe.Ck()から新しい暗号鍵SessionKeyを作成する。
            ' カード内部でも同じSessionKeyが計算されてMac_Aの計算に使用される。
            If Not Me.MakeSessionKey() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            Dim A As String = ""

            For i As Int16 = 0 To 13
                ' Rc1を暗号化の初期値に設定する。
                Me.Iv = Me.Rc1.Reverse.ToArray

                ' Mac_A付き読込を行う（暗号化が一致しない場合はエラーとなる）。
                If Not Me.FelicaReadBinaryWithMac_A(i, chr1, chrlen) Then
                    Me.Disconnect()
                    Me.ReleaseContext()
                    A += "Mac Error"
                    Return False
                End If


                A += "|" + i.ToString("X2") + "h|"
                A += " S_PAD" + i.ToString("D2") + " "
                A += "|" + BitConverter.ToString(chr1, 0, chrlen - 2) + "|"
                For j As Integer = 0 To 15
                    If chr1(j) = &H0 Then
                        A += " "
                    Else
                        If chr1(j) >= 20 And chr1(j) <= 126 Then
                            A += Chr(chr1(j))
                        Else
                            A += "･"
                        End If
                    End If
                Next
                A += "|" + vbNewLine
            Next

            'If Not Me.FelicaReadBinaryWithMac(1, chr1, chrlen) Then
            '    Me.Disconnect()
            '    Me.ReleaseContext()
            '    A += "Error"
            '    Return False
            'End If
            'For j As Integer = 0 To 15
            '    If chr1(j) = &H0 Then
            '        A += " "
            '    Else
            '        If chr1(j) >= 20 And chr1(j) <= 126 Then
            '            A += Chr(chr1(j))
            '        Else
            '            A += "･"
            '        End If
            '    End If
            'Next
            'A += vbNewLine


            ' (7)FelicaDumpBinary
            ' ⇒Felicaカードの内容をダンプ
            'If Not Me.FelicaDumpBinary(A) Then
            '    Me.Disconnect()
            '    Me.ReleaseContext()
            '    Return False
            'End If
            Me._S_PAD0 = A

            ' (8)SessionEnd
            ' ⇒コマンドモードのセッション終了
            If Not Me.SessionEnd() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (9)SCardDisconnect
            ' ⇒カードとの通信を切断
            Me.Disconnect()

            ' (10)SCardReleaseContext
            ' ⇒リソースマネージャのハンドルを解放
            Me.ReleaseContext()

            Return True
        Catch ex As Exception
            Me._ErrorMsg = "[Error] " + ex.Message
            Return False
        End Try

    End Function




    '======================
    ' カードの個別化カード鍵を作成し、カードのCkブロックに書き込む
    '======================
    Public Function makeCardKey() As Boolean
        Dim ID As Byte() = New Byte(0) {}
        'Dim IDlen As Integer
        '------------------
        ' プロパティ初期化
        '------------------
        Me._CardType = String.Empty
        Me._IsFelica = False
        Me._IsMifare = False
        Me._IDm = String.Empty
        Me.bIDm = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Me._PMm = String.Empty
        Me.bPMm = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Me._UID = String.Empty
        Me._S_PAD0 = String.Empty
        Me._ErrorMsg = String.Empty
        Me.Mac = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Me.Mac_A = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        'Me.Ck1 = New Byte() {&H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7}
        'Me.Ck2 = New Byte() {&H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF}
        'Me.Rc1 = New Byte() {&H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7}
        'Me.Rc2 = New Byte() {&H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF}

        '------------------------------------
        ' FelicaのIDm,PMm・MifareのUIDを取得
        '------------------------------------
        Try
            ' (1)SCardEstablishContext
            ' ⇒リソースマネージャのハンドルを取得
            If Not Me.EstablishContext() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (2)SCardListReaders
            ' ⇒リーダー／ライターの名称を取得
            If Not Me.ListReaders() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If


            ' (3)SCardConnect
            ' ⇒カードなしでＲ／Ｗに接続
            If Not Me.ConnectDirect() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (4)SessionStart
            ' ⇒コマンドモードのセッション開始
            If Not Me.SessionStart() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (5)ChangeFelicaProtocol
            ' ⇒Felicaのプロトコルに変更
            If Not Me.ChangeFelicaProtocol() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (6)FelicaPolling
            ' ⇒Felicaカードのポーリングの開始（IDmとPMmの読み取り）
            Dim SysCode As Byte() = New Byte() {&H88, &HB4}
            If Not Me.FelicaPolling(SysCode) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (7)個別化カード鍵の作成
            If Not Me.CalcCKBlock() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (8)個別化カード鍵の書込
            If Not Me.FelicaUpdateBinary(&H87, Ck, Ck.Length) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (8)SessionEnd
            ' ⇒コマンドモードのセッション終了
            If Not Me.SessionEnd() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (9)SCardDisconnect
            ' ⇒カードとの通信を切断
            Me.Disconnect()

            ' (10)SCardReleaseContext
            ' ⇒リソースマネージャのハンドルを解放
            Me.ReleaseContext()

            Return True
        Catch ex As Exception
            Me._ErrorMsg = "[Error] " + ex.Message
            Return False
        End Try

    End Function


    Private Function CalcCKBlock() As Boolean
        ' 個別化カード鍵の作成

        ' カードからID番号を読み込む
        Dim ID(15) As Byte, IDlen As Integer
        If Not Me.FelicaReadBinary(&H82, ID, IDlen) Then
            Me.Disconnect()
            Me.ReleaseContext()
            Return False
        End If

        Dim Iv0 As Byte() = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Dim M1 As Byte() = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Dim Mm1 As Byte()
        Dim M2 As Byte() = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Dim K As Byte() = Me.CardMasterKey
        Dim K1 As Byte()

        Dim L As Byte()
        Dim B1 As Byte() = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H1B}
        Dim B2 As Byte() = New Byte() {&H80, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Dim C1 As Byte()
        Dim C2 As Byte()
        Dim T1 As Byte()
        Dim T2 As Byte()
        Dim T(15) As Byte

        ' 2key TripleDESオブジェクトの作成。Kはマスター鍵から作成したものを使用。Ivは{0}
        Dim des1 As New cTripleDES(K, Iv0)
        L = des1.Encrypt(Iv0)

        ReDim Preserve L(7)

        If L(0) And &H80 = &H80 Then
            K1 = Me.BitShift(L, 1)
            K1 = Me.ByteXor(K1, B1)
        Else
            K1 = Me.BitShift(L, 1)
        End If

        Array.Copy(ID, 0, M1, 0, 8)
        Array.Copy(ID, 8, M2, 0, 8)
        Mm1 = Me.ByteXor(M1, B2)
        M2 = Me.ByteXor(M2, K1)

        des1.Iv = Iv0
        C1 = des1.Encrypt(M1)
        ReDim Preserve C1(7)
        des1.Iv = C1
        T1 = des1.Encrypt(M2)
        ReDim Preserve T1(7)

        des1.Iv = Iv0
        C2 = des1.Encrypt(Mm1)
        ReDim Preserve C2(7)
        des1.Iv = C2
        T2 = des1.Encrypt(M2)
        ReDim Preserve T2(7)

        Array.Copy(T1, 0, T, 0, 8)
        Array.Copy(T2, 0, T, 8, 8)

        Me.Ck1 = T1
        Me.Ck2 = T2
        Me.Ck = T       ' 個別化カード鍵

        Return True

    End Function


    Private Function MakeSessionKey() As Boolean
        ' セッション鍵の作成

        ' ８バイトの乱数を２つ作成する。
        Dim rng As New RNGCryptoServiceProvider()
        rng.GetBytes(Me.Rc1)
        rng.GetBytes(Me.Rc2)
        ' ２つの乱数を１６バイトの１つにまとめる。
        Array.Copy(Me.Rc1, 0, Me.Rc, 0, 8)
        Array.Copy(Me.Rc2, 0, Me.Rc, 8, 8)
        ' felicaのRCブロック(&H80)に乱数を書き込む。（これは電源が切れると消去される。
        Dim cklen As Integer = 16
        If Not Me.FelicaUpdateBinary(&H80, Me.Rc, cklen) Then
            Me.Disconnect()
            Me.ReleaseContext()
            Return False
        End If

        Dim temp1 As Byte(), temp2 As Byte()
        ' &H0×8を初期値とする。
        Me.Iv = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        ' Ck1とCk2から24バイトのキーを作成する。Ck1とCk2は事前にFilicaのCkブロック(&H87)に書き込まれていること。
        Array.Copy(Me.Ck1.Reverse().ToArray(), 0, Me.Key, 0, 8)
        Array.Copy(Me.Ck2.Reverse().ToArray(), 0, Me.Key, 8, 8)
        Array.Copy(Me.Ck1.Reverse().ToArray(), 0, Me.Key, 16, 8)

        ' 2key Ttiple TDESオブジェクトdes1の作成する（keyとIvを与えて初期化）
        Dim des1 As New cTripleDES(Me.Key, Me.Iv)

        ' 乱数Rc1のバイトオーダーを反転させ、それをdes1を用いて暗号化する。
        temp1 = des1.Encrypt(Rc1.Reverse().ToArray())
        ' 結果は16バイトになるので前半の8バイトだけを残す。
        ReDim Preserve temp1(7)
        '暗号のバイトオーダーを反転させ、セッション鍵Sk1に保存する。
        Me.Sk1 = temp1.Reverse.ToArray

        ' 上記の計算結果を初期値にする。
        des1.Iv = temp1
        ' 乱数Rc2のバイトオーダーを反転させ、それをdes1を用いて暗号化する。
        temp2 = des1.Encrypt(Me.Rc2.Reverse.ToArray)
        ' 結果は16バイトになるので前半の8バイトだけを残す
        ReDim Preserve temp2(7)
        '暗号のバイトオーダーを反転させ、セッション鍵Sk2に保存する。
        Me.Sk2 = temp2.Reverse.ToArray

        ' Sk1とSk2から24バイトのキーを作成する。
        Array.Copy(Me.Sk1.Reverse().ToArray(), 0, Me.Key, 0, 8)
        Array.Copy(Me.Sk2.Reverse().ToArray(), 0, Me.Key, 8, 8)
        Array.Copy(Me.Sk1.Reverse().ToArray(), 0, Me.Key, 16, 8)

        Return True

    End Function


    '======================
    ' Mac付きリード
    '======================
    Private Function FelicaReadBinaryWithMac(ByVal adr As UInt16, ByRef chr As Byte(), ByRef chrlen As Integer) As Boolean

        Dim temp1 As Byte(), temp2 As Byte()
        Dim des1 As New cTripleDES(Me.Key, Me.Rc1.Reverse.ToArray)

        'des1.Key = Me.Key
        'des1.Iv = Me.Rc1.Reverse.ToArray

        If Not Me.FelicaReadBinary(adr, chr, chrlen) Then
            Me.Disconnect()
            Me.ReleaseContext()
            Return False
        End If

        Dim data1(7) As Byte, data2(7) As Byte, data3(7) As Byte, data4(7) As Byte
        data3 = Me.Mac

        Array.Copy(chr, 0, data1, 0, 8)
        Array.Copy(chr, 8, data2, 0, 8)

        temp1 = des1.Encrypt(data1.Reverse().ToArray())
        ReDim Preserve temp1(7)

        des1.Iv = temp1
        temp2 = des1.Encrypt(data2.Reverse.ToArray)
        ReDim Preserve temp2(7)

        'Me.Iv = temp2
        data4 = temp2.Reverse.ToArray

        For i As Integer = 0 To data3.Length - 1
            If data3(i) <> data4(i) Then
                Return False
            End If
        Next

        Return True

    End Function


    '======================
    ' Mac_A付きリード
    '======================
    Private Function FelicaReadBinaryWithMac_A(ByVal adr As UInt16, ByRef chr As Byte(), ByRef chrlen As Integer) As Boolean

        Dim temp1 As Byte(), temp2 As Byte()
        Dim des1 As New cTripleDES(Me.Key, Me.Rc1.Reverse.ToArray)

        Dim commandFeliCaRead As Byte() = New Byte() {&HFF, &HC2, &H0, &H1, _
                                      &H14, _
                                      &H95, _
                                      &H12, &H12, _
                                      &H6, _
                                      &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, _
                                      &H1, _
                                      &HB, &H0, _
                                      &H2, _
                                      &H80, &H0, _
                                      &H80, &H91, _
                                      &H0}
        '               &HFF, &HC2, &H0, &H1   ：直接Ｒ／Ｗにカード固有のコマンドを送るモード
        '               &H14                   ：以降のコマンドの長さ（バイト数、最後の&H0は無視）この場合は20個
        '               &H95                   ：送信及び受信コマンド
        '               &H12, &H12             ：Ｒ／Ｗに送信するコマンドの長さ（２回繰り返す、最後の&H0は無視）
        '               &H6                    ：FelicaのRead Without Encryptionコマンド
        '               &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0：カードのIDm（Pollingコマンドで取得したものをここに上書きする)
        '               &H1                    ：サービス数（通常は1）
        '               &HB, &H0               ：サービスコード（読込のみ:&HB, &H0、読み書き:&H9, &H0）
        '               &H1                    ：読み込むブロック数（この場合は１）
        '               &H80, &H0              ：読み込むブロック番号リスト（この場合は０番のブロック）複数の場合はここに追加する。
        '               &H80, &H91             ：Mac_Aのブロック番号。
        '               &H0                    ：最後は必ず&H0を入れる


        Dim recvBuffer As Byte() = New Byte(0) {}
        Dim recvBufferLen As Integer

        'コマンドにPollingで読み込んだIDmを書き込む
        Array.Copy(Me.bIDm, 0, commandFeliCaRead, 9, 8)
        ' 読み出すブロックの番号をコマンドに書き込む
        commandFeliCaRead(22) = adr And &HFF
        ' FelicaからブロックデータをMac_A付きで読み出す。
        If Not Me.Control(commandFeliCaRead, recvBuffer, recvBufferLen) Then
            Me.Disconnect()
            Me.ReleaseContext()
            Return False
        End If
        ReDim chr(15)
        chrlen = 16
        ' 目的のブロックのデータをchrにコピー
        Array.Copy(recvBuffer, 27, chr, 0, 16)
        ' Mac_AのデータをMe.Mac_Aにコピー
        Array.Copy(recvBuffer, 43, Me.Mac_A, 0, 8)

        ' 暗号化の初期値を作成
        Dim BB As Byte() = New Byte() {&HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &HFF}
        BB(0) = adr And &HFF
        BB(1) = &H0
        BB(2) = &H91
        BB(3) = &H0
        ' 初期値のバイトオーダーを反転させてIvに設定
        des1.Iv = BB.Reverse.ToArray
        ' 乱数Rc1を暗号化
        temp1 = des1.Encrypt(Me.Rc1.Reverse().ToArray())
        ' 結果は16バイトになるので前半の8バイトだけを残す。
        ReDim Preserve temp1(7)

        Dim data1(7) As Byte, data2(7) As Byte, data3(7) As Byte, data4(7) As Byte
        ' Mac_Aをdata3にコピー
        data3 = Me.Mac_A

        ' 読み込んだブロックデータを8バイトのdata1とdata2に分割
        Array.Copy(chr, 0, data1, 0, 8)
        Array.Copy(chr, 8, data2, 0, 8)

        ' 上記で求めた結果を初期値に設定
        des1.Iv = temp1
        ' data1のバイトオーダーを反転させて、それを暗号化する
        temp1 = des1.Encrypt(data1.Reverse().ToArray())
        ' 結果は16バイトになるので前半の8バイトだけを残す。
        ReDim Preserve temp1(7)

        ' 上記で求めた結果を初期値に設定
        des1.Iv = temp1
        ' data2のバイトオーダーを反転させて、それを暗号化する
        temp2 = des1.Encrypt(data2.Reverse.ToArray)
        ' 結果は16バイトになるので前半の8バイトだけを残す。
        ReDim Preserve temp2(7)

        '結果のバイトオーダーを反転させたものが計算結果（この値がMac_Aを一致すればOK）
        data4 = temp2.Reverse.ToArray

        ' バイト毎に一致するかどうかをチェックする（ひとつでも異なればFalseをReturnして終了）
        For i As Integer = 0 To data3.Length - 1
            If data3(i) <> data4(i) Then
                Return False
            End If
        Next

        ' Felicaから読み込んだMac_Aと計算したMac_Aが等しいのでTrueをReturn
        Return True

    End Function




    Private Function SessionStart() As Boolean
        Dim commandSessionStart As Byte() = New Byte() {&HFF, &HC2, &H0, &H0, &H2, &H81, &H0, &H0}
        Dim recvBuffer As Byte() = New Byte(0) {}
        Dim recvBufferLen As Integer

        If Not Me.Control(commandSessionStart, recvBuffer, recvBufferLen) Then
            Me.Disconnect()
            Me.ReleaseContext()
            Return False
        End If
        Return True
    End Function

    Private Function SessionEnd() As Boolean
        Dim commandSessionEnd As Byte() = New Byte() {&HFF, &HC2, &H0, &H0, &H2, &H82, &H0, &H0}
        Dim recvBuffer As Byte() = New Byte(0) {}
        Dim recvBufferLen As Integer

        If Not Me.Control(commandSessionEnd, recvBuffer, recvBufferLen) Then
            Me.Disconnect()
            Me.ReleaseContext()
            Return False
        End If
        Return True
    End Function

    Private Function ChangeFelicaProtocol() As Boolean

        Dim commandSwitchFeliCa As Byte() = New Byte() {&HFF, &HC2, &H0, &H2, &H4, &H8F, &H2, &H3, &H0, &H0}
        '                                                &HFF, &HC2, &H0, &H2:Switch Protocal Command（コマンドのモード切替）
        '                                                &H4：以降のコマンドの長さ（バイト数、最後の&H0は無視）
        '                                                &H8F：Switch Protocal コマンド
        '                                                &H2, 以降のコマンドの長さ（バイト数、最後の&H0は無視）
        '                                                &H3：Felicaのコマンドモードに指定
        '                                                &H0, レイヤーの切替（0:レイヤーなしの場合）
        '                                                &H0 :最後は必ず&H0を入れる

        Dim recvBuffer As Byte() = New Byte(0) {}
        Dim recvBufferLen As Integer

        If Not Me.Control(commandSwitchFeliCa, recvBuffer, recvBufferLen) Then
            Me.Disconnect()
            Me.ReleaseContext()
            Return False
        End If

        Return True

    End Function


    Private Function FelicaPolling(ByRef SysCode As Byte()) As Boolean

        Dim commandFeliCaPolling As Byte() = New Byte() {&HFF, &HC2, &H0, &H1, &H8, &H95, &H6, &H6, &H0, &H88, &HB4, &H0, &H3, &H0}
        '           &HFF, &HC2, &H0, &H1    ：直接Ｒ／Ｗにカード固有のコマンドを送るモード
        '           &H8                     ：以降のコマンドの長さ（バイト数、最後の&H0は無視）
        '           &H95                    ：送信及び受信コマンド
        '           &H6, &H6                ：Ｒ／Ｗに送信するコマンドの長さ（２回繰り返す、最後の&H0は無視）
        '           &H0                     ：FelicaのPollingコマンド
        '           &H88, &HB4              ：Felica Lite-Sのシステムコード（&HFF,&HFFとするとすべてのカードが反応する）
        '           &H0                     ：リクエストコード（0:要求無し、1:システムコード要求、2:通信性能要求）
        '           &H3                     ：応答可能な最大スロット数の指定（マニュアル参照）
        '           &H0                     ：最後は必ず&H0を入れる

        Dim recvBuffer As Byte() = New Byte(0) {}
        Dim recvBufferLen As Integer

        '指定のシステムコードの書き込み
        Array.Copy(SysCode, 0, commandFeliCaPolling, 9, 2)

        Do
            If Not Me.Control(commandFeliCaPolling, recvBuffer, recvBufferLen) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If
            If recvBuffer(3) = &H90 Then Exit Do
            ' 指定のシステムコードのカードが接触するまで永遠に繰り返す
        Loop

        Array.Copy(recvBuffer, 16, Me.bIDm, 0, 8)
        Array.Copy(recvBuffer, 24, Me.bPMm, 0, 8)
        Me._IDm = BitConverter.ToString(Me.bIDm, 0, 8).Replace("-", String.Empty)
        Me._PMm = BitConverter.ToString(Me.bPMm, 0, 8).Replace("-", String.Empty)
        Me._IsFelica = True
        Me._CardType = "Felica"

        Return True

    End Function


    Private Function FelicaReadBinary(ByVal adr As UInt16, ByRef chr As Byte(), ByRef chrlen As Integer) As Boolean

        Dim commandFeliCaRead As Byte() = New Byte() {&HFF, &HC2, &H0, &H1, _
                                              &H14, _
                                              &H95, _
                                              &H12, &H12, _
                                              &H6, _
                                              &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, _
                                              &H1, _
                                              &HB, &H0, _
                                              &H2, _
                                              &H80, &H0, _
                                              &H80, &H81, _
                                              &H0}
        ''               &HFF, &HC2, &H0, &H1   ：直接Ｒ／Ｗにカード固有のコマンドを送るモード
        ''               &H12                   ：以降のコマンドの長さ（バイト数、最後の&H0は無視）この場合は18個
        ''               &H95                   ：送信及び受信コマンド
        ''               &H10, &H10             ：Ｒ／Ｗに送信するコマンドの長さ（２回繰り返す、最後の&H0は無視）
        ''               &H6                    ：FelicaのRead Without Encryptionコマンド
        ''               &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0：カードのIDm（Pollingコマンドで取得したものをここに上書きする)
        ''               &H1                    ：サービス数（通常は1）
        ''               &HB, &H0               ：サービスコード（読込のみ:&HB, &H0、読み書き:&H9, &H0）
        ''               &H1                    ：読み込むブロック数（この場合は１）
        ''               &H80, &H0              ：読み込むブロック番号リスト（この場合は０番のブロック）複数の場合はここに追加する。
        ''               &H0                    ：最後は必ず&H0を入れる


        Dim recvBuffer As Byte() = New Byte(0) {}
        Dim recvBufferLen As Integer
        Array.Copy(Me.bIDm, 0, commandFeliCaRead, 9, 8)
        commandFeliCaRead(22) = adr And &HFF

        If Not Me.Control(commandFeliCaRead, recvBuffer, recvBufferLen) Then
            Me.Disconnect()
            Me.ReleaseContext()
            Return False
        End If
        ReDim chr(15)
        chrlen = 16
        Array.Copy(recvBuffer, 27, chr, 0, 16)
        Array.Copy(recvBuffer, 43, Me.Mac, 0, 8)

        Return True
    End Function


    Private Function readBinary(ByVal adr As UInt16, ByRef chr As Byte(), ByRef chrlen As Integer) As Boolean
        Dim sendBuffer As Byte() = New Byte() {&HFF, &HB0, &H80, &H1, &H3, &H0, &H0, &H0, &H10}

        Dim recvBuffer As Byte() = New Byte(0) {}
        Dim recvBufferLen As Integer

        sendBuffer(6) = adr And &HFF
        sendBuffer(7) = (adr >> 8) And &HFF
        If Not Me.Transmit(sendBuffer, recvBuffer, recvBufferLen) Then
            Me.Disconnect()
            Me.ReleaseContext()
            Return False
        End If
        chr = recvBuffer
        chrlen = recvBufferLen
        Return True
    End Function

    Private Function updateBinary(ByVal adr As UInt16, ByRef chr As Byte(), ByRef chrlen As Integer) As Boolean

        'Private writeBuffer_SPAD0 As Byte() = New Byte() {&HFF, &HD6, &H0, &H0, 16, &H30, &H31, &H32, &H33, _
        '                                          &H34, &H36, &H37, &H38, &H39, &H2A, &H2A, &H2A, &H2A, &H2A, &H2A, &H2A, &H0} 'UPDATE BINARY
        'Private clearBuffer_SPAD0 As Byte() = New Byte() {&HFF, &HD6, &H0, &H0, 16, &H0, &H0, &H0, &H0, _
        '                                                  &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0} 'UPDATE BINARY

        Dim sendBuffer(255) As Byte   '{&HFF, &HD6, &H80, &H1, &H3, &H0, &H0, &H0, &H10}
        Dim recvBuffer As Byte() = New Byte(0) {}
        Dim recvBufferLen As Integer

        sendBuffer(0) = &HFF
        sendBuffer(1) = &HD6
        sendBuffer(3) = adr And &HFF
        sendBuffer(2) = (adr >> 8) And &HFF
        sendBuffer(4) = 16
        For i As Int16 = 0 To 16
            sendBuffer(i + 5) = &H0
        Next
        If chrlen > 0 Then
            For i As Int16 = 0 To chrlen - 1
                sendBuffer(i + 5) = chr(i)
            Next
        End If

        If Not Me.Transmit(sendBuffer, recvBuffer, recvBufferLen) Then
            Me.Disconnect()
            Me.ReleaseContext()
            Return False
        End If
        'chr = recvBuffer
        'chrlen = recvBufferLen
        Return True
    End Function

    Private Function FelicaUpdateBinary(ByVal adr As UInt16, ByRef chr As Byte(), ByRef chrlen As Integer) As Boolean

        Dim commandFeliCaWrite As Byte() = New Byte() {&HFF, &HC2, &H0, &H1, _
                                              &H22, _
                                              &H95, _
                                              &H20, &H20, _
                                              &H8, _
                                              &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, _
                                              &H1, _
                                              &H9, &H0, _
                                              &H1, _
                                              &H80, &H0, _
                                              &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, _
                                              &H0}
        ''               &HFF, &HC2, &H0, &H1   ：直接Ｒ／Ｗにカード固有のコマンドを送るモード
        ''               &H12                   ：以降のコマンドの長さ（バイト数、最後の&H0は無視）この場合は18個
        ''               &H95                   ：送信及び受信コマンド
        ''               &H10, &H10             ：Ｒ／Ｗに送信するコマンドの長さ（２回繰り返す、最後の&H0は無視）
        ''               &H8                    ：FelicaのWrite Without Encryptionコマンド
        ''               &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0：カードのIDm（Pollingコマンドで取得したものをここに上書きする)
        ''               &H1                    ：サービス数（通常は1）
        ''               &HB, &H0               ：サービスコード（読込のみ:&HB, &H0、読み書き:&H9, &H0）
        ''               &H1                    ：書き込むブロック数（この場合は１）
        ''               &H80, &H0              ：書き込むブロック番号リスト（この場合は０番のブロック）複数の場合はここに追加する。
        '                &H0....&H0(16Byte)     ：書き込むデータ（16バイト）             
        ''               &H0                    ：最後は必ず&H0を入れる


        Dim recvBuffer As Byte() = New Byte(0) {}
        Dim recvBufferLen As Integer
        Array.Copy(Me.bIDm, 0, commandFeliCaWrite, 9, 8)
        commandFeliCaWrite(22) = adr And &HFF

        For i As Int16 = 0 To 15
            commandFeliCaWrite(i + 23) = &H0
        Next

        If chrlen > 0 Then
            For i As Int16 = 0 To chrlen - 1
                commandFeliCaWrite(i + 23) = chr(i)
            Next
        End If

        If Not Me.Control(commandFeliCaWrite, recvBuffer, recvBufferLen) Then
            Me.Disconnect()
            Me.ReleaseContext()
            Return False
        End If
        'ReDim chr(15)
        'chrlen = 16
        'Array.Copy(recvBuffer, 27, chr, 0, 16)

        Return True
    End Function


    Public Function dumpBinary(ByRef dumpText As String) As Boolean
        Dim chr1 As Byte() = New Byte(0) {}
        Dim chrlen As Integer
        Dim i As UInt16
        dumpText = ""
        dumpText += New String("-"c, 80) + vbNewLine
        dumpText += "|adr|BlockName|                     DATA                      |      ASCII     |" + vbNewLine

        dumpText += New String("="c, 80) + vbNewLine
        For i = 0 To 13
            dumpText += "|" + i.ToString("X2") + "h|"
            dumpText += " S_PAD" + i.ToString("D2") + " "
            If Not Me.readBinary(i, chr1, chrlen) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If
            dumpText += "|" + BitConverter.ToString(chr1, 0, chrlen - 2) + "|"

            For j As Integer = 0 To 15
                If chr1(j) = &H0 Then
                    dumpText += " "
                Else
                    If chr1(j) >= 20 And chr1(j) <= 126 Then
                        dumpText += Chr(chr1(j))
                    Else
                        dumpText += "･"
                    End If
                End If
            Next
            dumpText += "|" + vbNewLine
        Next

        For i = 14 To 14
            dumpText += "|" + i.ToString("X2") + "h|"
            dumpText += Left(" REG" + "       ", 9)
            If Not Me.readBinary(i, chr1, chrlen) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If
            dumpText += "|" + BitConverter.ToString(chr1, 0, chrlen - 2) + "|"

            For j As Integer = 0 To 15
                If chr1(j) = &H0 Then
                    dumpText += " "
                Else
                    If chr1(j) >= 20 And chr1(j) <= 126 Then
                        dumpText += Chr(chr1(j))
                    Else
                        dumpText += "･"
                    End If
                End If
            Next
            dumpText += "|" + vbNewLine
        Next

        Dim st As Int16 = &H80
        For i = st To st + 8
            dumpText += "|" + i.ToString("X2") + "h|"
            Select Case i
                Case st
                    dumpText += Left(" RC" + "       ", 9)
                Case st + 1
                    dumpText += Left(" MAC" + "       ", 9)
                Case st + 2
                    dumpText += Left(" ID" + "       ", 9)
                Case st + 3
                    dumpText += Left(" D_ID" + "       ", 9)
                Case st + 4
                    dumpText += Left(" SER_C" + "       ", 9)
                Case st + 5
                    dumpText += Left(" SYS_C" + "       ", 9)
                Case st + 6
                    dumpText += Left(" CKY" + "       ", 9)
                Case st + 7
                    dumpText += Left(" CK" + "       ", 9)
                Case st + 8
                    dumpText += Left(" MC" + "       ", 9)
            End Select

            If Not Me.readBinary(i, chr1, chrlen) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If
            dumpText += "|" + BitConverter.ToString(chr1, 0, chrlen - 2) + "|"

            For j As Integer = 0 To 15
                If chr1(j) = &H0 Then
                    dumpText += " "
                Else
                    If chr1(j) >= 20 And chr1(j) <= 126 Then
                        dumpText += Chr(chr1(j))
                    Else
                        dumpText += "･"
                    End If
                End If
            Next
            dumpText += "|" + vbNewLine
        Next

        st = &H90
        For i = st To st + 2
            dumpText += "|" + i.ToString("X2") + "h|"
            Select Case i
                Case st
                    dumpText += Left(" WCNT" + "       ", 9)
                Case st + 1
                    dumpText += Left(" MAC_A" + "       ", 9)
                Case st + 2
                    dumpText += Left(" STATE" + "       ", 9)
            End Select

            If Not Me.readBinary(i, chr1, chrlen) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If
            dumpText += "|" + BitConverter.ToString(chr1, 0, chrlen - 2) + "|"

            For j As Integer = 0 To 15
                If chr1(j) = &H0 Then
                    dumpText += " "
                Else
                    If chr1(j) >= 20 And chr1(j) <= 126 Then
                        dumpText += Chr(chr1(j))
                    Else
                        dumpText += "･"
                    End If
                End If
            Next
            dumpText += "|" + vbNewLine
        Next

        st = &HA0
        For i = st To st
            dumpText += "|" + i.ToString("X2") + "h|"
            Select Case i
                Case st
                    dumpText += Left("CRC_CHECK" + "       ", 9)
            End Select

            If Not Me.readBinary(i, chr1, chrlen) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If
            dumpText += "|" + BitConverter.ToString(chr1, 0, chrlen - 2) + "|"

            For j As Integer = 0 To 15
                If chr1(j) = &H0 Then
                    dumpText += " "
                Else
                    If chr1(j) >= 20 And chr1(j) <= 126 Then
                        dumpText += Chr(chr1(j))
                    Else
                        dumpText += "･"
                    End If
                End If
            Next
            dumpText += "|" + vbNewLine
        Next
        dumpText += New String("-"c, 80) + vbNewLine
        Return True
    End Function

    Public Function FelicaDumpBinary(ByRef dumpText As String) As Boolean
        Dim chr1 As Byte() = New Byte(0) {}
        Dim chrlen As Integer
        Dim i As UInt16
        dumpText = ""
        'dumpText += "|" + New String("-"c, 78) + "|" + vbNewLine
        dumpText += "[DumpData]" + vbNewLine
        dumpText += "|------------------------------------------------------------------------------|" + vbNewLine
        dumpText += "|adr|BlockName|                     DATA                      |      ASCII     |" + vbNewLine
        dumpText += "|===|=========|===============================================|================|" + vbNewLine
        'dumpText += New String("="c, 80) + vbNewLine
        For i = 0 To 13
            dumpText += "|" + i.ToString("X2") + "h|"
            dumpText += " S_PAD" + i.ToString("D2") + " "
            If Not Me.FelicaReadBinary(i, chr1, chrlen) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If
            dumpText += "|" + BitConverter.ToString(chr1, 0, chrlen) + "|"

            For j As Integer = 0 To 15
                If chr1(j) = &H0 Then
                    dumpText += " "
                Else
                    If chr1(j) >= 20 And chr1(j) <= 126 Then
                        dumpText += Chr(chr1(j))
                    Else
                        dumpText += "･"
                    End If
                End If
            Next
            dumpText += "|" + vbNewLine
        Next

        For i = 14 To 14
            dumpText += "|" + i.ToString("X2") + "h|"
            dumpText += Left(" REG" + "       ", 9)
            If Not Me.FelicaReadBinary(i, chr1, chrlen) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If
            dumpText += "|" + BitConverter.ToString(chr1, 0, chrlen) + "|"

            For j As Integer = 0 To 15
                If chr1(j) = &H0 Then
                    dumpText += " "
                Else
                    If chr1(j) >= 20 And chr1(j) <= 126 Then
                        dumpText += Chr(chr1(j))
                    Else
                        dumpText += "･"
                    End If
                End If
            Next
            dumpText += "|" + vbNewLine
        Next

        Dim st As Int16 = &H80
        For i = st To st + 8
            dumpText += "|" + i.ToString("X2") + "h|"
            Select Case i
                Case st
                    dumpText += Left(" RC" + "       ", 9)
                Case st + 1
                    dumpText += Left(" MAC" + "       ", 9)
                Case st + 2
                    dumpText += Left(" ID" + "       ", 9)
                Case st + 3
                    dumpText += Left(" D_ID" + "       ", 9)
                Case st + 4
                    dumpText += Left(" SER_C" + "       ", 9)
                Case st + 5
                    dumpText += Left(" SYS_C" + "       ", 9)
                Case st + 6
                    dumpText += Left(" CKY" + "       ", 9)
                Case st + 7
                    dumpText += Left(" CK" + "       ", 9)
                Case st + 8
                    dumpText += Left(" MC" + "       ", 9)
            End Select

            If Not Me.FelicaReadBinary(i, chr1, chrlen) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If
            dumpText += "|" + BitConverter.ToString(chr1, 0, chrlen) + "|"

            For j As Integer = 0 To 15
                If chr1(j) = &H0 Then
                    dumpText += " "
                Else
                    If chr1(j) >= 20 And chr1(j) <= 126 Then
                        dumpText += Chr(chr1(j))
                    Else
                        dumpText += "･"
                    End If
                End If
            Next
            dumpText += "|" + vbNewLine
        Next

        st = &H90
        For i = st To st + 2
            dumpText += "|" + i.ToString("X2") + "h|"
            Select Case i
                Case st
                    dumpText += Left(" WCNT" + "       ", 9)
                Case st + 1
                    dumpText += Left(" MAC_A" + "       ", 9)
                Case st + 2
                    dumpText += Left(" STATE" + "       ", 9)
            End Select

            If Not Me.FelicaReadBinary(i, chr1, chrlen) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If
            dumpText += "|" + BitConverter.ToString(chr1, 0, chrlen) + "|"

            For j As Integer = 0 To 15
                If chr1(j) = &H0 Then
                    dumpText += " "
                Else
                    If chr1(j) >= 20 And chr1(j) <= 126 Then
                        dumpText += Chr(chr1(j))
                    Else
                        dumpText += "･"
                    End If
                End If
            Next
            dumpText += "|" + vbNewLine
        Next

        st = &HA0
        For i = st To st
            dumpText += "|" + i.ToString("X2") + "h|"
            Select Case i
                Case st
                    dumpText += Left("CRC_CHECK" + "       ", 9)
            End Select

            If Not Me.FelicaReadBinary(i, chr1, chrlen) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If
            dumpText += "|" + BitConverter.ToString(chr1, 0, chrlen) + "|"

            For j As Integer = 0 To 15
                If chr1(j) = &H0 Then
                    dumpText += " "
                Else
                    If chr1(j) >= 20 And chr1(j) <= 126 Then
                        dumpText += Chr(chr1(j))
                    Else
                        dumpText += "･"
                    End If
                End If
            Next
            dumpText += "|" + vbNewLine
        Next
        dumpText += "|------------------------------------------------------------------------------|"
        'dumpText += New String("-"c, 80) + vbNewLine
        Return True
    End Function


    '======================
    ' カードにデータを書き込む
    '======================
    Public Function setCardID(ByVal adr As UInt16, ByRef chr As Byte(), ByRef chrlen As Integer) As Boolean
        '------------------
        ' プロパティ初期化
        '------------------
        Me._CardType = String.Empty
        Me._IsFelica = False
        Me._IsMifare = False
        Me._IDm = String.Empty
        Me._PMm = String.Empty
        Me._UID = String.Empty
        Me._S_PAD0 = String.Empty
        Me._ErrorMsg = String.Empty

        '------------------------------------
        ' FelicaのIDm,PMm・MifareのUIDを取得
        '------------------------------------
        Try
            ' (1)SCardEstablishContext
            ' ⇒リソースマネージャのハンドルを取得
            If Not Me.EstablishContext() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (2)SCardListReaders
            ' ⇒リーダー／ライターの名称を取得
            If Not Me.ListReaders() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (3)SCardGetStatusChange
            ' ⇒リーダー／ライターの状態を取得
            If Not Me.GetStatusChange() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (4)SCardConnect
            ' ⇒カードに接続
            If Not Me.Connect() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (5)SCardStatus
            ' ⇒ATR(Answer To Reset)を取得しカード種別を判定
            If Not Me.Status() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (6)SCardTransmit
            ' ⇒カードにコマンド送信・データを受信
            Dim recvBuffer As Byte() = New Byte(0) {}
            Dim recvBufferLen As Integer
            If Me._IsFelica Then
                'FelicaのIDmを取得
                If Not Me.Transmit(Me.sendBuffer_IDm, recvBuffer, recvBufferLen) Then
                    Me.Disconnect()
                    Me.ReleaseContext()
                    Return False
                End If
                Me._IDm = BitConverter.ToString(recvBuffer, 0, recvBufferLen - 2).Replace("-", String.Empty)

                'FelicaのPMmを取得
                If Not Me.Transmit(Me.sendBuffer_PMm, recvBuffer, recvBufferLen) Then
                    Me.Disconnect()
                    Me.ReleaseContext()
                    Return False
                End If
                Me._PMm = BitConverter.ToString(recvBuffer, 0, recvBufferLen - 2).Replace("-", String.Empty)

                'FelicaのS_PAD0にデータを書き込む
                If Not Me.updateBinary(adr, chr, chrlen) Then
                    Me.Disconnect()
                    Me.ReleaseContext()
                    Return False
                End If
                'Me._S_PAD0 = BitConverter.ToString(recvBuffer, 0, recvBufferLen - 2).Replace("-", String.Empty)
            End If

            ' (7)SCardDisconnect
            ' ⇒カードとの通信を切断
            Me.Disconnect()

            ' (8)SCardReleaseContext
            ' ⇒リソースマネージャのハンドルを解放
            Me.ReleaseContext()

            Return True
        Catch ex As Exception
            Me._ErrorMsg = "[Error] " + ex.Message
            Return False
        End Try
    End Function


    '======================
    ' カードにデータを書き込む
    '======================
    Public Function setFelicaCard(ByVal adr As UInt16, ByRef chr As Byte(), ByRef chrlen As Integer) As Boolean
        '------------------
        ' プロパティ初期化
        '------------------
        Me._CardType = String.Empty
        Me._IsFelica = False
        Me._IsMifare = False
        Me._IDm = String.Empty
        Me.bIDm = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Me._PMm = String.Empty
        Me.bPMm = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Me._UID = String.Empty
        Me._S_PAD0 = String.Empty
        Me._ErrorMsg = String.Empty

        '------------------------------------
        ' FelicaのIDm,PMm・MifareのUIDを取得
        '------------------------------------
        Try
            ' (1)SCardEstablishContext
            ' ⇒リソースマネージャのハンドルを取得
            If Not Me.EstablishContext() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (2)SCardListReaders
            ' ⇒リーダー／ライターの名称を取得
            If Not Me.ListReaders() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If


            ' (3)SCardConnect
            ' ⇒カードなしでＲ／Ｗに接続
            If Not Me.ConnectDirect() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (4)SessionStart
            ' ⇒コマンドモードのセッション開始
            If Not Me.SessionStart() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (5)ChangeFelicaProtocol
            ' ⇒Felicaのプロトコルに変更
            If Not Me.ChangeFelicaProtocol() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (6)FelicaPolling
            ' ⇒Felicaカードのポーリングの開始（IDmとPMmの読み取り）
            Dim SysCode As Byte() = New Byte() {&H88, &HB4}
            If Not Me.FelicaPolling(SysCode) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (7)FelicaDumpBinary
            ' ⇒Felicaカードの内容をダンプ
            If Not Me.FelicaUpdateBinary(adr, chr, chrlen) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If
            
            ' (8)SessionEnd
            ' ⇒コマンドモードのセッション終了
            If Not Me.SessionEnd() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (9)SCardDisconnect
            ' ⇒カードとの通信を切断
            Me.Disconnect()

            ' (10)SCardReleaseContext
            ' ⇒リソースマネージャのハンドルを解放
            Me.ReleaseContext()

            Return True
        Catch ex As Exception
            Me._ErrorMsg = "[Error] " + ex.Message
            Return False
        End Try

    End Function


    Private Function ByteXor(ByRef b1 As Byte(), b2 As Byte()) As Byte()
        ' バイト配列同士の排他的論理和(Xor) を行う。

        Dim b1len As Integer, b2len As Integer, b3() As Byte
        b1len = b1.Length
        b2len = b2.Length

        If b1len > 0 And b1len = b2len Then ' 2つのバイト配列が同じ長さで　かつ　　長さが0でない場合に計算を行う。
            ReDim b3(b1len - 1)
            For i As Integer = 0 To b1len - 1
                b3(i) = b1(i) Xor b2(i)
            Next

        Else
            ReDim b3(0)
            b3(0) = 0
        End If
        Return b3

    End Function

    Private Function BitShift(ByRef b1 As Byte(), ByVal shift_n As Integer) As Byte()
        ' バイト配列のビットシフトを行う。

        Dim A As UInt64
        Dim B As Byte()
        A = BitConverter.ToUInt64(b1, 0)    ' バイト配列を64ビット整数に変換
        A = A << shift_n                    ' ビットをシフト
        B = BitConverter.GetBytes(A)        ' 64ビット整数をバイト配列に変換
        Return B

    End Function


    '======================
    ' カードのIDを取得する
    '======================
    Public Function getNoWithMac_A() As Boolean
        Dim chr1 As Byte() = New Byte(0) {}
        Dim chrlen As Integer
        '------------------
        ' プロパティ初期化
        '------------------
        Me._CardType = String.Empty
        Me._IsFelica = False
        Me._IsMifare = False
        Me._IDm = String.Empty
        Me.bIDm = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Me._PMm = String.Empty
        Me.bPMm = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Me._UID = String.Empty
        Me._S_PAD0 = String.Empty
        Me._ErrorMsg = String.Empty
        Me.Mac = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        Me.Mac_A = New Byte() {&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0}
        'Me.Ck1 = New Byte() {&H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7}
        'Me.Ck2 = New Byte() {&H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF}
        'Me.Rc1 = New Byte() {&H0, &H1, &H2, &H3, &H4, &H5, &H6, &H7}
        'Me.Rc2 = New Byte() {&H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF}

        '------------------------------------
        ' FelicaのIDm,PMm・MifareのUIDを取得
        '------------------------------------
        Try
            ' (1)SCardEstablishContext
            ' ⇒リソースマネージャのハンドルを取得
            If Not Me.EstablishContext() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (2)SCardListReaders
            ' ⇒リーダー／ライターの名称を取得
            If Not Me.ListReaders() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If


            ' (3)SCardConnect
            ' ⇒カードなしでＲ／Ｗに接続
            If Not Me.ConnectDirect() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (4)SessionStart
            ' ⇒コマンドモードのセッション開始
            If Not Me.SessionStart() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (5)ChangeFelicaProtocol
            ' ⇒Felicaのプロトコルに変更
            If Not Me.ChangeFelicaProtocol() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (6)FelicaPolling
            ' ⇒Felicaカードのポーリングの開始（IDmとPMmの読み取り）
            Dim SysCode As Byte() = New Byte() {&H88, &HB4}
            If Not Me.FelicaPolling(SysCode) Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' カードに書かれたIDと暗号鍵から個別暗号鍵を作成し、Me.Ck()変数にコピーする。
            ' この個別暗号鍵と同じものが事前にカードに書き込まれている。
            If Not Me.CalcCKBlock() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' 乱数Me.Rc1とMe.Rc2を作成し、カードのRcブロックに書き込む。
            ' 上記のMe.Ck()から新しい暗号鍵SessionKeyを作成する。
            ' カード内部でも同じSessionKeyが計算されてMac_Aの計算に使用される。
            If Not Me.MakeSessionKey() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            Dim A As String = ""

            For i As Int16 = 0 To 0
                ' Rc1を暗号化の初期値に設定する。
                Me.Iv = Me.Rc1.Reverse.ToArray

                ' Mac_A付き読込を行う（暗号化が一致しない場合はエラーとなる）。
                If Not Me.FelicaReadBinaryWithMac_A(i, chr1, chrlen) Then
                    Me.Disconnect()
                    Me.ReleaseContext()
                    A += "Mac Error"
                    Return False
                End If


                For j As Integer = 0 To chrlen - 1
                    If chr1(j) = &H0 Then
                        A += " "
                    Else
                        If chr1(j) >= 20 And chr1(j) <= 126 Then
                            A += Chr(chr1(j))
                        End If
                    End If
                Next
            Next


            Me._S_PAD0 = A

            ' (8)SessionEnd
            ' ⇒コマンドモードのセッション終了
            If Not Me.SessionEnd() Then
                Me.Disconnect()
                Me.ReleaseContext()
                Return False
            End If

            ' (9)SCardDisconnect
            ' ⇒カードとの通信を切断
            Me.Disconnect()

            ' (10)SCardReleaseContext
            ' ⇒リソースマネージャのハンドルを解放
            Me.ReleaseContext()

            Return True
        Catch ex As Exception
            Me._ErrorMsg = "[Error] " + ex.Message
            Return False
        End Try

    End Function


End Class
