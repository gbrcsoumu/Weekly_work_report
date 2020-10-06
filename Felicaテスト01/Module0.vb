Module Module0
    Public LoginID As String                                ' FileMakerのログインID
    Public LoginPassWord As String                          ' FileMakerのログインパスワード
    Public Const DataBaseName As String = "退勤管理test01"  ' FileMakerのデータベース名
    Public Const MemberNameTable As String = "職員一覧"     ' 職員名簿のテーブル名
    Public Const MemberNameTable2 As String = "職員所属部署"     ' 職員名簿のテーブル名
    Public Const MemberLogTable As String = "出退勤一覧"      ' 退勤管理のテーブル名
    Public Const CardMasterKeyString = "GBRC 2020"          ' Felicaカードの暗号化のキー
    Public FelicaRW As Boolean
End Module
