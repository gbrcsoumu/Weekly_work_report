'===================================================================================================
'
'   Felica カードへの書込フォーム
'
'   2019/12 CODED By kanyama
'
'===================================================================================================

Imports System.Threading
Imports Microsoft.Office.Interop

Public Class FelicaRWForm

    'Private CardMasterKeyString As String
    Private kind1() As String, kind2() As String, kind3() As String
    Private No() As String, Name1() As String, Affiliation1() As String, Affiliation2() As String, Affiliation3() As String, Post() As String, IDm() As String
    'Private Kind4() As String, Date1() As String, Time1() As String
    Private Kubun2() As String, Kubun As String
    Private Date1() As Date, Calender1() As String, Huzai() As String, Jiyuu() As String, Kubun1() As String, Time1() As DateTime, MC1() As String, Time2() As DateTime, MC2() As String
    Private GaishutuCount1() As String, GaishutuCount2() As String, ShoteiTime() As TimeSpan, Hounai() As TimeSpan, Hougai() As TimeSpan
    Private HouteiH() As TimeSpan, HougaiH() As TimeSpan, shinya() As TimeSpan, Hougai40() As TimeSpan
    Private Chikoku() As TimeSpan, Soutai() As TimeSpan, Gaishutu() As TimeSpan, Youbi() As String, Name2() As String, YukyuJikan() As TimeSpan
    Private ShoteiTime_sum As TimeSpan, Hounai_sum As TimeSpan, Hougai_sum As TimeSpan, HouteiH_sum As TimeSpan, HougaiH_sum As TimeSpan, shinya_sum As TimeSpan, Hougai40_sum As TimeSpan
    Private Chikoku_sum As TimeSpan, Soutai_sum As TimeSpan, Gaishutu_sum As TimeSpan
    Private 所定日数 As Double, 有給日数 As Double


    '====================================
    ' Invokeメソッドで使用するデリゲート
    '====================================
    Delegate Sub txtMessage_Text_Delegate(ByVal value As String)
    Delegate Sub txtMessage_Scroll_Delegate()
    Delegate Sub btnCardID_Enable_Delegate(ByVal value As Boolean)

    Private dic As Dictionary(Of String, Int16)
    ' 
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim db As New OdbcDbIf
        Dim tb As DataTable
        Dim Sql_Command As String
        Dim i As Integer
        Dim Path1 As String
        Path1 = Application.StartupPath

        Dim n As Integer

        LoginID = ""
        LoginPassWord = ""

        ' ログインダイアログの表示

        Dim res As DialogResult
        Dim f1 As New LoginForm1
        res = f1.ShowDialog()
        f1.Dispose()
        If res = System.Windows.Forms.DialogResult.Cancel Then
            ' CANCELされた場合は閉じる。
            Me.Close()
            Exit Sub
        End If


        Me.Width = 1000     ' フォームんの幅を設定
        Me.Height = 600     ' フォームんの高さを設定

        Try
            db.Connect(DataBaseName, LoginID, LoginPassWord, -1)

        Catch ex As Exception
            MsgBox("ID又はパスワードが違います。")
            Me.Close()
            db.Disconnect()
            Exit Sub
        End Try


        'Sql_Command = "SELECT ID,""試験名称"",""依頼者名"" FROM ""見積書"" WHERE (ID LIKE 'H%%%%%%%%%%') AND (""見積作成日"">= DATE" & day3 & ") ORDER BY ""見積作成日"" DESC"
        'Sql_Command = "SELECT ID,""試験名称"",""依頼者名"" FROM ""見積書"" WHERE (ID LIKE 'H29:001-001') AND (""見積作成日"">= DATE" & day3 & ") ORDER BY ""見積作成日"" DESC"

        Sql_Command = "SELECT ""所属センター"" FROM """ + MemberNameTable2 + """ ORDER BY ""所属センター"""
        tb = db.ExecuteSql(Sql_Command)
        n = tb.Rows.Count
        If n > 0 Then
            ReDim Me.Affiliation1(n - 1)
            For i = 0 To n - 1
                Me.Affiliation1(i) = tb.Rows(i).Item("所属センター").ToString()
            Next

        End If
        Me.kind1 = Nkind(Me.Affiliation1)
        Me.Label_Center.Text = "所属センター(" + Me.kind1.Length.ToString("0") + ")"
        For Each s In Me.kind1
            Me.ComboBoxCenter.Items.Add(s)
        Next

        Dim this_Year As String = DateTime.Today.Year
        Dim year_n As Integer = 5
        For i = 1 To year_n
            Me.ComboBoxYear.Items.Add((this_Year - i + 1).ToString)
        Next
        Me.ComboBoxYear.Text = this_Year.ToString()

        Dim this_Month As String = DateTime.Today.Month
        For i = 1 To 12
            Me.ComboBoxMonth.Items.Add(i.ToString)
        Next
        Me.ComboBoxMonth.Text = (this_Month - 1).ToString()






        'Sql_Command = "SELECT * FROM ""職員一覧"""
        'tb = db.ExecuteSql(Sql_Command)
        'Dim n As Integer = tb.Rows.Count
        'If n > 0 Then
        '    ReDim No(n - 1), Name(n - 1), Affiliation1(n - 1), Affiliation2(n - 1), Affiliation3(n - 1), Post(n - 1), IDm(n - 1)
        '    Me.ComboBox1.Items.Clear()
        '    For i = 0 To n - 1
        '        No(i) = tb.Rows(i).Item("職員番号").ToString()
        '        Name(i) = tb.Rows(i).Item("氏名").ToString()
        '        Affiliation1(i) = tb.Rows(i).Item("所属センター").ToString()
        '        Affiliation2(i) = tb.Rows(i).Item("所属部").ToString()
        '        Affiliation3(i) = tb.Rows(i).Item("所属室").ToString()
        '        Post(i) = tb.Rows(i).Item("役職").ToString()
        '        IDm(i) = tb.Rows(i).Item("IDm").ToString()
        '    Next
        'End If
        db.Disconnect()
        tb.Dispose()


        'dic = New Dictionary(Of String, Int16)
        'For a As Int16 = 0 To 13
        '    'Me.ComboBox1.Items.Add("S_PAD" + a.ToString("D2"))
        '    dic.Add("S_PAD" + a.ToString("D2"), a)
        'Next
        'For Each s In dic
        '    Me.ComboBox1.Items.Add(s.Key)
        'Next
        'Me.ComboBox1.Text = ComboBox1.GetItemText(ComboBox1.Items(0))

        Const C_width As Integer = 100
        With Me.DataGridView1
            .Width = 100 + C_width * 7.5 + 60
            .Height = 300
            .ColumnCount = 7
            .ColumnHeadersVisible = True
            .ColumnHeadersHeight = 14
            .ScrollBars = ScrollBars.Both

            Dim columnHeaderStyle As New DataGridViewCellStyle()
            columnHeaderStyle.BackColor = Color.White
            columnHeaderStyle.Font = New Font("MSゴシック", 10, FontStyle.Bold)
            columnHeaderStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ColumnHeadersDefaultCellStyle = columnHeaderStyle
            .Columns(0).Name = "職員番号"
            .Columns(1).Name = "氏名"
            .Columns(2).Name = "センター"
            .Columns(3).Name = "所属部"
            .Columns(4).Name = "所属室"
            .Columns(5).Name = "役職"
            .Columns(6).Name = "IDm"
            '                  .Columns(5).Name = "On/Off"
            .RowHeadersVisible = True
            .Columns(0).Width = 80
            .Columns(1).Width = C_width
            .Columns(2).Width = C_width
            .Columns(3).Width = C_width
            .Columns(4).Width = C_width
            .Columns(5).Width = C_width
            .Columns(6).Width = C_width * 1.5
            '                    .Columns(5).Width = C_width

            'For i = 0 To n - 1
            '    Row = {No(i), Name(i), Affiliation1(i), Affiliation2(i), Affiliation3(i), Post(i), IDm(i)}
            '    .Rows.Add(Row)
            'Next

            'DataGridViewButtonColumnの作成
            Dim column As New DataGridViewButtonColumn()
            '列の名前を設定
            column.Name = "就業週報"
            '全てのボタンに"詳細閲覧"と表示する
            column.UseColumnTextForButtonValue = True
            column.Text = "作成"
            'DataGridViewに追加する
            .Columns.Add(column)
        End With
        'Dim column1 As New DataGridViewCheckBoxColumn
        'DataGridView1.Columns.Add(column1)
        'DataGridView1.Columns(5).Name = "風向"
        'DataGridView1.Columns(5).Width = C_width / 2

        'Dim column2 As New DataGridViewCheckBoxColumn
        'DataGridView1.Columns.Add(column2)
        'DataGridView1.Columns(6).Name = "風速"
        'DataGridView1.Columns(6).Width = C_width / 2


        'day1 = Now
        'day2 = day1.AddDays(-60)
        'day3 = "'" & Format(day2.Date, "yyyy-MM-dd") & "'"

        'db.Connect()
        ''Sql_Command = "SELECT ID,""試験名称"",""依頼者名"" FROM ""見積書"" WHERE (ID LIKE 'H%%%%%%%%%%') AND (""見積作成日"">= DATE" & day3 & ") ORDER BY ""見積作成日"" DESC"
        ''Sql_Command = "SELECT ID,""試験名称"",""依頼者名"" FROM ""見積書"" WHERE (ID LIKE 'H29:001-001') AND (""見積作成日"">= DATE" & day3 & ") ORDER BY ""見積作成日"" DESC"
        'Sql_Command = "SELECT ID,""試験名称"",""依頼者名"" FROM ""見積書"" WHERE (""見積作成日"">= DATE" & day3 & ") ORDER BY ""見積作成日"" DESC"
        'tb = db.ExecuteSql(Sql_Command)

        'Me.ComboBox1.Items.Clear()
        'For i = 0 To tb.Rows.Count - 1
        '    Me.ComboBox1.Items.Add(tb.Rows(i).Item("ID").ToString() & ":" & _
        '                           tb.Rows(i).Item("依頼者名").ToString())
        'Next
        'db.Disconnect()
        'Me.ComboBox1.SelectedIndex = 0
        'Me.ComboBox1.Refresh()
    End Sub



    Public Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        'Me.CardMasterKeyString = "GBRC 2020"

    End Sub

    Private Function Nkind(ByRef x() As String) As String()
        Dim xn As Integer, yn() As String
        Dim i As Integer, kn As Integer
        xn = x.Length
        If xn > 0 Then
            ReDim yn(xn - 1)
            kn = 1
            yn(kn - 1) = x(0)
            For i = 1 To xn - 1
                If x(i) <> yn(kn - 1) Then
                    kn += 1
                    yn(kn - 1) = x(i)
                End If
            Next
            ReDim Preserve yn(kn - 1)

        Else
            ReDim yn(0)
            yn(0) = "NO DATA"

        End If

        Return yn
    End Function


    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxCenter.SelectedIndexChanged
        Dim db As New OdbcDbIf
        Dim tb As DataTable
        Dim Sql_Command As String
        Dim i As Integer
        Dim n As Integer
        Dim A As String

        A = ComboBoxCenter.Text
        If A <> "" Then
            db.Connect()

            'Sql_Command = "SELECT ID,""試験名称"",""依頼者名"" FROM ""見積書"" WHERE (ID LIKE 'H%%%%%%%%%%') AND (""見積作成日"">= DATE" & day3 & ") ORDER BY ""見積作成日"" DESC"
            'Sql_Command = "SELECT ID,""試験名称"",""依頼者名"" FROM ""見積書"" WHERE (ID LIKE 'H29:001-001') AND (""見積作成日"">= DATE" & day3 & ") ORDER BY ""見積作成日"" DESC"

            Sql_Command = "SELECT ""所属部"" FROM """ + MemberNameTable2 + """ WHERE ""所属センター"" = '" & A & "' ORDER BY ""所属部"""
            tb = db.ExecuteSql(Sql_Command)
            n = tb.Rows.Count
            If n > 0 Then
                ReDim Me.Affiliation2(n - 1)
                For i = 0 To n - 1
                    Me.Affiliation2(i) = tb.Rows(i).Item("所属部").ToString()
                Next

            End If
            Me.kind2 = Nkind(Me.Affiliation2)
            Me.ComboBoxDepartment.Items.Clear()
            For Each s In Me.kind2
                Me.ComboBoxDepartment.Items.Add(s)
            Next
            If Me.kind2.Length = 1 Then
                Me.ComboBoxDepartment.Text = Me.kind2(0)
            Else
                Me.ComboBoxDepartment.Text = ""
            End If
            Me.Label_Department.Text = "所属部(" + Me.kind2.Length.ToString("0") + ")"

            db.Disconnect()
            tb.Dispose()
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxDepartment.SelectedIndexChanged
        Dim db As New OdbcDbIf
        Dim tb As DataTable
        Dim Sql_Command As String
        Dim i As Integer
        Dim n As Integer
        Dim A As String, B As String

        A = ComboBoxCenter.Text
        B = ComboBoxDepartment.Text

        If A <> "" And B <> "" Then
            db.Connect()

            'Sql_Command = "SELECT ID,""試験名称"",""依頼者名"" FROM ""見積書"" WHERE (ID LIKE 'H%%%%%%%%%%') AND (""見積作成日"">= DATE" & day3 & ") ORDER BY ""見積作成日"" DESC"
            'Sql_Command = "SELECT ID,""試験名称"",""依頼者名"" FROM ""見積書"" WHERE (ID LIKE 'H29:001-001') AND (""見積作成日"">= DATE" & day3 & ") ORDER BY ""見積作成日"" DESC"

            Sql_Command = "SELECT ""所属室"" FROM """ + MemberNameTable2 + """ WHERE ""所属センター"" = '" + A + "' AND ""所属部"" = '" + B + "' ORDER BY ""所属室"""
            tb = db.ExecuteSql(Sql_Command)
            n = tb.Rows.Count
            If n > 0 Then
                ReDim Me.Affiliation3(n - 1)
                For i = 0 To n - 1
                    Me.Affiliation3(i) = tb.Rows(i).Item("所属室").ToString()
                Next

            End If
            Me.kind3 = Nkind(Me.Affiliation3)

            Me.ComboBoxSection.Items.Clear()
            For Each s In Me.kind3
                Me.ComboBoxSection.Items.Add(s)
            Next
            If Me.kind3.Length = 1 Then
                Me.ComboBoxSection.Text = Me.kind3(0)
            Else
                Me.ComboBoxSection.Text = ""
            End If
            Me.Label_Section.Text = "所属室(" + Me.kind3.Length.ToString("0") + ")"

            db.Disconnect()
            tb.Dispose()
        End If

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxSection.SelectedIndexChanged, ComboBoxYear.SelectedIndexChanged, ComboBoxMonth.SelectedIndexChanged
        Dim db As New OdbcDbIf
        Dim tb As DataTable
        Dim Sql_Command As String
        Dim i As Integer
        'Dim n As Integer
        Dim A As String, B As String, C As String
        Dim Row() As String

        A = ComboBoxCenter.Text
        B = ComboBoxDepartment.Text
        C = ComboBoxSection.Text

        If A <> "" And B <> "" And C <> "" Then

            db.Connect()

            Dim Year1 As String = Me.ComboBoxYear.Text
            Dim Month1 As String = Me.ComboBoxMonth.Text
            Dim M2 As Integer = Integer.Parse(Me.ComboBoxMonth.Text) + 1
            If M2 > 12 Then M2 = 1
            Dim Month2 As String = M2.ToString
            Month1 = ("0" + Month1).Substring(Month1.Length - 1, 2)
            Month2 = ("0" + Month2).Substring(Month2.Length - 1, 2)

            '{06/05/2019}
            'Dim Day1 As String = "DATE'" + Year1 + "-" + Month1 + "-01'"
            Dim Day1 As String = "{" + Month1 + "/01/" + Year1 + "}"
            Dim period As String = " (""異動日前日"" > " + Day1 + "AND ""異動日"" <= " + Day1 + ")"
            'Dim period As String = " AND ""異動日"" <= " + Day1

            'Sql_Command = "SELECT * FROM """ + MemberNameTable + """ WHERE ""所属センター"" = '" + A + "' AND ""所属部"" = '" + B + "' AND ""所属室"" = '" + C + "' ORDER BY ""職員番号"""
            'Sql_Command = "SELECT * FROM """ + MemberNameTable2 + """ WHERE ""所属センター"" = '" + A + "' AND ""所属部"" = '" + B + "' AND ""所属室"" = '" + C + "'" + period + " ORDER BY ""職員番号"""
            Sql_Command = "SELECT * FROM """ + MemberNameTable2 + """ WHERE " + period + " AND ""所属センター"" = '" + A + "' AND ""所属部"" = '" + B + "' AND ""所属室"" = '" + C + "'" + " ORDER BY ""職員番号"""
            'Sql_Command = "SELECT * FROM """ + MemberNameTable + """ WHERE ""所属センター"" = '" + A + "'"


            tb = db.ExecuteSql(Sql_Command)
            Dim n As Integer = tb.Rows.Count
            If n > 0 Then
                ReDim No(n - 1), Name1(n - 1), Affiliation1(n - 1), Affiliation2(n - 1), Affiliation3(n - 1), Post(n - 1), IDm(n - 1)
                'Me.ComboBox1.Items.Clear()
                For i = 0 To n - 1
                    No(i) = tb.Rows(i).Item("職員番号").ToString()
                    Name1(i) = tb.Rows(i).Item("氏名").ToString()
                    Affiliation1(i) = tb.Rows(i).Item("所属センター").ToString()
                    Affiliation2(i) = tb.Rows(i).Item("所属部").ToString()
                    Affiliation3(i) = tb.Rows(i).Item("所属室").ToString()
                    Post(i) = tb.Rows(i).Item("役職").ToString()
                    IDm(i) = tb.Rows(i).Item("IDm").ToString()
                Next
            End If

            With Me.DataGridView1
                .Rows.Clear()
                For i = 0 To n - 1
                    Row = {No(i), Name1(i), Affiliation1(i), Affiliation2(i), Affiliation3(i), Post(i), IDm(i)}
                    .Rows.Add(Row)
                Next
            End With

            db.Disconnect()
            tb.Dispose()
        End If

    End Sub

    'CellContentClickイベントハンドラ
    Private Sub DataGridView1_CellContentClick(ByVal sender As Object,
                ByVal e As DataGridViewCellEventArgs) _
                Handles DataGridView1.CellContentClick

        System.Diagnostics.Debug.WriteLine("OK")
        Dim dgv As DataGridView = CType(sender, DataGridView)
        '"Button"列ならば、ボタンがクリックされた
        If dgv.Columns(e.ColumnIndex).Name = "就業週報" Then
            'MessageBox.Show((e.RowIndex.ToString() +
            '    "行のボタンがクリックされました。"))


            '        Dim pcsc As New clsWinSCard(CardMasterKeyString)
            Dim msg As String = ""
            'Dim adr As Int16, chrlen As Integer, data_n As Integer
            'Dim chr As Byte()
            '        adr = 0
            Dim data_i As Integer = e.RowIndex
            'chr = System.Text.Encoding.UTF8.GetBytes(Me.No(data_n))
            'chrlen = Len(Me.No(data_n))
            Dim Year1 As String = Me.ComboBoxYear.Text
            Dim Month1 As String = Me.ComboBoxMonth.Text
            Dim M2 As Integer = Integer.Parse(Me.ComboBoxMonth.Text) + 1
            If M2 > 12 Then M2 = 1
            Dim Month2 As String = M2.ToString
            Month1 = ("0" + Month1).Substring(Month1.Length - 1, 2)
            Month2 = ("0" + Month2).Substring(Month2.Length - 1, 2)

            Dim Half As Boolean = Me.CheckBox1.Checked

            Make_Table(data_i, Year1, Month1, Month2, Half)

        End If

    End Sub

    Private Sub Make_Table(ByVal data_i As Integer, ByVal Year1 As String, ByVal Month1 As String, ByVal Month2 As String, ByVal Half As Boolean)

        Dim db As New OdbcDbIf
        Dim tb As DataTable
        Dim Sql_Command As String
        Dim i As Integer


        'Dim n As Integer
        'Dim A As String, B As String, C As String
        'Dim Row() As String

        'Dim Year1 As String = Me.ComboBoxYear.Text
        'Dim Month1 As String = Me.ComboBoxMonth.Text
        'Dim M2 As Integer = Integer.Parse(Me.ComboBoxMonth.Text) + 1
        'If M2 > 12 Then M2 = 1
        'Dim Month2 As String = M2.ToString
        'Month1 = ("0" + Month1).Substring(Month1.Length - 1, 2)
        'Month2 = ("0" + Month2).Substring(Month2.Length - 1, 2)

        'Dim Half As Boolean = Me.CheckBox1.Checked

        Dim period As String

        If Half Then
            period = """日付"" >= DATE'" + Year1 + "-" + Month1 + "-01" + "' AND ""日付""< DATE'" + Year1 + "-" + Month1 + "-16'"
        Else
            period = """日付"" >= DATE'" + Year1 + "-" + Month1 + "-01" + "' AND ""日付""< DATE'" + Year1 + "-" + Month2 + "-01'"
        End If


        db.Connect()

        'Sql_Command = "SELECT * FROM ""職員一覧"" WHERE ""所属センター"" = '" + A + "' AND ""所属部"" = '" + B + "' AND ""所属室"" = '" + C + "' ORDER BY ""職員番号"""

        'Sql_Command = "UPDATE """ + MemberNameTable + """ SET IDm = '" + Me.IDm(data_n) + "' WHERE ""職員番号"" = '" + Me.No(data_n) + "'"

        Sql_Command = "SELECT * FROM  """ + MemberLogTable + """ WHERE ""職員番号"" = '" + No(data_i) + "'"

        Sql_Command += " AND " + period

        tb = db.ExecuteSql(Sql_Command)
        Dim day_n As Integer = tb.Rows.Count
        If day_n > 0 Then
            Kubun = tb.Rows(0).Item("勤務区分2").ToString()
            ReDim Date1(day_n - 1), Calender1(day_n - 1), Huzai(day_n - 1), Jiyuu(day_n - 1), Kubun1(day_n - 1), Time1(day_n - 1), MC1(day_n - 1), Time2(day_n - 1), MC2(day_n - 1)
            ReDim GaishutuCount1(day_n - 1), GaishutuCount2(day_n - 1), ShoteiTime(day_n - 1), Hounai(day_n - 1), Hougai(day_n - 1)
            ReDim HouteiH(day_n - 1), HougaiH(day_n - 1), shinya(day_n - 1), Hougai40(day_n - 1)
            ReDim Chikoku(day_n - 1), Soutai(day_n - 1), Gaishutu(day_n - 1), Youbi(day_n - 1), Name2(day_n - 1), YukyuJikan(day_n - 1)

            'ReDim No(n - 1), Name1(n - 1), Kind4(n - 1), Date1(n - 1), Time1(n - 1)
            'Me.ComboBox1.Items.Clear()
            Dim t1 As String, t2 As TimeSpan
            For i = 0 To day_n - 1

                Name2(i) = tb.Rows(i).Item("職員名").ToString()

                Date1(i) = Date.Parse(tb.Rows(i).Item("日付").ToString().Substring(0, 10))

                Youbi(i) = Date1(i).ToString("ddd")

                Calender1(i) = tb.Rows(i).Item("カレンダ").ToString()
                'If t1 <> "" Then
                '    Calender1(i) = tb.Rows(i).Item("カレンダ")
                'End If

                Huzai(i) = tb.Rows(i).Item("不在理由").ToString()

                Jiyuu(i) = tb.Rows(i).Item("事由").ToString()

                Kubun1(i) = tb.Rows(i).Item("勤務区分").ToString()

                t1 = tb.Rows(i).Item("出勤時刻").ToString
                If t1 <> "" Then
                    Time1(i) = DateTime.Parse(Date1(i).ToString("yyyy/MM/dd") + " " + t1)
                Else
                    Time1(i) = DateTime.Parse("1900/1/1 12:00:00")
                End If

                MC1(i) = tb.Rows(i).Item("出勤MC").ToString()

                t1 = tb.Rows(i).Item("退勤時刻").ToString
                If t1 <> "" Then
                    Time2(i) = DateTime.Parse(Date1(i).ToString("yyyy/MM/dd") + " " + t1)
                Else
                    Time2(i) = DateTime.Parse("1900/1/1 12:00:00")
                End If

                MC2(i) = tb.Rows(i).Item("退勤MC").ToString()

                GaishutuCount1(i) = tb.Rows(i).Item("公用外出の回数").ToString()

                GaishutuCount2(i) = tb.Rows(i).Item("私用外出の回数").ToString()

                ShoteiTime(i) = tb.Rows(i).Item("その日の所定時間")

                Hounai(i) = tb.Rows(i).Item("法内残業時間")

                Hougai(i) = tb.Rows(i).Item("法外残業時間")

                shinya(i) = tb.Rows(i).Item("深夜残業時間")

                HouteiH(i) = tb.Rows(i).Item("法定休H")

                HougaiH(i) = tb.Rows(i).Item("法定外H")

                Hougai40(i) = tb.Rows(i).Item("法外40残業時間")

                Chikoku(i) = tb.Rows(i).Item("遅刻時間")

                Soutai(i) = tb.Rows(i).Item("早退時間")

                t1 = tb.Rows(i).Item("私用外出時間").ToString()
                If t1 <> "" Then Gaishutu(i) = tb.Rows(i).Item("私用外出時間")

                YukyuJikan(i) = tb.Rows(i).Item("臨時職員の有給時間")

            Next
        End If

        Dim xlApp As Microsoft.Office.Interop.Excel.Application
        Dim xlBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet

        xlApp = New Excel.Application()

        Select Case Kubun
            Case "管理職"


                xlBook = xlApp.Application.Workbooks.Add()
                xlSheet = CType(xlApp.Worksheets(1), Excel.Worksheet)
                xlSheet.Name = Name2(0)

                Dim xlPageSetup As Excel.PageSetup = xlSheet.PageSetup

                With xlPageSetup
                    .PaperSize = Excel.XlPaperSize.xlPaperA4 '用紙サイズをＡ４
                    .Orientation = Excel.XlPageOrientation.xlLandscape
                    .LeftMargin = xlApp.CentimetersToPoints(1.2) '左余白
                    .RightMargin = xlApp.CentimetersToPoints(1.3) '右余白
                    .TopMargin = xlApp.CentimetersToPoints(0.9)
                    .BottomMargin = xlApp.CentimetersToPoints(1.5)
                    .FooterMargin = xlApp.CentimetersToPoints(1)
                    .HeaderMargin = xlApp.CentimetersToPoints(0.8)
                    '.PrintTitleRows = "$14:$16"
                    .CenterHorizontally = True
                    .CenterVertically = False
                End With

                Dim Row1 As Integer = 5

                With xlSheet
                    .Cells().Font.Name = "ＭＳ 明朝"
                    .Cells().Font.Size = 9

                    ''全ての列の幅を変更する
                    'xlSheet.Columns.ColumnWidth = 5

                    ''全ての行の高さを変更する
                    'xlSheet.Rows.RowHeight = 10

                    'C～Fの列の幅を変更する
                    .Range("A:A").ColumnWidth = 3.0    '  余白
                    .Range("B:B").ColumnWidth = 4.4    '  日付
                    .Range("C:C").ColumnWidth = 2.7    '  曜日
                    .Range("D:D").ColumnWidth = 4.0    '  カレンダー
                    .Range("E:E").ColumnWidth = 3.9    '  不在
                    .Range("F:F").ColumnWidth = 5.8    '  事由
                    .Range("G:G").ColumnWidth = 6.0    '  勤務区分
                    .Range("H:H").ColumnWidth = 5.8    '  出勤時刻
                    .Range("I:I").ColumnWidth = 5.5    '  出勤MC
                    .Range("J:J").ColumnWidth = 5.8    '  退勤時刻
                    .Range("K:K").ColumnWidth = 3.3    '  退勤MC
                    .Range("L:M").ColumnWidth = 5.8
                    .Range("N:T").ColumnWidth = 6.6


                    '2～5の行の高さを変更する
                    .Range("1:1").RowHeight = 12.6
                    .Range("2:2").RowHeight = 16.2
                    .Range("3:3").RowHeight = 18.0
                    .Range("4:4").RowHeight = 24.0

                    Dim range1 As String
                    range1 = Row1.ToString + ":" + (Row1 - 1 + day_n + 2).ToString
                    .Range(range1).RowHeight = 11.4
                    range1 = (Row1 + day_n + 3).ToString + ":" + (Row1 + day_n + 3).ToString
                    .Range(range1).RowHeight = 3.6
                    range1 = (Row1 + day_n + 4).ToString + ":" + (Row1 + day_n + 5).ToString
                    .Range(range1).RowHeight = 12.0


                    .Cells(Row1, 2) = "日付"
                    .Cells(Row1, 3) = "曜"
                    .Cells(Row1, 4) = "ｶﾚﾝﾀﾞ"
                    .Cells(Row1, 5) = "不在"
                    .Cells(Row1, 6) = "理由"
                    .Cells(Row1, 7) = "勤務区分"
                    .Cells(Row1, 8) = "出勤時刻"
                    .Cells(Row1, 9) = "ＭＣ"
                    .Cells(Row1, 10) = "退勤時刻"
                    .Cells(Row1, 11) = "ＭＣ"
                    .Cells(Row1, 12) = "公用外出"
                    .Cells(Row1, 13) = "私用外出"
                    .Cells(Row1, 14) = "所定時間"
                    .Cells(Row1, 15) = "法定休H"
                    .Cells(Row1, 16) = "法定外H"
                    .Cells(Row1, 17) = "深夜残業"
                    .Cells(Row1, 18) = "法定外40"
                    .Cells(Row1, 19) = "遅早時間"
                    .Cells(Row1, 20) = "私用外出"

                    ' セルの書式を中央揃え
                    Dim range2 As String = "A" + Row1.ToString + ":T" + Row1.ToString
                    .Range(range2).Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                    .Range(range2).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

                    ' 日付の書き込み
                    ''Dim theDay = New DateTime
                    ''Dim youbi As String
                    Dim t0 = DateTime.Parse("1900/1/1 12:00:00")
                    ShoteiTime_sum = TimeSpan.Parse("00:00:00")
                    Hounai_sum = TimeSpan.Parse("00:00:00")
                    Hougai_sum = TimeSpan.Parse("00:00:00")
                    HouteiH_sum = TimeSpan.Parse("00:00:00")
                    HougaiH_sum = TimeSpan.Parse("00:00:00")
                    shinya_sum = TimeSpan.Parse("00:00:00")
                    Hougai40_sum = TimeSpan.Parse("00:00:00")
                    Chikoku_sum = TimeSpan.Parse("00:00:00")
                    Soutai_sum = TimeSpan.Parse("00:00:00")
                    Gaishutu_sum = TimeSpan.Parse("00:00:00")
                    所定日数 = 0.0
                    有給日数 = 0.0

                    For i = 0 To day_n - 1
                        '.Cells(Row1 + i, 2) = "'" + Month1 + "/" + i.ToString("00")
                        'theDay = New DateTime(Integer.Parse(Year1), Integer.Parse(Month1), i)
                        'youbi = theDay.ToString("ddd")
                        '.Cells(Row1 + i, 3) = youbi
                        'If youbi = "土" Then
                        '    .Cells(Row1 + i, 4) = "法外"
                        'End If
                        'If youbi = "日" Then
                        '    .Cells(Row1 + i, 4) = "法定"
                        'End If

                        .Cells(Row1 + i + 1, 2) = Date1(i)
                        '.Range(.Cells(Row1 + i + 1, 2), .Cells(Row1 + i + 1, 2)).NumberFormat = "mm/dd"
                        .Cells(Row1 + i + 1, 3) = Youbi(i)
                        .Cells(Row1 + i + 1, 4) = Calender1(i)

                        If Huzai(i) <> "" And Huzai(i) <> "なし" Then
                            .Cells(Row1 + i + 1, 5) = Huzai(i)
                            Select Case Huzai(i)
                                Case "年休"
                                    有給日数 += 1.0
                                Case Else
                                    If Huzai(i).Contains("有") Then
                                        有給日数 += 0.5
                                    End If
                            End Select
                        Else
                            .Cells(Row1 + i + 1, 5) = ""
                        End If

                        If Jiyuu(i) <> "" And Jiyuu(i) <> "なし" Then
                            .Cells(Row1 + i + 1, 6) = Jiyuu(i)
                        Else
                            .Cells(Row1 + i + 1, 6) = ""
                        End If

                        Select Case Kubun1(i)
                            Case "8時管理職"
                                .Cells(Row1 + i + 1, 7) = "8時管"
                            Case "10時管理職"
                                .Cells(Row1 + i + 1, 7) = "10時管"
                            Case "8時一般職"
                                .Cells(Row1 + i + 1, 7) = "8時一般"
                            Case "10時管理職"
                                .Cells(Row1 + i + 1, 7) = "10時一般"
                            Case Else
                                .Cells(Row1 + i + 1, 7) = ""
                        End Select

                        '.Cells(Row1 + i + 1, 7) = Kubun1(i).Replace("職", "")

                        If Time1(i) = t0 Then
                            .Cells(Row1 + i + 1, 8) = ""
                        Else
                            .Cells(Row1 + i + 1, 8) = Time1(i)
                        End If

                        If MC1(i) <> "" And MC1(i) <> "なし" Then
                            .Cells(Row1 + i + 1, 9) = MC1(i)
                        Else
                            .Cells(Row1 + i + 1, 9) = ""
                        End If

                        If Time2(i) = t0 Then
                            .Cells(Row1 + i + 1, 10) = ""
                        Else
                            .Cells(Row1 + i + 1, 10) = Time2(i)
                        End If

                        If MC2(i) <> "" And MC2(i) <> "なし" Then
                            .Cells(Row1 + i + 1, 11) = MC2(i)
                        Else
                            .Cells(Row1 + i + 1, 11) = ""
                        End If

                        .Cells(Row1 + i + 1, 12) = GaishutuCount1(i)

                        .Cells(Row1 + i + 1, 13) = GaishutuCount2(i)

                        If ShoteiTime(i) > TimeSpan.Parse("00:00:00") Then
                            .Cells(Row1 + i + 1, 14) = ShoteiTime(i).ToString("%h\:mm")
                            ShoteiTime_sum += ShoteiTime(i)
                            Select Case ShoteiTime(i)
                                Case >= TimeSpan.Parse("4:45:00")
                                    所定日数 += 1.0
                                Case >= TimeSpan.Parse("1:45:00")
                                    所定日数 += 0.5
                            End Select
                        Else
                            .Cells(Row1 + i + 1, 14) = "----"
                        End If

                        If HouteiH(i) > TimeSpan.Parse("00:00:00") Then
                            .Cells(Row1 + i + 1, 15) = HouteiH(i).ToString("%h\:mm")
                            HouteiH_sum += HouteiH(i)
                        Else
                            .Cells(Row1 + i + 1, 15) = "----"
                        End If

                        If HougaiH(i) > TimeSpan.Parse("00:00:00") Then
                            .Cells(Row1 + i + 1, 16) = HougaiH(i).ToString("%h\:mm")
                            HougaiH_sum += HougaiH(i)
                        Else
                            .Cells(Row1 + i + 1, 16) = "----"
                        End If

                        If shinya(i) > TimeSpan.Parse("00:00:00") Then
                            .Cells(Row1 + i + 1, 17) = shinya(i).ToString("%h\:mm")
                            shinya_sum += shinya(i)
                        Else
                            .Cells(Row1 + i + 1, 17) = "----"
                        End If

                        If Hougai40(i) > TimeSpan.Parse("00:00:00") Then
                            .Cells(Row1 + i + 1, 18) = Hougai40(i).ToString("%h\:mm")
                            Hougai40_sum += Hougai40(i)
                        Else
                            .Cells(Row1 + i + 1, 18) = "----"
                        End If

                        If Chikoku(i) + Soutai(i) > TimeSpan.Parse("00:00:00") Then
                            .Cells(Row1 + i + 1, 19) = (Chikoku(i) + Soutai(i)).ToString("%h\:mm")
                            Chikoku_sum += Chikoku(i)
                            Soutai_sum += Soutai(i)
                        Else
                            .Cells(Row1 + i + 1, 19) = "----"
                        End If

                        If Gaishutu(i) > TimeSpan.Parse("00:00:00") Then
                            .Cells(Row1 + i + 1, 20) = Gaishutu(i).ToString("%h\:mm")
                            Gaishutu_sum += Gaishutu(i)
                        Else
                            .Cells(Row1 + i + 1, 20) = "----"
                        End If

                    Next

                    .Range(.Cells(Row1 + 1, 2), .Cells(Row1 + day_n, 2)).NumberFormat = "mm/dd"     ' 日付のフォーマット "01/01"
                    .Range(.Cells(Row1 + 1, 8), .Cells(Row1 + day_n, 8)).NumberFormat = "h:mm"     ' 時刻のフォーマット "12:00"
                    .Range(.Cells(Row1 + 1, 10), .Cells(Row1 + day_n, 8)).NumberFormat = "h:mm"     ' 時刻のフォーマット "12:00"

                    .Cells(Row1 + day_n + 1, 14) = "所定時間"
                    .Cells(Row1 + day_n + 2, 14) = (Integer.Parse(ShoteiTime_sum.ToString("%d")) * 24 + Integer.Parse(ShoteiTime_sum.ToString("%h"))).ToString + ShoteiTime_sum.ToString("\:mm")
                    .Range(.Cells(Row1 + day_n + 2, 14), .Cells(Row1 + day_n + 2, 14)).NumberFormat = "[h]:mm"

                    .Cells(Row1 + day_n + 1, 15) = "法定休H"
                    .Cells(Row1 + day_n + 2, 15) = (Integer.Parse(HouteiH_sum.ToString("%d")) * 24 + Integer.Parse(HouteiH_sum.ToString("%h"))).ToString + HouteiH_sum.ToString("\:mm")
                    .Range(.Cells(Row1 + day_n + 2, 15), .Cells(Row1 + day_n + 2, 15)).NumberFormat = "[h]:mm"

                    .Cells(Row1 + day_n + 1, 16) = "法定外H"
                    .Cells(Row1 + day_n + 2, 16) = (Integer.Parse(HougaiH_sum.ToString("%d")) * 24 + Integer.Parse(HougaiH_sum.ToString("%h"))).ToString + HougaiH_sum.ToString("\:mm")
                    .Range(.Cells(Row1 + day_n + 2, 16), .Cells(Row1 + day_n + 2, 16)).NumberFormat = "[h]:mm"

                    .Cells(Row1 + day_n + 1, 17) = "深夜残業"
                    .Cells(Row1 + day_n + 2, 17) = (Integer.Parse(shinya_sum.ToString("%d")) * 24 + Integer.Parse(shinya_sum.ToString("%h"))).ToString + shinya_sum.ToString("\:mm")
                    .Range(.Cells(Row1 + day_n + 2, 17), .Cells(Row1 + day_n + 2, 17)).NumberFormat = "[h]:mm"

                    .Cells(Row1 + day_n + 1, 18) = "法定外40"
                    .Cells(Row1 + day_n + 2, 18) = (Integer.Parse(Hougai40_sum.ToString("%d")) * 24 + Integer.Parse(Hougai40_sum.ToString("%h"))).ToString + Hougai40_sum.ToString("\:mm")
                    .Range(.Cells(Row1 + day_n + 2, 18), .Cells(Row1 + day_n + 2, 18)).NumberFormat = "[h]:mm"

                    .Cells(Row1 + day_n + 1, 19) = "遅早時間"
                    .Cells(Row1 + day_n + 2, 19) = (Integer.Parse((Chikoku_sum + Soutai_sum).ToString("%d")) * 24 + Integer.Parse((Chikoku_sum + Soutai_sum).ToString("%h"))).ToString + (Chikoku_sum + Soutai_sum).ToString("\:mm")
                    .Range(.Cells(Row1 + day_n + 2, 19), .Cells(Row1 + day_n + 2, 19)).NumberFormat = "[h]:mm"

                    .Cells(Row1 + day_n + 1, 20) = "私用外出"
                    .Cells(Row1 + day_n + 2, 20) = (Integer.Parse(Gaishutu_sum.ToString("%d")) * 24 + Integer.Parse(Gaishutu_sum.ToString("%h"))).ToString + Gaishutu_sum.ToString("\:mm")
                    .Range(.Cells(Row1 + day_n + 2, 20), .Cells(Row1 + day_n + 2, 20)).NumberFormat = "[h]:mm"


                    ' 罫線の書き込み
                    Dim border As Excel.Border = Nothing
                    ' 
                    Dim range3 As String = "A" + (Row1).ToString + ":T" + (Row1 + day_n + 2).ToString
                    ' 書式を中央揃え
                    .Range(range3).Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                    .Range(range3).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

                    ' レンジ内の横線を細線に設定
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlInsideHorizontal)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlThin

                    ' レンジ内の縦線を細線に設定
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlInsideVertical)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlThin

                    ' 表の上端を中細線に設定
                    range3 = "A" + (Row1).ToString + ":T" + (Row1).ToString
                    .Range(range3).Interior.Pattern = Excel.XlPattern.xlPatternGray25
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlMedium
                    ' 表の2段目の下端を中細線に設定
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlMedium

                    ' 表の下端を中細線に設定
                    range3 = "A" + (Row1 + day_n).ToString + ":T" + (Row1 + day_n).ToString
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlMedium
                    range3 = "A" + (Row1 + day_n).ToString + ":T" + (Row1 + day_n + 1).ToString
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlMedium
                    range3 = "A" + (Row1 + day_n).ToString + ":T" + (Row1 + day_n + 2).ToString
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlMedium

                    range3 = "A" + (Row1 + day_n + 1).ToString + ":T" + (Row1 + day_n + 1).ToString
                    .Range(range3).Interior.Pattern = Excel.XlPattern.xlPatternGray25

                    ' 表の左端を中細線に設定
                    range3 = "A" + (Row1).ToString + ":A" + (Row1 + day_n + 2).ToString
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlMedium

                    ' 表の右端を中細線に設定
                    range3 = "T" + (Row1).ToString + ":T" + (Row1 + day_n + 2).ToString
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlMedium

                    ' 表の欄外の別表のセルを結合
                    range3 = "A" + (Row1 + day_n + 4).ToString + ":B" + (Row1 + day_n + 4).ToString
                    .Range(range3).Merge()
                    .Cells(Row1 + day_n + 4, 1) = "所定日数"
                    range3 = "C" + (Row1 + day_n + 4).ToString + ":D" + (Row1 + day_n + 4).ToString
                    .Range(range3).Merge()
                    .Cells(Row1 + day_n + 5, 1) = 所定日数
                    .Range(.Cells(Row1 + day_n + 5, 1), .Cells(Row1 + day_n + 5, 1)).NumberFormat = "0.00"

                    .Cells(Row1 + day_n + 4, 3) = "有給日数"
                    range3 = "A" + (Row1 + day_n + 5).ToString + ":B" + (Row1 + day_n + 5).ToString
                    .Range(range3).Merge()
                    range3 = "C" + (Row1 + day_n + 5).ToString + ":D" + (Row1 + day_n + 5).ToString
                    .Range(range3).Merge()
                    .Cells(Row1 + day_n + 5, 3) = 有給日数
                    .Range(.Cells(Row1 + day_n + 5, 3), .Cells(Row1 + day_n + 5, 3)).NumberFormat = "0.00"

                    '　欄外の別表の罫線を設定
                    range3 = "A" + (Row1 + day_n + 4).ToString + ":D" + (Row1 + day_n + 5).ToString
                    .Range(range3).Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                    .Range(range3).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlInsideHorizontal)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlThin
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlInsideVertical)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlThin


                    range3 = "A" + (Row1 + day_n + 4).ToString + ":D" + (Row1 + day_n + 4).ToString
                    .Range(range3).Interior.Pattern = Excel.XlPattern.xlPatternGray25
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlMedium

                    range3 = "A" + (Row1 + day_n + 5).ToString + ":D" + (Row1 + day_n + 5).ToString
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlMedium

                    range3 = "A" + (Row1 + day_n + 4).ToString + ":A" + (Row1 + day_n + 5).ToString
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlMedium

                    range3 = "D" + (Row1 + day_n + 4).ToString + ":D" + (Row1 + day_n + 5).ToString
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlMedium





                    ' 表の欄外の承認欄のセルを結合
                    range3 = "R2:R3"
                    .Range(range3).Merge()
                    .Cells(1, 18) = "承認3"

                    range3 = "S2:S3"
                    .Range(range3).Merge()
                    .Cells(1, 19) = "承認2"

                    range3 = "T2:T3"
                    .Range(range3).Merge()
                    .Cells(1, 20) = "承認1"

                    '　欄外の承認欄の罫線を設定
                    range3 = "R1:T3"
                    .Range(range3).Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                    .Range(range3).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlInsideHorizontal)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlThin
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlInsideVertical)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlThin

                    range3 = "R1:T3"
                    '.Range(range3).Interior.Pattern = Excel.XlPattern.xlPatternGray25
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlThin

                    'range3 = "A" + (Row1 + day_n + 5).ToString + ":D" + (Row1 + day_n + 5).ToString
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlThin

                    'range3 = "A" + (Row1 + day_n + 4).ToString + ":A" + (Row1 + day_n + 5).ToString
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlThin

                    'range3 = "D" + (Row1 + day_n + 4).ToString + ":D" + (Row1 + day_n + 5).ToString
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlThin





                    ' 作成日の印字
                    range3 = "A1:D1"
                    .Range(range3).Merge()
                    .Range(range3).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                    .Range(range3).Font.Size = 7
                    '.Range(range3).Style = "yyyy/mm/dd"
                    .Cells(1, 1) = DateTime.Now

                    ' タイトルの印字
                    range3 = "I2:O2"
                    .Range(range3).Merge()
                    .Range(range3).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .Range(range3).Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                    .Range(range3).Font.Size = 14
                    .Range(range3).Font.FontStyle = "太字"
                    .Cells(2, 9) = "就業週報"
                    border = .Range(range3).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlMedium

                    ' 職員番号、氏名の印字
                    range3 = "A4:F4"
                    .Range(range3).Merge()
                    .Range(range3).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                    .Range(range3).Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                    .Range(range3).Font.Size = 9
                    .Range(range3).Font.Underline = True
                    .Cells(4, 1) = "個人コード：" + No(data_i) + "  " + Name2(0)

                    ' 所属部署の印字
                    range3 = "H4:L4"
                    .Range(range3).Merge()
                    .Range(range3).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                    .Range(range3).Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                    .Range(range3).Font.Size = 9
                    .Range(range3).Font.Underline = True
                    Dim aff1 As String = "所属 ： "
                    If Affiliation2(data_i) <> "なし" Then
                        aff1 += Affiliation2(data_i)
                    End If
                    If Affiliation3(data_i) <> "なし" Then
                        aff1 += " " + Affiliation3(data_i)
                    End If
                    .Cells(4, 8) = aff1

                    ' 期間の印字
                    range3 = "P4:T4"
                    .Range(range3).Merge()
                    .Range(range3).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    .Range(range3).Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                    .Range(range3).Font.Size = 9
                    .Range(range3).Font.Underline = True
                    .Cells(4, 16) = "処理期間　：　" + Year1 + "/" + Month1 + "/01 ～ " + Year1 + "/" + Month1 + "/" + day_n.ToString("00")

                End With

                Dim newFn As String = "C:\temp\" + Year1 + "_" + Month1 + ".pdf"

                xlBook.ExportAsFixedFormat(Type:=0,
                                           Filename:=newFn,
                                           Quality:=0,
                                           IncludeDocProperties:=True,
                                           IgnorePrintAreas:=True,
                                           OpenAfterPublish:=False)

                xlApp.DisplayAlerts = False
                xlBook.Close()

        End Select



        xlApp.Visible = False

        '3秒待機する
        System.Threading.Thread.Sleep(3000)

        xlApp.Quit()

        'COMオブジェクトを解放する
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
        'UPDATE " 従業員名簿 " SET " 給与 " =32000, " 控除 " =1 WHERE " 従業員番号 " = 'E10001'

        '                A = ComboBox2.Text
        '                B = ComboBox3.Text
        '                C = ComboBox4.Text
        '                'db.Connect()

        '                Sql_Command = "SELECT Idm FROM """ + MemberNameTable + """ WHERE ""職員番号"" = '" + Me.No(data_n) + "'"
        '                tb = db.ExecuteSql(Sql_Command)
        '                Dim n As Integer = tb.Rows.Count
        '                If n > 0 Then
        '                    Me.IDm(data_n) = tb.Rows(0).Item("IDm").ToString()
        '                    Me.DataGridView1.CurrentCell = DataGridView1(6, data_n)
        '                    Me.DataGridView1.CurrentCell.Value = Me.IDm(data_n)
        '                End If

        '                db.Disconnect()
        '                tb.Dispose()
        '            End If

        '            'メッセージを画面に表示
        '            Me.txtMessage_Text(msg + vbNewLine)
        '            'Me.Invoke(msg_txt, New Object() {msg + vbNewLine})
        '        Else
        '            'エラーメッセージを画面に表示
        '            Me.txtMessage_Text(pcsc.ErrorMsg + vbNewLine)
        '        End If

        '        MsgBox("カードに職員番号(" + Me.No(data_n) + ")が書き込まれました。")

        db.Disconnect()

    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Dim CartInput1 As New CardInputForm
        'CartInput1.Show()
        'Me.DialogResult = DialogResult.OK
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Name_Find()
    End Sub


    Private Sub TexBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles NameTextBox.KeyPress
        If e.KeyChar = Chr(13) Then 'chr(13)はEnterキー

            'Dim a As String
            'a = Me.NameTextBox.Text
            'コード()
            Me.Name_Find()
            'e.KeyChar(13) = "" 'キーをクリアする(必要であれば)
        End If
    End Sub

    Private Sub Name_Find()
        If Me.NameTextBox.Text <> "" Then
            Dim db As New OdbcDbIf
            Dim tb As DataTable
            Dim Sql_Command As String
            Dim i As Integer
            'Dim n As Integer
            Dim A As String
            Dim Row() As String

            A = Me.NameTextBox.Text

            db.Connect()

            Sql_Command = "SELECT * FROM """ + MemberNameTable + """ WHERE ""氏名"" LIKE '%" + A + "%' ORDER BY ""氏名"""
            tb = db.ExecuteSql(Sql_Command)
            Dim n As Integer = tb.Rows.Count
            If n > 0 Then
                ReDim No(n - 1), Name1(n - 1), Affiliation1(n - 1), Affiliation2(n - 1), Affiliation3(n - 1), Post(n - 1), IDm(n - 1)
                'Me.ComboBox1.Items.Clear()
                For i = 0 To n - 1
                    No(i) = tb.Rows(i).Item("職員番号").ToString()
                    Name1(i) = tb.Rows(i).Item("氏名").ToString()
                    Affiliation1(i) = tb.Rows(i).Item("所属センター").ToString()
                    Affiliation2(i) = tb.Rows(i).Item("所属部").ToString()
                    Affiliation3(i) = tb.Rows(i).Item("所属室").ToString()
                    Post(i) = tb.Rows(i).Item("役職").ToString()
                    IDm(i) = tb.Rows(i).Item("IDm").ToString()
                Next
            Else
                MsgBox("そのような名前は存在しません！")
                tb.Dispose()
                db.Disconnect()
                Exit Sub
            End If

            With Me.DataGridView1
                .Rows.Clear()
                For i = 0 To n - 1
                    Row = {No(i), Name1(i), Affiliation1(i), Affiliation2(i), Affiliation3(i), Post(i), IDm(i)}
                    .Rows.Add(Row)
                Next
            End With
            db.Disconnect()
            tb.Dispose()
        End If
    End Sub

End Class
