Public Class LoginForm1

    ' TODO: 指定されたユーザー名およびパスワードを使用して、カスタム認証を実行するコードを挿入します 
    ' ( http://go.microsoft.com/fwlink/?LinkId=35339 を参照してください)。  
    ' カスタム プリンシパルは、以下のように現在のスレッドのプリンシパルにアタッチできます:
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' この場合 CustomPrincipal は、認証を実行するのに使われる IPrincipal 実装です。
    ' これにより My.User は、ユーザー名および表示名などの CustomPrincipal オブジェクトに要約された ID 情報を
    ' 返します。

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        LoginID = Me.UsernameTextBox.Text
        LoginPassWord = Me.PasswordTextBox.Text
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub LoginForm1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.UsernameTextBox.Text = "kanyama"
        Me.PasswordTextBox.Text = "soumu2049"
    End Sub
End Class
