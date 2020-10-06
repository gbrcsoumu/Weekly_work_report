<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FelicaRWForm
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.ComboBoxCenter = New System.Windows.Forms.ComboBox()
        Me.ComboBoxDepartment = New System.Windows.Forms.ComboBox()
        Me.ComboBoxSection = New System.Windows.Forms.ComboBox()
        Me.Label_Center = New System.Windows.Forms.Label()
        Me.Label_Department = New System.Windows.Forms.Label()
        Me.Label_Section = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.NameTextBox = New System.Windows.Forms.TextBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.ComboBoxYear = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBoxMonth = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(44, 285)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 33
        Me.DataGridView1.Size = New System.Drawing.Size(1358, 752)
        Me.DataGridView1.TabIndex = 9
        '
        'ComboBoxCenter
        '
        Me.ComboBoxCenter.Font = New System.Drawing.Font("MS UI Gothic", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.ComboBoxCenter.FormattingEnabled = True
        Me.ComboBoxCenter.Location = New System.Drawing.Point(44, 218)
        Me.ComboBoxCenter.Name = "ComboBoxCenter"
        Me.ComboBoxCenter.Size = New System.Drawing.Size(446, 45)
        Me.ComboBoxCenter.TabIndex = 10
        '
        'ComboBoxDepartment
        '
        Me.ComboBoxDepartment.Font = New System.Drawing.Font("MS UI Gothic", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.ComboBoxDepartment.FormattingEnabled = True
        Me.ComboBoxDepartment.Location = New System.Drawing.Point(567, 218)
        Me.ComboBoxDepartment.Name = "ComboBoxDepartment"
        Me.ComboBoxDepartment.Size = New System.Drawing.Size(371, 45)
        Me.ComboBoxDepartment.TabIndex = 11
        '
        'ComboBoxSection
        '
        Me.ComboBoxSection.Font = New System.Drawing.Font("MS UI Gothic", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.ComboBoxSection.FormattingEnabled = True
        Me.ComboBoxSection.Location = New System.Drawing.Point(1019, 218)
        Me.ComboBoxSection.Name = "ComboBoxSection"
        Me.ComboBoxSection.Size = New System.Drawing.Size(384, 45)
        Me.ComboBoxSection.TabIndex = 12
        '
        'Label_Center
        '
        Me.Label_Center.AutoSize = True
        Me.Label_Center.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label_Center.Location = New System.Drawing.Point(55, 176)
        Me.Label_Center.Name = "Label_Center"
        Me.Label_Center.Size = New System.Drawing.Size(177, 33)
        Me.Label_Center.TabIndex = 13
        Me.Label_Center.Text = "所属センター"
        '
        'Label_Department
        '
        Me.Label_Department.AutoSize = True
        Me.Label_Department.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label_Department.Location = New System.Drawing.Point(572, 176)
        Me.Label_Department.Name = "Label_Department"
        Me.Label_Department.Size = New System.Drawing.Size(111, 33)
        Me.Label_Department.TabIndex = 13
        Me.Label_Department.Text = "所属部"
        '
        'Label_Section
        '
        Me.Label_Section.AutoSize = True
        Me.Label_Section.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label_Section.Location = New System.Drawing.Point(1014, 176)
        Me.Label_Section.Name = "Label_Section"
        Me.Label_Section.Size = New System.Drawing.Size(159, 33)
        Me.Label_Section.TabIndex = 13
        Me.Label_Section.Text = "所属室・課"
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("MS UI Gothic", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Button3.Location = New System.Drawing.Point(1859, 18)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(255, 154)
        Me.Button3.TabIndex = 14
        Me.Button3.Text = "閉じる"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'NameTextBox
        '
        Me.NameTextBox.Font = New System.Drawing.Font("MS UI Gothic", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.NameTextBox.Location = New System.Drawing.Point(1240, 96)
        Me.NameTextBox.Name = "NameTextBox"
        Me.NameTextBox.Size = New System.Drawing.Size(352, 45)
        Me.NameTextBox.TabIndex = 16
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("MS UI Gothic", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Button4.Location = New System.Drawing.Point(1240, 18)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(276, 72)
        Me.Button4.TabIndex = 17
        Me.Button4.Text = "名前検索"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'ComboBoxYear
        '
        Me.ComboBoxYear.Font = New System.Drawing.Font("MS UI Gothic", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.ComboBoxYear.FormattingEnabled = True
        Me.ComboBoxYear.Location = New System.Drawing.Point(1435, 218)
        Me.ComboBoxYear.Name = "ComboBoxYear"
        Me.ComboBoxYear.Size = New System.Drawing.Size(214, 45)
        Me.ComboBoxYear.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(1430, 176)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 33)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "年度"
        '
        'ComboBoxMonth
        '
        Me.ComboBoxMonth.Font = New System.Drawing.Font("MS UI Gothic", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.ComboBoxMonth.FormattingEnabled = True
        Me.ComboBoxMonth.Location = New System.Drawing.Point(1677, 218)
        Me.ComboBoxMonth.Name = "ComboBoxMonth"
        Me.ComboBoxMonth.Size = New System.Drawing.Size(214, 45)
        Me.ComboBoxMonth.TabIndex = 12
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(1672, 176)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(47, 33)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "月"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(1932, 218)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(182, 28)
        Me.CheckBox1.TabIndex = 19
        Me.CheckBox1.Text = "半月分の表示"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'FelicaRWForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(13.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(2132, 1054)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.NameTextBox)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label_Section)
        Me.Controls.Add(Me.Label_Department)
        Me.Controls.Add(Me.Label_Center)
        Me.Controls.Add(Me.ComboBoxMonth)
        Me.Controls.Add(Me.ComboBoxYear)
        Me.Controls.Add(Me.ComboBoxSection)
        Me.Controls.Add(Me.ComboBoxDepartment)
        Me.Controls.Add(Me.ComboBoxCenter)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "FelicaRWForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "就業週報作成システム v1.0"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents ComboBoxCenter As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxDepartment As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxSection As System.Windows.Forms.ComboBox
    Friend WithEvents Label_Center As System.Windows.Forms.Label
    Friend WithEvents Label_Department As System.Windows.Forms.Label
    Friend WithEvents Label_Section As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents NameTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents ComboBoxYear As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents ComboBoxMonth As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents CheckBox1 As CheckBox
End Class
