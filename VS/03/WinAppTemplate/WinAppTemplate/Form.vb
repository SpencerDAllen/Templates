Public Class Form
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
Friend WithEvents Label As System.Windows.Forms.Label
Friend WithEvents TextBox As System.Windows.Forms.TextBox
Friend WithEvents Button As System.Windows.Forms.Button
Friend WithEvents ListBox As System.Windows.Forms.ListBox
Friend WithEvents butDel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.Label = New System.Windows.Forms.Label
Me.TextBox = New System.Windows.Forms.TextBox
Me.Button = New System.Windows.Forms.Button
Me.ListBox = New System.Windows.Forms.ListBox
Me.butDel = New System.Windows.Forms.Button
Me.SuspendLayout()
'
'Label
'
Me.Label.Location = New System.Drawing.Point(8, 8)
Me.Label.Name = "Label"
Me.Label.Size = New System.Drawing.Size(272, 23)
Me.Label.TabIndex = 0
'
'TextBox
'
Me.TextBox.Location = New System.Drawing.Point(8, 49)
Me.TextBox.Name = "TextBox"
Me.TextBox.TabIndex = 1
Me.TextBox.Text = ""
'
'Button
'
Me.Button.Location = New System.Drawing.Point(8, 87)
Me.Button.Name = "Button"
Me.Button.TabIndex = 2
Me.Button.Text = "Click"
'
'ListBox
'
Me.ListBox.Location = New System.Drawing.Point(8, 128)
Me.ListBox.Name = "ListBox"
Me.ListBox.Size = New System.Drawing.Size(120, 95)
Me.ListBox.TabIndex = 3
'
'butDel
'
Me.butDel.Location = New System.Drawing.Point(104, 88)
Me.butDel.Name = "butDel"
Me.butDel.TabIndex = 4
Me.butDel.Text = "Delete"
'
'Form
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(292, 266)
Me.Controls.Add(Me.butDel)
Me.Controls.Add(Me.ListBox)
Me.Controls.Add(Me.Button)
Me.Controls.Add(Me.TextBox)
Me.Controls.Add(Me.Label)
Me.Name = "Form"
Me.Text = "Form"
Me.ResumeLayout(False)

    End Sub

#End Region

Private Sub Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button.Click
ListBox.Items.Clear()
GetSubDir(TextBox.Text)
End Sub

Private Sub TextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox.TextChanged

End Sub

Private Sub butDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butDel.Click
  Remove()
End Sub

Private Sub ListBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox.SelectedIndexChanged

End Sub
End Class
