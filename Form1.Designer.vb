<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Button1 = New Button()
        Button2 = New Button()
        btnUnzip = New Button()
        Button3 = New Button()
        Button4 = New Button()
        DataGridView1 = New DataGridView()
        Button5 = New Button()
        DateTimePicker1 = New DateTimePicker()
        CType(DataGridView1, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(420, 30)
        Button1.Name = "Button1"
        Button1.Size = New Size(75, 23)
        Button1.TabIndex = 0
        Button1.Text = "Connect"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' Button2
        ' 
        Button2.Location = New Point(501, 30)
        Button2.Name = "Button2"
        Button2.Size = New Size(75, 23)
        Button2.TabIndex = 1
        Button2.Text = "Close"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' btnUnzip
        ' 
        btnUnzip.Location = New Point(582, 30)
        btnUnzip.Name = "btnUnzip"
        btnUnzip.Size = New Size(75, 23)
        btnUnzip.TabIndex = 2
        btnUnzip.Text = "7-zip"
        btnUnzip.UseVisualStyleBackColor = True
        ' 
        ' Button3
        ' 
        Button3.Location = New Point(663, 30)
        Button3.Name = "Button3"
        Button3.Size = New Size(75, 23)
        Button3.TabIndex = 3
        Button3.Text = "Transfer"
        Button3.UseVisualStyleBackColor = True
        ' 
        ' Button4
        ' 
        Button4.Location = New Point(744, 30)
        Button4.Name = "Button4"
        Button4.Size = New Size(75, 23)
        Button4.TabIndex = 4
        Button4.Text = "LoadList"
        Button4.UseVisualStyleBackColor = True
        ' 
        ' DataGridView1
        ' 
        DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridView1.Location = New Point(12, 59)
        DataGridView1.Name = "DataGridView1"
        DataGridView1.Size = New Size(1355, 150)
        DataGridView1.TabIndex = 5
        ' 
        ' Button5
        ' 
        Button5.Location = New Point(825, 30)
        Button5.Name = "Button5"
        Button5.Size = New Size(95, 23)
        Button5.TabIndex = 6
        Button5.Text = "Transfer 1 row"
        Button5.UseVisualStyleBackColor = True
        ' 
        ' DateTimePicker1
        ' 
        DateTimePicker1.Location = New Point(1164, 28)
        DateTimePicker1.Name = "DateTimePicker1"
        DateTimePicker1.Size = New Size(200, 23)
        DateTimePicker1.TabIndex = 7
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1376, 283)
        Controls.Add(DateTimePicker1)
        Controls.Add(Button5)
        Controls.Add(DataGridView1)
        Controls.Add(Button4)
        Controls.Add(Button3)
        Controls.Add(btnUnzip)
        Controls.Add(Button2)
        Controls.Add(Button1)
        Name = "Form1"
        Text = "Form1"
        CType(DataGridView1, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents btnUnzip As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Button5 As Button
    Friend WithEvents DateTimePicker1 As DateTimePicker

End Class
