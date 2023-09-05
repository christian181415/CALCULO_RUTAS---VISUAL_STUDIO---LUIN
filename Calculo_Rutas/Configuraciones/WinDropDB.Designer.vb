<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WinDropDB
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(WinDropDB))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.BtnCCloseUp = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.BtnEliminar = New System.Windows.Forms.Button()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Maroon
        Me.Panel1.Controls.Add(Me.BtnCCloseUp)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.ForeColor = System.Drawing.Color.White
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(533, 54)
        Me.Panel1.TabIndex = 0
        '
        'BtnCCloseUp
        '
        Me.BtnCCloseUp.BackgroundImage = CType(resources.GetObject("BtnCCloseUp.BackgroundImage"), System.Drawing.Image)
        Me.BtnCCloseUp.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.BtnCCloseUp.FlatAppearance.BorderSize = 0
        Me.BtnCCloseUp.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(CType(CType(90, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.BtnCCloseUp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(90, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.BtnCCloseUp.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnCCloseUp.Location = New System.Drawing.Point(505, 3)
        Me.BtnCCloseUp.Name = "BtnCCloseUp"
        Me.BtnCCloseUp.Size = New System.Drawing.Size(15, 15)
        Me.BtnCCloseUp.TabIndex = 6
        Me.BtnCCloseUp.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label1.Location = New System.Drawing.Point(0, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(533, 54)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "LIBERACIÓN DE ESPACION EN DB"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(45, 91)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(443, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "¿Desea eliminar la informacion contenida en la base de datos?"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BtnEliminar
        '
        Me.BtnEliminar.BackColor = System.Drawing.SystemColors.Control
        Me.BtnEliminar.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.BtnEliminar.FlatAppearance.BorderSize = 0
        Me.BtnEliminar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Maroon
        Me.BtnEliminar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnEliminar.ForeColor = System.Drawing.Color.Red
        Me.BtnEliminar.Location = New System.Drawing.Point(0, 300)
        Me.BtnEliminar.Name = "BtnEliminar"
        Me.BtnEliminar.Size = New System.Drawing.Size(533, 40)
        Me.BtnEliminar.TabIndex = 2
        Me.BtnEliminar.Text = "ELIMINAR INFORMACIÓN"
        Me.BtnEliminar.UseVisualStyleBackColor = False
        '
        'txtPassword
        '
        Me.txtPassword.BackColor = System.Drawing.SystemColors.Control
        Me.txtPassword.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtPassword.Location = New System.Drawing.Point(48, 249)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(440, 15)
        Me.txtPassword.TabIndex = 3
        Me.txtPassword.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(90, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label3.Location = New System.Drawing.Point(45, 117)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(443, 42)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Una vez elimine la informacion de la base de datos no hay vuelta atrás. Por favor" &
    ", asegúrese de hacerlo."
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(90, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label4.Location = New System.Drawing.Point(45, 195)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(347, 16)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Ingrese la contraseña para confirmar su solicitud."
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(45, 221)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(90, 16)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "Contraseña:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'WinDropDB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(533, 340)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.BtnEliminar)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "WinDropDB"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "WinDropDB"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents BtnEliminar As Button
    Friend WithEvents txtPassword As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents BtnCCloseUp As Button
End Class
