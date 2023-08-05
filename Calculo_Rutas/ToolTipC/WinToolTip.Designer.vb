<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WinToolTip
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(WinToolTip))
        Me.PBoxPicture = New System.Windows.Forms.PictureBox()
        Me.PTop = New System.Windows.Forms.Panel()
        Me.LHeigth = New System.Windows.Forms.Label()
        Me.LWidth = New System.Windows.Forms.Label()
        Me.LMilisegundos = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.PText = New System.Windows.Forms.Panel()
        Me.LText = New System.Windows.Forms.Label()
        CType(Me.PBoxPicture, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PTop.SuspendLayout()
        Me.PText.SuspendLayout()
        Me.SuspendLayout()
        '
        'PBoxPicture
        '
        Me.PBoxPicture.BackColor = System.Drawing.SystemColors.Control
        Me.PBoxPicture.BackgroundImage = CType(resources.GetObject("PBoxPicture.BackgroundImage"), System.Drawing.Image)
        Me.PBoxPicture.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.PBoxPicture.Dock = System.Windows.Forms.DockStyle.Top
        Me.PBoxPicture.Location = New System.Drawing.Point(0, 11)
        Me.PBoxPicture.Name = "PBoxPicture"
        Me.PBoxPicture.Size = New System.Drawing.Size(150, 88)
        Me.PBoxPicture.TabIndex = 10
        Me.PBoxPicture.TabStop = False
        '
        'PTop
        '
        Me.PTop.BackColor = System.Drawing.SystemColors.Control
        Me.PTop.Controls.Add(Me.LHeigth)
        Me.PTop.Controls.Add(Me.LWidth)
        Me.PTop.Controls.Add(Me.LMilisegundos)
        Me.PTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.PTop.Location = New System.Drawing.Point(0, 0)
        Me.PTop.Name = "PTop"
        Me.PTop.Size = New System.Drawing.Size(150, 11)
        Me.PTop.TabIndex = 11
        '
        'LHeigth
        '
        Me.LHeigth.AutoSize = True
        Me.LHeigth.ForeColor = System.Drawing.SystemColors.ScrollBar
        Me.LHeigth.Location = New System.Drawing.Point(146, 1)
        Me.LHeigth.Name = "LHeigth"
        Me.LHeigth.Size = New System.Drawing.Size(0, 13)
        Me.LHeigth.TabIndex = 10
        Me.LHeigth.Visible = False
        '
        'LWidth
        '
        Me.LWidth.AutoSize = True
        Me.LWidth.ForeColor = System.Drawing.SystemColors.ScrollBar
        Me.LWidth.Location = New System.Drawing.Point(146, -1)
        Me.LWidth.Name = "LWidth"
        Me.LWidth.Size = New System.Drawing.Size(0, 13)
        Me.LWidth.TabIndex = 9
        Me.LWidth.Visible = False
        '
        'LMilisegundos
        '
        Me.LMilisegundos.AutoSize = True
        Me.LMilisegundos.ForeColor = System.Drawing.SystemColors.ScrollBar
        Me.LMilisegundos.Location = New System.Drawing.Point(3, 1)
        Me.LMilisegundos.Name = "LMilisegundos"
        Me.LMilisegundos.Size = New System.Drawing.Size(31, 13)
        Me.LMilisegundos.TabIndex = 8
        Me.LMilisegundos.Text = "1000"
        Me.LMilisegundos.Visible = False
        '
        'Timer1
        '
        '
        'PText
        '
        Me.PText.BackColor = System.Drawing.SystemColors.Control
        Me.PText.Controls.Add(Me.LText)
        Me.PText.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PText.Location = New System.Drawing.Point(0, 99)
        Me.PText.Name = "PText"
        Me.PText.Size = New System.Drawing.Size(150, 38)
        Me.PText.TabIndex = 12
        '
        'LText
        '
        Me.LText.AutoSize = True
        Me.LText.BackColor = System.Drawing.SystemColors.Control
        Me.LText.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LText.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LText.Location = New System.Drawing.Point(0, 0)
        Me.LText.MaximumSize = New System.Drawing.Size(600, 0)
        Me.LText.Name = "LText"
        Me.LText.Padding = New System.Windows.Forms.Padding(5, 5, 10, 15)
        Me.LText.Size = New System.Drawing.Size(131, 37)
        Me.LText.TabIndex = 2
        Me.LText.Text = "Corporativo LUIN"
        Me.LText.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'WinToolTip
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(150, 137)
        Me.Controls.Add(Me.PText)
        Me.Controls.Add(Me.PBoxPicture)
        Me.Controls.Add(Me.PTop)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MinimumSize = New System.Drawing.Size(150, 33)
        Me.Name = "WinToolTip"
        Me.Opacity = 0.9R
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "WinToolTip"
        Me.TransparencyKey = System.Drawing.Color.FromArgb(CType(CType(91, Byte), Integer), CType(CType(91, Byte), Integer), CType(CType(91, Byte), Integer))
        CType(Me.PBoxPicture, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PTop.ResumeLayout(False)
        Me.PTop.PerformLayout()
        Me.PText.ResumeLayout(False)
        Me.PText.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents PBoxPicture As PictureBox
    Friend WithEvents PTop As Panel
    Friend WithEvents LMilisegundos As Label
    Friend WithEvents Timer1 As Timer
    Friend WithEvents PText As Panel
    Public WithEvents LText As Label
    Friend WithEvents LWidth As Label
    Friend WithEvents LHeigth As Label
End Class
