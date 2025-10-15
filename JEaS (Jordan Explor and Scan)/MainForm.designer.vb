' Copyright (c) Microsoft Corporation. All rights reserved.
Partial Public Class MainForm
    Inherits System.Windows.Forms.Form

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

    End Sub

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.Explorer = New System.Windows.Forms.Button()
        Me.Simple = New System.Windows.Forms.Button()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.fileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.InfoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.exitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Explorer
        '
        Me.Explorer.AccessibleDescription = "Button with text ""Launch Explorer-Style Viewer"""
        Me.Explorer.AccessibleName = "Launch Explorer-Style Viewer"
        Me.Explorer.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Explorer.Location = New System.Drawing.Point(250, 33)
        Me.Explorer.Name = "Explorer"
        Me.Explorer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Explorer.Size = New System.Drawing.Size(232, 23)
        Me.Explorer.TabIndex = 5
        Me.Explorer.Text = "Lancia il J-Esplora Risorse"
        '
        'Simple
        '
        Me.Simple.AccessibleDescription = "Button with text ""Launch Directory Scanner"""
        Me.Simple.AccessibleName = "Launch Directory Scanner"
        Me.Simple.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Simple.Location = New System.Drawing.Point(12, 33)
        Me.Simple.Name = "Simple"
        Me.Simple.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Simple.Size = New System.Drawing.Size(232, 23)
        Me.Simple.TabIndex = 4
        Me.Simple.Text = "Lancia la J-Scansione delle Directory"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.fileToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(494, 24)
        Me.MenuStrip1.TabIndex = 6
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'fileToolStripMenuItem
        '
        Me.fileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.InfoToolStripMenuItem, Me.exitToolStripMenuItem})
        Me.fileToolStripMenuItem.Name = "fileToolStripMenuItem"
        Me.fileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.fileToolStripMenuItem.Text = "&File"
        '
        'InfoToolStripMenuItem
        '
        Me.InfoToolStripMenuItem.Name = "InfoToolStripMenuItem"
        Me.InfoToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.InfoToolStripMenuItem.Text = "Info"
        '
        'exitToolStripMenuItem
        '
        Me.exitToolStripMenuItem.Name = "exitToolStripMenuItem"
        Me.exitToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.exitToolStripMenuItem.Text = "&Esci"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 59)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(213, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "I colori che verranno visualizzati significano:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.Color.Red
        Me.Label2.Location = New System.Drawing.Point(9, 107)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(89, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Rosso= > 500MB"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.Color.Green
        Me.Label3.Location = New System.Drawing.Point(9, 81)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Verde = 0MB - 200MB"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.Color.Gold
        Me.Label4.Location = New System.Drawing.Point(9, 94)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(113, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Oro = 200MB - 500MB"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(247, 59)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(235, 35)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Potrebbero passare alcuni secondi, ma non ti preoccupare non si blocca niente. Tu" & _
    "tto OK!!!"
        '
        'MainForm
        '
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(494, 125)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.Simple)
        Me.Controls.Add(Me.Explorer)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.Name = "MainForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Scegli l'applicazione desiderata"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Explorer As System.Windows.Forms.Button
    Friend WithEvents Simple As System.Windows.Forms.Button
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents fileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents exitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents InfoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
