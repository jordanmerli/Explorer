' Copyright (C) Microsoft Corporation. Tutti i diritti riservati.
Public Class MainForm
    ' Gestisce l'evento Click per il pulsante "Avvia scanner directory".
    Private Sub btnSimple_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Simple.Click
        Dim sdv As New DirectoryScanner()
        sdv.Show()
    End Sub

    ' Gestisce l'evento Click per il pulsante "Avvia visualizzatore di tipo Internet Explorer".
    Private Sub btnExplorer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Explorer.Click
        Dim esfv As New ExplorerStyleViewer()
        esfv.Show()
    End Sub
    Private Sub exitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles exitToolStripMenuItem.Click
        Form1.Close()
        Me.Close()
    End Sub

    Private Sub InfoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InfoToolStripMenuItem.Click
        Form1.Show()
    End Sub
End Class
