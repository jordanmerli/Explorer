' Copyright (C) Microsoft Corporation. Tutti i diritti riservati.
Imports System.IO

Public Class ExplorerStyleViewer
    ' Dichiarare le variabili per contenere le istanze per ognuna delle classi personalizzate
    Private dtvwDirectory As DirectoryTreeView
    Private flvFiles As FileListView
    Private mivChecked As MenuItemView

    ' Gestisce l'evento AfterSelect per DirectoryTreeView, che fa in modo che
    ' l'oggetto FileListView visualizzi i contenuti della directory selezionata.
    Sub DirectoryTreeViewOnAfterSelect(ByVal obj As Object, ByVal tvea As TreeViewEventArgs)
        flvFiles.ShowFiles(tvea.Node.FullPath)
    End Sub

    ' Questa subroutine gestisce l'evento FormLoad.
    Private Sub ExplorerStyleViewer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Creare un'istanza di flvFilesView.
        flvFiles = New FileListView()
        SplitContainer1.Panel2.Controls.Add(flvFiles)
        flvFiles.Dock = DockStyle.Fill

        ' Creare un'istanza di DirectoryTreeView e aggiungere un gestore eventi OnAfterSelect.
        dtvwDirectory = New DirectoryTreeView()
        'dtvwDirectory.Parent = Me
        SplitContainer1.Panel1.Controls.Add(dtvwDirectory)
        dtvwDirectory.Dock = DockStyle.Left
        ' Aggiungere dinamicamente un gestore eventi AfterSelect.
        AddHandler dtvwDirectory.AfterSelect, _
            AddressOf DirectoryTreeViewOnAfterSelect

        ' Aggiungere un comando di menu View al menu principale esistente.
        Dim menuView As New ToolStripMenuItem("&View")
        MenuStrip1.Items.Add(menuView)

        ' Aggiungere quattro voci di menu al menu View. innanzitutto, creare matrici per impostare
        ' le proprietà di ogni voce di menu.
        Dim astrView As String() = {"L&arge Icons", "S&mall Icons", "&List", "&Details"}
        Dim aview As View() = {View.LargeIcon, View.SmallIcon, View.List, View.Details}
        ' Creare un gestore eventi per le voci di menu.
        Dim eh As New EventHandler(AddressOf MenuOnViewSelect)

        Dim i As Integer
        For i = 0 To 3
            ' Utilizzare una classe personalizzata MenuItemView, che estende MenuItem per supportare una
            ' proprietà View.
            Dim miv As New MenuItemView()
            miv.Text = astrView(i)
            miv.View = aview(i)
            miv.Checked = False
            ' Associare il gestore creato in precedenza all'evento Click.
            AddHandler miv.Click, eh

            ' Impostare la visualizzazione Details come visualizzazione predefinita.
            If i = 3 Then
                mivChecked = miv
                mivChecked.Checked = True
                flvFiles.View = mivChecked.View
            End If
            ' Aggiungere la nuova voce di menu al menu View.
            menuView.DropDownItems.Add(miv)
        Next i
    End Sub

    ' Gestisce l'evento OnViewSelect per le voci del menu View.
    Sub MenuOnViewSelect(ByVal obj As Object, ByVal ea As EventArgs)
        ' Deselezionare la voce correntemente selezionata.
        mivChecked.Checked = False
        ' Eseguire il cast del mittente dell'evento e selezionarlo.
        mivChecked = CType(obj, MenuItemView)
        mivChecked.Checked = True
        ' Modificare la modalità di visualizzazione dei file nel controllo FileListView.
        flvFiles.View = mivChecked.View
    End Sub

    Private Sub exitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles exitToolStripMenuItem.Click
        Me.Close()
    End Sub
End Class