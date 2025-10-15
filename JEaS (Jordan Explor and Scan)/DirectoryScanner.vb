' Copyright (C) Microsoft Corporation. Tutti i diritti riservati.
Imports System.IO

Public Class DirectoryScanner

    Const MB As Long = 1024 * 1024

    ''' <summary>
    ''' Gestisce l'evento Load per DirectoryScanner.
    ''' </summary>
    Private Sub DirectoryScanner_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Impostare le colonne iniziali di ListView
        ListView1.Columns.Add("Size", 90, HorizontalAlignment.Left)
        ListView1.Columns.Add("Folder Name", 400, HorizontalAlignment.Left)
    End Sub

    ''' <summary>
    ''' Questa subroutine aggiunge l'oggetto strSubDirectory selezionato dall'utente in TreeView 
    ''' a ListView e imposta testo, dimensione e colore.
    ''' </summary>
    Private Sub AddToListView(ByVal strSize As String, ByVal strFolderName As String)
        Dim listViewItem As New ListViewItem()
        Dim listViewSubItem As ListViewItem.ListViewSubItem

        listViewItem.Text = strSize
        listViewItem.ForeColor = GetSizeColor(strSize)

        listViewSubItem = New ListViewItem.ListViewSubItem()
        listViewSubItem.Text = strFolderName

        listViewItem.SubItems.Add(listViewSubItem)
        ListView1.Items.Add(listViewItem)
    End Sub

    ''' <summary>
    ''' Questa subroutine restituisce un colore in base alla combinazione delle dimensioni della directory 
    ''' e di tutte le relative sottodirectory. Si tratta di uno dei due overload disponibili.
    ''' </summary>
    Private Function GetSizeColor(ByVal strSize As String) As System.Drawing.Color
        Return GetSizeColor(CLng(CDbl(strSize.Substring(0, _
            strSize.LastIndexOf("M") - 1)) * MB))
    End Function

    ''' <summary>
    ''' Questa funzione restituisce un colore in base alla combinazione delle dimensioni della directory 
    ''' e di tutte le relative sottodirectory. Si tratta del secondo dei due overload disponibili.
    ''' </summary>
    Private Function GetSizeColor(ByVal intSize As Long) As System.Drawing.Color
        Select Case intSize
            Case 200 * MB To 500 * MB
                Return System.Drawing.Color.Gold
            Case Is > 500 * MB
                Return System.Drawing.Color.Red
            Case Else
                Return System.Drawing.Color.Green
        End Select
    End Function

    ''' <summary>
    ''' Questa funzione restituisce le dimensioni di una directory e di tutte le relative sottodirectory.
    ''' </summary>
    Public Function GetDirectorySize(ByVal strDirPath As String, _
        ByVal dnDriveOrDirectory As DirectoryNode) As Long

        Try
            Dim astrSubDirectories As String() = Directory.GetDirectories(strDirPath)
            Dim strSubDirectory As String

            ' La dimensione della directory corrente dipende dalle dimensioni
            ' delle sottodirectory nella matrice astrSubDirectories. Scorrere pertanto
            ' la matrice e utilizzare la ricorsione per ottenere le dimensioni
            ' totali della directory corrente e di tutte le relative sottodirectory.
            For Each strSubDirectory In astrSubDirectories
                Dim dnSubDirectoryNode As DirectoryNode
                dnSubDirectoryNode = New DirectoryNode()

                ' Impostare il testo del nodo utilizzando solo l'ultima parte del percorso completo.
                dnSubDirectoryNode.Text = _
                    strSubDirectory.Remove(0, strSubDirectory.LastIndexOf("\") + 1)

                ' Si noti che la seguente riga è ricorsiva.
                dnDriveOrDirectory.Size += _
                    GetDirectorySize(strSubDirectory, dnSubDirectoryNode)
                dnDriveOrDirectory.Nodes.Add(dnSubDirectoryNode)
            Next

            ' Aggiungere alla dimensione calcolata in precedenza tutti i file nella directory
            ' corrente.
            Dim astrFiles As String() = Directory.GetFiles(strDirPath)
            Dim strFileName As String
            Dim Size As Long = 0

            For Each strFileName In astrFiles
                dnDriveOrDirectory.Size += New FileInfo(strFileName).Length
            Next

            ' Impostare il colore di TreeNode in base alla dimensione totale calcolata.
            dnDriveOrDirectory.ForeColor = _
                GetSizeColor(dnDriveOrDirectory.Size)

        Catch exc As Exception
            ' Non eseguire alcuna operazione. Ignorare semplicemente le directory che non è possibile leggere. Il controllo
            ' passa alla prima riga dopo End Try.
        End Try

        ' Restituire la dimensione totale di questa directory.
        Return dnDriveOrDirectory.Size

    End Function

    ''' <summary>
    ''' Quando il nodo di una directory viene espanso, aggiungere le relative sottodirectory al controllo ListView. 
    ''' </summary>
    Public Sub ShowSubDirectories(ByVal dnDrive As DirectoryNode)
        Dim strSubDirectory As DirectoryNode

        ListView1.Items.Clear()

        For Each strSubDirectory In dnDrive.Nodes
            AddToListView(Format(strSubDirectory.Size / MB, "F") + "MB", _
                strSubDirectory.Text)
        Next
    End Sub

    ''' <summary>
    ''' Chiudere l'applicazione
    ''' </summary>
    Private Sub exitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles exitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub AllDirectoriesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllDirectoriesToolStripMenuItem.Click
        Me.Cursor = Cursors.WaitCursor

        ' Ottenere una matrice di tutte le unità logiche.
        Dim drives As String() = Directory.GetLogicalDrives()
        Dim drive As String

        TreeView1.Nodes.Clear()
        ListView1.Items.Clear()

        For Each drive In drives
            Dim dnDrive As DirectoryNode

            Try
                ' Creare una classe DirectoryNode che rappresenti ogni unità logica e aggiungerla
                ' a TreeView.
                dnDrive = New DirectoryNode()
                dnDrive.Text = drive.Remove(Len(drive) - 1, 1)
                TreeView1.Nodes.Add(dnDrive)

                ' Calcolare la dimensione dell'unità sommando le dimensioni di tutte le relative
                ' sottodirectory.
                dnDrive.Size += GetDirectorySize(drive, dnDrive)
            Catch exc As Exception
                ' Non eseguire alcuna operazione. Ignorare semplicemente le directory che non è possibile leggere. Il controllo
                ' passa alla prima riga dopo End Try.
            End Try
        Next
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub FromOneDirectoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FromOneDirectoryToolStripMenuItem.Click
        Me.Cursor = Cursors.WaitCursor
        ' Visualizzare la finestra di dialogo FolderBrowser e impostare la directory iniziale in base alla
        ' selezione dell'utente.
        Dim SelectedDirectory As String = FolderBrowser.ShowDialog()
        Dim SelectedDirectoryNode As DirectoryNode

        TreeView1.Nodes.Clear()
        ListView1.Items.Clear()

        Try
            ' Aggiungere la classe DirectoryNode che rappresenta la directory selezionata a
            ' TreeView.
            SelectedDirectoryNode = New DirectoryNode()
            SelectedDirectoryNode.Text = SelectedDirectory
            TreeView1.Nodes.Add(SelectedDirectoryNode)

            '	Calcolare la dimensione della directory selezionata sommando le dimensioni di tutte le relative
            ' sottodirectory.
            SelectedDirectoryNode.Size += GetDirectorySize(SelectedDirectory, _
                SelectedDirectoryNode)

        Catch exc As Exception
            ' Non eseguire alcuna operazione. Ignorare semplicemente le directory che non è possibile leggere. Il controllo
            ' passa alla prima riga dopo End Try.
        End Try
        Me.Cursor = Cursors.Arrow
    End Sub

    ''' <summary>
    ''' Gestisce l'evento AfterExpand per TreeView, che non si verifica dopo la
    ''' selezione di TreeView, bensì dopo che tramite l'applicazione è stato stabilito che il tentativo dell'utente
    ''' di espandere il nodo può essere consentito. Il gestore eventi BeforeExpand 
    ''' viene utilizzato per prendere questa decisione, se lo si desidera. Tutti gli eventi Before______ passano un oggetto TreeViewCancelEventArgs contenente una proprietà Cancel.
    ''' Questa proprietà può essere utilizzata per respingere l'azione dell'utente. L'evento "AfterExpand"
    ''' potrebbe pertanto essere denominato anche "AfterExpandApproval".
    ''' </summary>
    Private Sub TreeView1_AfterExpand(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterExpand
        e.Node.Expand()
        ShowSubDirectories(CType(e.Node, DirectoryNode))
    End Sub

    ''' <summary>
    ''' Gestisce l'evento AfterSelect per TreeView, che non si verifica dopo la 
    ''' selezione di TreeView, bensì dopo che tramite l'applicazione è stato stabilito che il tentativo dell'utente 
    ''' di selezionare il nodo può essere consentito. Il gestore eventi BeforeSelect 
    ''' viene utilizzato per prendere questa decisione, se lo si desidera. Tutti gli eventi Before______ passano un oggetto TreeViewCancelEventArgs contenente una proprietà Cancel.
    ''' Questa proprietà può essere utilizzata per respingere l'azione dell'utente. L'evento "AfterSelect"
    ''' potrebbe pertanto essere denominato anche "AfterSelectApproval".
    ''' </summary>
    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        Dim SubDirectory As DirectoryNode = CType(e.Node, DirectoryNode)
        ListView1.Items.Clear()
        AddToListView(Format(SubDirectory.Size / (1024 * 1024), "F") + "MB", _
            SubDirectory.Text)
    End Sub
End Class