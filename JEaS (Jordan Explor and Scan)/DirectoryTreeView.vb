' Copyright (C) Microsoft Corporation. Tutti i diritti riservati.
Imports System.IO

Class DirectoryTreeView
    Inherits TreeView

    ' Questo è il costruttore della classe.
    Public Sub New()
        ' Creare un po' più di spazio per i nomi di directory più lunghi.
        Me.Width *= 2

        ' Ottenere le immagini per la struttura.
        Me.ImageList = New ImageList()
        With Me.ImageList.Images
            .Add(My.Resources.FLOPPY)
            .Add(My.Resources.CLSDFOLD)
            .Add(My.Resources.OPENFOLD)
        End With

        ' Creare la struttura.
        RefreshTree()
    End Sub

    ' Gestisce l'evento BeforeExpand per il controllo TreeView sottoclassato. Per ulteriori informazioni
    ' sulle coppie di eventi Before_____ e After_______ TreeView, vedere
    ' /DirectoryScanner/DirectoryScanner.vb.
    Protected Overrides Sub OnBeforeExpand(ByVal tvcea As TreeViewCancelEventArgs)
        MyBase.OnBeforeExpand(tvcea)

        ' Per garantire buone prestazioni e per evitare che il controllo TreeView sia instabile durante
        ' l'aggiornamento di un nodo di grandi dimensioni, è consigliabile eseguire il wrapping del codice di aggiornamento nelle istruzioni BeginUpdate...
        ' EndUpdate.
        Me.BeginUpdate()

        Dim tn As TreeNode
        ' Aggiungere nodi figlio per ogni nodo figlio nel nodo in cui l'utente fa clic. 
        ' Per garantire buone prestazioni, ogni nodo in DirectoryTreeView 
        ' contiene solo il livello successivo di nodi figlio e viene visualizzato il segno +
        ' per indicare se l'utente può espandere il nodo. Quando pertanto un utente espande
        ' un nodo, affinché il segno + venga correttamente visualizzato per il
        ' livello successivo di nodi figlio, è necessario aggiungere i *relativi* nodi figlio.
        For Each tn In tvcea.Node.Nodes
            AddDirectories(tn)
        Next tn

        Me.EndUpdate()
    End Sub

    ' Questa subroutine viene utilizzata per aggiungere un nodo figlio a ogni directory sotto al
    ' relativo nodo padre, che viene passato come argomento. Per ulteriori informazioni, vedere
    ' il gestore eventi OnBeforeExpand.
    Sub AddDirectories(ByVal tn As TreeNode)
        tn.Nodes.Clear()

        Dim strPath As String = tn.FullPath
        Dim diDirectory As New DirectoryInfo(strPath)
        Dim adiDirectories() As DirectoryInfo

        Try
            ' Ottenere una matrice di tutte le sottodirectory come oggetti DirectoryInfo.
            adiDirectories = diDirectory.GetDirectories()
        Catch exp As Exception
            Exit Sub
        End Try

        Dim di As DirectoryInfo
        For Each di In adiDirectories
            ' Creare un nodo figlio per ogni sottodirectory, passando il nome della
            ' directory e le immagini che verranno utilizzate dal relativo nodo.
            Dim tnDir As New TreeNode(di.Name, 1, 2)
            ' Aggiungere il nuovo nodo figlio al nodo padre.
            tn.Nodes.Add(tnDir)

            ' A questo punto, è possibile completare l'intera struttura chiamando in modo ricorsivo
            ' AddDirectories():
            '
            '   AddDirectories(tnDir)
            '
            ' Questo metodo è lento, ma è possibile provarlo.
        Next
    End Sub

    ' Questa subroutine cancella gli oggetti TreeNode esistenti e ricostruisce
    ' DirectoryTreeView, mostrando le unità logiche.
    Public Sub RefreshTree()

        ' Per garantire buone prestazioni e per evitare che il controllo TreeView sia instabile durante 
        ' l'aggiornamento di un nodo di grandi dimensioni, è consigliabile eseguire il wrapping del codice di aggiornamento nelle istruzioni BeginUpdate...
        ' EndUpdate.
        BeginUpdate()

        Nodes.Clear()

        ' Impostare le unità disco come nodi principali. 
        Dim astrDrives As String() = Directory.GetLogicalDrives()

        Dim strDrive As String
        For Each strDrive In astrDrives
            Dim tnDrive As New TreeNode(strDrive, 0, 0)
            Nodes.Add(tnDrive)
            AddDirectories(tnDrive)

            ' Impostare l'unità C come selezione predefinita.
            If strDrive = "C:\" Then
                Me.SelectedNode = tnDrive
            End If
        Next

        EndUpdate()
    End Sub
End Class
