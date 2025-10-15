' Copyright (C) Microsoft Corporation. Tutti i diritti riservati.
Imports System.Diagnostics ' Per Process.Start
Imports System.IO

Class FileListView
    Inherits ListView
    Private strDirectory As String

    Public Sub New()
        ' Impostare l'enumerazione View predefinita su Details.
        Me.View = System.Windows.Forms.View.Details

        ' Ottenere le immagini da utilizzare come icone per alcuni dei tipi di file comuni.
        Dim img As New ImageList()
        With img.Images
            .Add(My.Resources.DOC)
            .Add(My.Resources.EXE)
        End With

        ' I controlli SmallImageList e LargImageList utilizzano lo stesso set di
        ' immagini.
        Me.SmallImageList = img
        Me.LargeImageList = img

        ' Creare le colonne.
        With Columns
            .Add("Nome", 110, HorizontalAlignment.Left)
            .Add("Dimensione (in Byte)", 120, HorizontalAlignment.Right)
            .Add("Modificato il", 120, HorizontalAlignment.Left)
            .Add("Attributi", 50, HorizontalAlignment.Left)
        End With
    End Sub

    ''' <summary>
    ''' Esegue l'override del gestore eventi OnItemActivate della classe di base. Estende l'implementazione della
    ''' classe di base per consentire l'esecuzione di qualsiasi file con estensione exe o di qualsiasi file a cui sia associato un eseguibile.
    ''' </summary>
    Protected Overrides Sub OnItemActivate(ByVal ea As EventArgs)
        MyBase.OnItemActivate(ea)

        Dim lvi As ListViewItem
        For Each lvi In SelectedItems
            Process.Start(Path.Combine(strDirectory, lvi.Text))
        Next lvi
    End Sub

    ''' <summary>
    ''' Questa subroutine è utilizzata per visualizzare un elenco di tutti i file nella directory
    ''' correntemente selezionata dall'utente nel controllo TreeView personalizzato.
    ''' </summary>
    Public Sub ShowFiles(ByVal strDirectory As String)
        ' Salvare il nome della directory come campo.
        Me.strDirectory = strDirectory

        Items.Clear()

        Dim diDirectories As New DirectoryInfo(strDirectory)
        Dim afiFiles() As FileInfo

        Try
            ' Chiamare il comodo metodo GetFiles per ottenere una matrice di tutti i file
            ' nella directory.
            afiFiles = diDirectories.GetFiles()
        Catch
            Return
        End Try

        Dim fi As FileInfo
        For Each fi In afiFiles
            ' Creare ListViewItem.
            Dim lvi As New ListViewItem(fi.Name)

            ' Assegnare ImageIndex in base all'estensione del nome di file.
            Select Case Path.GetExtension(fi.Name).ToUpper()
                Case ".EXE"
                    lvi.ImageIndex = 1
                Case Else
                    lvi.ImageIndex = 0
            End Select

            ' Aggiungere le voci secondarie relative alla lunghezza e all'ultima modifica del file.
            lvi.SubItems.Add(fi.Length.ToString("N0"))
            lvi.SubItems.Add(fi.LastWriteTime.ToString())

            ' Aggiungere la voce secondaria relativa all'attributo.
            Dim strAttr As String = ""

            If (fi.Attributes And FileAttributes.Archive) <> 0 Then
                strAttr += "A"
            End If
            If (fi.Attributes And FileAttributes.Hidden) <> 0 Then
                strAttr += "H"
            End If
            If (fi.Attributes And FileAttributes.ReadOnly) <> 0 Then
                strAttr += "R"
            End If
            If (fi.Attributes And FileAttributes.System) <> 0 Then
                strAttr += "S"
            End If
            lvi.SubItems.Add(strAttr)

            ' Aggiungere la classe ListViewItem completa a FileListView.
            Items.Add(lvi)
        Next fi
    End Sub

End Class