' Copyright (C) Microsoft Corporation. Tutti i diritti riservati.
Imports System.Windows.Forms

' Questa classe è necessaria in quanto
' System.Windows.Forms.Design.FolderNameEditor.FolderBrowser è protetto e pertanto
' non è accessibile in questo contesto. Derivando una classe pubblica diventa possibile
' utilizzare la finestra di dialogo nel codice.
Public Class FolderBrowser
    Inherits System.Windows.Forms.Design.FolderNameEditor

    Public Shared Function ShowDialog() As String
        Dim folders As New FolderBrowser()
        folders.Description = "Seleziona la Directory che vuoi che Jordan analizzi..."
        folders.Style = Design.FolderNameEditor.FolderBrowserStyles.RestrictToFilesystem
        folders.ShowDialog()

        Return folders.DirectoryPath

    End Function
End Class
