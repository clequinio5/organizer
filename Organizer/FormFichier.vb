Imports System.IO
Public Class FormFichier
    Dim Reference As Integer
    Dim DateReference As String

    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Reference = VP_Fichier
        DateReference = Jour(Reference).Tag

        Try
            Dim items() As String = DatasFichiers(Reference).Split("*")
            For i = 0 To items.Length - 1
                ListView1.Items.Add(Path.GetFileName(items(i)))
                ListView1.Items(i).Tag = items(i)
            Next
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        Dim lines As List(Of String) = File.ReadAllLines(Fichierpath).ToList

        Dim IndexRemove As Integer
        IndexRemove = -1
        For i = 0 To lines.Count - 1
            If lines(i).StartsWith(DateReference) Then
                IndexRemove = i
                Exit For
            End If
        Next

        If IndexRemove <> -1 Then
            lines.RemoveAt(IndexRemove)
        End If

        Dim NewStr As String = ""
        For Each item As ListViewItem In ListView1.Items
            NewStr &= item.Tag & "*"
        Next



        If NewStr <> "" Then
            NewStr = NewStr.Substring(0, NewStr.Length - 1)
            Fichier(Reference).BackgroundImage = My.Resources.doc
            lines.Add((DateReference & "|" & NewStr))
        Else
            Fichier(Reference).BackgroundImage = Nothing
        End If

        DatasFichiers(Reference) = NewStr
        File.WriteAllLines(Fichierpath, lines.ToArray)

        Me.Close()

    End Sub

    Private Sub Listview1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListView1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub
    Private Sub Listview1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListView1.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim MyFiles() As String
            Dim i As Integer
            Dim it As ListViewItem
            Dim newpath As String

            ' Assign the files to an array.
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
            ' Loop through the array and add the files to the list.
            For i = 0 To MyFiles.Length - 1
                newpath = ""
                newpath = Application.StartupPath & "\Documents\" & Path.GetFileName(MyFiles(i))
                If File.Exists(newpath) = False Then
                    File.Copy(MyFiles(i), newpath, False)
                Else
                    Dim compt As Integer = 0
                    Dim trouve As Boolean = False
                    Dim pathe As String = ""
                    Do
                        compt += 1
                        pathe = Path.GetFileNameWithoutExtension(newpath) & " (" & compt & ")" & Path.GetExtension(newpath)
                        If File.Exists(pathe) = False Then
                            newpath = pathe
                            File.Copy(MyFiles(i), newpath, False)
                            trouve = True
                        End If
                    Loop Until trouve = True
                End If

                If CopieDeplace = False Then
                    File.Delete(MyFiles(i))
                End If

                it = New ListViewItem
                it.Tag = newpath
                it.Text = Path.GetFileName(newpath)
                ListView1.Items.Add(it)

            Next
        End If
    End Sub
    Private Sub ListView1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListView1.KeyDown
        If e.KeyCode = 46 Then
            For Each item As ListViewItem In ListView1.SelectedItems
                ListView1.Items.Remove(item)
            Next
        End If
    End Sub

    Private Sub ListView1_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        For Each item As ListViewItem In ListView1.SelectedItems
            System.Diagnostics.Process.Start(item.Tag)
        Next
    End Sub
End Class