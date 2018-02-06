Imports System.IO
Imports System.Text
Public Class FormEditerAgenda
    Dim Reference As Integer
    Dim DateReference As String
    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Reference = VP_Agenda
        RichTextBox1.Text = Agenda(Reference).Text
        DateReference = Jour(Reference).Tag

    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'CHARGEMENT
        Dim lines As List(Of String) = File.ReadAllLines(Agendapath).ToList

        'INDEX A SUPPRIMER
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

        If RichTextBox1.Text <> "" Then
            lines.Add(DateReference & "|" & CompileVbcrlf(RichTextBox1.Lines))
        End If

        File.WriteAllLines(Agendapath, lines.ToArray)

        Agenda(Reference).Text = RichTextBox1.Text

        Me.Close()

    End Sub
    Private Function CompileVbcrlf(ByVal str() As String) As String
        CompileVbcrlf = String.Join("\", str)
    End Function

    Private Sub Form2_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        loadagenda = False
    End Sub
End Class