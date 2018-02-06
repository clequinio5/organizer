Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices

Public Class FormPrincipale

    Dim Editer As FormEditerAgenda
    Dim Parametre As FormSettings
    Dim Fichiers As FormFichier

    Public Const MOD_ALT As Integer = &H1
    Public Const WM_HOTKEY As Integer = &H312

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Timer1.Interval = (86400 - CInt(DateAndTime.Timer) + 1) * 1000
        Timer1.Enabled = True

        RegisterHotKey(Me.Handle, 100, MOD_ALT, Keys.F11)
        RegisterHotKey(Me.Handle, 200, MOD_ALT, Keys.F12)
        FormSettings.AddStartup()

        If Directory.Exists(Destdir) = False Then
            Directory.CreateDirectory(Destdir)
        End If
        Agendapath = Destdir & "\Agenda.txt"
        If File.Exists(Agendapath) = False Then
            File.WriteAllText(Agendapath, "")
        End If
        Fichierpath = Destdir & "\Fichier.txt"
        If File.Exists(Fichierpath) = False Then
            File.WriteAllText(Fichierpath, "")
        End If
        Settingspath = Destdir & "\Settings.txt"
        If Directory.Exists(Destdir & "\Documents") = False Then
            Directory.CreateDirectory(Destdir & "\Documents")
        End If

        Chargement_StringJourMois()

        Chargement_Settings()

        Me.Location = New Point(Screen.PrimaryScreen.Bounds.Width - Me.Size.Width, 0)

        loadsettings = False
        loadagenda = False

        BIG_CHARGEMENT()

        Label1.Text = Date2MoisAnnee(Jour(1).Tag)
    End Sub
    Public Sub BIG_CHARGEMENT()
        TableLayoutPanel1.Visible = False
        Chargement_Layout()
        Reset()
        Chargement_Jours(Now)
        Chargement_Agenda()
        Chargement_Fichiers()
        Chargement_Meteo(Settings(9))

        TableLayoutPanel1.Visible = True
    End Sub
    Public Sub Chargement_Layout()

        Label1.BackColor = Color_Mois
        Label2.BackColor = Color_Mois
        PictureBox1.BackColor = Color_Mois

        Label1.ForeColor = Color_FontMois

        TableLayoutPanel1.Controls.Clear()
        Dim nbrcolonne As Integer = JourAffiches + 1
        Dim wdth As Integer = TableLayoutPanel1.Width - nbrcolonne - 1
        With TableLayoutPanel1
            .ColumnCount = nbrcolonne
            .BackColor = Color_Bordures
        End With

        TableLayoutPanel1.ColumnStyles.Item(0).SizeType = SizeType.Percent
        TableLayoutPanel1.ColumnStyles.Item(0).Width = 100.0! / nbrcolonne

        For i = 0 To JourAffiches
            TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0! / nbrcolonne))
        Next


        ReDim Jour(JourAffiches)
        ReDim Meteo(JourAffiches)
        ReDim Tblytpnl(JourAffiches)
        ReDim Agenda(JourAffiches)
        ReDim Fichier(JourAffiches)

        For i = 0 To JourAffiches

            Jour(i) = New Label
            With Jour(i)
                .Dock = DockStyle.Fill
                .TextAlign = ContentAlignment.MiddleCenter
                .Font = New Font("Arial", 10, FontStyle.Bold)
                .ForeColor = Color_FontJour
                .Margin = New Padding(LBords, LBords, LBords, LBords)
                .AutoEllipsis = True
            End With

            Agenda(i) = New Label
            With Agenda(i)
                .Dock = DockStyle.Fill
                .Font = New Font("Arial", 9, FontStyle.Bold)
                .ForeColor = Color_FontAgenda
                .Margin = New Padding(LBords, LBords, LBords, LBords)
                .AutoEllipsis = True
                .Tag = i
                AddHandler .Click, AddressOf Agenda_Click
            End With

            Fichier(i) = New Panel
            With Fichier(i)
                .Dock = DockStyle.Fill
                .AllowDrop = True
                .BackgroundImageLayout = ImageLayout.Center
                .Margin = New Padding(LBords, LBords, 0, LBords)
                .Tag = i
                AddHandler .Click, AddressOf fichierclick
                AddHandler .DragEnter, AddressOf Fichieri_DragEnter
                AddHandler .DragDrop, AddressOf Fichieri_DragDrop
            End With

            Meteo(i) = New PictureBox
            With Meteo(i)
                .BackColor = Color_Fond
                .Dock = DockStyle.Fill
                .BackgroundImageLayout = ImageLayout.Stretch
                .Margin = New Padding(0, LBords, LBords, LBords)
                .Tag = i
            End With

            Tblytpnl(i) = New TableLayoutPanel
            With Tblytpnl(i)
                .RowCount = 1
                .ColumnCount = 2
                .Margin = New Padding(0)
                For j = 0 To 1
                    .ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
                Next
                .RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
                With .Controls
                    .Add(Fichier(i), 0, 0)
                    .Add(Meteo(i), 1, 0)
                End With
                .Dock = DockStyle.Fill
                .CellBorderStyle = TableLayoutPanelCellBorderStyle.None
            End With

            With TableLayoutPanel1.Controls
                .Add(Jour(i), i, 0)
                .Add(Tblytpnl(i), i, 1)
                .Add(Agenda(i), i, 2)
            End With

        Next
    End Sub
    Private Sub Agenda_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        VP_Agenda = sender.tag
        If loadagenda = False Then
            Editer = New FormEditerAgenda
            Editer.Show()
            loadagenda = True
        Else
            Editer.Activate()
        End If
    End Sub
    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

        If loadsettings = False Then
            Parametre = New FormSettings
            Parametre.Show()
            loadsettings = True
        Else
            Parametre.Activate()
        End If

    End Sub
    Public Sub fichierclick(ByVal sender As Panel, ByVal e As System.EventArgs)

        VP_Fichier = sender.Tag
        Fichiers = New FormFichier
        Fichiers.Show()

    End Sub
    Private Sub Fichieri_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub
    Private Sub Fichieri_DragDrop(ByVal sender As Panel, ByVal e As System.Windows.Forms.DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim MyFiles() As String
            Dim newpath, Newdata As String
            MyFiles = e.Data.GetData(DataFormats.FileDrop)

            newpath = ""
            Newdata = ""
            For i = 0 To MyFiles.Length - 1
                newpath = Destdir & "\Documents\" & Path.GetFileName(MyFiles(i))
                File.Copy(MyFiles(i), newpath, True)

                If CopieDeplace = False Then
                    File.Delete(MyFiles(i))
                End If
                Newdata &= "*" & newpath
            Next

            Dim Reference As Integer = sender.Tag
            Dim Datereference As String = Jour(Reference).Tag
            Dim Oldline As String = ""
            Dim Newline As String = ""
            Dim IndexRemove As Integer
            IndexRemove = -1

            Dim lines As List(Of String) = File.ReadAllLines(Fichierpath).ToList
            For i = 0 To lines.Count - 1
                If lines(i).StartsWith(DateReference) Then
                    IndexRemove = i
                    Oldline = lines(i)
                    Exit For
                End If
            Next

            If IndexRemove <> -1 Then
                lines.RemoveAt(IndexRemove)
                Newline = Oldline & Newdata
            Else
                Newline = Datereference & "|" & Mid(Newdata, 2)
            End If

            lines.Add(Newline)
            Fichier(Reference).BackgroundImage = My.Resources.doc
            DatasFichiers(Reference) = Newline.Split("|")(1)
            File.WriteAllLines(Fichierpath, lines.ToArray)


        End If
    End Sub
    Private Sub OuvrirToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OuvrirToolStripMenuItem.Click
        Me.Show()
        Me.Activate()
    End Sub
    Private Sub FermerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FermerToolStripMenuItem.Click
        Application.Exit()
    End Sub
    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        Me.Hide()
    End Sub
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If m.Msg = WM_HOTKEY Then
            Dim id As IntPtr = m.WParam
            Select Case (id.ToString)
                Case "100"
                    Me.Hide()
                Case "200"
                    OuvrirToolStripMenuItem.PerformClick()
            End Select
        End If
        MyBase.WndProc(m)
    End Sub
    Private Sub Form1_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        UnregisterHotKey(Me.Handle, 100)
        UnregisterHotKey(Me.Handle, 200)

    End Sub
    <DllImport("User32.dll")> _
    Public Shared Function RegisterHotKey(ByVal hwnd As IntPtr, ByVal id As Integer, ByVal fsModifiers As Integer, ByVal vk As Integer) As Integer
    End Function
    <DllImport("User32.dll")> _
    Public Shared Function UnregisterHotKey(ByVal hwnd As IntPtr, ByVal id As Integer) As Integer
    End Function
    Private Sub DateTimePicker1_CloseUp(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.CloseUp
        TableLayoutPanel1.Visible = False
        Reset()
        Chargement_Jours(DateTimePicker1.Value)
        Chargement_Agenda()
        Chargement_Fichiers()
        Chargement_Meteo(Settings(9))
        TableLayoutPanel1.Visible = True
        Label1.Text = Date2MoisAnnee(Jour(1).Tag)
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer2.Enabled = True
        Timer1.Enabled = False
        'CHARGEMENT
        BIG_CHARGEMENT()
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        'CHARGEMENT
        BIG_CHARGEMENT()
    End Sub
End Class
