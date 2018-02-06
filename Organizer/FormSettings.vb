Imports System.IO
Public Class FormSettings
    Dim firstload As Boolean

    Private Sub Form3_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        loadsettings = False

    End Sub
    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        firstload = True

        Labell1.BackColor = String2Color(Settings(1))
        Labell2.BackColor = String2Color(Settings(2))
        Labell3.BackColor = String2Color(Settings(3))
        Labell4.BackColor = String2Color(Settings(4))
        Labell5.BackColor = String2Color(Settings(5))
        Labell6.BackColor = String2Color(Settings(11))
        Labell7.BackColor = String2Color(Settings(10))
        NumericUpDown1.Value = Settings(6)
        NumericUpDown2.Value = Settings(0) + 1
        NumericUpDown3.Value = Settings(7)
        NumericUpDown4.Value = Settings(8)
        If Settings(9) <> "" Then
            TextBoxmeteo.Text = Settings(9)
        Else
            TextBoxmeteo.Text = GeoIp()
        End If

        If CopieDeplace = True Then
            CheckBox2.Checked = True
        Else
            CheckBox2.Checked = False
        End If
        ComboBox1.SelectedItem = Settings(13)
        If StartWithWindows = True Then
            CheckBox1.Checked = True
        Else
            CheckBox1.Checked = False
        End If

        firstload = False

    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Settings(0) = NumericUpDown2.Value - 1
        Settings(1) = Color2String(Labell1.BackColor)
        Settings(2) = Color2String(Labell2.BackColor)
        Settings(3) = Color2String(Labell3.BackColor)
        Settings(4) = Color2String(Labell4.BackColor)
        Settings(5) = Color2String(Labell5.BackColor)
        Settings(6) = NumericUpDown1.Value
        Settings(7) = NumericUpDown3.Value
        Settings(8) = NumericUpDown4.Value
        Settings(9) = TextBoxmeteo.Text
        Settings(11) = Color2String(Labell6.BackColor)
        Settings(10) = Color2String(Labell7.BackColor)
        Settings(12) = CheckBox2.Checked.ToString
        Settings(13) = ComboBox1.Text
        Settings(14) = CheckBox1.Checked.ToString

        File.WriteAllLines(Settingspath, Settings)

        Me.Close()

    End Sub
    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Labell1.Click
        If ColorDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Labell1.BackColor = ColorDialog1.Color
            Color_Fond = ColorDialog1.Color
            For i = 0 To JourAffiches
                If Jour(i).Tag <> Now.ToShortDateString Then
                    Jour(i).BackColor = Color_Fond
                    Fichier(i).BackColor = Color_Fond
                    Meteo(i).BackColor = Color_Fond
                    Agenda(i).BackColor = Color_Fond
                End If
            Next
        End If
    End Sub
    Private Sub Labell2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Labell2.Click
        If ColorDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Labell2.BackColor = ColorDialog1.Color
            Color_Bordures = ColorDialog1.Color
            FormPrincipale.TableLayoutPanel1.BackColor = Color_Bordures
        End If
    End Sub

    Private Sub Labell3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Labell3.Click
        If ColorDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Labell3.BackColor = ColorDialog1.Color
            Color_FontAgenda = ColorDialog1.Color
            For i = 0 To JourAffiches
                Agenda(i).ForeColor = Color_FontAgenda
            Next
        End If
    End Sub

    Private Sub Labell4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Labell4.Click
        If ColorDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Labell4.BackColor = ColorDialog1.Color
            Color_FontJour = ColorDialog1.Color
            For i = 0 To JourAffiches
                Jour(i).ForeColor = Color_FontJour
            Next
        End If
    End Sub
    Private Sub Labell5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Labell5.Click
        If ColorDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Labell5.BackColor = ColorDialog1.Color
            Color_DDay = ColorDialog1.Color
            For i = 0 To JourAffiches
                If Jour(i).Tag = Now.ToShortDateString Then
                    Jour(i).BackColor = Color_DDay
                    Meteo(i).BackColor = Color_DDay
                    Fichier(i).BackColor = Color_DDay
                    Agenda(i).BackColor = Color_DDay
                End If
            Next
        End If
    End Sub
    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        If firstload = False Then
            LBords = NumericUpDown1.Value
            For i = 0 To JourAffiches
                Fichier(i).Margin = New Padding(LBords, LBords, 0, LBords)
                Meteo(i).Margin = New Padding(0, LBords, LBords, LBords)
                Jour(i).Margin = New Padding(LBords)
                Agenda(i).Margin = New Padding(LBords)
            Next
        End If

    End Sub
    Private Sub NumericUpDown2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown2.ValueChanged
        If NumericUpDown2.Value <> 3 Then
            If firstload = False Then
                JourAffiches = NumericUpDown2.Value - 1
                FormPrincipale.BIG_CHARGEMENT()
            End If
        End If
    End Sub

    Private Sub NumericUpDown3_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown3.Validated
        If firstload = False Then
            FormPrincipale.Width = NumericUpDown3.Value
            FormPrincipale.Location = New Point(Screen.PrimaryScreen.WorkingArea.Right - FormPrincipale.Width, Screen.PrimaryScreen.WorkingArea.Top)
        End If
    End Sub
    Private Sub NumericUpDown4_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown4.Validated
        If firstload = False Then
            FormPrincipale.Height = NumericUpDown4.Value
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            BDD = ComboBox1.Text
            Chargement_Meteo(TextBoxmeteo.Text)
        Catch ex As Exception
            MsgBox("Une erreur s'est produite lors du chargement de la météo." & vbCrLf & "La ville n'a peut être pas été trouvée ou le serveur est peut être temporairement indisponible", MsgBoxStyle.Exclamation, "Erreur")
        End Try

    End Sub

    Private Sub TextBoxmeteo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxmeteo.Click
        TextBoxmeteo.SelectAll()
    End Sub

    Private Sub Label13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Labell6.Click
        If ColorDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Labell6.BackColor = ColorDialog1.Color
            Color_Mois = ColorDialog1.Color
            FormPrincipale.Label1.BackColor = Color_Mois
            FormPrincipale.Label2.BackColor = Color_Mois
            FormPrincipale.PictureBox1.BackColor = Color_Mois
        End If
    End Sub

    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Labell7.Click
        If ColorDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Labell7.BackColor = ColorDialog1.Color
            Color_FontMois = ColorDialog1.Color
            FormPrincipale.Label1.ForeColor = Color_FontMois
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            CopieDeplace = True
        Else
            CopieDeplace = False
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            AddStartup()
            StartWithWindows = True
        Else
            RemoveStartup()
            StartWithWindows = False
        End If
    End Sub
    Public Sub AddStartup()
        My.Computer.Registry.CurrentUser.CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Run").SetValue("Organizer", Application.ExecutablePath)
    End Sub
    Sub RemoveStartup()
        My.Computer.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run", True).DeleteValue("Organizer")
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Settings(15) = InputBox("ID OpenWeather:" & vbCrLf & "ex:ca00e30d3a1bfdaca866311f96bee9d6", "ID OpenWeather", Settings(15))
        Settings(16) = InputBox("ID WunderWeather:" & vbCrLf & "ex:59b04c00d6a8b5bc", "ID WunderWeather", Settings(16))
    End Sub
End Class