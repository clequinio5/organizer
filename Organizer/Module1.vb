Imports System.IO
Imports System.Text
Imports System.Xml

Module Module1
    'FONCTIONNEMENT DE L'APP
    Public MesDocPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
    Public Destdir As String = MesDocPath & "\Organizer"
    Public Agendapath, Fichierpath, Settingspath As String
    Public Jour(), Agenda() As Label
    Public Fichier() As Panel
    Public Meteo() As PictureBox
    Public Tblytpnl() As TableLayoutPanel
    Public DatasFichiers() As String
    Public Settings() As String
    Public loadsettings, loadagenda As Boolean
    Public VP_Agenda, VP_Fichier As String

    'CONSTANTE
    Public Jourdelasemaine() As String
    Public Jourdelasemainel() As String
    Public Moisdelannee() As String

    'SETTINGS
    Public Color_Fond, Color_DDay, Color_FontJour, Color_FontAgenda, Color_Bordures, Color_FontMois, Color_Mois As Color
    Public LBords As Integer
    Public CopieDeplace, StartWithWindows As Boolean
    Public JourAffiches As Integer
    Public BDD As String
    Public AppIDOpenWeather As String
    Public AppIDWunderWeather As String
    Public Sub Chargement_Settings()
        Try
            ReDim Settings(16)
            If Not File.Exists(Settingspath) Then
                Settings(0) = CStr(9)
                Settings(1) = Color2String(Color.DarkGray) 'fond
                Settings(2) = Color2String(Color.LightGray) 'bordures
                Settings(3) = Color2String(Color.Black) 'agenda
                Settings(4) = Color2String(Color.Green) 'jour
                Settings(5) = Color2String(Color.LightBlue) 'J
                Settings(6) = CStr(1) 'larg bordures
                Settings(7) = CStr(FormPrincipale.Width)
                Settings(8) = CStr(FormPrincipale.Height)
                Settings(9) = ""
                Settings(10) = Color2String(Color.White)
                Settings(11) = Color2String(Color.DimGray)
                Settings(12) = "False"
                Settings(13) = "OpenWeather"
                Settings(14) = "True"
                Settings(15) = "ca00e30d3a1bfdaca866311f96bee9d6"
                Settings(16) = "59b04c00d6a8b5bc"
                Dim str As String = ""
                For i = 0 To Settings.Length - 1
                    str &= Settings(i) & vbCrLf
                Next
                File.WriteAllText(Settingspath, str)
            End If

            Dim lecture As New StreamReader(Settingspath)
            With lecture
                For i = 0 To Settings.Length - 1
                    Settings(i) = .ReadLine
                Next
            End With
            lecture.Close()

            JourAffiches = CInt(Settings(0))
            Color_Fond = String2Color(Settings(1))
            Color_Bordures = String2Color(Settings(2))
            Color_FontAgenda = String2Color(Settings(3))
            Color_FontJour = String2Color(Settings(4))
            Color_DDay = String2Color(Settings(5))
            LBords = CInt(Settings(6))
            FormPrincipale.Size = New Size(CInt(Settings(7)), CInt(Settings(8)))
            Color_FontMois = String2Color(Settings(10))
            Color_Mois = String2Color(Settings(11))
            If Settings(12) = "True" Then
                CopieDeplace = True
            Else
                CopieDeplace = False
            End If
            BDD = Settings(13)
            If Settings(14) = "True" Then
                StartWithWindows = True
            Else
                StartWithWindows = False
            End If
            AppIDOpenWeather = Settings(15)
            AppIDWunderWeather = Settings(16)

        Catch ex As Exception
            MsgBox("Erreur de chargement des settings")
        End Try
1:
    End Sub
    Public Function String2Color(ByVal str As String) As Color
        Dim Ints() As String
        Ints = str.Split("*")
        String2Color = Color.FromArgb(CInt(Ints(0)), CInt(Ints(1)), CInt(Ints(2)))
    End Function
    Public Function Color2String(ByVal col As Color) As String
        Color2String = col.R & "*" & col.G & "*" & col.B
    End Function
    Public Sub Chargement_Jours(ByVal dat As Date)

        'LABEL JOUR
        Dim Datetemp As String
        For i = 0 To JourAffiches
            Datetemp = dat.AddDays(i - 1).ToShortDateString
            With Jour(i)
                .Tag = Datetemp
                .Text = Date2LabelText(Datetemp)
                .BackColor = Color_Fond
            End With
            If Datetemp = Now.ToShortDateString Then
                Jour(i).BackColor = Color_DDay
                Meteo(i).BackColor = Color_DDay
                Agenda(i).BackColor = Color_DDay
                Fichier(i).BackColor = Color_DDay
            End If
        Next

    End Sub
    Public Sub Chargement_Meteo(ByVal str As String)

        Try

            For i = 0 To JourAffiches
                Meteo(i).BackgroundImage = Nothing
            Next
            FormPrincipale.MeteoDetails.RemoveAll()
            Dim location As String = ""
            If str = "" Then
                location = GeoIp()
            Else
                location = str
            End If
            Dim tooltip As String
            Dim datet As Date
            Dim results As List(Of WeatherDetails)
            results = Weather.GetWeather(location, BDD)

            For i = 0 To JourAffiches
                datet = CDate(Jour(i).Tag)
                tooltip = ""
                For j = 0 To results.Count - 1
                    If Jour(i).Tag = results(j).WeatherDay Then
                        With results(j)
                            Meteo(i).BackgroundImage = My.Resources.ResourceManager.GetObject(.WeatherIcon)
                            'Meteo(i).BackgroundImage = Image.FromFile(Iconspath & "\" & .WeatherIcon)
                            Select Case BDD
                                Case "OpenWeather"
                                    tooltip = Jourdelasemainel(datet.DayOfWeek) & " " & datet.Day & " " & Moisdelannee(datet.Month) & " " & datet.Year & vbCrLf & _
                                        location & vbCrLf & vbCrLf & _
                                        "Weather: " & .Weather & vbCrLf & _
                                        "Temperature Moy: " & .Temperature & vbCrLf & _
                                        "Temperature Max: " & .MaxTemperature & vbCrLf & _
                                        "Temperature Min: " & .MinTemperature & vbCrLf & _
                                        "Humidity: " & .Humidity & vbCrLf & _
                                        "Wind Direction: " & .WindDirection & vbCrLf & _
                                        "Wind Speed: " & .WindSpeed & vbCrLf & _
                                        "Pressure: " & .Pressure & vbCrLf & _
                                        "Clouds: " & .Clouds
                                Case "WunderWeather"
                                    tooltip = Jourdelasemainel(datet.DayOfWeek) & " " & datet.Day & " " & Moisdelannee(datet.Month) & " " & datet.Year & vbCrLf & _
                                        location & vbCrLf & vbCrLf & _
                                    "Weather: " & .Weather & vbCrLf & _
                                    "Temperature Max: " & .MaxTemperature & vbCrLf & _
                                        "Temperature Min: " & .MinTemperature & vbCrLf & _
                                        "Humidity: " & .Humidity & vbCrLf & _
                                        "Wind Direction: " & .WindDirection & vbCrLf & _
                                        "Wind Speed: " & .WindSpeed & vbCrLf & _
                                        "Snow: " & .Snow
                            End Select

                        End With
                    End If
                Next
                FormPrincipale.MeteoDetails.SetToolTip(Meteo(i), tooltip)
            Next
        Catch ex As Exception

        End Try


    End Sub
    Public Sub Chargement_Fichiers()
        'FICHIER
        Try
            ReDim DatasFichiers(JourAffiches)
            Dim line As String
            Dim Split1() As String

            Dim lecture As StreamReader
            lecture = New StreamReader(Fichierpath, Encoding.Default)
            While lecture.Peek <> -1
                line = lecture.ReadLine
                Split1 = line.Split("|")
                For i = 0 To JourAffiches
                    If Split1(0) = Jour(i).Tag Then
                        Fichier(i).BackgroundImage = My.Resources.doc
                        DatasFichiers(i) = Split1(1)
                        Exit For
                    End If
                Next
            End While
            lecture.Close()
        Catch ex As Exception
            MsgBox("Une erreur s'est produite dans le chargement des fichiers")
        End Try
    End Sub
    Public Sub Chargement_Agenda()

        'AGENDA
        Try
            Dim line As String
            Dim Split1() As String
            Dim Split2() As String
            Dim Labeltext As String

            Dim lecture As StreamReader
            lecture = New StreamReader(Agendapath, Encoding.Default)
            While lecture.Peek <> -1
                line = lecture.ReadLine

                Split1 = line.Split("|")
                For i = 0 To JourAffiches
                    If Split1(0) = Jour(i).Tag Then
                        Split2 = Split1(1).Split("\")
                        Labeltext = ""
                        For j = 0 To Split2.Length - 1
                            Labeltext &= Split2(j) & vbCrLf
                        Next
                        With Agenda(i)
                            .Text = Labeltext
                        End With
                        Exit For
                    End If
                Next

            End While
            lecture.Close()
        Catch ex As Exception
            MsgBox("Une erreur s'est produite dans le chargement de l'agenda")
        End Try

    End Sub
    Public Sub Reset()
        'EFFACER
        For i = 0 To JourAffiches
            Agenda(i).Text = ""
            Agenda(i).BackColor = Color_Fond
            Fichier(i).BackgroundImage = Nothing
            Fichier(i).BackColor = Color_Fond
            Meteo(i).BackgroundImage = Nothing
            Meteo(i).BackColor = Color_Fond
            Erase DatasFichiers
        Next
    End Sub
    Public Function Date2LabelText(ByVal str As String) As String
        Date2LabelText = Jourdelasemaine(CDate(str).DayOfWeek) & " " & CInt(str.Split("/")(0)).ToString
    End Function
    Public Function Date2MoisAnnee(ByVal str As String) As String
        Dim Datetemp As Date
        Datetemp = CDate(str)
        Date2MoisAnnee = Moisdelannee(Datetemp.Month - 1) & " " & Datetemp.Year
    End Function
    Public Sub Chargement_StringJourMois()

        ReDim Jourdelasemaine(6)
        Jourdelasemaine(0) = "Dim."
        Jourdelasemaine(1) = "Lun."
        Jourdelasemaine(2) = "Mar."
        Jourdelasemaine(3) = "Mer."
        Jourdelasemaine(4) = "Jeu."
        Jourdelasemaine(5) = "Ven."
        Jourdelasemaine(6) = "Sam."

        ReDim Jourdelasemainel(6)
        Jourdelasemainel(0) = "Dimanche"
        Jourdelasemainel(1) = "Lundi"
        Jourdelasemainel(2) = "Mardi"
        Jourdelasemainel(3) = "Mercredi"
        Jourdelasemainel(4) = "Jeudi"
        Jourdelasemainel(5) = "Vendredi"
        Jourdelasemainel(6) = "Samedi"

        ReDim Moisdelannee(11)
        Moisdelannee(0) = "Janvier"
        Moisdelannee(1) = "Février"
        Moisdelannee(2) = "Mars"
        Moisdelannee(3) = "Avril"
        Moisdelannee(4) = "Mai"
        Moisdelannee(5) = "Juin"
        Moisdelannee(6) = "Juillet"
        Moisdelannee(7) = "Aout"
        Moisdelannee(8) = "Septembre"
        Moisdelannee(9) = "Octobre"
        Moisdelannee(10) = "Novembre"
        Moisdelannee(11) = "Décembre"
    End Sub
    Public Function GeoIp() As String
        Try
            Dim xmldoc As New XmlDocument
            Dim xmlnode As XmlNodeList
            xmldoc.Load("http://freegeoip.net/xml/")
            xmlnode = xmldoc.GetElementsByTagName("Response")
            For i = 0 To xmlnode.Count - 1
                xmlnode(i).ChildNodes.Item(0).InnerText.Trim()
                Select Case BDD
                    Case "OpenWeather"
                        Return xmlnode(i).ChildNodes.Item(5).InnerText.Trim() & "," & xmlnode(i).ChildNodes.Item(4).InnerText.Trim() & "," & xmlnode(i).ChildNodes.Item(2).InnerText.Trim()
                    Case "WunderWeather"
                        Return xmlnode(i).ChildNodes.Item(1).InnerText.Trim() & "/" & xmlnode(i).ChildNodes.Item(5).InnerText.Trim()
                End Select

                'Label1.Text = "IP Address : " & xmlnode(i).ChildNodes.Item(0).InnerText.Trim()
                'Label2.Text = "Country Code : " & xmlnode(i).ChildNodes.Item(1).InnerText.Trim()
                'Label3.Text = "Country Name : " & xmlnode(i).ChildNodes.Item(2).InnerText.Trim()
                'Label4.Text = "Region Code : " & xmlnode(i).ChildNodes.Item(3).InnerText.Trim()
                'Label5.Text = "Region Name : " & xmlnode(i).ChildNodes.Item(4).InnerText.Trim()
                'Label6.Text = "City : " & xmlnode(i).ChildNodes.Item(5).InnerText.Trim()
                'Label7.Text = "Zip Code : " & xmlnode(i).ChildNodes.Item(6).InnerText.Trim()
                'Label8.Text = "Latitude : " & xmlnode(i).ChildNodes.Item(7).InnerText.Trim()
                'Label9.Text = "Longitude : " & xmlnode(i).ChildNodes.Item(8).InnerText.Trim()
                'Label10.Text = "Metro Code : " & xmlnode(i).ChildNodes.Item(9).InnerText.Trim()
            Next
            Return ""
        Catch ex As Exception
            Return ""
        End Try
       
    End Function
   
End Module
