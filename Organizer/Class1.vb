Imports System.Net
Imports System.IO

Public Class Weather

    Public Shared Function GetWeather(ByVal location As String, ByVal bdd1 As String) As List(Of WeatherDetails)

        Dim url As String = ""
        Select Case bdd1
            Case "OpenWeather"
                url = String.Format("http://api.openweathermap.org/data/2.5/forecast/daily?q={0}&type=accurate&mode=xml&units=metric&cnt=14&appid={1}", location, AppIDOpenWeather)
            Case "WunderWeather"
                url = String.Format("http://api.wunderground.com/api/{0}/forecast10day/q/{1}.xml", AppIDWunderWeather, location)
        End Select

        Dim Client As New WebClient
        Client.Proxy = Nothing ' Est nécessaire pour éviter quelques bugs

        Dim Response As String = ""
        Try
            Response = Client.DownloadString(url)
        Catch ex As Exception
            MsgBox("Connexion au serveur impossible" & vbCrLf & "Vérifier votre connexion internet ou l'ID associé à la BDD Météo choisie", MsgBoxStyle.Exclamation, "Récupération de la météo impossible")
        End Try

        If Not (Response.Contains("message") And Response.Contains("cod")) Then
            Dim xEl = XElement.Load(New StringReader(Response))
            Return GetWeatherInfo(xEl, bdd1)
        Else
            Return New List(Of WeatherDetails)
        End If


    End Function

    Private Shared Function GetWeatherInfo(ByVal xEl As XElement, ByVal bdd2 As String) As List(Of WeatherDetails)

        Select Case bdd2
            Case "OpenWeather"
                Dim w = xEl...<time>.Select(Function(el) New WeatherDetails With {.Humidity = el.<humidity>.@value & "%",
                                                                        .MaxTemperature = el.<temperature>.@max & " °C",
                                                                        .MinTemperature = el.<temperature>.@min & " °C",
                                                                        .Temperature = el.<temperature>.@day & " °C",
                                                                        .Weather = el.<symbol>.@name,
                                                                        .WeatherDay = DayOfTheWeek(el),
                                                                        .WeatherIcon = WeatherIconPath(el),
                                                                        .WindDirection = el.<windDirection>.@name,
                                                                        .WindSpeed = el.<windSpeed>.@mps & " mps",
                                                                         .Clouds = el.<clouds>.@value & " (" & el.<clouds>.@all & "%)",
                                                                         .Pressure = el.<pressure>.@value & " hPa"})
                Return (w.ToList)
            Case "WunderWeather"
                Dim xEl1 As XElement
                xEl1 = xEl...<simpleforecast>(0)
                Dim w = xEl1...<forecastday>.Select(Function(el) New WeatherDetails With {.WeatherDay = DateT(el.<date>.<day>.Value & "/" & el.<date>.<month>.Value & "/" & el.<date>.<year>.Value),
                                                                                         .Humidity = el.<avehumidity>.Value & " %",
                                                                                          .MaxTemperature = el.<high>.<celsius>.Value & " °C",
                                                                                          .MinTemperature = el.<low>.<celsius>.Value & " °C",
                                                                                          .Weather = el.<conditions>.Value,
                                                                                          .WeatherIcon = el.<icon>.Value,
                                                                                          .WindSpeed = el.<avewind>.<kph>.Value & " km/h",
                                                                                          .WindDirection = el.<avewind>.<dir>.Value,
                                                                                          .Snow = el.<snow_allday>.<cm>.Value})
                Return (w.ToList)
            Case Else
                Return New List(Of WeatherDetails)
        End Select

    End Function
    Private Shared Function DateT(ByVal str As String) As String
        DateT = CDate(str).ToShortDateString
    End Function
    Private Shared Function DayOfTheWeek(ByVal el As XElement) As String
        Dim tab() As String
        tab = Split(el.@day, "-")
        Dim dW = tab(2) & "/" & tab(1) & "/" & tab(0)
        Return dW
    End Function

    Private Shared Function WeatherIconPath(ByVal el As XElement) As String
        Dim symbolVar = el.<symbol>.@var
        Dim symbolNumber = el.<symbol>.@number
        Dim dayOrNight = symbolVar.ElementAt(2) ' d or n
        Return (String.Format("_{0}{1}", symbolNumber, dayOrNight))
    End Function
End Class
