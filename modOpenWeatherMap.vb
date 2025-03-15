Imports System.Xml

Module modOpenWeatherMap
    Dim tmrOWMCheckTimer As System.Timers.Timer

    Function Disable() As String
        Unload()
        My.Settings.OpenWeatherMap_Enable = False
        My.Application.Log.WriteEntry("OpenWeatherMap module is disabled")
        Return "OpenWeatherMap module disabled"
    End Function

    Function Enable() As String
        My.Settings.OpenWeatherMap_Enable = True
        My.Application.Log.WriteEntry("OpenWeatherMap module is enabled")
        Load()
        Return "OpenWeatherMap module enabled"
    End Function

    Function GatherWeatherData() As String
        If My.Settings.OpenWeatherMap_Enable = True Then
            Dim SpeechString As String
            Dim WeatherRequestString As String

            If My.Settings.GPS_Enable = False AndAlso My.Settings.OpenWeatherMap_CityID = "" Then
                My.Application.Log.WriteEntry("No City ID set, asking for it")
                My.Settings.OpenWeatherMap_CityID = InputBox("Enter City ID. You can look up your city at http://openweathermap.org/city and find the ID in the URL of the resulting page.", "City ID")
            End If

            Dim WeatherData As XmlDocument = New XmlDocument
            Dim WeatherNode As XmlNode
            If My.Settings.GPS_Enable = True AndAlso (modGPS.CurrentLatitude <> 0 OrElse modGPS.CurrentLongitude <> 0) Then
                WeatherRequestString = "http://api.openweathermap.org/data/2.5/weather?lat=" + CStr(modGPS.CurrentLatitude) + "&lon=" + CStr(modGPS.CurrentLongitude) + "&appid=" + My.Settings.OpenWeatherMap_APIKey + "&mode=xml&units=imperial"
            Else
                WeatherRequestString = "http://api.openweathermap.org/data/2.5/weather?id=" + My.Settings.OpenWeatherMap_CityID + "&appid=" + My.Settings.OpenWeatherMap_APIKey + "&mode=xml&units=imperial"
            End If

            My.Application.Log.WriteEntry("Requesting OpenWeatherMap data")
            Try
                WeatherData.Load(WeatherRequestString)

                WeatherNode = WeatherData.SelectSingleNode("/current/weather")
                Dim strWeather As String = WeatherNode.Attributes.GetNamedItem("value").Value
                SpeechString = "The current outside weather condition is " + strWeather

                WeatherNode = WeatherData.SelectSingleNode("/current/temperature")
                Dim dblTemperature As Double = CDbl(WeatherNode.Attributes.GetNamedItem("value").Value)
                SpeechString = SpeechString & ". The current outside temperature is " + CStr(Int(dblTemperature)) & " degrees Fahrenheit."

                WeatherNode = WeatherData.SelectSingleNode("/current/humidity")
                Dim dblHumidity As Double = CDbl(WeatherNode.Attributes.GetNamedItem("value").Value)

                WeatherNode = WeatherData.SelectSingleNode("/current/city")
                Dim strCityName As String = WeatherNode.Attributes.GetNamedItem("name").Value

                WeatherNode = WeatherData.SelectSingleNode("/current/lastupdate")
                Dim dteLastUpdate As DateTime = CDate(WeatherNode.Attributes.GetNamedItem("value").Value)

                My.Application.Log.WriteEntry(SpeechString)

                Dim result As Integer = New Integer

                modDatabase.ExecuteScalar("SELECT Id FROM ENVIRONMENT WHERE Date = """ + dteLastUpdate.ToString("u") + """ AND Source = ""OWM""", result)
                If result <> 0 Then
                    My.Application.Log.WriteEntry("Not entering duplicate OpenWeatherMap data")
                Else
                    modDatabase.Execute("INSERT INTO ENVIRONMENT (Date, Source, Location, Temperature, Humidity, Condition) VALUES('" + dteLastUpdate.ToString("u") + "', 'OWM', '" + strCityName + "', " + CStr(Int(dblTemperature)) + ", " + CStr(Int(dblHumidity)) + ", '" + strWeather + "')")
                End If

                My.Settings.Global_LastKnownOutsideTemp = CInt(Int(dblTemperature))
                My.Settings.Global_LastKnownOutsideCondition = strWeather

                Return SpeechString
            Catch NetExcep As System.Net.WebException
                My.Application.Log.WriteException(NetExcep)

                Return "Error getting weather data"
            Catch XmlExcep As System.Xml.XmlException
                My.Application.Log.WriteException(XmlExcep)

                Return "Error parsing weather data"
            End Try
        Else
            My.Application.Log.WriteEntry("OpenWeatherMap module is disabled")
            Return "Disabled"
        End If
    End Function

    Function Load() As String
        If My.Settings.OpenWeatherMap_Enable = True Then
            My.Application.Log.WriteEntry("Loading OpenWeatherMap module")
            If My.Settings.OpenWeatherMap_APIKey = "" Then
                My.Application.Log.WriteEntry("No OpenWeatherMap API key, asking for it")
                My.Settings.OpenWeatherMap_APIKey = InputBox("Enter OpenWeatherMap API Key. You can get an API key at http://openweathermap.org/appid by signing up for a free account.", "OpenWeatherMap API")
            End If

            My.Application.Log.WriteEntry("Scheduling automatic OpenWeatherMap checks")
            tmrOWMCheckTimer = New System.Timers.Timer
            AddHandler tmrOWMCheckTimer.Elapsed, AddressOf GatherWeatherData
            tmrOWMCheckTimer.Interval = 600000 ' 10min
            tmrOWMCheckTimer.Enabled = True

            Dim InitialOWMCheck As New Threading.Thread(AddressOf GatherWeatherData)
            InitialOWMCheck.Start()
            Return "OpenWeatherMap module loaded"
        Else
            My.Application.Log.WriteEntry("OpenWeatherMap module is disabled, module not loaded")
            Return "OpenWeatherMap module is disabled, module not loaded"
        End If
    End Function

    Function Unload() As String
        My.Application.Log.WriteEntry("Unloading OpenWeatherMap module")
        If tmrOWMCheckTimer IsNot Nothing Then
            tmrOWMCheckTimer.Enabled = False
        End If
        Return "OpenWeatherMap module unloaded"
    End Function
End Module