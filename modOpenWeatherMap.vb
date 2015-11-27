Imports System.Xml

Module modOpenWeatherMap
    Sub Disable()
        My.Application.Log.WriteEntry("Unloading OpenWeatherMap module")
        Unload()
        My.Settings.OpenWeatherMap_Enable = False
        My.Application.Log.WriteEntry("OpenWeatherMap module is disabled")
    End Sub

    Sub Enable()
        My.Settings.OpenWeatherMap_Enable = True
        My.Application.Log.WriteEntry("OpenWeatherMap module is enabled")
        My.Application.Log.WriteEntry("Loading OpenWeatherMap module")
        Load()
    End Sub

    Sub GatherWeatherData()
        Dim SpeechString As String

        If My.Settings.OpenWeatherMap_APIKey = "" Then
            My.Application.Log.WriteEntry("No OpenWeatherMap API key, asking for it")
            My.Settings.OpenWeatherMap_APIKey = InputBox("Enter OpenWeatherMap API Key. You can get an API key at http://openweathermap.org/appid by signing up for a free account.", "OpenWeatherMap API")
        End If

        If My.Settings.OpenWeatherMap_CityID = "" Then
            My.Application.Log.WriteEntry("No City ID set, asking for it")
            My.Settings.OpenWeatherMap_CityID = InputBox("Enter City ID. You can look up your city at http://openweathermap.org/city and find the ID in the URL of the resulting page.", "City ID")
        End If

        Dim WeatherData As XmlDocument = New XmlDocument
        Dim WeatherNode As XmlNode
        Dim WeatherRequestString As String = "http://api.openweathermap.org/data/2.5/weather?id=" + My.Settings.OpenWeatherMap_CityID + "&appid=" + My.Settings.OpenWeatherMap_APIKey + "&mode=xml&units=imperial"
        My.Application.Log.WriteEntry("Requesting OpenWeatherMap data")
        Try
            WeatherData.Load(WeatherRequestString)

            WeatherNode = WeatherData.SelectSingleNode("/current/weather")
            Dim strWeather As String = WeatherNode.Attributes.GetNamedItem("value").Value
            SpeechString = "The current outside weather condition is " + strWeather
            My.Application.Log.WriteEntry(SpeechString)
            modSpeech.Say(SpeechString)

            WeatherNode = WeatherData.SelectSingleNode("/current/temperature")
            Dim dblTemperature As Double = WeatherNode.Attributes.GetNamedItem("value").Value
            SpeechString = "The current outside temperature is " + CStr(Int(dblTemperature)) + " degrees Fahrenheit"
            My.Application.Log.WriteEntry(SpeechString)
            modSpeech.Say(SpeechString)
        Catch NetExcep As System.Net.WebException
            My.Application.Log.WriteException(NetExcep)
        End Try
    End Sub

    Sub Load()
        If My.Settings.OpenWeatherMap_Enable = True Then
            Dim WeatherThread As New Threading.Thread(AddressOf GatherWeatherData)
            WeatherThread.Start()
        Else
            My.Application.Log.WriteEntry("OpenWeatherMap module is disabled, module not loaded")
        End If
    End Sub

    Sub Unload()

    End Sub
End Module