Imports System.Xml

Module modOpenWeatherMap
    Sub GatherWeatherData()
        Dim SpeechString As String

        If My.Settings.OWMAPIKey = "" Then
            My.Application.Log.WriteEntry("No OpenWeatherMap API key, asking for it")
            My.Settings.OWMAPIKey = InputBox("Enter OpenWeatherMap API Key", "OpenWeatherMap API")
        End If

        If My.Settings.CityID = "" Then
            My.Application.Log.WriteEntry("No City ID set, asking for it")
            My.Settings.CityID = InputBox("Enter City ID", "City ID")
        End If

        Dim WeatherData As XmlDocument = New XmlDocument
        Dim WeatherNode As XmlNode
        Dim WeatherRequestString As String = "http://api.openweathermap.org/data/2.5/weather?id=" + My.Settings.CityID + "&appid=" + My.Settings.OWMAPIKey + "&mode=xml&units=imperial"
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
        Dim WeatherThread As New Threading.Thread(AddressOf GatherWeatherData)
        WeatherThread.Start()
    End Sub
End Module