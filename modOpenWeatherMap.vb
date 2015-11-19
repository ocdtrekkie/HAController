Imports System.Xml

Module modOpenWeatherMap
    Sub GatherWeatherData()
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
        Dim WeatherRequestString As String = "http://api.openweathermap.org/data/2.5/weather?id=" + My.Settings.CityID + "&appid=" + My.Settings.OWMAPIKey + "&mode=xml"
        WeatherData.Load(WeatherRequestString)

        WeatherNode = WeatherData.SelectSingleNode("/current/weather")
        Dim strWeather As String = WeatherNode.Attributes.GetNamedItem("value").Value
        My.Application.Log.WriteEntry("Current outside weather condition is " + strWeather)

        WeatherNode = WeatherData.SelectSingleNode("/current/temperature")
        Dim dblTemperature As Double = WeatherNode.Attributes.GetNamedItem("value").Value
        My.Application.Log.WriteEntry("Outside temperature in Kelvin is " + dblTemperature.ToString)
    End Sub
End Module