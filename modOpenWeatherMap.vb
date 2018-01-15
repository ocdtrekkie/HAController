Imports Quartz
Imports Quartz.Impl
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
                Dim dblTemperature As Double = WeatherNode.Attributes.GetNamedItem("value").Value
                SpeechString = SpeechString & ". The current outside temperature is " + CStr(Int(dblTemperature)) & " degrees Fahrenheit."

                WeatherNode = WeatherData.SelectSingleNode("/current/humidity")
                Dim dblHumidity As Double = WeatherNode.Attributes.GetNamedItem("value").Value

                WeatherNode = WeatherData.SelectSingleNode("/current/city")
                Dim strCityName As String = WeatherNode.Attributes.GetNamedItem("name").Value

                WeatherNode = WeatherData.SelectSingleNode("/current/lastupdate")
                Dim dteLastUpdate As DateTime = WeatherNode.Attributes.GetNamedItem("value").Value

                My.Application.Log.WriteEntry(SpeechString)

                Dim result As Integer = New Integer

                modDatabase.ExecuteScalar("SELECT Id FROM ENVIRONMENT WHERE Date = """ + dteLastUpdate.ToString("u") + """ AND Source = ""OWM""", result)
                If result <> 0 Then
                    My.Application.Log.WriteEntry("Not entering duplicate OpenWeatherMap data")
                Else
                    modDatabase.Execute("INSERT INTO ENVIRONMENT (Date, Source, Location, Temperature, Humidity, Condition) VALUES('" + dteLastUpdate.ToString("u") + "', 'OWM', '" + strCityName + "', " + CStr(Int(dblTemperature)) + ", " + CStr(Int(dblHumidity)) + ", '" + strWeather + "')")
                End If

                My.Settings.Global_LastKnownOutsideTemp = Int(dblTemperature)
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

    Sub Load()
        If My.Settings.OpenWeatherMap_Enable = True Then
            If My.Settings.OpenWeatherMap_APIKey = "" Then
                My.Application.Log.WriteEntry("No OpenWeatherMap API key, asking for it")
                My.Settings.OpenWeatherMap_APIKey = InputBox("Enter OpenWeatherMap API Key. You can get an API key at http://openweathermap.org/appid by signing up for a free account.", "OpenWeatherMap API")
            End If

            My.Application.Log.WriteEntry("Scheduling automatic OpenWeatherMap checks")
            Dim WeatherCheckJob As IJobDetail = JobBuilder.Create(GetType(WeatherUpdateSchedule)).WithIdentity("checkjob", "modopenweathermap").Build()
            Dim WeatherCheckTrigger As ISimpleTrigger = TriggerBuilder.Create().WithIdentity("checktrigger", "modopenweathermap").WithSimpleSchedule(Sub(x) x.WithIntervalInMinutes(10).RepeatForever()).Build()

            Try
                modScheduler.sched.ScheduleJob(WeatherCheckJob, WeatherCheckTrigger)
            Catch QzExcep As Quartz.ObjectAlreadyExistsException
                My.Application.Log.WriteException(QzExcep)
            End Try
        Else
            My.Application.Log.WriteEntry("OpenWeatherMap module is disabled, module not loaded")
        End If
    End Sub

    Sub Unload()

    End Sub

    Public Class WeatherUpdateSchedule : Implements IJob
        Public Async Function Execute(context As Quartz.IJobExecutionContext) As Task Implements Quartz.IJob.Execute
            GatherWeatherData()
            Await Task.Delay(1)
        End Function
    End Class
End Module