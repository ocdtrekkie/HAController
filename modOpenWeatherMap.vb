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

    Sub GatherWeatherData(Optional ByVal Silent As Boolean = True)
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
            If Silent = False Then
                modSpeech.Say(SpeechString)
            End If

            WeatherNode = WeatherData.SelectSingleNode("/current/temperature")
            Dim dblTemperature As Double = WeatherNode.Attributes.GetNamedItem("value").Value
            SpeechString = "The current outside temperature is " + CStr(Int(dblTemperature)) + " degrees Fahrenheit"
            My.Application.Log.WriteEntry(SpeechString)
            If Silent = False Then
                modSpeech.Say(SpeechString)
            End If

            WeatherNode = WeatherData.SelectSingleNode("/current/humidity")
            Dim dblHumidity As Double = WeatherNode.Attributes.GetNamedItem("value").Value

            modDatabase.Execute("INSERT INTO ENVIRONMENT (Date, Source, Location, Temperature, Humidity, Condition) VALUES('" + Now.ToString + "', 'OWM', 'Local', " + CStr(Int(dblTemperature)) + ", " + CStr(Int(dblHumidity)) + ", '" + strWeather + "')")
        Catch NetExcep As System.Net.WebException
            My.Application.Log.WriteException(NetExcep)
        End Try
    End Sub

    Sub Load()
        If My.Settings.OpenWeatherMap_Enable = True Then
            My.Application.Log.WriteEntry("Scheduling automatic OpenWeatherMap checks")
            Dim WeatherCheckJob As IJobDetail = JobBuilder.Create(GetType(WeatherUpdateSchedule)).WithIdentity("checkjob", "modopenweathermap").Build()
            Dim WeatherCheckTrigger As ISimpleTrigger = TriggerBuilder.Create().WithIdentity("checktrigger", "modopenweathermap").WithSimpleSchedule(Sub(x) x.WithIntervalInMinutes(60).RepeatForever()).Build()

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
        Public Sub Execute(context As Quartz.IJobExecutionContext) Implements Quartz.IJob.Execute
            GatherWeatherData(True)
        End Sub
    End Class
End Module