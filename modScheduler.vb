Imports Quartz
Imports Quartz.Impl

' modScheduler cannot be disabled, Disable and Enable methods will be skipped

Module modScheduler
    ' construct a scheduler factory
    Dim schedFact As ISchedulerFactory = New StdSchedulerFactory()
    Public sched As IScheduler

    Public Class ScheduleJob : Implements IJob
        Public Async Function Execute(context As Quartz.IJobExecutionContext) As Task Implements Quartz.IJob.Execute
            Dim dataMap As JobDataMap = context.JobDetail.JobDataMap
            Dim response As String = ""
            'write your schedule job
            modSpeech.Say("The time is now " & dataMap.GetString("intHour") & ":" & dataMap.GetString("intMinute"))
            modInsteon.InsteonAlarmControl(My.Settings.Insteon_AlarmAddr, response, "on", 4)
            Threading.Thread.Sleep(2000)
            modInsteon.InsteonLightControl(My.Settings.Insteon_WakeLightAddr, response, "on")
            ' Tested: Four seconds is EXACTLY enough to make you want to rip the BuzzLinc out of the wall, but too short to let you do so.
            Await Task.Delay(1)
        End Function
    End Class

    Async Sub Load()
        My.Application.Log.WriteEntry("Getting scheduler")
        sched = Await schedFact.GetScheduler()
        My.Application.Log.WriteEntry("Starting scheduler")
        Await sched.Start()

        ' TEMP WAKE ALARM CODE
        ' construct job info
        Dim intHour As Integer = 7
        Dim intMinute As Integer = 5
        Dim job As IJobDetail = JobBuilder.Create(GetType(ScheduleJob)).UsingJobData("intHour", CStr(intHour)).UsingJobData("intMinute", CStr(intMinute)).WithIdentity("wakejob", "modscheduler").Build()
        ' construct trigger

        Dim tempTrigger As ITrigger = TriggerBuilder.Create().WithIdentity("waketrigger", "modscheduler").StartNow().WithSchedule(CronScheduleBuilder.DailyAtHourAndMinute(intHour, intMinute)).Build()
        Dim trigger As ICronTrigger = DirectCast(tempTrigger, ICronTrigger)

        Await sched.ScheduleJob(job, trigger)
    End Sub

    Async Sub Unload()
        My.Application.Log.WriteEntry("Shutting down scheduler")
        Await sched.Shutdown()
    End Sub
End Module
