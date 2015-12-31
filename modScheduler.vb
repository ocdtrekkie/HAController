Imports Quartz
Imports Quartz.Impl

' modScheduler cannot be disabled, Disable and Enable methods will be skipped

Module modScheduler
    ' construct a scheduler factory
    Dim schedFact As ISchedulerFactory = New StdSchedulerFactory()

    ' get a scheduler
    Public sched As IScheduler = schedFact.GetScheduler()

    Public Class ScheduleJob : Implements IJob
        Public Sub Execute(context As Quartz.IJobExecutionContext) Implements Quartz.IJob.Execute
            Dim dataMap As JobDataMap = context.JobDetail.JobDataMap
            Dim response As String = ""
            'write your schedule job
            modSpeech.Say("The time is now " & dataMap.GetString("intHour") & ":" & dataMap.GetString("intMinute"))
            modInsteon.InsteonAlarmControl(My.Settings.Insteon_AlarmAddr, response, "On", 2)
        End Sub
    End Class

    Sub Load()
        My.Application.Log.WriteEntry("Starting scheduler")
        sched.Start()

        ' TEMP WAKE ALARM CODE
        ' construct job info
        Dim intHour As Integer = 6
        Dim intMinute As Integer = 50
        Dim job As IJobDetail = JobBuilder.Create(GetType(ScheduleJob)).UsingJobData("intHour", CStr(intHour)).UsingJobData("intMinute", CStr(intMinute)).WithIdentity("wakejob", "modscheduler").Build()
        ' construct trigger

        Dim tempTrigger As ITrigger = TriggerBuilder.Create().WithIdentity("waketrigger", "modscheduler").StartNow().WithSchedule(CronScheduleBuilder.DailyAtHourAndMinute(intHour, intMinute)).Build()
        Dim trigger As ICronTrigger = DirectCast(tempTrigger, ICronTrigger)

        sched.ScheduleJob(job, trigger)
    End Sub

    Sub Unload()
        My.Application.Log.WriteEntry("Shutting down scheduler")
        sched.Shutdown()
    End Sub
End Module
