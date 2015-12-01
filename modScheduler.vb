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
            Dim response As String = ""
            'write your schedule job
            modSpeech.Say("The time is now 6:44")
            modInsteon.InsteonAlarmControl(My.Settings.Insteon_AlarmAddr, response, "On", 10)
        End Sub
    End Class

    Sub Load()
        My.Application.Log.WriteEntry("Starting scheduler")
        sched.Start()

        ' construct job info
        Dim job As IJobDetail = JobBuilder.Create(GetType(ScheduleJob)).WithIdentity("job2", "group2").Build()
        ' construct trigger

        Dim tempTrigger As ITrigger = TriggerBuilder.Create().WithIdentity("Trigger1").StartNow().WithSchedule(CronScheduleBuilder.DailyAtHourAndMinute(18, 44)).Build()
        Dim trigger As ICronTrigger = DirectCast(tempTrigger, ICronTrigger)

        sched.ScheduleJob(job, trigger)
    End Sub

    Sub Unload()
        My.Application.Log.WriteEntry("Shutting down scheduler")
        sched.Shutdown()
    End Sub
End Module
