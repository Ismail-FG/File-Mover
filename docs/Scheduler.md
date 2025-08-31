# Excel with Visual Basic code

## Scheduler
```groovy
Dim NextRunTime As Date
Dim TaskScheduled As Boolean

Sub ScheduleTask()
    TaskScheduled = True
    NextRunTime = Now + TimeValue("00:30:00")
    Application.OnTime NextRunTime, "RunTask"
    Debug.Print "Lagi jalan"
End Sub

Sub RunTask()

    PindahkanDokumen

    ' Place the function you want to run here
    Debug.Print "Function executed!"

    ' Reschedule the task
    If TaskScheduled Then ScheduleTask
End Sub

Sub StopTask()
    On Error Resume Next
    If TaskScheduled Then
        Application.OnTime NextRunTime, "RunTask", , False
        TaskScheduled = False
        Debug.Print "Scheduled task has been stopped."
    Else
        Debug.Print "No scheduled task found."
    End If
End Sub

---

