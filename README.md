Option Explicit

Sub InitializeSchedule()
    ' Initialize the schedule with empty data for two weeks
    Dim schedule As Object
    Set schedule = CreateObject("Scripting.Dictionary")
    
    Dim startDate As Date
    startDate = Date
    
    Dim i As Integer
    For i = 1 To 14
        Dim dateString As String
        dateString = Format(startDate + i - 1, "yyyy-mm-dd")
        schedule(dateString) = CreateObject("Scripting.Dictionary")
    Next i
    
    ' Save the initialized schedule to a JSON file
    SaveScheduleToJson schedule, ThisWorkbook.Path & "\schedule.json"
    
    MsgBox "Schedule initialized successfully!"
End Sub

Sub AddTaskAndPeople(dateString As String, task As String, people() As String)
    ' Add a task and people to a specific date in the schedule
    Dim schedule As Object
    Set schedule = LoadScheduleFromJson(ThisWorkbook.Path & "\schedule.json")
    
    Dim taskData As Object
    Set taskData = CreateObject("Scripting.Dictionary")
    taskData("task") = task
    taskData("people") = people
    
    schedule(dateString)(task) = taskData
    
    ' Save the updated schedule to a JSON file
    SaveScheduleToJson schedule, ThisWorkbook.Path & "\schedule.json"
    
    MsgBox "Task and people added successfully!"
End Sub

Function LoadScheduleFromJson(filePath As String) As Object
    ' Load the schedule from a JSON file
    Dim jsonContent As String
    Dim fileNum As Integer
    
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    jsonContent = Input(LOF(fileNum), fileNum)
    Close #fileNum
    
    Dim jsonObject As Object
    Set jsonObject = JsonConverter.ParseJson(jsonContent)
    
    Set LoadScheduleFromJson = jsonObject
End Function

Sub SaveScheduleToJson(schedule As Object, filePath As String)
    ' Save the schedule to a JSON file
    Dim jsonContent As String
    jsonContent = JsonConverter.ConvertToJson(schedule)
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Output As #fileNum
    Print #fileNum, jsonContent
    Close #fileNum
End Sub
