set ws = createObject("WScript.Shell")
set fso = createObject("Scripting.FileSystemObject")
set wshShell = createObject("WScript.Shell")

path = fso.getFile(WScript.ScriptFullName).ParentFolder & "\"

input = msgBox("1. Reminder on every boot: Abort" & vbCrLf & "2. Specify date and time: Retry" & vbCrLf & "3. Daily reminder at specified time: Ignore", 2 + 64, "Reminder")
content = inputBox("Enter the reminder content", "Reminder Content")

' Generate a unique and valid filename for the reminder file
rawFileName = CStr(Date) & CStr(Time)
validFileName = replace(replace(replace(replace(rawFileName, "/", ""), ":", ""), "-", ""), " ", "")
validFileName = left(validFileName,len(validFileName)-2)
fileName = path & validFileName & ".vbs"

fso.createTextFile(fileName)

' Write the reminder file content
set edit = fso.OpenTextFile(fileName, 2, True)

' Speak the reminder
edit.writeLine "set sapi = createObject(""SAPI.SpVoice"")"
edit.writeLine "sapi.speak(""" & content & """)"

edit.writeLine "msgBox """ & content & """, 4096 + 64, ""Reminder"""
edit.Close

Select Case input

    Case 3 ' Reminder on every boot
        targetFolder = wshShell.specialFolders("Startup") & "\"
        fso.moveFile fileName, targetFolder
        ' ws.Run "schtasks /create /sc ONSTART /tn RemindTask /tr " & fileName

    Case 4 ' Specify date and time
        targetTime = inputBox("Specify the time for the reminder (in HH:MM format e.g. 01:23 (1:23 is not allowed), 24-hour clock)",, ":")
        targetDate = inputBox("Specify the date (in DD/MM/YYYY format (1/1/2024 is not allowed. Please use the format 01/01/2024 instead) )",, "//")
        ws.Run "schtasks /create /sc ONCE /tn " & validFileName & " /st " & targetTime & " /sd " & targetDate & " /tr " & fileName

    Case 5 ' Daily reminder at specified time
        targetTime = inputBox("Specify the time for the daily reminder (in HH:MM format e.g. 01:23 (1:23 is not allowed), 24-hour clock)",, ":")
        ws.Run "schtasks /create /sc DAILY /tn " & validFileName & " /st " & targetTime & " /tr " & fileName

End Select
