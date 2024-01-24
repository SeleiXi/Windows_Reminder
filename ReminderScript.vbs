Set ws = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")

path = fso.GetFile(WScript.ScriptFullName).ParentFolder & "\"

input = MsgBox("1. Reminder on every boot: Abort" & vbCrLf & "2. Specify date and time: Retry" & vbCrLf & "3. Daily reminder at specified time: Ignore", 2 + 64, "Reminder")
content = InputBox("Enter the reminder content", "Reminder Content")

' Generate a unique and valid filename for the reminder file
rawFileName = CStr(Date) & CStr(Time)
validFileName = Replace(Replace(Replace(Replace(rawFileName, "/", ""), ":", ""), "-", ""), " ", "")
validFileName = Left(validFileName,len(validFileName)-2)
fileName = path & validFileName & ".vbs"

fso.CreateTextFile(fileName)

' Write the reminder file content
Set edit = fso.OpenTextFile(fileName, 2, True)

' Speak the reminder
edit.WriteLine "Set sapi = CreateObject(""SAPI.SpVoice"")"
edit.WriteLine "sapi.Speak(""" & content & """)"

edit.WriteLine "MsgBox """ & content & """, 4096 + 64, ""Reminder"""
edit.Close

Select Case input

    Case 3 ' Reminder on every boot
        targetFolder = wshShell.SpecialFolders("Startup") & "\"
        fso.MoveFile fileName, targetFolder
        ' ws.Run "schtasks /create /sc ONSTART /tn RemindTask /tr " & fileName

    Case 4 ' Specify date and time
        targetTime = InputBox("Specify the time for the reminder (in HH:MM format, 24-hour clock)",, ":")
        targetDate = InputBox("Specify the date (in DD/MM/YYYY format)",, "//")
        ws.Run "schtasks /create /sc ONCE /tn " & validFileName & " /st " & targetTime & " /sd " & targetDate & " /tr " & fileName

    Case 5 ' Daily reminder at specified time
        targetTime = InputBox("Specify the time for the daily reminder (in HH:MM format, 24-hour clock)",, ":")
        ws.Run "schtasks /create /sc DAILY /tn " & validFileName & " /st " & targetTime & " /tr " & fileName

End Select
