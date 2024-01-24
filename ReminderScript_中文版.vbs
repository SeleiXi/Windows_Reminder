'如出現報錯，將文件以ANSI編碼保存，詳細見README.MD文件

set ws = wscript.createObject("Wscript.Shell")
set fso = createObject("scripting.filesystemobject")
Set wshShell = createObject("WScript.Shell")

path = fso.getFile(wscript.scriptFullName).parentFolder + "\"

input = msgBox("1.每次開機提醒：Abort" & vbCrLf & "2.指定日期及時間：Retry " & vbCrLf & "3.每天的指定時間提醒：Ignore",2+64,"定時提醒器")
content = inputBox("請輸入要提醒的事項","提醒內容")

'提醒文件起名
rawFileName = Cstr(date) & Cstr(time)
validFileName = Replace(Replace(Replace(Replace(rawFileName, "/", ""), ":", ""), "-", ""), " ", "")
validFileName = Left(validFileName,len(validFileName)-2)
fileName = path & validFileName & ".vbs"

fso.createTextFile(filename)

'編寫提醒文件內容
set edit = fso.openTextFile(filename,2, True)

'提醒時發聲
edit.writeLine "set sapi = CreateObject(""SAPI.SpVoice"")" 
edit.writeLine "sapi.Speak(""" & content & """)" 

edit.writeLine "msgBox """ & content & """,4096+64,""提醒"""
edit.close

select case input

case 3: '開機提醒
    targetFolder = WshShell.SpecialFolders("Startup") & "\"
    fso.MoveFile fileName, targetFolder
    'ws.run "schtasks /create /sc ONSTART /tn RemindTask /tr " & fileName 

case 4: '指定日期及時間
    targetTime = inputbox("指定每日特定時間提醒-請輸入指定的小時及分鐘（用冒號分隔 e.g. 08:00,二十四小時制用英文半角冒號）",,":")
    targetDate = inputbox("指定日期（用/分隔 e.g. 24/01/2024）",,"//")
    ws.run "schtasks /create /sc ONCE /tn " & validFileName &" /st " & targetTime & " /sd " & targetDate  & " /tr " & fileName

case 5: '每日運行，可指定具體時間
    targetTime = inputbox("指定每日特定時間提醒-請輸入指定的小時及分鐘（用冒號分隔 e.g. 08:00,二十四小時制用英文半角冒號）",,":")
    ws.run "schtasks /create /sc DAILY /tn " & validFileName &" /st " & targetTime &  " /tr " & fileName

end select
