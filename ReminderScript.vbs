'����F���e�����ļ���ANSI���a����

set ws = wscript.createObject("Wscript.Shell")
set fso = createObject("scripting.filesystemobject")
Set wshShell = createObject("WScript.Shell")

path = fso.getFile(wscript.scriptFullName).parentFolder + "\"

input = msgBox("1.�_�C���ѣ�Abort" & vbCrLf & "2.��ԇ��Retry " & vbCrLf & "3.ÿ���ָ���r�g���ѣ�Ignore",2+64,"���r������")
content = inputBox("Ոݔ��Ҫ���ѵ����","���у���")

'�����ļ�����
rawFileName = Cstr(date) & Cstr(time)
validFileName = Replace(Replace(Replace(Replace(rawFileName, "/", ""), ":", ""), "-", ""), " ", "")
validFileName = Left(validFileName, 13)
fileName = path & validFileName & ".vbs"

fso.createTextFile(filename)

'���������ļ�����
set edit = fso.openTextFile(filename,2, True)

'���ѕr�l
edit.writeLine "set sapi = CreateObject(""SAPI.SpVoice"")" 
edit.writeLine "sapi.Speak(""" & content & """)" 

edit.writeLine "msgBox """ & content & """,4096+64,""����"""
edit.close

select case input

case 3: '�_�C����
    targetFolder = WshShell.SpecialFolders("Startup") & "\"
    fso.MoveFile fileName, targetFolder
    'ws.run "schtasks /create /sc ONSTART /tn RemindTask /tr " & fileName 

case 4: 'ָ�����ڼ��r�g
    targetTime = inputbox("ָ��ÿ���ض��r�g����-Ոݔ��ָ����С�r����犣���ð̖�ָ� e.g. 08:00,��ʮ��С�r����Ӣ�İ��ð̖��",,":")
    targetDate = inputbox("ָ�����ڣ���/�ָ� e.g. 24/01/2024��",,"//")
    ws.run "schtasks /create /sc ONCE /tn " & validFileName &" /st " & targetTime & " /sd " & targetDate  & " /tr " & fileName

case 5: 'ÿ���\�У���ָ�����w�r�g
    targetTime = inputbox("ָ��ÿ���ض��r�g����-Ոݔ��ָ����С�r����犣���ð̖�ָ� e.g. 08:00,��ʮ��С�r����Ӣ�İ��ð̖��",,":")
    ws.run "schtasks /create /sc DAILY /tn " & validFileName &" /st " & targetTime &  " /tr " & fileName

end select
