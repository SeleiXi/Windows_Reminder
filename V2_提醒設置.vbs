on error resume next
input = msgbox("選擇中止=開機提醒；重試=指定日期時間提醒；忽略=每天的指定時間提醒，進入設置請輸入setting",2+64,"定時提醒器")
content = inputbox("請輸入要提醒的事項","提醒內容")
set fso = createobject("scripting.filesystemobject")
dim paths(4)
paths(0) = fso.getfile(wscript.scriptfullname).parentfolder + "\"
paths(1) = "C:\"
paths(2) = "D:\"
paths(3) = "E:\"

err.description = False
continue = False
for each path in paths
  test = path + "test.txt"
  fso.createtextfile(test)
  if fso.fileexists(test) then
    continue = True
  end if
  if continue = True then
    fso.DeleteFile(test)
    exit for
  end if
Next
Last_filename = ""
creation_moment = Cstr(date+time)
filename = path+creation_moment +".vbs"
filename = Cstr(filename)
Last_Filename = (Replace(filename,"/","-"))
Last_Filename = (Replace(Last_Filename,":","-",4))
Last_filename = Replace(Last_Filename," ","")
Last_Filename = Cstr(Last_Filename)
fso.createtextfile("C:\" + Last_filename)
fso.createtextfile("")

set edit = fso.opentextfile(Last_Filename,8,False)
edit.writeline "set sapi = CreateObject("&chr(34)&"SAPI.SpVoice"&chr(34)&")"
edit.writeline "sapi.Speak(" &chr(34)& content &chr(34)&")"
edit.writeline "msgbox "&chr(34)&content&chr(34)&",4096+64,"&chr(34)&"提醒"&chr(34)&""
edit.close
set path = fso.getfile("C:\" + Last_Filename)
msgbox last_filename
path.attributes = 2

if input = 3 then '開機提醒
'ws.run"schtasks /create /sc DAILY /tn "&task& " /tr " & filename  

elseif input = 4 then '指定日期及時間
'ws.run"schtasks /create /sc DAILY /tn "&task& " /tr " & filename 
input_items = inputbox("請輸入指定的日期和時間(day month hour min),24小時制,記得用空格分隔開")
task_time = split(input_items) ' 4格的array


elseif input = 5 then '指定每天時間
SetTime = inputbox("指定每日特定時間提醒-請輸入指定的小時及分鐘（用冒號分隔 e.g. 08:00,二十四小時制用英文半角冒號）",,":")
order = "schtasks /create /tn ParticularTimeEveryDay /tr C:\"&Last_Filename& "/sc daily /st " & SetTime
x = inputbox("x",,order)
msgbox("test")
ws.run order
end if