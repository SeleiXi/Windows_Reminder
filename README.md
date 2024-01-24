【繁中介紹】
只係適用於Windows系統，download咗之後可以直接運行

中文版script可能會出現"String未結束"呢個問題，
  1.右鍵檔案 2.選擇編輯 
  3.選擇左上角的"檔案"選項 4.選擇"另存為" 5.將右下部分的編碼改為"ANSI" 6.選擇"保存"
  
如果想要del咗之前創建咗嘅task：
  1.打開工作排程器
  2.打開程序庫
  3.刪除創建的任務

如果係想刪除設置咗開機就執行嘅task，同時點擊 Win鍵 + R，輸入shell:startup , 刪除對應時間嘅file（注意，file名係創建呢個文件的時間, 241202415339.vbs = 24/12/2024,15:33:09 可以用notepad嘅打開方式，打開個文件黎確認下入邊寫嘅提醒事項係咩）

輸入時注意：日期唔可以寫成 1/1/2024，而係要寫成 01/01/2024，填寫時間同理要寫成 01:01

【Introduction in English】
Applicable for Windows only. 
Download to your computer and run with a click.

Delete the task you have created:
  1.Open 'Task Scheduler'.
  2.Open the 'Task Library'.
  3.Delete the created task.

If the task is set to run on startup: Press 'Win + R' simultaneously, type 'shell:startup', Delete the file corresponding to the created time (Note: The file name represents the time the file was created. For example, 241202415339.vbs = 24/12/2024, 15:33:09). You can open the file with Notepad to check the reminder content.

Note when entering input: Dates should be written as '01/01/2024' instead of '1/1/2024', and time should also be written as '01:01'."

【简中介绍】
仅适用于Windows系统，下载到电脑点击即可运行

中文版脚本出现"字符串未结束"问题的解决方法:
  1.右键文件 2.选择编辑 
  3.选择左上角的"文件"选项 4.选择"另存为" 5.将右下部分的编码改为"ANSI" 6."保存"

删除创建的任务：
  1.打开任务计划程序
  2.打开程序库
  3.删除创建的任务

如果是想刪除设置为开机自启动的任务：同时点击 Win键 + R , 输入shell:startup , 删除对应时间的文件（注意，文件名是创建这个文件的时间。241202415339.vbs = 24/12/2024,15:33:09 以及可以以记事本的方式打开文件以确认里面的提醒事项是什麽）

输入时注意：日期不能写成 1/1/2024而是要写成 01/01/2024，填写时间同理要写成 01:01



