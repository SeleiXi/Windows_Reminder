僅適用於Windows，下載到電腦點擊即可運行

中文版腳本出現"字符串未結束"問題的解決方法:
  1.右鍵文件 2.選擇編輯 
  3.選擇左上角的"文件"選項 4.選擇"另存為" 5.將右下部分的編碼改為"ANSI" 6."保存"

刪除創建的任務：
  1.打開task scheduler（英文版） /  工作排程器（繁中版）/ 任务计划程序（簡中版）
  2.打開程序庫
  3.刪除創建的任務

如果是設置為開機自啟動的任務：同時點擊win + r , 輸入shell:startup , 刪除對應時間的文件（注意，文件名是創建這個文件的時間。241202415339.vbs = 24/12/2024,15:33:09 以及可以以記事本的方式打開文件以確認裡面的提醒事項是什麼）

輸入時注意：日期不能寫成 1/1/2024而是要寫成 01/01/2024，填寫時間同理要寫成 01:01


Applicable for Windows only. 
Download to your computer and run with a click.

Delete the task you have created:
  1.Open 'Task Scheduler'.
  2.Open the 'Task Library'.
  3.Delete the created task.

If the task is set to run on startup: Press 'Win + R' simultaneously, type 'shell:startup', Delete the file corresponding to the created time (Note: The file name represents the time the file was created. For example, 241202415339.vbs = 24/12/2024, 15:33:09). You can open the file with Notepad to check the reminder content.

Note when entering input: Dates should be written as '01/01/2024' instead of '1/1/2024', and time should also be written as '01:01'."




