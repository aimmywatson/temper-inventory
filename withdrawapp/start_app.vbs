Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

' สั่งให้ระบบชี้มาที่โฟลเดอร์ปัจจุบันที่ไฟล์ vbs นี้ตั้งอยู่
currentFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = currentFolder

' สั่งรัน Streamlit แบบซ่อนจอ (ใช้คำสั่ง python -m แก้ปัญหา Windows หาไม่เจอ)
WshShell.Run "cmd /c python -m streamlit run add.py", 0, False