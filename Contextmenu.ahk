#SingleInstance Ignore
#NoTrayIcon
SetWorkingDir(A_ScriptDir)

ClipSaved := ClipboardAll()
A_Clipboard := ""
Send("^c")
ClipWait(1)
args := A_Args[1]
temp := A_Clipboard
A_Clipboard := ClipSaved

if temp
{
    for i in StrSplit(temp, "`r`n")
        args .= ' "' i '"'
} else
{
    ToolTip("未复制到路径,请重试")
    Sleep(1500)
    ExitApp
}

if FileExist("SmartZip.ahk")
    RunWait("SmartZip.ahk " args)
else if FileExist("SmartZip.exe")
    RunWait("SmartZip.exe " args)
else
    MsgBox("主脚本不存在")