;@Ahk2Exe-SetName         SmartZip
;@Ahk2Exe-SetDescription  7-zip的功能扩展
;@Ahk2Exe-SetCopyright    Copyright (c) since 2022
;@Ahk2Exe-SetCompanyName  viv
;@Ahk2Exe-SetOrigFilename SmartZip.exe
;@Ahk2Exe-SetMainIcon     ico.ico
;@Ahk2Exe-SetFileVersion 2.11
;@Ahk2Exe-SetProductVersion %A_AhkVersion%
;@Ahk2Exe-ExeName SmartZip.exe

#SingleInstance off
#NoTrayIcon

ini := A_ScriptDir "\SmartZip.ini"
IniCreate

zip := SmartZip(IniRead(ini, "set", "7zipDir", ""))

;https://www.iconfont.cn/collections/detail?spm=a313x.7781069.0.da5a778a4&cid=24599

if !FileExist(icon := IniRead(ini, "set", "icon", ""))
    icon := zip.7zG

TraySetIcon(icon)

if A_Args.Length
    zip.Init(A_Args).Exec()
else
{
    IniSetting
    ContextMenu
}

ExitApp

class SmartZip
{
    ;用于命令行7Z检测 返回0正常 1错误

    __New(sevenZipDir)
    {
        this.now := A_TickCount
        sevenZipDir := RTrim(sevenZipDir, "\")

        if !DirExist(sevenZipDir)
        {
            MsgBox("请在 SmartZip.ini 中 7zipDir 选项中设置 7zip 的文件夹")
            ExitApp
        }

        this.7z := sevenZipDir "\7z.exe"
        this.7zG := sevenZipDir "\7zG.exe"
        this.7zFM := sevenZipDir "\7zFM.exe"
        if !FileExist(this.7z) || !FileExist(this.7zG) || !FileExist(this.7zFM)
        {
            MsgBox("7z文件夹中必需包含 7z.exe 7zG.exe 7zFM.exe`n请检测文件夹是否设置正确")
            Run(ini)
            ExitApp
        }

        this.continue := this.guiShow := false
        this.ext := Map()
        this.extExp := []
    }

    Init(argsArr)
    {
        if RegExMatch(argsArr[1], "^[xoa]$")
            this.to := argsArr[1], argsArr.RemoveAt(1)	;根据第一个传入参数决定动作
        else
            this.to := "x"
        this.arr := []

        for i in argsArr
        {
            if FileExist(i)
                loop files RTrim(i, "\"), "DF"
                    this.arr.Push(A_LoopFileFullPath)
        }

        this.muilt := this.arr.Length > 1	;多文件
        this.IniReadLoop("ext", this.ext, true)
        this.IniReadLoop("extExp", this.extExp)
        this.logLevel := IniRead(ini, "set", "logLevel", 0)
        this.cmdLog := IniRead(ini, "set", "cmdLog", 0)
        this.hideRunSize := IniRead(ini, "set", "hideRunSize", 0x7FFFFFFFFFFFFFFF)

        SplitPath(this.arr[1], , &dir)
        SetWorkingDir(this.dir := dir)
        return this
    }

    Exec()
    {
        switch this.to
        {
            case "x": this.Unzip()
            case "o": this.OpenZip()
            case "a": this.CreateZip()
            default:this.Unzip()
        }
    }

    Unzip()
    {
        this.password := ["", FormatPassword(A_Clipboard), IniRead(ini, "temp", "lastPass", "")]
        this.delSource := IniRead(ini, "set", "delSource", 0)
        this.delWhenHasPass := IniRead(ini, "set", "delWhenHasPass", 0)
        this.IniReadLoop("password", this.password)
        this.succesSpercent := IniRead(ini, "set", "succesSpercent", 0)

        ;批量解压文件中某项为嵌套压缩包时不显示7zip界面
        guiShow := IniRead(ini, "temp", "guiShow", "")

        for i in this.arr
        {
            this.tmpDir := tmpDir := '__7z' A_Now
            this.index := A_Index

            this.size := FileGetSize(i)
            hideBool := this.size / 1024 / 1024 < this.hideRunSize

            if this.muilt && !this.guiShow && !hideBool
                IniWrite(1, ini, "temp", "guiShow"), this.Gui()

            zipx(i)

            if !DirExist(tmpDir)	;密码错误以及未输入正确密码
                continue

            loop files tmpDir "\*.*", "RDF"
                AfterUnzip(A_LoopFileFullPath)

            loop files tmpDir "\*.*", "DF"
            {
                count := A_Index, souceFile := A_LoopFileFullPath
                if A_Index > 2
                    break
            }

            notDir := 0
            loop files tmpDir "\*.*", "FR"
            {
                if !DirExist(A_LoopFileFullPath)
                    notDir++ , souceFile2 := A_LoopFileFullPath
                if notDir = 2
                    break
            }

            if count = 1 || notDir = 1	;只有一个文件或文件夹
            {
                if notDir = count	;单个文件夹包含单个文件
                    souceFile := souceFile2

                isDir := DirExist(souceFile)
                SplitPath(souceFile, &name, , &ext)

                outFile := this.MoveItem(souceFile, this.dir "\" name, isDir, A_LineNumber)

                this.RecycleItem(tmpDir, A_LineNumber, true)
                if !isDir
                    UnZipNesting(outFile, ext)
            } else	;多个文件
            {
                SplitPath(i, , , , &nameNoEXT)
                outFile := this.MoveItem(tmpDir, this.dir "\" nameNoEXT, 1, A_LineNumber)

                if IniRead(ini, "set", "muiltNesting", false)
                    loop files outFile "\*.*", "F"
                        UnZipNesting(A_LoopFileFullPath, A_LoopFileExt)
            }
        }

        ; if hideBool	;隐藏运行时可以不会刷新,发送刷新到当前窗口
        ;     Send("{F5}")
        if this.muilt
            IniWrite("", ini, "temp", "guiShow")

        ;解压嵌套
        UnZipNesting(path, ext)
        {
            if !this.IsArchive(ext)
                return

            IniWrite(1, ini, "temp", "isLoop")
            RunWait('"' A_ScriptFullPath '" x "' path '"')
            this.Loging("解压嵌套 <--> " path, A_LineNumber)
            IniWrite("", ini, "temp", "isLoop")
            this.RecycleItem(path, A_LineNumber)	;删除嵌套文件
        }

        ;执行解压
        zipx(path)
        {
            pass := ""
            this.error := true
            for i in this.password
            {
                cmd := this.7z ' t "' path '" -aou -bsp1 -p"' i '"'
                log .= cmd "`n"
                this.CheckCMD()
                this.RunCmd(cmd)

                if this.continue
                    return
                if !this.error
                {
                    if i
                    {
                        pass := ' -p"' i '"'
                        IniWrite(i, ini, "temp", "lastPass")
                    }
                    break
                }
            }
            this.Loging(RTrim(log, "`n"), A_LineNumber, 3)

            if !this.error	;密码正确或无需密码
            {
                if hideBool
                    RunWait(this.7z ' x "' path '" -aou -o' tmpDir pass, , "hide")
                else
                    RunWait(this.7zG ' x "' path '" -aou -o' tmpDir pass, , this.guiShow || guiShow ? "hide" : "")

                if IsSuccess()
                {
                    if IniRead(ini, "temp", "isLoop", "")
                        this.RecycleItem(path, A_LineNumber, true)
                    else if pass && this.delWhenHasPass
                        this.RecycleItem(path, A_LineNumber)
                }
            } else	;密码错误需手动输入密码
            {
                RunWait(this.7zG ' x "' path '" -aou -o' tmpDir)
                if IsSuccess() && this.delSource
                    this.RecycleItem(path, A_LineNumber)
            }

            IsSuccess()
            {
                if !DirExist(tmpDir)
                    return false
                isTrue := ""
                folderSize := ComObject("Scripting.FileSystemObject").GetFolder(tmpDir).Size

                if folderSize >= this.size
                    isTrue := true
                else if this.size - folderSize <= this.size / 100 * this.succesSpercent
                    isTrue := true

                if isTrue
                {
                    this.Loging("解压 <--> " path, A_LineNumber)
                    return true
                }
                this.RecycleItem(tmpDir, A_LineNumber, true)
                return false
            }
        }

        ; 解压后处理
        AfterUnzip(path)
        {

            static isRead := false, obj := { rename: { ext: Map(), name: Map(), exp: Map() },
                delete: { ext: [], name: [], exp: [] } }

            if !isRead
            {
                this.IniReadLoop("renameExt", obj.rename.ext, , true)
                this.IniReadLoop("renameName", obj.rename.name, , true)
                this.IniReadLoop("renameExp", obj.rename.exp, , true)

                this.IniReadLoop("deleteExt", obj.delete.ext)
                this.IniReadLoop("deleteName", obj.delete.name)
                this.IniReadLoop("deleteExp", obj.delete.exp)

                isRead := true
            }

            if (isDir := DirExist(path)) && ComObject("Scripting.FileSystemObject").GetFolder(path).Size = 0	;空文件夹
            {
                return this.RecycleItem(path, A_LineNumber)
            }

            SplitPath(path, &name, &dir, &ext, &nameNoExt)

            for i in obj.delete.ext
                if ext = i
                    return this.RecycleItem(path, A_LineNumber)
            for i in obj.delete.name
                if InStr(name, i)
                    return this.RecycleItem(path, A_LineNumber)
            for i in obj.delete.exp
                if name ~= i
                    return this.RecycleItem(path, A_LineNumber)

            for ori, out in obj.rename.ext
                if !isDir && ext = ori
                    path := this.MoveItem(path, dir '\' nameNoExt '.' (ext := out), 0, A_LineNumber)

            for needle, replaceText in obj.rename.name
                if InStr(name, needle)
                    SplitPath(path := this.MoveItem(path, dir '\' StrReplace(name, needle, replaceText), isDir, A_LineNumber), &name, , , &nameNoExt)

            for needle, replaceText in obj.rename.exp
                if name ~= needle
                    SplitPath(path := this.MoveItem(path, dir '\' RegExReplace(name, needle, replaceText), isDir, A_LineNumber), &name, , , &nameNoExt)
        }

        FormatPassword(str)
        {
            if StrLen(str) < 100
            {
                str := RegExReplace(str, "(\R*)")	;移除所有换行符
                str := RegExReplace(str, "^[ \t]+|[ \t]+$")	;移除首尾所有空格或制表符
            } else
                str := ""
            return str
        }
    }

    OpenZip()
    {
        SplitPath(this.arr[1], , , &ext, &nameNoExt)
        this.CheckCMD("openZip")

        if !this.muilt
        {
            extForOpen := Map()
            this.IniReadLoop("extForOpen", extForOpen, true)

            path := ' "' this.arr[1] '" '

            If DirExist(this.arr[1])
                this.error := true
            else if (this.IsArchive(ext) || extForOpen.Has(ext))
                this.error := false
            else
                this.RunCmd(this.7z ' l ' path)

            if !this.error
            {
                Run(this.7zFM path, , , &pid)
                ; hwnd := WinWaitActive("ahk_pid" pid)
                ; if WinActive("ahk_class #32770 ahk_pid " pid)
                ;     WinClose(), this.error := true
                this.Loging("打开 <--> " path, A_LineNumber)
            }
        }

        if this.muilt || this.error
        {
            args := IniRead(ini, "7z", "openAdd")
            path := ""
            for i in this.arr
                path .= ' "' i '" '
            zipname := this.muilt ? StrReplace(RegExReplace(this.dir, ".+\\"), ":") : nameNoExt
            zipname := this.AUO(zipname, RegExReplace(args, '(.+?)".+', "$1"))

            RunWait(this.7zG ' a "' zipname args path)

            this.Loging("新建压缩 <--> " path, A_LineNumber)
        }
    }

    CreateZip()
    {
        SplitPath(this.arr[1], , , , &nameNoExt)

        count := 0
        for i in this.arr
            if DirExist(i)
                count++
        hideBool := IsHide()

        args := IniRead(ini, "7z", "add")
        ext := RegExReplace(args, '(.+?)".+', "$1")

        if count = this.arr.Length	;全是文件夹,单独添加
        {

            for i in this.arr
            {
                hideBool := IsHide(i)
                if !hideBool && !this.guiShow
                    this.Gui

                this.index := A_Index
                RunWait((hideBool ? this.7z : this.7zG) ' a "' this.AUO(RegExReplace(i, ".*\\"), ext) args ' "' i '\*"', , (hideBool || count > 1) ? "hide" : "")
                this.Loging("压缩 <--> " i, A_LineNumber)
            }
            return

        } else if this.arr.Length = 1	;单个文件
            RunWait((hideBool ? this.7z : this.7zG) ' a "' this.AUO(nameNoExt, ext) args ' "' this.arr[1] '"', , hideBool ? "hide" : ""), this.Loging("压缩 <--> " this.arr[1], A_LineNumber)
        else	;文件文件夹混合
        {
            for i in this.arr
                path .= ' "' i '" '
            RunWait((hideBool ? this.7z : this.7zG) ' a "' this.AUO(RegExReplace(this.dir, ".+\\"), ext) args path, , hideBool ? "hide" : "")
            this.Loging("压缩 <--> " path, A_LineNumber)
        }
        ; if hideBool
        ;     Send("{F5}")

        IsHide(dir := "")
        {
            countSzie := 0
            if dir
            {
                loop files dir "\*.*", "RDF"
                    if (countSzie += A_LoopFileSizeMB) > this.hideRunSize
                        return false
                return true
            }

            for i in this.arr
            {
                if DirExist(i)
                {
                    loop files i "\*.*", "RDF"
                        if (countSzie += A_LoopFileSizeMB) > this.hideRunSize
                            return false
                } else
                    countSzie += FileGetSize(i, "M")
                if countSzie > this.hideRunSize
                    return false
            }
            return true
        }
    }

    AUO(name, ext) => StrReplace(this.PathDupl(this.dir "\" name ext, 0), ext)

    Gui()
    {
        this.guiShow := true
        DetectHiddenWindows(1)

        g := Gui("+LastFound")
        hWnd := WinExist()
        DllCall("RegisterShellHookWindow", "UInt", hWnd)
        MsgNum := DllCall("RegisterWindowMessage", "Str", "SHELLHOOK")
        OnMessage(MsgNum, ShellMessage)

        sub := "ahk_exe 7zG.exe"
        g.SetFont(, "Segoe UI")
        g.BackColor := "FFFFFF"
        left1 :=

        已用时间1 := g.Add("text", "w100"), 已用时间2 := leftVaule()
        总大小1 := rightTitle(), 总大小2 := rightVaule()

        剩余时间1 := leftTitle(), 剩余时间2 := leftVaule()
        速度1 := rightTitle(), 速度2 := rightVaule()

        文件1 := leftTitle(), 文件2 := leftVaule()
        已处理1 := rightTitle(), 已处理2 := rightVaule()

        leftTitle(), 文件3 := leftVaule()
        压缩后大小1 := rightTitle(), 压缩后大小2 := rightVaule()

        总进度1 := leftTitle(), 总进度2 := leftVaule()
        压缩率1 := rightTitle(), 压缩率2 := rightVaule()

        g.Add("text", "h1 w1")

        处理1 := g.Add("text", "xs w500")
        if this.to = "x"
            处理2 := g.Add("text", "w500")
        处理3 := g.Add("text", "w500")
        进度 := g.Add("Progress", "w500 h20")

        leftTitle() => g.Add("text", "xs w100")
        leftVaule() => g.Add("text", "yp w100 x120 Right")
        rightTitle() => g.Add("text", "yp w100 x300")
        rightVaule() => g.Add("text", "yp w100 x420 Right")

        g.Add("text")

        g.AddButton("w100 x10", "显示原始界面").OnEvent("Click", ButtonShowHide)
        暂停 := g.AddButton("x300 w100 yp")
        暂停.OnEvent("Click", ButtonPause)
        取消 := g.AddButton("w100 yp")
        取消.OnEvent("Click", ButtonCance)

        g.OnEvent("Close", Close)
        g.OnEvent("Escape", Close)
        g.Show("AutoSize")
        ; SetTimer(update, 500)

        Close(*)
        {
            if ProcessExist("7zG.exe")
                ProcessClose("7zG.exe"), ProcessWaitClose("7zG.exe")
            if this.HasOwnProp("tmpDir")
                this.RecycleItem(this.tmpDir, A_LineNumber, true)
            IniWrite("", ini, "temp", "guiShow")
            ExitApp
        }

        ButtonShowHide(GuiCtrlObj, *)
        {
            DetectHiddenWindows(0)

            if !ProcessExist("7zG.exe")
                return

            if WinExist(sub)
                WinHide(sub), GuiCtrlObj.Text := "显示原始界面"
            else
                WinShow(sub), WinActivate(sub), GuiCtrlObj.Text := "隐藏原始界面"

            DetectHiddenWindows(1)
        }

        ButtonPause(GuiCtrlObj, Info)
        {
            if !WinExist(sub)
                return

            textSave := GuiCtrlObj.Text
            ControlClick("Button2", sub)
            ShellMessage
            while WinExist(sub) && GuiCtrlObj.Text = textSave
                ControlClick("Button2", sub), Sleep(500), ShellMessage()
        }

        ButtonCance(GuiCtrlObj, Info)
        {
            if !WinExist(sub)
                return
            ControlClick("Button3", sub)
            ShellMessage
            while WinExist(sub) && !InStr(WinGetText(sub), "否(&N)")
                ControlClick("Button3", sub), Sleep(500)

            WinWaitActive(sub)
            WinWaitClose
            num := 0
            loop 5
                if !WinExist(sub)
                    num++

            if num = 5
                Close
        }

        ; update()
        ShellMessage(wParam := 6, *)
        {
            static timeSave := A_TickCount
            if A_TickCount - timeSave < 100 || wParam != 6 || !WinExist(sub)
                return

            IsChanged(总进度1, "总进度:")
            , IsChanged(总进度2, this.index "\" this.arr.Length)
            try
            {
                if g.Title != WinGetTitle(sub)
                    g.Title := WinGetTitle(sub)

                arr := StrSplit(WinGetText(sub), "`n")
                for i in arr
                {
                    if InStr(i, "您真的要取消吗")
                        return
                    arr[A_Index] := SubStr(i, 1, -1)
                }

                IsChanged(进度, RegExReplace(g.Title, "(.+ )?(\d+)%.+", "$2"))
                , IsChanged(暂停, arr[2], 1)
                , IsChanged(取消, arr[3], 1)

                , IsChanged(已用时间1, arr[4])
                , IsChanged(剩余时间1, arr[5])
                , IsChanged(文件1, arr[6])

                ; IsChanged(发生错误, arr[7])

                , IsChanged(总大小1, arr[8])
                , IsChanged(速度1, arr[9])
                , IsChanged(已处理1, arr[10])
                , IsChanged(压缩后大小1, arr[11])
                , IsChanged(压缩率1, arr[12])

                , IsChanged(已用时间2, arr[13])
                , IsChanged(剩余时间2, arr[14])
                , IsChanged(文件2, arr[15])

                index := 16
                if this.to = "a"
                    IsChanged(文件3, arr[index++ ])

                IsChanged(总大小2, arr[index++ ])	;16
                , IsChanged(速度2, arr[index++ ])	;17
                , IsChanged(已处理2, arr[index++ ])	;18
                , IsChanged(压缩后大小2, arr[index++ ])	;19
                , IsChanged(压缩率2, arr[index++ ])	;20

                , IsChanged(处理1, arr[index++ ])	;21

                index++
                if this.to = "x"
                    IsChanged(处理2, arr[index - 1])	;22

                IsChanged(处理3, arr[index])	;23
            }
            timeSave := A_TickCount

            IsChanged(obj, value, text := 0)
            {
                if !text
                {
                    if obj.Value != value
                        obj.Value := value
                } else if obj.Text != value
                    obj.Text := value
            }
        }
    }

    IniReadLoop(Section, arrMap, lower := false, twoVar := false)
    {
        loop
        {
            if !(var := IniRead(ini, Section, A_Index, ""))
                break

            if lower
                var := StrLower(var)

            if Type(arrMap) = "Array"
                arrMap.Push(var)
            else if Type(arrMap) = "Map"
            {
                if twoVar
                    arrMap[RegExReplace(var, "<--->.*")] := RegExReplace(var, ".+<--->")
                else
                    arrMap[var] := true
            }
        }
        if Type(arrMap) = "Array"
        {
            loop
            {
                if arrMap.Length < A_Index
                    break

                var := arrMap[A_Index]
                temp := []
                for i in arrMap
                    if var = i
                        temp.Push(A_Index)

                temp.RemoveAt(1)
                loop temp.Length
                    arrMap.RemoveAt(temp[temp.Length - A_Index + 1])
            }
        }
    }

    RecycleItem(souce, lineNum, delete := false)
    {
        try
        {
            if delete
                DirExist(souce) ? DirDelete(souce, 1) : FileDelete(souce)
            else
                FileRecycle(souce)
            this.Loging(souce, lineNum, 1)
        }
    }

    MoveItem(souce, dest, isdir, lineNum)
    {
        try
        {
            (isDir ? DirMove : FileMove)(souce, oPath := this.PathDupl(dest, isdir))
            this.Loging(souce " <--> " oPath, lineNum, 2)
            return oPath
        } catch
            return souce
    }

    PathDupl(path, isdir := 0)
    {
        if FileExist(path)	;目标文件重复
        {
            SplitPath(path, , &dir, &ext, &nameNoExt)

            if isdir && ext	;文件夹包含.被识别为文件  示例 " D:\1.2"
                nameNoExt := RegExReplace(path, ".*\\")

            ext := (isdir || !ext) ? "" : "." ext	;目标为文件夹 或 目标为文件但无 ext 时 ext 为空
            while FileExist(path)
                path := dir '\' nameNoExt '_' A_Index ext
        }
        return path
    }

    IsArchive(ext)
    {
        ext := StrLower(ext)

        if !ext
            return true

        if this.ext.Has(zip)
            return true

        for i, n in this.ext
            if InStr(i, ext)
                return true

        for i in this.extExp
            if ext ~= "i)" i
                return true

        return false
    }

    ;https://www.autohotkey.com/boards/viewtopic.php?t=93944
    RunCmd(CmdLine, Codepage := "CP0") {
        DllCall("CreatePipe", "PtrP", &hPipeR := 0, "PtrP", &hPipeW := 0, "Ptr", 0, "Int", 0)
        , DllCall("SetHandleInformation", "Ptr", hPipeW, "Int", 1, "Int", 1)
        , DllCall("SetNamedPipeHandleState", "Ptr", hPipeR, "UIntP", &PIPE_NOWAIT := 1, "Ptr", 0, "Ptr", 0)

        , P8 := (A_PtrSize = 8)
        , SI := Buffer(P8 ? 104 : 68, 0)	; STARTUPINFO structure
        , NumPut("UInt", P8 ? 104 : 68, SI)	; size of STARTUPINFO
        , NumPut("UInt", STARTF_USESTDHANDLES := 0x100, SI, P8 ? 60 : 44)	; dwFlags
        , NumPut("Ptr", hPipeW, SI, P8 ? 88 : 60)	; hStdOutput
        , NumPut("Ptr", hPipeW, SI, P8 ? 96 : 64)	; hStdError
        , PI := Buffer(P8 ? 24 : 16)	; PROCESS_INFORMATION structure

        If !DllCall("CreateProcess", "Ptr", 0, "Str", CmdLine, "Ptr", 0, "Int", 0, "Int", True, "Int", 0x08000000 | DllCall("GetPriorityClass", "Ptr", -1, "UInt"), "Int", 0
            , "Ptr", 0, "Ptr", SI.ptr, "Ptr", PI.ptr)
            Return Format("{1:}", "", -1, DllCall("CloseHandle", "Ptr", hPipeW), DllCall("CloseHandle", "Ptr", hPipeR))

        DllCall("CloseHandle", "Ptr", hPipeW)
        , this.CMDPID := NumGet(PI, P8 ? 16 : 8, "UInt")
        , File := FileOpen(hPipeR, "h", Codepage)
        , LineNum := 1, sOutput := ""

        While (this.CMDPID + DllCall("Sleep", "Int", 0)) && DllCall("PeekNamedPipe", "Ptr", hPipeR, "Ptr", 0, "Int", 0, "Ptr", 0, "Ptr", 0, "Ptr", 0)
            While this.CMDPID && !File.AtEOF
                Line := File.ReadLine(), sOutput .= this.CheckCMD(LineNum++ , Line)

        this.CMDPID := 0
        , hProcess := NumGet(PI, 0, "Ptr")
        , hThread := NumGet(PI, A_PtrSize, "Ptr")

        , DllCall("CloseHandle", "Ptr", hProcess)
        , DllCall("CloseHandle", "Ptr", hThread)
        , DllCall("CloseHandle", "Ptr", hPipeR)

        Return sOutput
    }

    CheckCMD(what := "unZip", line := "")
    {
        static check, checkSave := "", whatSave := what

        if Type(what) = "String"
        {
            if !checkSave
            {
                checkSave := { error: Map(), errorExp: Map(), errorrContinueExp: Map(), success: Map(), successExp: Map() }
                this.IniReadLoop(what "CheckError", checkSave.error, , true)
                this.IniReadLoop(what "CheckErrorExp", checkSave.errorExp, , true)
                this.IniReadLoop(what "CheckErrorContinueExp", checkSave.errorrContinueExp, , true)
                this.IniReadLoop(what "CheckSuccess", checkSave.success, , true)
                this.IniReadLoop(what "CheckSuccessExp", checkSave.successExp, , true)
            }
            check := objCloneMap(checkSave)
            this.isCmdReturn := false
        } else if line
        {
            if this.isCmdReturn
                return

            if this.cmdLog
            {
                static cmdobj := FileOpen("__SmartZipCmdLog.txt", "a", "UTF-8")
                cmdobj.WriteLine("[" what "]`n" line)
            }

            for i in check.errorrContinueExp
                if line ~= i && --check.errorrContinueExp[i] < 1
                    return (this.continue := true, LogAndReturn(1, A_LineNumber))

            for i in check.error
                if InStr(line, i) && --check.error[i] < 1
                    return LogAndReturn(2, A_LineNumber)

            for i in check.errorExp
                if line ~= i && --check.errorExp[i] < 1
                    return LogAndReturn(3, A_LineNumber)

            for i in check.success
                if InStr(line, i) && --check.success[i] < 1
                    return LogAndReturn(4, A_LineNumber)

            for i in check.successExp
                if line ~= i && --check.successExp[i] < 1
                    return LogAndReturn(5, A_LineNumber)

            if whatSave = "unZip"
            {
                if (line ~= "^ +[1-9]+%") && !(line ~= "^ +[1-9]+%.+Open$" || line ~= "^ +[1-9]+%$")
                    return LogAndReturn(6, A_LineNumber)
            }

            LogAndReturn(num, lineNum)
            {
                this.isCmdReturn := true
                ProcessClose(this.CMDPID), ProcessWaitClose(this.CMDPID)
                if num < 4
                    this.error := 1
                else
                    this.error := 0
                this.Loging(what "<-->" line, lineNum, this.error ? 3 : 4)
            }
        }

        objCloneMap(obj)
        {
            tempObj := {}
            for i in Obj.OwnProps()
                tempObj.%i% := obj.%i%.Clone()
            return tempObj
        }
    }

    ; 关闭0/删除1/重命名2/命令行错误3/命令行正确4/其他5
    Loging(log, lineNum, level := 5)
    {
        if !this.logLevel || level > this.logLevel
            return

        static logObj := FileOpen(A_ScriptDir "\log.txt", "a", "UTF-8")

        switch level
        {
            case 5: msg := "其他"
            case 4: msg := "命令行正确"
            case 3: msg := "命令行错误"
            case 2: msg := "重命名"
            case 1: msg := "删除"
        }

        logObj.WriteLine(Format("[{1}] [{2}ms] [{3}] [{4}]`n{5}", msg, A_TickCount - this.now, A_Now, lineNum, log))
        this.now := A_TickCount
    }

}

IniCreate()
{

    iniExist := FileExist(ini)
    version := IniRead(ini, "other", "version", "2.0")
    iniEx := A_ScriptDir "`\ini说明.txt"

    if !iniExist
    {
        IniWrite(A_ScriptDir "\7-zip", ini, "set", "7zipDir")

        IniWrite("10", ini, "set", "successPercent")
        IniWrite("0", ini, "set", "delSource")
        IniWrite("0", ini, "set", "delWhenHasPass")
        IniWrite("5", ini, "set", "logLevel")
        IniWrite(A_ScriptDir "\ico.ico", ini, "set", "icon")

        IniWrite("123456", ini, "password", "1")
        IniWrite("666888", ini, "password", "2")
        IniWrite("1024", ini, "password", "3")
        IniWrite("++", ini, "password", "4")
        IniWrite("", ini, "password", "5")
        IniWrite("", ini, "password", "6")
        IniWrite("", ini, "password", "7")

        IniWrite("zip", ini, "ext", "1")
        IniWrite("rar", ini, "ext", "2")
        IniWrite("7z", ini, "ext", "3")
        IniWrite("001", ini, "ext", "4")
        IniWrite("cab", ini, "ext", "5")
        IniWrite("bz2", ini, "ext", "6")
        IniWrite("gz", ini, "ext", "7")
        IniWrite("gzip", ini, "ext", "8")
        IniWrite("tar", ini, "ext", "9")

        IniWrite("^\d+$", ini, "extExp", "1")
        IniWrite("zi", ini, "extExp", "2")
        IniWrite("7", ini, "extExp", "3")
        IniWrite("z", ini, "extExp", "4")

        IniWrite("iso", ini, "extForOpen", "1")
        IniWrite("apk", ini, "extForOpen", "2")
        IniWrite("wim", ini, "extForOpen", "3")
        IniWrite("exe", ini, "extForOpen", "4")

        IniWrite("mp+3<--->mp3", ini, "renameExt", "1")

        IniWrite("666666-<--->", ini, "renameName", "1")

        IniWrite("^[ 	]+<--->", ini, "renameExp", "1")

        IniWrite("", ini, "deleteExt", "1")

        IniWrite("来自666666.org", ini, "deleteName", "1")
        IniWrite("关注666666网.txt", ini, "deleteName", "2")
        IniWrite("扫码关注公众号.jpg", ini, "deleteName", "3")
        IniWrite("前往_666666", ini, "deleteName", "4")
        IniWrite("自行添加文件后缀.7z.txt", ini, "deleteName", "5")

        IniWrite("", ini, "deleteExp", "1")

        IniWrite("Wrong password<--->1", ini, "unZipCheckError", "1")
        IniWrite("ERROR:<--->10", ini, "unZipCheckError", "2")
        IniWrite("No files to process<--->1", ini, "unZipCheckError", "3")
        IniWrite("Cannot open encrypted archive<--->1", ini, "unZipCheckError", "4")

        IniWrite("", ini, "unZipCheckErrorExp", "1")

        IniWrite("Cannot open the file as archive<--->1", ini, "unZipCheckErrorContinueExP", "1")

        IniWrite("Everything is Ok<--->1", ini, "unZipCheckSuccess", "1")

        IniWrite("", ini, "unZipCheckSuccessExp", "1")

        IniWrite("ERROR:<--->1", ini, "openZipCheckError", "1")

        IniWrite("", ini, "openZipCheckErrorExp", "1")

        IniWrite("Enter password (will not be echoed):<--->1", ini, "openZipCheckSuccess", "1")	;需要输入密码则可能是压缩文件

        IniWrite("\d*-\d*-\d* *\d*:\d*:\d* *\d* *\d* *(\d*) files(, (\d*) folders)?<--->1", ini, "openZipCheckSuccessExp", "1")	;多少个文件多少个文件夹则可能是压缩文件

        IniWrite('.zip" -tzip -mx=0 -aou -ad', ini, "7z", "openAdd")
        IniWrite('.zip"', ini, "7z", "add")

        IniWrite("1", ini, "menu", "openZip")
        IniWrite("用7-Zip打开", ini, "menu", "openZipName")
        IniWrite("1", ini, "menu", "unZip")
        IniWrite("智能解压", ini, "menu", "unZipName")
        IniWrite("1", ini, "menu", "addZip")
        IniWrite("压缩", ini, "menu", "addZipName")
    }

    if !iniExist || VersionsCompare("2.10")
    {
        IniWrite("1", ini, "set", "muiltNesting")
        IniWrite("10", ini, "set", "hideRunSize")
        IniWrite("0", ini, "set", "cmdLog")
    }

    if !iniExist || VersionsCompare("2.11")
        IniWrite("2.11", ini, "other", "version")

    if !iniExist
        IniSetting

    VersionsCompare(str)
    {
        if str = version
            return false
        if RegExReplace(str, "\..*") < RegExReplace(version, "\..*")
            return true
        if RegExReplace(str, ".+?\.") > RegExReplace(version, ".+?\.")
            return true
    }

}

IniSetting()
{
    static hasShow := 0
    if hasShow
        return

    CoordMode("ToolTip", "Window")
    SetTimer(fn, 100)
    RunWait(ini)
    SetTimer(fn, 0)
    hasShow++

    fn()
    {
        if WinExist("SmartZip.ini") && WinActive()
            ToolTip("设置完 ini 后会继续运行", 0, 0)
        else
            ToolTip
    }
}

ContextMenu()
{
    openZip := IniRead(ini, "menu", "openZip", "")
    unZip := IniRead(ini, "menu", "unZip", "")
    addZip := IniRead(ini, "menu", "addZip", "")

    if !openZip && !unZip && !addZip
        return

    keyPath := "HKCU\SOFTWARE\Classes\AllFilesystemObjects\shell"

    answer := MsgBox("`t点击是注册右键菜单,否删除右键菜单,取消退出`t", "SmartZip", "YNC")
    if answer = "Yes"
    {
        menuPath := A_ScriptDir "\Contextmenu"

        if FileExist(menuPath ".ahk")
            menuPath := '"' A_AhkPath '" "' menuPath '.ahk" '
        else if FileExist(menuPath ".exe")
            menuPath := '"' menuPath '.exe" '
        else
            MsgBox("右键菜单所需要文件不存在"), ExitApp()

        if openZip
        {
            RegWrite(icon, "REG_SZ", keyPath "\OpenZip", "Icon")
            RegWrite(IniRead(ini, "menu", "openZipName"), "REG_SZ", keyPath "\OpenZip")
            RegWrite(menuPath "o", "REG_SZ", keyPath "\OpenZip\command")
        }

        if unZip
        {
            RegWrite(icon, "REG_SZ", "HKCU\SOFTWARE\Classes\*\shell\UnZip", "Icon")
            RegWrite(IniRead(ini, "menu", "unZipName"), "REG_SZ", "HKCU\SOFTWARE\Classes\*\shell\UnZip")
            RegWrite(menuPath "x", "REG_SZ", "HKCU\SOFTWARE\Classes\*\shell\UnZip\command")
        }

        if addZip
        {
            RegWrite(icon, "REG_SZ", keyPath "\AddZip", "Icon")
            RegWrite(IniRead(ini, "menu", "addZipName"), "REG_SZ", keyPath "\AddZip")
            RegWrite(menuPath "a", "REG_SZ", keyPath "\AddZip\command")
        }

        TrayTip("已注册右键菜单", "SmartZip")
        ToolTip("已注册右键菜单")

    } else if answer = "No"
    {
        try RegDeleteKey(keyPath "\OpenZip")
        try RegDeleteKey("HKCU\SOFTWARE\Classes\*\shell\UnZip")
        try RegDeleteKey(keyPath "\AddZip")
        TrayTip("右键菜单已删除", "SmartZip")
        ToolTip("右键菜单已删除")
    }
    Sleep(2000)
}