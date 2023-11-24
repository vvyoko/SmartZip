;@Ahk2Exe-SetName         SmartZip
;@Ahk2Exe-SetDescription  7-zip的功能扩展
;@Ahk2Exe-SetCopyright    Copyright (c) since 2022
;@Ahk2Exe-SetCompanyName  viv
;@Ahk2Exe-SetOrigFilename SmartZip.exe
;@Ahk2Exe-SetMainIcon     ico.ico
;@Ahk2Exe-SetFileVersion 3.4
;@Ahk2Exe-SetProductVersion 17
;@Ahk2Exe-ExeName SmartZip.exe
buildVersion := 18
MainVersion := "3.4"
;Msgbox FormatTime(A_Now, "yyyy/M/d H:m:s")
buileTime := "2022/8/3 14:27:28"
app := "SmartZip"
#SingleInstance off
#NoTrayIcon

ini.Init(A_ScriptDir "\" app ".ini")
IniCreate
zip := SmartZip(RelativePath(ini.zipDir))

;https://www.iconfont.cn/collections/detail?spm=a313x.7781069.0.da5a778a4&cid=24599
icon := FileExist(icon := RelativePath(ini.icon)) ? icon : ""

TraySetIcon(icon)

if A_Args.Length
    zip.Init(A_Args).Exec()
else
    Setting

class SmartZip
{
    __New(sevenZipDir)
    {
        this.now := A_TickCount
        this.exitCode := -1
        this.setShow := false

        sevenZipDir := sevenZipDir ~= "i)^[a-z]:\\$" ? sevenZipDir : RTrim(sevenZipDir, "\")

        if !DirExist(sevenZipDir)
            return MsgBox("7-zip 文件夹不存在,请设置其路径")

        this.7z := sevenZipDir "\7z.exe"
        this.7zG := sevenZipDir "\7zG.exe"
        this.7zFM := sevenZipDir "\7zFM.exe"

        if !FileExist(this.7z) || !FileExist(this.7zG) || !FileExist(this.7zFM)
            return MsgBox("7-zip 文件夹中必需包含 7z.exe,7zG.exe,7zFM.exe`n请检测文件夹是否设置正确")
    }

    Init(argsArr)
    {
        this.codePage := ""
        if argsArr[1] = "xc"
            SetCodePage(), argsArr.RemoveAt(1)

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

        if !this.arr.Length
            ExitApp(2)
        this.isRunning := true
        this.muilt := this.arr.Length > 1	;多文件

        SplitPath(this.arr[1], , &dir)
        SetWorkingDir(this.defaultDir := dir)

        this.continue := this.guiShow := this.cmdHide := false
        this.pid := this.log := this.testLog := ''

        this.ext := Map()
        this.extExp := []

        ini.ReadLoop("ext", this.ext, true)
        ini.ReadLoop("extExp", this.extExp)

        this.logLevel := ini.logLevel

        this.cmdLog := ini.cmdLog
        this.hideRunSize := ini.hideRunSize

        if this.logLevel || this.cmdLog
            OnExit(ExitLog)

        return this

        ExitLog(*)
        {
            if this.log
                FileAppend(this.log "`n", A_ScriptDir "\log.txt", "UTF-8")
            if this.testLog
                FileAppend(this.testLog "`n", A_ScriptDir "\cmdLog.txt", "UTF-8")
        }

        SetCodePage()
        {
            ini.ReadLoop("codepage", cpCustomArr := [])
            arr := ["简体中文（GBK）", "繁体中文（大五码）", "日文（Shift_JIS）", "韩文（EUC-KR）", "UTF-8 Unicode"]
            for i in cpCustomArr
                arr.Push(i)
            cpArr2 := [936, 950, 932, 949, 65001]

            cpG := gui("+AlwaysOnTop +ToolWindow", "请选择或输入你需要的代码页")
            cpG.AddText()
            cpG.SetFont(, "Segoe UI")
            cpG.AddLink("", '<a href="https://docs.microsoft.com/zh-cn/windows/win32/intl/code-page-identifiers">其他代码页</a>')
            v := cpG.AddComboBox("", arr)
            v.ToolTip := "如需添加其他常用代码页,请参考上方链接`n输入 数字(标识符) 点击右边的 添加`n删除只能移除添加的项目"
            cpG.AddButton("yp", "添加").OnEvent("Click", Add)
            cpG.AddButton("yp", "删除").OnEvent("Click", Delete)
            cpG.AddText()
            cpG.AddButton("", "确定").OnEvent("Click", DetectCp)
            cpG.OnEvent("Escape", Close)
            cpG.OnEvent("Close", Close)
            cpG.Show()
            OnMessage(0x200, WM_MOUSEMOVE)
            WinWaitClose(cpg.Hwnd)

            Close(*) => (OnMessage(0x200, WM_MOUSEMOVE, 0), cpg.Destroy())

            Add(*)
            {
                if !(text := v.Text) || !IsNumber(text)
                    return
                for i in arr
                    if i = text
                        return
                cpCustomArr.Push(text), arr.push(Text), v.Delete(), v.Add(arr)
            }

            Delete(*)
            {
                if v.Value < 6 || !(text := v.Text) || !IsNumber(text)
                    return
                cpCustomArr.RemoveAt(v.Value - 5), arr.RemoveAt(v.Value), arr.Text := "", v.Delete(), v.Add(arr)

            }

            DetectCp(*)
            {
                if IsNumber(text := v.Text)
                    this.codePage := " -mcp=" text
                else
                {
                    for i in arr
                    {
                        if text = i
                        {
                            this.codePage := " -mcp=" cpArr2[A_Index]
                            break
                        }
                    }
                }

                for i in cpCustomArr
                {
                    if ini.Read(A_Index, , "codepage") != i
                        ini.Write(i, A_Index, "codepage")
                }
                loop
                {
                    if !(vaf := ini.Read(cpCustomArr.Length + A_Index, , "codepage"))
                        break
                    ini.Delete("codepage", cpCustomArr.Length + A_Index)
                }

                Close
            }

        }
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
        this.isRunning := false
        if this.cmdHide && !this.guiShow
        {
            ToolTip("处理完成")
            Sleep(2000)
            ToolTip()
        }
        if !this.setShow
            ExitApp
    }

    Unzip(loopPath := "")
    {
        if !loopPath
        {
            arr := this.arr
            this.autoAddPass := ini.autoAddPass
            this.dynamicPassSort := ini.dynamicPassSort
            this.test := ini.test
            this.partSkip := ini.partSkip
            this.delSource := ini.delSource
            this.delWhenHasPass := ini.delWhenHasPass
            this.nesting := ini.nesting
            this.nestingMuilt := ini.nestingMuilt
            this.succesSpercent := ini.successPercent
            this.autoRemovePass := ini.autoRemovePass
            if (targetDir := ini.targetDir)
                targetDir := targetDir ~= "i)^[a-z]:\\$" ? targetDir : RTrim(targetDir, "\")

            this.addDir2Pass := ini.Read("addDir2Pass", , "set")

            if targetDir && DirExist(targetDir)
                SetWorkingDir(this.defaultDir := targetDir)

            this.password := ["", ini.lastPass, FormatPassword(A_Clipboard)]

            ini.ReadLoop("password", this.password)

            excludeExt := []
            ini.ReadLoop("excludeExt", excludeExt)
            excludeName := []
            ini.ReadLoop("excludeName", excludeName)
            this.excludeArgs := ""
            if excludeExt.Length || excludeName.Length
            {
                for i in excludeExt
                    this.excludeArgs .= ' -x!*.' i
                for i in excludeName
                    this.excludeArgs .= ' -x!*' i '*'
            }
            if this.excludeArgs
                this.excludeArgs .= " -r"

            if this.dynamicPassSort || this.autoAddPass
            {
                ini.ReadLoop("password", this.dynamicPassArr := [])
                this.passwordMap := Map()

                for i in this.dynamicPassArr
                {
                    this.passwordMap[i] := A_Index
                    this.dynamicPassArr[A_Index] := [i, ini.Read(A_Index, 0, "passwordSort")]
                    ; if !IsNumber(this.dynamicPassArr[A_Index][2])
                    ; ini.Write(this.dynamicPassArr[A_Index][2] := 0, A_Index, "passwordSort")
                }
            }

            this.fileSystemObject := ComObject("Scripting.FileSystemObject")

        } else
        {
            arr := [loopPath]
            SplitPath(loopPath, , &dir)
            SetWorkingDir(dir)
        }

        for i in arr
        {
            if !loopPath && A_WorkingDir != this.defaultDir
                SetWorkingDir(this.defaultDir)
            if !loopPath
                this.index := A_Index

            this.temp := tmpDir := '__7z' A_Now

            this.currentSize := FileGetSize(i)
            hideBool := this.currentSize / 1024 / 1024 < this.hideRunSize

            part := IsPart(i)
            if this.partSkip && !part
                continue

            if this.muilt && !this.guiShow && !hideBool
                this.Gui()

            if this.addDir2Pass
                SplitPath(i, , &dir), this.password.Push(RegExReplace(dir, ".+\\"))
            zipx(i)
            if this.addDir2Pass
                this.password.RemoveAt(this.password.Length)

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

            ;解压后没有文件
            if !IsSet(count)
            {
                this.RecycleItem(tmpDir, A_LineNumber, true)
                continue
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

                outFile := this.MoveItem(souceFile, A_WorkingDir "\" name, isDir, A_LineNumber)

                this.RecycleItem(tmpDir, A_LineNumber, true)

                if !this.nesting || !this.nestingMuilt
                    continue

                if !isDir
                {
                    if this.nesting
                        UnZipNesting(outFile, ext)
                } else if this.nestingMuilt
                    loop files outFile "\*.*", "F"
                        UnZipNesting(A_LoopFileFullPath, A_LoopFileExt)

            } else	;多个文件
            {
                SplitPath(i, , , , &nameNoEXT)
                outFile := this.MoveItem(tmpDir, A_WorkingDir "\" nameNoEXT, 1, A_LineNumber)

                if this.nestingMuilt
                    loop files outFile "\*.*", "F"
                        UnZipNesting(A_LoopFileFullPath, A_LoopFileExt)
            }
        }

        if loopPath
            return

        if this.autoRemovePass && (this.dynamicPassSort || this.autoAddPass)
        {
            if this.dynamicPassSort
                PasswordSort

            for i in this.dynamicPassArr
                if A_Index > this.autoRemovePass
                    PasswordClear(A_Index, true)

            if this.dynamicPassArr.Length > this.autoRemovePass
                this.dynamicPassArr.RemoveAt(this.autoRemovePass + 1, this.dynamicPassArr.Length - this.autoRemovePass)
        } else if this.dynamicPassSort
            PasswordSort

        ;执行解压
        zipx(path)
        {
            if this.logLevel
                this.log .= '`n#####`n' path '`n'

            pass := ""
            this.continue := false

            for i in this.password
            {
                if A_Index = 1
                {
                    this.isFile := this.isCmdReturn := false
                    this.needPass := 4
                    cmdArgs := this.7z ' l -slt -bsp1  "' path '"'
                    if this.cmdLog
                        this.testLog .= '`n#####`n' cmdArgs '`n'
                    this.RunCmd(cmdArgs, , CheckEncrypted)

                    switch this.needPass
                    {
                        case 0: break
                        case 1: continue
                        case 2:
                        {
                            for n in this.password
                            {
                                if n
                                {
                                    this.isFile := this.isCmdReturn := false
                                    this.needPass := 5
                                    cmdArgs := this.7z ' l -slt -bsp1 -p"' n '" "' path '"'
                                    this.RunCmd(cmdArgs, , CheckEncrypted)
                                    if this.cmdLog
                                        this.testLog .= '`n#####`n' cmdArgs '`n'
                                    if this.needPass = 1
                                    {
                                        this.error := 0
                                        pass := ' -p"' AddPass(n) '"'
                                        break 2
                                    }
                                }
                            }
                        }
                        case 3: return
                    }
                }

                this.CheckCMD(, this.7z ' t -bsp1 "' path '" -p"' i '"')

                if this.continue
                    return

                if !this.error
                {
                    if i
                        pass := ' -p"' AddPass(i) '"'
                    break
                }
            }

            if !this.needPass || !this.error	;密码正确或无需密码
            {
                this.Run7z(hideBool, 'x', path, '" -aou -o' tmpDir pass this.excludeArgs this.codePage, hideBool || this.guiShow, () => IsSuccess(), A_LineNumber)

                if IsSuccess()
                {
                    if part != -1
                        return
                    if loopPath
                        this.RecycleItem(path, A_LineNumber, true)
                    else if this.delSource || (pass && this.delWhenHasPass)
                        this.RecycleItem(path, A_LineNumber)
                }
            } else
            {
                this.tryPasssword := ""
                if this.autoAddPass
                    SetTimer(TrackPass, 10)
                this.Run7z(false, 'x', path, '" -aou -o' tmpDir this.excludeArgs this.codePage, , () => IsSuccess(), A_LineNumber)
                SetTimer(TrackPass, 0), ToolTip()

                if this.tryPasssword && this.exitCode != 255
                {
                    this.error := true
                    this.CheckCMD(, this.7z ' t -bsp1 "' path '" -p"' this.tryPasssword '"')
                    if !this.error
                        AddPass(this.tryPasssword)
                }

                if IsSuccess() && part = -1 && this.delWhenHasPass
                    this.RecycleItem(path, A_LineNumber)	;密码错误需手动输入密码
            }

            TrackPass()
            {
                title := "ahk_pid " this.pid
                if WinExist(title) && WinActive(title)
                {
                    try
                        if InStr(ControlGetText("Button1", title), "&S")
                        {
                            if !ControlGetChecked("Button1", title)
                                return (this.tryPasssword := "", ToolTip())
                        } else if !ControlGetText("Static14", title)
                            return (SetTimer(TrackPass, 0), ToolTip())

                    try
                        if (str := ControlGetText("Edit1", title))
                            this.tryPasssword := str

                    try
                        if ControlGetText("Static14", title)
                            this.tryPasssword := ""

                    if this.tryPasssword
                        ToolTip "当前密码 : " this.tryPasssword
                } else
                    ToolTip()
            }

            AddPass(pass)
            {
                static notLoopPass := ""	;用以确保嵌套和源文件同密码时不会重复记录

                if !pass
                    return

                if !loopPath
                    notLoopPass := pass
                else if pass = notLoopPass
                    return pass

                if ini.lastPass != pass
                    ini.Write(pass, "lastPass", "temp")

                if this.dynamicPassSort || this.autoAddPass
                {
                    if !this.passwordMap.Has(pass) && this.password.Length > 2 && this.password[3] = pass
                        return pass

                    if !this.passwordMap.Has(pass)
                    {
                        this.dynamicPassArr.Push([pass, 0]), this.passwordMap[pass] := this.dynamicPassArr.Length
                        if this.autoAddPass
                            ini.Write(pass, this.passwordMap.Count, "password")
                    } else
                        this.dynamicPassArr[this.passwordMap[pass]][2]++
                }
                return pass
            }

            CheckEncrypted(LineNum, Line)
            {
                if this.isCmdReturn
                    return

                if this.cmdLog
                    this.testLog .= "[" LineNum "] " line '`n'

                if !this.isFile && InStr(Line, "Attributes = A") || Line ~= "CRC = [A-Z0-9]+"
                    this.isFile := true
                else if this.isFile && InStr(Line, "Attributes = D") || Line ~= "CRC = *?$"
                    this.isFile := false
                else if this.isFile && InStr(Line, "Encrypted = -")
                    LogAndReturn(0, A_LineNumber)

                else if InStr(Line, "Encrypted = +") || InStr(Line, "Wrong password?")
                    LogAndReturn(1, A_LineNumber)
                else if InStr(Line, "Enter password (will not be echoed):")
                    LogAndReturn(2, A_LineNumber)
                else if this.needPass = 5 && InStr(Line, "Errors: 1")
                    LogAndReturn(2, A_LineNumber)
                else if InStr(Line, "Errors: 1") || InStr(Line, "Cannot open the file as archive") || InStr(Line, "Unexpected end of archive")
                    this.continue := true, LogAndReturn(3, A_LineNumber)

                LogAndReturn(num := "", logLineNum := "")
                {
                    this.isCmdReturn := true
                    this.needPass := num
                    ProcessClose(this.CMDPID), ProcessWaitClose(this.CMDPID)
                    this.CMDPID := 0
                    this.Loging(cmdArgs '`n[' LineNum '] ' line, logLineNum, this.needPass > 2 ? 3 : 4)
                }
            }

            IsSuccess()
            {
                if !this.exitCode
                    return true

                if !DirExist(tmpDir)
                    return false

                if this.exitCode != 255
                {
                    folderSize := this.fileSystemObject.GetFolder(tmpDir).Size
                    this.Loging("文件大小: " this.currentSize " 临时文件夹大小: " folderSize, A_LineNumber)

                    if folderSize >= this.currentSize
                        return true
                    else if folderSize / this.currentSize * 100 > this.succesSpercent
                        return true
                }

                this.RecycleItem(tmpDir, A_LineNumber, true)
                return false
            }
        }

        ;解压嵌套
        UnZipNesting(path, ext)
        {
            if !this.IsArchive(ext) || !(part := IsPart(path))
                return

            timeSave := FileGetTime(path), sizeSave := FileGetSize(path)
            this.exitCode := -1
            this.Unzip(path)
            this.Loging("解压嵌套 <--> " path, A_LineNumber)

            if !this.exitCode && part = -1 && FileExist(path) && FileGetTime(path) = timeSave && FileGetSize(path) = sizeSave	;!exitCode &&
                this.RecycleItem(path, A_LineNumber)
        }

        ; 解压后处理
        AfterUnzip(path)
        {
            static isRead := false, obj := { rename: { ext: Map(), name: Map(), exp: Map() },
                deleteExp: [] }

            if !isRead
            {
                ini.ReadLoop("renameExt", obj.rename.ext, , true)
                ini.ReadLoop("renameName", obj.rename.name, , true)
                ini.ReadLoop("renameExp", obj.rename.exp, , true)
                ini.ReadLoop("deleteExp", obj.deleteExp)

                isRead := true
            }

            if (isDir := DirExist(path)) && this.fileSystemObject.GetFolder(path).Size = 0	;空文件夹
                return this.RecycleItem(path, A_LineNumber)

            SplitPath(path, &name, &dir, &ext, &nameNoExt)

            for i in obj.deleteExp
                if name ~= i
                    return this.RecycleItem(path, A_LineNumber)

            for ori, out in obj.rename.ext
            {
                if !isDir && ext = ori
                {
                    name := nameNoExt '.' out
                    break
                }
            }

            for needle, replaceText in obj.rename.name
                if InStr(name, needle)
                    name := StrReplace(name, needle, replaceText)

            for needle, replaceText in obj.rename.exp
                if name ~= needle
                    name := RegExReplace(name, needle, replaceText)

            if path != dir "\" name
                this.MoveItem(path, dir "\" name, isDir, A_LineNumber)
        }

        IsPart(path)
        {
            SplitPath(path, &name)

            if name ~= "i)\.part\d\.rar$"
            {
                if InStr(name, ".part1.rar")	;第一卷
                    return 1
                this.Loging("可能是分卷包  <--> " path, A_LineNumber, 5)
                return 0
            } else if name ~= "\..+\.\d+$"
            {
                if name ~= "\..+.0+1"	;第一卷
                    return 1
                this.Loging("可能是分卷包  <--> " path, A_LineNumber, 5)
                return 0
            }
            return -1
        }

        PasswordSort()
        {
            ;排序
            i := 0
            while (++ i <= this.dynamicPassArr.Length)
            {
                j := 0
                while (++ j <= this.dynamicPassArr.Length - i)
                {
                    if this.dynamicPassArr[j][2] < this.dynamicPassArr[j + 1][2]
                    {
                        temp := this.dynamicPassArr[j]
                        this.dynamicPassArr[j] := this.dynamicPassArr[j + 1]
                        this.dynamicPassArr[j + 1] := temp
                    }
                }
            }

            for i in this.dynamicPassArr
            {
                if ini.Read("A_Index", , "passwordSort") != i[2]
                    ini.Write(i[2], A_Index, "passwordSort")

                if ini.Read("A_Index", , "password") != i[1]
                    ini.Write(i[1], A_Index, "password")
            }

            arr := []
            loop	;清除重复密码
            {
                if (pwd := ini.Read(this.dynamicPassArr.Length + A_Index))
                {
                    if !this.passwordMap.Has(pwd)
                        arr.Push(pwd)
                    PasswordClear(this.dynamicPassArr.Length + A_Index, true)
                } else
                    break
            }

            ; if arr.Length
            ;     for i in arr
            ;         ini.Write(i, this.dynamicPassArr.Length + A_Index, "password")
        }

        PasswordClear(index, delete := false)
        {
            if delete
                IniDelete(ini, "password", index), IniDelete(ini, "passwordSort", index)
            else
                ini.Write(, index, "password"), ini.Write(, index, "passwordSort")
        }

        FormatPassword(str) => StrLen(str) < 100 ? Trim(RegExReplace(str, "(\R*)")) : ""	;移除所有换行符及首尾所有空格或制表符
    }

    OpenZip()
    {
        SplitPath(this.arr[1], , &dir, &ext, &nameNoExt)

        if !this.muilt
        {
            extForOpen := Map()
            ini.ReadLoop("extForOpen", extForOpen, true)

            path := ' "' this.arr[1] '" '

            If DirExist(this.arr[1])
                this.error := true
            else if (this.IsArchive(ext) || extForOpen.Has(ext))
                this.error := false
            else
            {
                if this.logLevel
                    this.log .= '`n#####`n' this.arr[1] '`n'
                this.CheckCMD("openZip", this.7z ' l ' path)
            }

            if !this.error
            {
                Run(this.7zFM path, , , &pid)
                this.Loging("打开 <--> " path, A_LineNumber)
            }
        }

        if this.muilt || this.error
        {
            args := ini.openAdd
            path := ""
            for i in this.arr
                path .= ' "' i '" '
            zipName := this.muilt ? StrReplace(RegExReplace(A_WorkingDir, ".+\\"), ":") : DirExist(this.arr[1]) ? RegExReplace(this.arr[1], ".+\\") : nameNoExt
            ext := RegExReplace(args, '(.+?)".*', "$1")
            this.Run7z(, 'a', this.AUO(zipName, ext), args path, , , A_LineNumber)
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

        args := ini.add
        ext := RegExReplace(args, '(.+?)".*', "$1")

        if count = this.arr.Length	;全是文件夹,单独添加
        {
            for i in this.arr
            {
                hideBool := IsHide(i)
                if count > 1 && !hideBool && !this.guiShow
                    this.Gui
                zipName := this.AUO(RegExReplace(i, ".*\\"), ext)
                this.temp := zipName ext
                this.index := A_Index
                this.Run7z(hideBool, 'a', zipName, args ' "' i '\*"', hideBool || count > 1, , A_LineNumber)

            }
            return

        } else if this.arr.Length = 1	;单个文件
            this.Run7z(hideBool, 'a', this.AUO(nameNoExt, ext), args ' "' this.arr[1] '"', hideBool, , A_LineNumber)
        else	;文件文件夹混合
        {
            for i in this.arr
                path .= ' "' i '" '
            this.Run7z(hideBool, 'a', this.AUO(RegExReplace(A_WorkingDir, ".+\\"), ext), args path, hideBool, , A_LineNumber)
        }

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

    Gui()
    {
        this.guiShow := true
        DetectHiddenWindows(1)

        g := Gui("+LastFound")
        DllCall("RegisterShellHookWindow", "UInt", WinExist())
        msgNum := DllCall("RegisterWindowMessage", "Str", "SHELLHOOK")
        OnMessage(msgNum, ShellMessage)

        g.SetFont(, "Segoe UI")
        g.BackColor := "FFFFFF"

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

        sub() => "ahk_pid " this.pid
        g.Show("AutoSize")

        Close(*)
        {
            if ProcessExist(this.pid)
                ProcessClose(this.pid), ProcessWaitClose(this.pid)
            if this.HasOwnProp("temp")
                this.RecycleItem(this.temp, A_LineNumber, true)
            if !this.setShow
                ExitApp(255)
            OnMessage(msgNum, ShellMessage, 0), g.Destroy()
        }

        ButtonShowHide(GuiCtrlObj, *)
        {
            DetectHiddenWindows(0)
            if !ProcessExist(this.pid)
                return

            if WinExist(sub())
                WinHide(sub()), GuiCtrlObj.Text := "显示原始界面"
            else
                WinShow(sub()), WinActivate(sub()), GuiCtrlObj.Text := "隐藏原始界面"
        }

        ButtonPause(GuiCtrlObj, Info)
        {
            DetectHiddenWindows(1)
            if !WinExist(sub())
                return

            textSave := GuiCtrlObj.Text
            ControlClick("Button2", sub())
            ShellMessage
            while WinExist(sub()) && GuiCtrlObj.Text = textSave
                ControlClick("Button2", sub()), Sleep(500), ShellMessage()
        }

        ButtonCance(GuiCtrlObj, Info)
        {
            DetectHiddenWindows(1)
            if !WinExist(sub())
                return
            ControlClick("Button3", sub())
            ShellMessage
            while WinExist(sub()) && !InStr(WinGetText(sub()), "否(&N)")
                ControlClick("Button3", sub()), Sleep(500)

            WinWaitActive(sub())
            WinWaitClose
            num := 0
            loop 5
                if !WinExist(sub())
                    num++

            if num = 5
                Close
        }

        DetectError()
        {
            DetectHiddenWindows(0)
            if WinExist(sub())
                return
            else
                WinShow(sub())
        }

        ShellMessage(wParam := 6, *)
        {
            ListLines(0)
            DetectHiddenWindows(1)
            static timeSave := A_TickCount

            if !this.isRunning
                Close()

            if A_TickCount - timeSave < 50 || wParam != 6 || !WinExist(sub())
                return

            IsChanged(总进度1, "总进度:")
            , IsChanged(总进度2, this.index "\" this.arr.Length)
            try
            {
                if g.Title != WinGetTitle(sub())
                    g.Title := WinGetTitle(sub())

                arr := StrSplit(WinGetText(sub()), "`n")
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

                IsChanged(总大小1, arr[8])
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

                if IsNumber(arr[index++ ])	;发生错误
                    DetectError()	; IsChanged(发生错误2, arr[index -1])
                else
                    index--

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
            ListLines(1)
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

    Run7z(is7z := false, xa := "x", path := "", args := "", hide := false, log := true, linenum := "")
    {
        this.pid := ""
        this.cmdHide := false
        if !is7z
            SetTimer(WinGetPID, 10)
        else if !this.guiShow
            this.cmdHide := true
        this.exitCode := RunWait((is7z ? this.7z : this.7zG) ' ' xa ' "' path args, , hide ? "hide" : "")

        SetTimer(WinGetPID, 0)
        if log
            this.Loging('[' this.exitCode '] ' (xa = 'x' ? "解压" : "压缩") " <--> " path, linenum)

        WinGetPID()
        {
            DetectHiddenWindows(1)
            static winmgmts := ComObjGet("winmgmts:")

            WinWait("ahk_exe 7zG.exe", , 3)
            winmgmts.ExecQuery('Select * from Win32_Process where Name="7zG.exe" and CommandLine like "%' StrReplace(path, "\", "\\") '%"')._NewEnum()(&proc)
            if (this.pid := IsSet(proc) ? proc.ProcessID : "")
            {
                if this.to = "x" && this.excludeArgs
                {

                    while (!GetSize())
                    {
                        if A_TickCount - this.now > 1000
                            break
                    }
                    if RegExMatch(GetSize(), "(.+) MB$", &size)
                        this.currentSize := size[1] * 1024 * 1024
                }
                SetTimer(WinGetPID, 0)
            }

            GetSize()
            {
                size := ""
                try
                    size := ControlGetText("Static15", "ahk_pid " this.pid)
                return size
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

    AUO(name, ext) => StrReplace(this.PathDupl(A_WorkingDir "\" name ext, 0), ext)

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
    RunCmd(CmdLine, Codepage := "CP0", fn := "") {
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
                Line := File.ReadLine(), sOutput .= fn ? Fn.Call(LineNum++ , Line) : this.CheckCMD(LineNum++ , Line)

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
        static check, checkSave := "", whatSave := what, cmdArgs

        if Type(what) = "String"
        {
            if !checkSave
            {
                ; checkSave := { error: Map(), errorExp: Map(), errorrContinueExp: Map(), success: Map(), successExp: Map() }
                checkSave := { error: Map(), errorExp: Map(), success: Map(), successExp: Map() }
                ini.ReadLoop(what "CheckError", checkSave.error, , true)
                ini.ReadLoop(what "CheckErrorExp", checkSave.errorExp, , true)
                ; ini.ReadLoop(what "CheckErrorContinueExp", checkSave.errorrContinueExp, , true)
                ini.ReadLoop(what "CheckSuccess", checkSave.success, , true)
                ini.ReadLoop(what "CheckSuccessExp", checkSave.successExp, , true)
            }
            check := {}
            for i in checkSave.OwnProps()
                check.%i% := checkSave.%i%.Clone()

            this.isCmdReturn := false
            this.error := true
            cmdArgs := line
            if this.cmdLog
                this.testLog .= '`n#####`n' cmdArgs '`n'

            this.RunCmd(cmdArgs)

        } else if line
        {
            if this.isCmdReturn
                return

            if this.cmdLog
                this.testLog .= "[" what "] " line '`n'

            ; for i in check.errorrContinueExp
            ;     if line ~= i && --check.errorrContinueExp[i] < 1
            ;         return (this.continue := true, LogAndReturn(1, A_LineNumber))

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

            if whatSave = "unZip" && (line ~= "^ +[1-9]+%") && !(line ~= "^ +[1-9]+%.+Open$" || line ~= "^ +[1-9]+%$")
                return LogAndReturn(6, A_LineNumber)

            LogAndReturn(num, lineNum)
            {
                this.isCmdReturn := true
                ProcessClose(this.CMDPID), ProcessWaitClose(this.CMDPID)
                this.CMDPID := 0
                if num < 4
                    this.error := 1
                else
                    this.error := 0
                this.Loging(cmdArgs "`n[" what '] ' line, lineNum, this.error ? 3 : 4)
            }
        }
    }

    ; 关闭0/删除1/重命名2/命令行错误3/命令行正确4/其他5
    Loging(log, lineNum, level := 5)
    {
        if !this.logLevel || level > this.logLevel
            return

        switch level
        {
            case 5: msg := "其他"
            case 4: msg := "命令行正确"
            case 3: msg := "命令行错误"
            case 2: msg := "重命名"
            case 1: msg := "删除"
        }

        this.log .= Format("[{1}] [{2}] [{3}ms] [{4}]`n{5}", msg, lineNum, A_TickCount - this.now, FormatTime(A_Now, "yyyy/M/d H:m:s"), log) '`n'
        this.now := A_TickCount
    }
}

Setting()
{
    static hwnd := ""

    keyPathForAll := "HKCU\SOFTWARE\Classes\AllFilesystemObjects\shell"
    keyPathForFile := "HKCU\SOFTWARE\Classes\*\shell"
    sendToDir := A_StartMenu "/../SendTo/"
    openZipLnk := sendToDir ini.openZipName ".lnk"
    unZipNameLnk := sendToDir ini.unZipName ".lnk"
    addZipNameLnk := sendToDir ini.addZipName ".lnk"
    unZipCPNameLnk := sendToDir ini.unZipCPName ".lnk"

    if zip.setShow
        return WinActivate(app " ahk_class AutoHotkeyGUI")

    if WinExist(app " ahk_class AutoHotkeyGUI")
        WinActivate(), ExitApp(0)

    if WinExist(hwnd)
        return WinActivate()

    zip.setShow := true
    set := Gui(, app)
    set.isChange := false
    set.SetFont("s11", "Segoe UI")

    var := { password: [],
        ext: [],
        extExp: [],
        extForOpen: [],
        renameExt: [],
        renameName: [],
        renameExp: [],
        excludeExt: [],
        excludeName: [],
        deleteExp: [] }
    ini.ReadLoop("password", var.password)
    passwordMap := Map()
    for i in var.password
        passwordMap[i] := ini.Read(A_Index, 0, "passwordSort")
    ini.ReadLoop("ext", var.ext)
    ini.ReadLoop("extExp", var.extExp)
    ini.ReadLoop("extForOpen", var.extForOpen)
    ini.ReadLoop("renameExt", var.renameExt)
    ini.ReadLoop("renameName", var.renameName)
    ini.ReadLoop("renameExp", var.renameExp)
    ini.ReadLoop("excludeExt", var.excludeExt)
    ini.ReadLoop("excludeName", var.excludeName)
    ini.ReadLoop("deleteExp", var.deleteExp)

    Tab := set.AddTab3(, ["主要", "处理", "其他", "自定义", "关联", "关于"])
    Tab.OnEvent("Change", TabChange)

    lineGeneration
    GuiEdit("zipDir", "7-zip路径", ini.zipDir, "%SmartZipDir% 为相对路径`n其代表 SmartZip 所在文件夹,不包括最后的 \", "Section")
    GuiEdit("targetDir", "解压路径", ini.targetDir, "为空时默认为当前文件夹")
    pwdList := GuiComboBox("密码列表", var.password)
    pwdList.OnEvent("Change",(ctrl,info)=> ctrl.ToolTip := passwordMap.Has(ctrl.Text) ? "当前密码使用次数 : " passwordMap[ctrl.Text] : "密码列表")

    lineGeneration("xs")
    GuiCheckBox("nesting", ini.nesting, "解压嵌套压缩包", "解压成功删除源文件,只针对单文件")
    GuiCheckBox("nestingMuilt", ini.nestingMuilt, "解压嵌套文件夹", "只检查第一层文件夹,解压成功删除源文件", "x+170 yp")
    GuiCheckBox("delSource", ini.delSource, "解压后删除源文件", "仅在解压成功时删除")
    GuiCheckBox("delWhenHasPass", ini.delWhenHasPass, "仅删除包含密码的源文件", "不需要选中 解压后删除源文件", "yp x+90")

    GuiCheckBox("autoAddPass", ini.autoAddPass, "自动添加密码", "在7-Zip输入密码框选中显示密码保存")
    GuiCheckBox("dynamicPassSort", ini.dynamicPassSort, "密码动态排序", "把使用次数最多的排在前面")
    GuiUpDownEdit("autoRemovePass", "删除非常用密码", ini.autoRemovePass, , "为0时禁用`n密码总数超过值的部分会被删除`n需要启用上面两项任意一项")
    pwdImportBtn := set.AddButton("xs", "从文本文件中导入密码")
    pwdImportBtn.OnEvent("Click", pwdImport)
    pwdImportBtn.ToolTip := "也可导入旧版本ini密码"

    Tab.UseTab(2)
    lineGeneration("Section")
    GuiComboBox("排除后缀名", var.excludeExt, "不解压后缀名为此的文件")
    GuiComboBox("排除文件名", var.excludeName, "不解压文件名包含此的文件")
    GuiComboBox("删除正则", var.deleteExp, "在解压后删除符合的文件")
    lineGeneration("xs")
    GuiComboBox("改名后缀名", var.renameExt, "示例:mp+3<--->mp3,将后缀名为mp+3的替换为mp3")
    GuiComboBox("改名文件名", var.renameName, "示例6666<--->将文件名6666的部分替换为空")
    GuiComboBox("改名正则", var.renameExp, "示例:^[ 	]+<--->,将文件名前面的空格替换为空")
    set.AddText()
    set.AddText("xs", "改名部分必需包含 <---> `n符号前面是搜索词,后面是替换词")

    Tab.UseTab(3)
    lineGeneration()
    GuiCheckBox("partSkip", ini.partSkip, "跳过分卷压缩包", "第一卷会被解压,其他的跳过`n分卷不会自动删除", "Section")
    GuiCheckBox("test", ini.test, "启用测试中的功能", "当前没有测试中功能")
    GuiCheckBox("cmdLog", ini.cmdLog, "启用测试日志", "检查文件时的测试日志,与下文的日志等级无关")
    lineGeneration("xs")
    GuiUpDownEdit("logLevel", "日志等级", ini.logLevel, 5, "关闭0/删除1/重命名2/命令行错误3/命令行正确4/其他5")
    GuiUpDownEdit("hideRunSize", "隐藏运行", ini.hideRunSize, , "源文件大小(单位 MB)小于此值的会隐藏运行", "xs")
    GuiUpDownEdit("successPercent", "判断解压成功百分比", ini.successPercent, 100, "部分文件可能解压后大小会小于源文件`n只要解压到一定百分比就判断解压成功", "xs")
    set.AddText("xs")
    if FileExist(A_ScriptDir "\log.txt")
        set.AddButton("", "查看日志").OnEvent("Click", (*) => Run(A_ScriptDir "\log.txt"))
    lineGeneration("xs")
    if FileExist(A_ScriptDir "\cmdLog.txt")
        set.AddButton(, "查看测试日志").OnEvent("Click", (*) => Run(A_ScriptDir "\cmdLog.txt"))
    lineGeneration("xs")
    set.AddButton("yp x400", "更多设置").OnEvent("Click", (*) => Run(A_ScriptDir "\SmartZip.ini"))

    Tab.UseTab(4)
    lineGeneration
    GuiEdit("icon", "图标路径", ini.icon, "为空则无图标, %SmartZipDir% 为相对路径", "Section")
    GuiEdit2("openAdd", "openZip参数", ini.openAdd, "当打开非压缩包会弹出新建压缩包的窗口,此为其参数")
    GuiEdit2("add", "addZip参数", ini.add, "压缩时的默认参数,两项都不能修改文件名`nzip为其后缀名可修改,不要移除引号")
    GuiEdit2("openZipName", "openZip名称", ini.openZipName)
    GuiEdit2("unZipName", "unZip名称", ini.unZipName)
    GuiEdit2("addZipName", "addZip名称", ini.addZipName)
    GuiEdit2("unZipCPName", "unZipCP名称", ini.unZipCPName)

    Tab.UseTab(5)
    lineGeneration
    GuiComboBox("格式", var.ext, "此格式会被当作压缩包打开或解压", "Section")
    GuiComboBox("格式正则", var.extExp, "此格式会尝试解压,`n举例: ^\d+$  代表后缀名由纯数字组成")
    GuiComboBox("打开格式", var.extForOpen, "此格式会当作压缩包打开")
    lineGeneration("xs")
    set.Add("GroupBox", "Section  w200 h220", "右键菜单").GetPos(&x, &y, &w, &h)
    o1 := set.AddCheckbox("xs+10 ys+20 r1.5 Checked" IsContextMenuVisible("openZip"), ini.openZipName)
    x1 := set.AddCheckbox("r1.5 Checked" IsContextMenuVisible("unZip"), ini.unZipName)
    a1 := set.AddCheckbox("r1.5 Checked" IsContextMenuVisible("addZip"), ini.addZipName)
    xc1 := set.AddCheckbox("r1.5 Checked" IsContextMenuVisible("unZipCP"), ini.unZipCPName)
    set.AddButton(" y+10", "注册").OnEvent("Click", (*) => ContextMenu(x1.Value, o1.Value, a1.Value, xc1.Value))
    set.Add("GroupBox", " Section  w200 h220 x" x + w + 30 ' y' y, "发送到菜单")
    o2 := set.AddCheckbox("xs+10 ys+20 r1.5 Checked" IsSendToVisible(openZipLnk, "o"), ini.openZipName)
    x2 := set.AddCheckbox("r1.5 Checked" IsSendToVisible(unZipNameLnk, "x"), ini.unZipName)
    a2 := set.AddCheckbox("r1.5 Checked" IsSendToVisible(addZipNameLnk, "a"), ini.addZipName)
    xc2 := set.AddCheckbox("r1.5 Checked" IsSendToVisible(unZipCPNameLnk, "xc"), ini.unZipCPName)
    sendToBtn := set.AddButton(" y+10", "注册")
    sendToBtn.OnEvent("Click", (*) => SendTo(x2.Value, o2.Value, a2.Value, xc2.Value))
    sendToBtn.ToolTip := "如要修改菜单名称请在修改前先卸载(取消所有选中注册)`n否则可能出现多个菜单"

    Tab.UseTab(6)
    set.AddText()
    set.AddText("", app " " MainVersion " (" buildVersion ")")
    lineGeneration
    set.AddText("", "修改时间 " buileTime)
    set.AddText()
    set.AddText(, "相关链接")
    set.AddLink(, '<a id="GitHub" href="https://github.com/vvyoko/SmartZip">GitHub</a>')
    set.AddLink("yp", '<a id="7-zip" href="https://www.7-zip.org/">7-zip</a>').ToolTip := "测试基于 21.07 版本"
    set.AddLink("yp", '<a id="AutoHotkey" href="https://www.autohotkey.com/">AutoHotkey</a>')
    set.AddLink("yp", '<a id="GitHub2" href="https://github.com/vvyoko/SmartZip/issues/new">建议反馈</a>').ToolTip := "在GitHub上新建issues建议或反馈"
    set.AddLink("yp", '<a id="GitHub2" href="https://meta.appinn.net/t/topic/33555">论坛反馈</a>').ToolTip := "也可选择在小众软件反馈"
    donateBtn := set.AddButton("y+80 x+50", "支持作者")
    donateBtn.OnEvent("Click", (*) => Donate())
    donateBtn.ToolTip := "如软件对您有帮助,可考虑捐助"

    Tab.UseTab()
    set.AddButton("", "取消").OnEvent("Click", Close)
    reloadBtn := set.AddButton("yp x+10", "重载")
    reloadBtn.OnEvent("Click", (*) => Reload())
    reloadBtn.ToolTip := "重载应用以查看修改,`n导入密码需新增一个密码或重载才能在列表中显示(不影响保存)`n修改菜单名称需要重载才能在关联时修改"
    set.AddButton("yp x385", "应用").OnEvent("Click", Apply)
    set.AddButton("yp x+10 default", "确定").OnEvent("Click", Apply)
    SB := set.AddStatusBar("", "拖入文件触发智能解压")
    set.Show()
    hwnd := set.Hwnd
    set.OnEvent("Close", Close)
    set.OnEvent("Escape", Close)
    set.OnEvent("DropFiles", CallX)
    OnMessage(0x200, WM_MOUSEMOVE)

    Close(*)
    {
        if zip.HasOwnProp("temp")
            zip.RecycleItem(zip.temp, A_LineNumber, true)
        ExitApp(0)
    }

    CallX(GuiObj, GuiCtrlObj, FileArray, X, Y)
    {
        set.Opt("+OwnDialogs")
        if !zip.HasProp("7z")
            return SB.Text := "设置未完成,请设置完成后重启应用再尝试"

        if zip.HasProp("isRunning") && zip.isRunning
            return SB.Text := "正在处理中,请等待结束后再拖入"
        SB.Text := "处理中..."
        zip.Init(FileArray).Exec()
        SB.Text := "处理完成"
    }

    Apply(GuiCtrlObj, Info)
    {
        v := set.Submit(false)
        for i, n in v.OwnProps()
            if ini.%i% != n
                ini.setWrite(i, n)

        for section in var.OwnProps()
        {
            loop
            {
                if !ini.Read(A_Index, , section)
                    break
                if A_Index > var.%section%.Length
                    ini.Delete(section, A_Index)
            }
            for n in var.%section%
                if ini.Read(A_Index, "", section) != n
                    ini.Write(n, A_Index, section)
        }

        for i in var.password
        {
            if passwordMap.Has(i)
            {
                if passwordMap[i] != ini.Read(A_Index, , "passwordSort")
                    ini.Write(passwordMap[i], A_Index, "passwordSort")
            } else
                ini.Write(0, A_Index, "passwordSort")
        }

        loop
        {
            if IsNumber(ini.Read(var.password.Length + A_Index, , "passwordSort"))
                ini.Delete("passwordSort", var.password.Length + A_Index)
            else
                break
        }

        set.isChange := true
        if GuiCtrlObj.Text = "确定"
            ExitApp
    }

    pwdImport(*)
    {
        set.Opt("+OwnDialogs")
        MsgBox("每行一个密码,请确保文件编码为UTF-8`n`n为防止编码错误,第一行不导入`n`n请在第一行写入中文测试文字,`n`n接下来请选择保存密码的文件或旧版本ini")
        path := FileSelect("1", , app, "*.txt;*.ini")

        if path
        {
            SplitPath(path, , , &ext)
            if ext = "ini"
            {
                arr := []
                loop
                {
                    if !(pwd := IniRead(path, "password", A_Index, ""))
                        break
                    arr.Push(pwd)
                }
            } else
            {
                arr := StrSplit(FileRead(path, "UTF-8"), ["`n", "`r`n"])
                if MsgBox("文件第一行的内容为此吗?`n`n" arr[1] "`n`n乱码点否退出导入,确定点是继续导入", app, "YN") = "No"
                    return
                arr.RemoveAt(1)
            }

            lengthSave := var.password.Length
            for i in arr
            {
                if !i
                    continue
                isdp := false
                for n in var.password
                {
                    if i = n
                    {
                        isdp := true
                        break
                    }
                }
                if !isdp
                    var.password.Push(i)
            }

            pwdList.Delete(), pwdList.Add(var.password)
            MsgBox("共导入" var.password.Length - lengthSave " 项,点击应用或确定以保存密码")
        }

    }

    TabChange(Ctrl, Info)
    {
        if set.isChange && Ctrl.Value = 5
        {
            if x1.Text != ini.openZipName
                x1.Text := x2.Text := ini.openZipName
            if o1.Text != ini.unZipName
                o1.Text := o2.Text := ini.unZipName
            if a1.Text != ini.addZipName
                a1.Text := a2.Text := ini.addZipName
            if xc1.Text != ini.unZipCPName
                xc1.Text := xc2.Text := ini.unZipCPName
            set.isChange := false
        }
    }

    GuiEdit(var, title, text := "", tips := "", opt := "xs")
    {
        set.AddText(opt, title)
        v := set.AddEdit("v" var " yp x150 w260", text)
        set.AddButton("w50 yp", "载入").OnEvent("Click", PathLoader)
        v.ToolTip := tips ? tips : title
        PathLoader(*)
        {
            isIcon := title = "图标路径"
            path := FileSelect(isIcon ? 1 : "D1", , app, isIcon ? "*.ico;*.exe" : "")
            if path
                v.Value := path
        }
    }

    GuiEdit2(var, title, text := "", tips := "", opt := "xs")
    {
        set.AddText(opt, title)
        v := set.AddEdit("v" var " yp x150 w260", text)
        v.ToolTip := tips ? tips : title
    }

    GuiCheckBox(var, checkeditem, title, tips := "", opt := "xs")
    {
        v := set.AddCheckbox(opt " r1.5 Checked" checkeditem " v" var, title)
        v.ToolTip := tips ? tips : title
        return v
    }

    GuiUpDownEdit(var, title, choose, rangeMax := 3600, tips := "", opt := "")
    {
        set.AddText(opt, title)
        v := set.AddEdit("v" var " Number yp w80 x250")
        set.AddUpDown("+0x80 Range0-" rangeMax, choose)
        v.ToolTip := tips ? tips : title
        return v
    }

    GuiComboBox(title, arr, tips := "", opt := "xs")
    {
        set.AddText(opt, title)
        v := set.AddComboBox("AltSubmit yp x150  w220", arr)
        set.AddButton("w40 yp", "增加").OnEvent("Click", ArrAdd)
        set.AddButton("w40 yp", "删除").OnEvent("Click", (*) => v.Text ? (arr.RemoveAt(v.Value), v.Delete(v.Value), v.Text := "") : "")
        v.ToolTip := tips ? tips : title
        return v

        ArrAdd(*)
        {
            If !v.Text
                return
            for i in arr
                if i = v.Text
                    return
            arr.push(v.Text), v.Delete(), v.Add(arr)
        }
    }

    IsSendToVisible(lnk, to)
    {
        if !FileExist(lnk) || (FileGetShortcut(lnk, &target, , &args), target != A_ScriptFullPath || args != to)
            return false
        return true
    }
    SendTo(x, o, a, xc)
    {
        set.Opt("+OwnDialogs")

        if x
            CreateLnk(unZipNameLnk, "x")
        else
            try FileDelete(unZipNameLnk)
        if o
            CreateLnk(openZipLnk, "o")
        else
            try FileDelete(openZipLnk)
        if a
            CreateLnk(addZipNameLnk, "a")
        else
            try FileDelete(addZipNameLnk)

        if xc
            CreateLnk(unZipCPNameLnk, "xc")
        else
            try FileDelete(unZipCPNameLnk)

        if x || o || a || xc
            MsgBox("已注册发送到菜单")
        else
            MsgBox("已删除发送到菜单")

        CreateLnk(lnk, to)
        {
            if !IsSendToVisible(lnk, to)
            {
                try
                    FileDelete(lnk)
                FileCreateShortcut(A_ScriptFullPath, lnk, , to)
            }
        }
    }

    IsContextMenuVisible(what)
    {
        if what = "UnZip" || what = "unZipCP"
            return InStr(RegRead(keyPathForFile "\" what "\command", , ""), A_ScriptDir "\Contextmenu") ? true : false
        else
            return InStr(RegRead(keyPathForAll "\" what "\command", , ""), A_ScriptDir "\Contextmenu") ? true : false
    }
    ContextMenu(x, o, a, xc)
    {
        set.Opt("+OwnDialogs")

        if x || o || a || xc
        {
            menuPath := A_ScriptDir "\Contextmenu"
            if FileExist(menuPath ".ahk")
                menuPath := '"' A_AhkPath '" "' menuPath '.ahk" '
            else if FileExist(menuPath ".exe")
                menuPath := '"' menuPath '.exe" '
            else
                MsgBox("右键菜单所需要文件不存在"), ExitApp(1)
            icon := FileExist(icon := RelativePath(ini.icon)) ? icon : ""
        }

        if x
        {
            RegWrite(icon, "REG_SZ", keyPathForFile "\UnZip", "Icon")
            RegWrite(ini.unZipName, "REG_SZ", keyPathForFile "\UnZip")
            RegWrite(menuPath "x", "REG_SZ", keyPathForFile "\UnZip\command")
        } else
            try RegDeleteKey(keyPathForFile "\UnZip")

        if o
        {
            RegWrite(icon, "REG_SZ", keyPathForAll "\OpenZip", "Icon")
            RegWrite(ini.openZipName, "REG_SZ", keyPathForAll "\OpenZip")
            RegWrite(menuPath "o", "REG_SZ", keyPathForAll "\OpenZip\command")
        } else
            try RegDeleteKey(keyPathForAll "\OpenZip")

        if a
        {
            RegWrite(icon, "REG_SZ", keyPathForAll "\AddZip", "Icon")
            RegWrite(ini.addZipName, "REG_SZ", keyPathForAll "\AddZip")
            RegWrite(menuPath "a", "REG_SZ", keyPathForAll "\AddZip\command")
        } else
            try RegDeleteKey(keyPathForAll "\AddZip")

        if xc
        {
            RegWrite(icon, "REG_SZ", keyPathForFile "\unZipCP", "Icon")
            RegWrite(ini.unZipCPName, "REG_SZ", "HKCU\SOFTWARE\Classes\*\shell\unZipCP")
            RegWrite(menuPath "xc", "REG_SZ", "HKCU\SOFTWARE\Classes\*\shell\unZipCP\command")
        } else
            try RegDeleteKey("HKCU\SOFTWARE\Classes\*\shell\unZipCP")

        if x || o || a || xc
            MsgBox("已注册右键")
        else
            MsgBox("已删除右键")
    }

    Donate()
    {
        static hwnd := 0
        if WinExist(hwnd)
            WinActivate
        wexin := A_Temp "\05330e88467ebffcb9b614d091ab1297e3396e063e28074c92c69e8eb36acc32"
        alipay := A_Temp "\dd64f3fcf719e9e263e0aa116e65aa1eb8ebb81ca2615841b23d7e1a902f10df"
        if !FileExist(wexin)
            FileInstall("donate\wexin.png", wexin)
        if !FileExist(alipay)
            FileInstall("donate\alipay.jpg", alipay)

        donateG := Gui("+ToolWindow +Owner" set.Hwnd, "捐助")
        donateG.SetFont(, "Segoe UI")
        donateG.AddText("", "如软件对您有帮助`n可扫描下方二维码通过微信或支付宝捐助`n感谢支持")
        ; donateG.AddLink("", '<a id="donate" href="https://github.com/vvyoko/SmartZip/blob/main/donate.md">支持者名单</a>').ToolTip := "如有隐私考虑可在备注上说明"
        donateG.AddText()
        donateG.AddPicture("w150 h-1", wexin)
        donateG.AddPicture("yp x+30 w150 h-1", alipay)
        donateG.AddText()
        donateG.Show()
        hwnd := donateG.Hwnd
        donateG.OnEvent("Escape", (*) => donateG.Destroy())
        donateG.OnEvent("Close", (*) => donateG.Destroy())
    }

    lineGeneration(opt := "") => set.AddText(opt " h1 BackgroundD8D8D8 w400")
}

WM_MOUSEMOVE(wParam, lParam, msg, Hwnd)
{
    ListLines(0)
    static PrevHwnd := 0
    if (Hwnd != PrevHwnd)
    {
        static PrevHwnd := 0
        static HoverControl := 0
        currControl := GuiCtrlFromHwnd(Hwnd)
        if (Hwnd != PrevHwnd)
        {
            Text := "", ToolTip()
            if CurrControl
            {
                if !CurrControl.HasProp("ToolTip")
                    return
                SetTimer(CheckHoverControl, 50)
                SetTimer(DisplayToolTip, -700)
            }
            PrevHwnd := Hwnd
        }
        return
        CheckHoverControl() => hwnd != prevHwnd ? (SetTimer(DisplayToolTip, 0), SetTimer(CheckHoverControl, 0)) : ""
        DisplayToolTip() => (ToolTip(CurrControl.ToolTip), SetTimer(CheckHoverControl, 0))
    }
    ListLines(1)
}

class ini
{
    static map := {
            zipDir: ["", "set"],
            icon: ["", "set"],
            targetDir: ["", "set"],
            delSource: [0, "set"],
            delWhenHasPass: [0, "set"],
            successPercent: [0, "set"],
            logLevel: [0, "set"],
            nesting: [0, "set"],
            nestingMuilt: [0, "set"],
            hideRunSize: [0x7FFFFFFFFFFFFFFF, "set"],
            cmdLog: [0, "set"],
            dynamicPassSort: [0, "set"],
            test: [0, "set"],
            partSkip: [0, "set"],
            autoRemovePass: [0, "set"],
            autoAddPass: [0, "set"],
            openZipName: ["", "menu"],
            unZipName: ["", "menu"],
            unZipCPName: ["", "menu"],
            addZipName: ["", "menu"],
            add: ["", "7z"],
            openAdd: ["", "7z"],
            lastPass: ["", "temp"],
            version: [0, "temp"],
        }

    static __Get(Key, Params)
        {
            if this.map.HasProp(key)
                return this.Read(key, this.map.%key%[1], this.map.%key%[2])
        }

    static Init(path) => this.path := path

    static Read(Key, default := "", section := "") => IniRead(this.path, section, key, default)

    static setWrite(key, value := "") => this.Write(value, key, this.map.%key%[2])

    static Write(value := "", Key := "", section := "")
        {
            static sectionSave := section
            if section
                sectionSave := section
            IniWrite(value, this.path, sectionSave, key)
        }

    static Delete(section, key)
        {
            IniDelete(this.path, section, key)
        }

    static ReadLoop(Section, arrMap, lower := false, twoVar := false)
    {
        loop
        {
            if !(var := this.Read(A_Index, , Section))
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
}

RelativePath(str) => StrReplace(str, "%SmartZipDir%", A_ScriptDir)

IniCreate()
{
    iniExist := FileExist(ini.path)
    version := ini.version
    VersionsCompare(num) => !iniExist || version < num

    if !iniExist
    {
        ini.setWrite("zipDir", "%SmartZipDir%\7-zip")
        ini.setWrite("icon", A_IsCompiled ? "%SmartZipDir%\SmartZip.exe" : "%SmartZipDir%\ico.ico")
        ini.setWrite("nesting", 1)
        ini.setWrite("nestingMuilt", 0)
        ini.setWrite("partSkip", 1)
        ini.setWrite("delSource", 0)
        ini.setWrite("delWhenHasPass", 0)
        ini.setWrite("autoAddPass", 0)
        ini.setWrite("dynamicPassSort", 0)
        ini.setWrite("autoRemovePass", 0)
        ini.setWrite("targetDir")
        ini.setWrite("test", 0)
        ini.setWrite("successPercent", 90)
        ini.setWrite("logLevel", 0)
        ini.setWrite("cmdLog", 0)
        ini.setWrite("hideRunSize", 10)

        ini.Write(0, "addDir2Pass", "set")

        ini.Write(, 1, "password")

        ini.setWrite("openZipName", "用7-Zip打开")
        ini.setWrite("unZipName", "智能解压")
        ini.setWrite("unZipCPName", "手动指定代码页解压")
        ini.setWrite("addZipName", "压缩")

        ini.Write("zip", 1, "ext")
        ini.Write("rar", 2)
        ini.Write("7z", 3)
        ini.Write("001", 4)
        ini.Write("cab", 5)
        ini.Write("bz2", 6)
        ini.Write("gz", 7)
        ini.Write("gzip", 8)
        ini.Write("tar", 9)

        ini.Write("^\d+$", 1, "extExp")
        ini.Write("zi", 2)
        ini.Write("7", 3)
        ini.Write("z", 4)

        ini.Write("iso", 1, "extForOpen")
        ini.Write("apk", 2)
        ini.Write("wim", 3)
        ini.Write("exe", 4)

        ini.Write("mp+3<--->mp3", 1, "renameExt")

        ini.Write("666666<--->", 1, "renameName")

        ini.Write("^[ 	]+<--->", 1, "renameExp")
        ini.Write("[ 	]+$<--->", 2)

        ini.Write(, 1, "excludeExt")

        ini.Write(, 1, "excludeName")

        ini.Write(, 1, "deleteExp")

        ini.setWrite("openAdd", '.zip" -tzip -mx=0 -aou -ad')
        ini.setWrite("add", '.zip"')

        ini.Write("Wrong password<--->1", 1, "unZipCheckError")
        ini.Write("Cannot open encrypted archive<--->1", 2)
        ini.Write("No files to process<--->1", 3)
        ini.Write("ERROR:<--->10", 4)

        ini.Write(, 1, "unZipCheckErrorExp")

        ; ini.Write("Errors: 1<--->1", 1, "unZipCheckErrorContinueExP")
        ; ini.Write("Data Error :<--->1", 2)
        ; ini.Write("Cannot open the file as archive<--->1", 3)

        ini.Write("Everything is Ok<--->1", 1, "unZipCheckSuccess")

        ini.Write(, 1, "unZipCheckSuccessExp")

        ini.Write("Errors: 1<--->1", 1, "openZipCheckError")
        ini.Write("ERROR:<--->1", 2)

        ini.Write(, 1, "openZipCheckErrorExp")

        ini.Write("Enter password (will not be echoed):<--->1", 1, "openZipCheckSuccess")	;需要输入密码则可能是压缩文件

        ini.Write("\d*-\d*-\d* *\d*:\d*:\d* *\d* *\d* *(\d*) files(, (\d*) folders)?<--->1", 1, "openZipCheckSuccessExp")	;多少个文件多少个文件夹则可能是压缩文件
    }

    if VersionsCompare(buildVersion)
        ini.setWrite("version", buildVersion)
}
