## 功能介绍
 - 智能解压
    - 单文件时解压至当前文件夹
    - 多文件时解压到当前文件夹下的某个文件夹
    - 压缩包包含密码时,遍历设置的密码,密码正确解压,不正确提示手动输入密码并解压
      - 自带两个密码,上次使用的密码,剪贴板复制的内容(移除了首尾空格和换行)
      - 如不想添加可以直接复制密码然后运行智能解压
      - 自动新增使用过的密码 **2.20+**
      - 动态排序密码 **2.20+**
    - 解压完成后按照指定规则处理压缩后的文件,如重命名,删除
    - 解压嵌套压缩包
      - 文件后缀名符合`ini-ext,extExp` 规则则解压
      - **嵌套压缩包解压后会删除**
 - 智能打开
   - 如果是压缩包则打开,如果不是则显示添加到压缩包界面
 - 压缩
   - 全是文件夹则每个文件夹生成一个压缩包, 否则生成单个压缩包

## 设置方式
 - 直接运行 `SmartZip.exe` 会显示设置界面 **3.0+**
 - **建议清空所有 `password` `rename` `delete` 然后按照需求添加**
 - 可批量从`tx`t或旧版本`ini`设置中导入密码  **3.0+**
 - 更多自定义请直接编辑ini,参考以下链接设置,后续可能不再更新ini文档
     - [INI设置](ini.md)

## 运行方式
 - 如果启用了右键,可在资源管理器中右键文件使用
    - 右键实现方式不完美
       - 由于右键菜单单次只能传递一个文件,传递多文件过于复杂
       - 目前方法为在当前窗口发送 复制(Ctrl+C) 快捷键,可能会扰乱剪贴板
       -  右键菜单有15个文件限制,解除限制访问下方链接按说明操作
          - [context-menus-shortened-select-over-15-files](https://docs.microsoft.com/zh-cn/troubleshoot/windows-client/shell-experience/context-menus-shortened-select-over-15-files)
 - 右键发送到菜单 **2.14+**
    - 不影响剪贴板
    - 不受15个文件限制影响
    - 如使用资源管理器可用此代替
    - 缺点是在二级目录里
 - 通过直接传递参数运行(推荐但比较繁杂)
   - 智能解压: `SmartZip.exe  x  file1 file2 file3 ....`
   - 使用7-zip打开: `SmartZip.exe  o  file1`
   - 压缩: `SmartZip.exe  a  file1 file2 file3 ....`
 - Directory Opus 示例
   - 智能解压: `SmartZip.exe x {allfilepath}`
   - 使用7-zip打开: `SmartZip.exe o {allfilepath} `
   - 压缩: `SmartZip.exe a {allfilepath} `
 - 向 `Contextmenu.exe` 传递参数或直接运行
    - 它会在运行时执行复制,然后将其传给主脚本执行
    - 选中文件然后以快捷键或其他方法调用`Contextmenu.exe`
    - 无参时默认智能解压
   - 智能解压: `Contextmenu.exe  x`
   - 使用7-zip打开: `Contextmenu.exe  o`
   - 压缩: `Contextmenu.exe  a`
 - 直接运行 `SmartZip.exe` 然后拖拽文件到界面上会触发智能解压 **3.0+**
 - 拖拽文件到 `SmartZip.exe` 上会触发智能解压

## 提示
 - **更新版本建议备份 ini以防出错**

## 预览图
 - 设置界面

![set](pic\set.gif)

 - 资源管理器右键界面

![2](https://user-images.githubusercontent.com/2145741/173320643-509a43e2-fb9f-4ca5-981f-c99b7f020f1e.jpg)

 - 发送到界面

![3](https://user-images.githubusercontent.com/2145741/173808930-bcce4273-c930-4e84-9a40-c52349760fc0.jpg)

 - 批量解压界面

![addZip](pic\addZip.jpg)


 - 批量压缩界面

![unZip](pic\unZip.jpg)


## 相关链接
  - [7-zip](https://www.7-zip.org/)
    - 测试基于 7-Zip 21.07 版本
  - [小众软件](https://www.appinn.com/smartzip-for-7zip/)
  - [小众软件发现频道](https://meta.appinn.net/t/topic/33555)