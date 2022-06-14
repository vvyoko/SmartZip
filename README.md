## 功能介绍
 - 智能解压
    - 单文件时解压至当前文件夹
    - 多文件时解压到当前文件夹下的某个文件夹
    - 压缩包包含密码时,遍历设置的密码,密码正确解压,不正确提示手动输入密码并解压
      - 自带两个密码,上次使用的密码,剪贴板复制的内容(移除了首尾空格和换行)
      - 也就是说如果不想添加可以直接复制密码然后运行智能解压
    - 解压完成后按照指定规则处理压缩后的文件,如重命名,删除
    - 解压嵌套压缩包
      - 文件后缀名符合`ini设置-ext,extExp` 标签则解压
      - 默认解压单个文件的嵌套压缩类似 `a.zip/b.zip`
      - 通过ini设置支持多个嵌套 `a.zip/b.zip, c.rar, d,7, ...` 不会遍历子文件夹
 - 智能打开
   - 如果是压缩包则打开,如果不是则显示添加到压缩包界面
 - 压缩
   - 全是文件夹则每个文件夹生成一个压缩包, 否则生成单个压缩包

## 设置方式
 - 首先运行 `SmartZip.exe`,会自动生成`SmartZip.ini`文件并打开
 - 然后参考`ini说明.txt`设置,必需设置 `7zipDir`
 - 建议清空所有 `password` `rename` `delete` 然后按照需求添加
 - 默认ini为了让功能能被人使用,默认开启了大部分功能
   - 比方说日志,右键菜单
 - 绝大多数功能都能自定义,具体查看注释

## 运行方式
 - 如果启用了右键,可在资源管理器中右键文件使用
    - 右键实现方式不完美,具体可查看`ini说明.txt`
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
 - 拖拽文件到 `SmartZip.exe` 上会触发智能解压


## 预览图
 - 右键关联界面

![2](https://user-images.githubusercontent.com/2145741/173320542-65ccfbbe-8e5a-4e97-80b0-a19f36a8881f.jpg)

 - 资源管理器右键界面

![3](https://user-images.githubusercontent.com/2145741/173320643-509a43e2-fb9f-4ca5-981f-c99b7f020f1e.jpg)


 - 批量解压界面

![4](https://user-images.githubusercontent.com/2145741/173320704-35a051a1-0f03-4172-b232-2e410b7a4311.jpg)


 - 批量压缩界面

![5](https://user-images.githubusercontent.com/2145741/173320771-15412318-05ef-4158-b01c-4ab828e12ec6.jpg)
