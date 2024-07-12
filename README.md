# SMS
一个Powershell脚本，用于群发SMS短信，使用ADB实现操控手机，使用Excel（.xlsx）存储短信内容信息。

## 准备要求
- Powershell
- `.xlsx` 的编辑器
- Android Platfrom-Tools (ADB)

环境问题请自行解决
## 用法
首先下载此存储库中的所有文件。

    git clone https://github.com/StevenShenPRC/SMS

打开下载的SMS文件夹，编辑 `sms.xlsx` 更改其中存储的可变信息。

编辑 `sms.ps1` 中的短信正文部分，使之符合需要。

    $message = "正文正文正文\ $var1\ 正文中间加空格或者符号需要转义\ $var3\ balabala\ $time\ balabalabala\ $var3$place\ balabalabala"

将准备发送短信的手机链接电脑，并连接adb

保存所有文件。在SMS文件夹处打开`Powershell`，或更改目录到 `SMS/` 文件夹下。

执行以下命令：

    ./SMS.ps1
等待执行完毕即可。

## 其他事项
- 为了防止快速发送短信被封号（断卡行动），设置每一条短信发送间隔为1s.
- 如有疑问 [联系邮箱：Steven@stevenshen.cn](mailto:Steven@stevenshen.cn) 非CS相关专业人员水平极其有限不保证能解答