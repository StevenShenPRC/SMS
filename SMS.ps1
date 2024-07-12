$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open('./sms.xlsx') # 指向存储数据的表格位置
$worksheet = $workbook.Sheets.Item(1)
$rowMax = ($worksheet.UsedRange.Rows).Count
for ($i = 1; $i -le $rowMax; $i++) {
    # 注意这里的$i,若表格首行为标题则应更改初始值为2
    # 以下为各个变量的名称 可根据需要自行更改
    $var1 = $worksheet.Cells.Item($i, 1).Value2
    $TargetPhoneNo = $worksheet.Cells.Item($i, 2).Value2 #此行导入发送短信的目标手机号
    $var3 = $worksheet.Cells.Item($i, 3).Value2
    
    #可根据需要进行不同条件的判断 比如判断性别 或群发通知时批次、地点、日期的确定
    #不需要直接注释掉就好
    if ($var3 -eq "Value1") {
         ($time = "ddmmyy") ; ($place = "somewhere") 
        }
    elseif ($var3 -eq "Value2")  {
         ($time = "ddmmyy") ; ($place = "somewhere") 
        } 
    else { 
        throw "Error: cannot decide condition"
     }
    
    # 或者简单粗暴一点 把上面的注释掉把下面的取消注释就能使用固定的内容
    # $time = $worksheet.Cells.Item($i, 4).Value2
    # $place = $worksheet.Cells.Item($i, 5).Value2
    
    #短信正文
    $message = "正文正文正文\ $var1\ 正文中间加空格或者符号需要转义\ $var3\ balabala\ $time\ balabalabala
    \ $var3$place\ balabalabala"
    

    # ADB发送短信部分
    adb shell am start -a android.intent.action.SENDTO -d sms:$TargetPhoneNo --es sms_body "$message" --ez exit_on_sent true
    # !!!注意 此行取决于你使用的Android设备UI布局 如果文本框和发送键之间有其他按键则需要执行
    adb shell input keyevent 22 #右方向键,光标移至表情
    # 文本框旁边直接就是发送的把上面那行注释掉就好
    adb shell input keyevent 22 #右方向键，光标移至发送
    adb shell input keyevent 66 #enter
    adb shell input keyevent 3  #回到桌面
    Write-Output "Sent message '$message' to '$var1' at '$TargetPhoneNo'." #输出发送成功消息
    Start-Sleep -Seconds 1 #设置延时，防止封号
}

# 关闭Excel和ADB进程
$workbook.Close()
$excel.Quit()