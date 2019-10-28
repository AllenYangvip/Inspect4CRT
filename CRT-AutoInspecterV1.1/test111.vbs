#$language = "VBScript"
#$interface = "1.0"
' Auth: AllenYang
' Desc: Cisco IOS device Inspect
' Date: 2019-10-17


crt.Screen.Synchronous = True


Sub Main
 ' 用户输入设备名称，2个作用
    ' 	1. 日志文件存储名称使用
    ' 	2. 输入命令后，用于捕获命令是否输入完毕。
    HostName = crt.Dialog.Prompt("请输入该设设备的hostname（hostname！！）")
    ' 用户输入日志全路径（包含文件名），方便日志存储
	LogPath = crt.Dialog.Prompt("请输入日志存储路径（如：D盘下的123文件夹--D:\123,禁止存放C:）")
	crt.Dialog.MessageBox("是否特权模式(#或]),如果不是，请进入特权模式，然后重新调用脚本","提示", 48|1)

	crt.Dialog.MessageBox('11')
	crt.Dialog.MessageBox("bbbb")
    ' 日志存储
    ' LogFullPath = LogPath + "\" + HostName + ".txt"
    ' crt.session.LogFileName = LogFullPath 
    ' crt.session.Log(true)

    ' ' 巡检命令数组
    ' Dim CMDArr(9)
    ' CMDArr(0) = "show version"
    ' CMDArr(1) = "show process cpu"
    ' CMDArr(2) = "show process memory"
    ' CMDArr(3) = "show inventory"
    ' CMDArr(4) = "show env all"
    ' CMDArr(5) = "show interface"
    ' CMDArr(6) = "show ntp status"
    ' CMDArr(7) = "show logging"
    ' CMDArr(8) = "show tech-support"
    
    ' ' 提示符
    ' Dprompt = HostName + "#"
    ' ' 清屏
    ' crt.Screen.Clear()
    ' crt.Screen.Send "ter len 0" & vbcrlf
    ' crt.Screen.WaitForString(Dprompt)
    ' For Each cmd In CMDArr
    '     r= GetResultByCmd(cmd, Dprompt)
    ' Next
    
    ' crt.Screen.WaitForString(Dprompt)
    crt.Sleep 1000
    crt.Screen.Send "quit" & vbcrlf
    crt.Session.Disconnect
End Sub


function GetResultByCmd(cmd, tip)
    crt.Screen.IgnoreEscape = True

    crt.Screen.Send cmd & vbcrlf
    crt.Screen.WaitForString cmd, 3 & vbcrlf
    crt.Screen.Send vbcr
    cmd_value = crt.Screen.ReadString(tip, 3, True)

    if (cmd_value = "") Then
        crt.Screen.Send vbcr
        cmd_value = crt.Screen.ReadString("More", 3, True)
        cmd_value = Replace(cmd_value, "-", "")
        cmd_value = Trim(Replace(cmd_value, chr(8), ""))

        if (cmd_value <> "") Then
            crt.Screen.SendKeys("^c")
            crt.Screen.Send vbcr
            crt.Screen.Send cmd & vbcrlf
            crt.Screen.WaitForString cmd, 3& vbcrlf
            cmd_value = crt.Screen.ReadString("More", 3, True)
            value = cmd_value
            do while (value <> "")
                crt.Screen.Send vbcr
                value = crt.Screen.ReadString("More", 3, True)
                cmd_value = cmd_value & value
            loop
        End if
    end if
    GetResultByCmd = Trim(cmd_value)
    
end function