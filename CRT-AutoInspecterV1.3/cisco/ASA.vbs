#$language = "VBScript"
#$interface = "1.0"
' Auth: Yangjh
' Desc: Cisco ASA device Inspect
' Date: 2019-10-25
' Version: 1.3
' Platform: SecureCRT
' Email: yangjh@szkingdom.com

' 使用说明：
'     1. 使用CRT连接到要巡检的设备，进入编辑状态(Cisco/Hillstone/Ruijie设备进入特权模式(#),H3C/Huawei/Juniper 进入普通模式(xxx>)即可)
'     2. 输入show version 或者display version 判断该设备是何类型设备什么系统
'     3. 选择 CRT中的菜单栏 “脚本” 选项  “脚本”——>“执行”——>然后找到要执行的相应厂商的相应系统的脚本（例如:CRT-AutoInspecterV1\cisco\IOS.vbs）
'     4. 这个时候，您需要耐心的等几秒钟(8-20秒)，程序正在加载日志存储信息、识别设备的提示符
'     5. 当发现已经开始自动巡检时，证明已经成功。可以完全放心进行其他设备的巡检。
' 注意：
'     1. 问题1：运行脚本后无反应
'         首先程序运行在8-20秒内属于正常现象，因为程序正在加载日志存储信息、识别设备的提示符。在V1.1版本中这些信息是需要人工输入的，所以运行特别快，
'         但CRT有时会发生延迟回显，导致程序捕获提示符困难，所以耗费时间。而且加上网络延迟问题，这种现象可能经常发生，建议您多等几十秒，期间您可以做其他工作。
'     2. 问题2：脚本运行超过60秒，仍然无反应
'         如果运行超过60秒，您可以单击菜单栏 “脚本” 选项  “脚本”——>“取消”，然后再试一次
'     3. 问题3：脚本运行发生错误
'         如果脚本发生错误，请重复试验几次(可能是CRT延迟问题)，如果超过3次仍然报错，请联系作者，谢谢
' 联系方式：
'     杨纪海    电话：18518461120(微信同号)  
'



Sub Main
    ' 窗口最大化
    crt.Window.Show 3
    ' 关闭窗口同步功能
    crt.Screen.Synchronous = False
    ' 通过本地配置文件conf.ini获取日志存储路径
    cp = createobject("Scripting.FileSystemObject").GetFolder(".").path
    ParentPath = ""
    for i=Lbound(split(cp,"\")) to Ubound(split(cp, "\")) - 1
       ParentPath = ParentPath & split(cp,"\")(i) & "\"
    next 
    ' 拼接路径
    confPath = ParentPath & "conf.ini"
    ' 获取设备日志存储路径
    LogPath = ReadIni(confPath, "Setting", "saveFolder")
    ' 获取提示符
    Dprompt = getPrompt()
    ' 获取hostname，日志文件名称
    HostName = mid(Dprompt, 1, len(Dprompt) - 1)
    ' 日志存储
    LogFullPath = LogPath + "\" + HostName + ".txt"
    crt.session.LogFileName = LogFullPath 
    crt.session.Log(true)

    ' 巡检命令数组
    Dim CMDArr(9)
    CMDArr(0) = "show version"
    CMDArr(1) = "show cpu usage"
    CMDArr(2) = "show memory"
    CMDArr(3) = "show inventory"
    CMDArr(4) = "show module"
    CMDArr(5) = "show interface"
    CMDArr(6) = "show ntp status"
    CMDArr(7) = "show tech-support"
    CMDArr(8) = "show logging"
    
    ' 提示符
    ' Dprompt = HostName + "#"
    ' 清屏
    crt.Screen.Clear()
    crt.Screen.Send "ter len 0" & vbcrlf
    crt.Screen.WaitForString(Dprompt)
    crt.Screen.Synchronous = True    
    For Each cmd In CMDArr
        r = GetResultByCmd(cmd, Dprompt)
    Next
    crt.Screen.Send vbcrlf
    crt.Screen.Send vbcrlf
    crt.Screen.WaitForString(Dprompt)
    crt.Screen.Synchronous = False    
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



function getPrompt()
        crt.Screen.Clear()
        
    do
        ' 用户输入设备名称，2个作用
        ' 	1. 日志文件存储名称使用
        ' 	2. 输入命令后，用于捕获命令是否输入完毕。
        ' HostName = crt.Dialog.Prompt("请输入该设设备的hostname（hostname！！）")
        ' 清屏
        ' 提示符
        crt.Screen.Send chr(13)
        crt.sleep 1000
        crt.Screen.Send chr(13)
        crt.sleep 1000
        crt.Screen.Send chr(13)
        crt.sleep 1000
        crt.Screen.Send chr(13)
        crt.sleep 1000
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.sleep(2000)
        Dprompt = crt.Screen.get(2,1,2,30)
    loop Until Trim(Dprompt) <> ""
    getPrompt = Trim(Dprompt)
end function