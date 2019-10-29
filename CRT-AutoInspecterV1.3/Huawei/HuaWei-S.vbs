#$language = "VBScript"
#$interface = "1.0"
' Auth: Yangjh
' Desc: Huawei S device Inspect
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
    HostName = Mid(Dprompt,2,len(Dprompt)-2)
    ' 日志存储
    LogFullPath = LogPath + "\" + HostName + ".txt"
    crt.session.LogFileName = LogFullPath 
    crt.session.Log(true)

    ' 巡检命令数组
    Dim CMDArr(14)
    CMDArr(0) = "display version"
    CMDArr(1) = "display cpu-usage"
    CMDArr(2) = "display memory-usage"
    CMDArr(3) = "display power"
    CMDArr(4) = "display fan"
    CMDArr(5) = "display ntp status"
    CMDArr(6) = "display interface"
    CMDArr(7) = "display environment"
    CMDArr(8) = "display logbuffer"
    CMDArr(9) = "display device"
    CMDArr(10) = "display esn"
    CMDArr(11) = "display health"
    CMDArr(12) = "display elabel"
    CMDArr(13) = "display diagnostic-information"
    
    ' 清屏
    crt.Screen.Clear()
    crt.Screen.Send vbcrlf
    crt.Screen.WaitForString(Dprompt)
    crt.Screen.Send vbcrlf
    crt.Screen.WaitForString(Dprompt)
    crt.Screen.Synchronous = True
    crt.Screen.Send "screen-length 0 temporary" & vbcrlf
    crt.Screen.WaitForString(Dprompt)
    crt.Screen.Send vbcrlf
    crt.Screen.WaitForString(Dprompt)
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


Function ReadIni( myFilePath, mySection, myKey )
    ' This function returns a value read from an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be returned
    '
    ' Returns:
    ' the [string] value for the specified key in the specified section
    '
    ' CAVEAT:     Will return a space if key exists but value is blank
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre and Rob van der Woude
    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection

    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ReadIni     = ""
    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )

    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, ForReading, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )

            ' Check if section is found in the current line
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                strLine = Trim( objIniFile.ReadLine )

                ' Parse lines until the next section is reached
                Do While Left( strLine, 1 ) <> "["
                    ' Find position of equal sign in the line
                    intEqualPos = InStr( 1, strLine, "=", 1 )
                    If intEqualPos > 0 Then
                        strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                        ' Check if item is found in the current line
                        If LCase( strLeftString ) = LCase( strKey ) Then
                            ReadIni = Trim( Mid( strLine, intEqualPos + 1 ) )
                            ' In case the item exists but value is blank
                            If ReadIni = "" Then
                                ReadIni = " "
                            End If
                            ' Abort loop when item is found
                            Exit Do
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = Trim( objIniFile.ReadLine )
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
        WScript.Echo strFilePath & " doesn't exists. Exiting..."
        Wscript.Quit 1
    End If
End Function

function getPrompt()   
    do
        crt.Screen.Clear()
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