#$language = "VBScript"
#$interface = "1.0"
' Auth: AllenYang
' Desc: Cisco IOS device Inspect
' Date: 2019-10-22


' 使用方法：
'     1. 使用CRT连接到要巡检的设备，进入特权模式
'     2. 输入show version 或者display version 判断该设备是何类型设备什么系统，然后复制好本设备的hostname
'     3. 选择 CRT中的菜单栏 “脚本” 选项  “脚本”——>“执行”——>然后找到要执行的相应厂商的相应系统的脚本（例如:CRT-AutoInspecterV1\cisco\IOS.vbs）
'     4. CRT会提示您 输入日志存储路径，如果您已经建立好了本次巡检的文件夹，请详细指定到文件夹。如：D:\123，也可以直接指定盘符 例如: D:
'     5. 输入设备的hostname,请您在进入设备后就复制好，这样方便您，在这里粘贴。
'     6. 确定后脚本会自己运行，然后自动保存到您所指定的目录，文件以hostname.txt保存
' 注意：
'     1. 以下为该脚本具体代码。以单引号(')开头的行为注释行，不会执行
'     2. 如果您没有其他需求，不建议您修改脚本。
'     3. 路径结尾不用跟路径分隔符(\)!!!!
'     4. 如果脚本存在问题，请您联系我  allenyangvip@126.com
'

'CRT开启屏幕同步
crt.Screen.Synchronous = True
crt.Window.Show 3


Sub Main
    ' 用户输入日志全路径（包含文件名），方便日志存储
    cp = createobject("Scripting.FileSystemObject").GetFolder(".").path
    ParentPath = ""
    for i=Lbound(split(cp,"\")) to Ubound(split(cp, "\")) - 1
    ParentPath = ParentPath & split(cp,"\")(i) & "\"
    next 
    confPath = ParentPath & "conf.ini"
    ' MsgBox confPath
    LogPath = ReadIni(confPath, "Setting", "saveFolder")
    ' 用户输入设备名称，2个作用
    ' 	1. 日志文件存储名称使用
    ' 	2. 输入命令后，用于捕获命令是否输入完毕。
    Dprompt = crt.Dialog.Prompt("请输入该设设备的提示符，方便捕获和存储文件")
    Dprompt = Trim(Dprompt)
    HostName = Mid(Dprompt,1,len(Dprompt)-1)
    ' 日志存储
    LogFullPath = LogPath + "\" + HostName + ".txt"
    crt.session.LogFileName = LogFullPath 
    crt.session.Log(true)

    ' 巡检命令数组
    Dim CMDArr(7)
    CMDArr(0) = "show version detail no-forwarding | no-more"
    CMDArr(1) = "show system uptime no-forwarding | no-more"
    CMDArr(2) = "show chassis environment no-forwarding | no-more"
    CMDArr(3) = "show chassis routing-engine no-forwarding | no-more"
    CMDArr(4) = "show interfaces extensive no-forwarding | no-more"
    CMDArr(5) = "show configuration | display set | no-more"
    CMDArr(6) = "request support information | no-more"
    
    ' 提示符
    ' Dprompt = HstName + ">"
    ' 清屏
    crt.Screen.Clear()
    crt.Screen.Send vbcrlf
    crt.Screen.WaitForString(Dprompt)
    For Each cmd In CMDArr
        r = GetResultByCmd(cmd, Dprompt)
    Next
    
    crt.Screen.WaitForString(Dprompt)
    crt.Sleep 1000
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