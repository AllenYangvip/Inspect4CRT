#$language = "VBScript"
#$interface = "1.0"
' Auth: AllenYang
' Desc: Cisco IOS device Inspect
' Date: 2019-10-


Sub Main
    cp = createobject("Scripting.FileSystemObject").GetFolder(".").path
    ParentPath = ""
    for i=Lbound(split(cp,"\")) to Ubound(split(cp, "\")) - 1
       ParentPath = ParentPath & split(cp,"\")(i) & "\"
    next 
    confPath = ParentPath & "conf.ini"
    ' MsgBox confPath
    dirPath = ReadIni(confPath, "Setting", "saveFolder")
    MsgBox dirPath
    ' crt.Screen.Send chr(13)
    ' crt.Screen.Send chr(13)
    ' crt.Screen.Send chr(13)
    ' crt.Screen.Send chr(13)
    ' crt.Screen.Send chr(13)
    ' crt.Screen.Send chr(13)
    ' crt.sleep(1000)
    ' tip = crt.Screen.get2(12,1,12,80)
    ' ' cisco/ RG/Hillstone/Juniper start: 1 end: len(tip) -3 后面有两个空格
    ' ' h3c/huawei  start: 2 end: len(tip) - 3
    ' tip3 = mid(tip, 1, len(tip)-3)
    ' MsgBox tip3
    Dprompt = getPrompt()
    MsgBox Dprompt

End Sub


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
        ' 用户输入设备名称，2个作用
        ' 	1. 日志文件存储名称使用
        ' 	2. 输入命令后，用于捕获命令是否输入完毕。
        ' HostName = crt.Dialog.Prompt("请输入该设设备的hostname（hostname！！）")
        ' 清屏
        crt.Screen.Clear()
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.Screen.Send chr(13)
        crt.sleep(3000)
        ' 提示符
        Dprompt = crt.Screen.get(12,1,12,80)
        ' 清屏
        crt.Screen.Clear()
        crt.Screen.Send chr(13)
    loop Until Trim(Dprompt) <> ""
    getPrompt = Trim(Dprompt)
end function