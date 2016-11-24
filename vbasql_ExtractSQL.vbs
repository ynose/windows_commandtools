Sub ExtractSQL(path, keywords)
    Dim objFile
    Dim buf
    Dim log
    Dim sql
    
    ' ファイルの最後に出現したkeywordの行を抽出する
    ' ファイルが存在していない場合は、空のログファイルが作成される
    Set objFile = objFileSys.OpenTextFile(path, ForReading, true)
        Do Until objFile.AtEndOfStream      ' 入力ファイルの終端まで繰り返し
            buf = objFile.ReadLine
            If InStrs(buf, keywords) > 0 Then
                log = Split(buf, Chr(9))    ' Tab区切りのログを分割して
                sql = log(6)                ' SQL部分を抜き出す
            End if
        Loop
    objFile.Close
    Set objFile = Nothing


    ' 結果表示(コマンドラインで実行したときのみ出力するため、WScript.StdErr.WriteLineを使う)
    On Error Resume Next
    WScript.StdErr.WriteLine keywords
    'WScript.StdErr.WriteLine path
    WScript.StdErr.WriteLine log(0) & " " & log(1) & " " & log(2) & " " & log(3) & " " & log(4) & " " & log(5) & " " & Left(log(6), 10) & " ..."
    'WScript.StdErr.WriteLine sql
    On Error Goto 0



    ' 抽出したSQLを一時ファイルに出力
    If sql <> "" Then
        Dim strFolderName: strFolderName = objFileSys.GetAbsolutePathName(".\")
        Dim strFullPath: strFullPath = objFileSys.BuildPath(strFolderName, TEMPFILE)
        
        Set objFile = objFileSys.OpenTextFile(strFullPath, ForWriting, true)
            objFile.WriteLine sql
        objFile.Close
        Set objFile = Nothing


        ' 抽出したファイルをCSEで開く
        If objFileSys.FileExists(CSEEXE) Then
            Call ShellRun("""" & CSEEXE & """" & " """ & strFullPath & """", 5)
            Call SendKeys("^(q)")  ' SQL崩し
            Call SendKeys("^(s)")  ' 上書き保存（SQL崩しの後にcseを閉じると保存ダイアログを表示されるため上書き保存する)
        Else
            EchoError "Not exist " & CSEEXE
            WScript.Quit
        End If
    End If

End Sub

Function InStrs(str, keywords)

    Dim keyword
    Dim keywordArray: keywordArray = Split(keywords, ",")
    dim pos: pos = 0

    For Each keyword In keywordArray
        pos = InStr(str, keyword)
        If keyword <> "" And pos > 0 Then
            InStrs = Pos
        End If
    Next

End Function

Sub ShellRun(command, windowStyle)

    objWScriptShell.Run command, windowStyle
    WScript.sleep(1) ' 少しスリープを入れないと後続処理でエラーになる場合がある

End Sub

Sub SendKeys(key)

    objWScriptShell.AppActivate("Common SQL Environment - [" & TEMPFILE & "]")
    WScript.sleep(100) ' 少しスリープを入れないと後続処理がされない場合がある
    objWScriptShell.SendKeys(key)

End Sub

Sub EchoUsage(str)
    WScript.echo "Usage: " & str
End Sub
