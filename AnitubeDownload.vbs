Dim URL : URL = InputBox("保存したいAnitubeのURLを入力してください。")

If URL = "" Then
    WScript.Quit
End If 

parts = split(URL,"/") 
saveTo = parts(ubound(parts)) + ".mp4"

Dim ScriptUrl : ScriptUrl = GetString(URL,"(http://www\.anitube\.se/player/config\.php\?key=[^""]+)")
Dim videoUrl : videoUrl = GetString(ScriptUrl,"""(http://.+/\d+\.mp4)""")
DownloadFile(videoUrl)

'MsgBox("できたー！"&saveTo)
' Done
WScript.Quit

Function GetString(url,pattern)

    ' Fetch the file
    Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")

    objXMLHTTP.open "GET", URL, false
    objXMLHTTP.send()

    If objXMLHTTP.Status = 200 Then

        Dim regEx : Set regEx = New RegExp
        regEx.Pattern = pattern
        regEx.IgnoreCase = True
        regEx.Global = True   

        ' 検索の実行
        Dim Matches : Set Matches = regEx.Execute(objXMLHTTP.responseText)

        Dim Match
        For Each Match in Matches
        GetString = Match.SubMatches(0)
        Next

    End if
    Set objXMLHTTP = Nothing

End Function

Sub DownloadFile(url)

    ' Fetch the file
    Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")

    objXMLHTTP.open "GET", URL, false
    objXMLHTTP.send()

    If objXMLHTTP.Status = 200 Then
        Set objADOStream = CreateObject("ADODB.Stream")
        objADOStream.Open
        objADOStream.Type = 1 'adTypeBinary

        objADOStream.Write objXMLHTTP.ResponseBody
        objADOStream.Position = 0    'Set the stream position to the start

        Set objFSO = Createobject("Scripting.FileSystemObject")
        If objFSO.Fileexists(saveTo) Then objFSO.DeleteFile saveTo
        Set objFSO = Nothing

        objADOStream.SaveToFile saveTo
        objADOStream.Close
        Set objADOStream = Nothing

    End if
    Set objXMLHTTP = Nothing

End Sub
