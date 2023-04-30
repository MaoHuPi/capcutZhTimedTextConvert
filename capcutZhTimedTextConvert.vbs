' 2023 (c) MaoHuPi
' capcutZhTimedTextConvert.vbs
' v1.0.0
' 
' 將剪映專案內的字幕在各類中文間轉換
' 此程式依賴「繁化姬」的API

Dim shell
Set shell = CreateObject("Wscript.Shell")
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Function changeDirectory(path)
    path = shell.expandEnvironmentStrings(path)
    shell.CurrentDirectory = path
    ' Call shell.Run("explorer " & shell.CurrentDirectory)
End Function

Function listDirectory(path, fileList, subFolderList)
    If Not fso.FolderExists(path) Then
        MsgBox "Directory D.N.E!" & vbNewLine & "(檔案夾不存在)"
        listDirectory = False
    End If
    Dim folder
    Set folder = fso.GetFolder(path)
    If IsObject(fileList) Then
        For Each file In folder.Files
            fileList.Add file.Name
        Next
    End If
    If IsObject(subFolderList) Then
        For Each subFolder In folder.SubFolders
            subFolderList.Add subFolder.Name
        Next
    End If
End Function

Function readFile(path, encoding)
    path = shell.expandEnvironmentStrings(path)

    If Not fso.FileExists(path) Then
        MsgBox "File D.N.E!" & vbNewLine & "(檔案不存在)"
        readFile = False
    Else
        Dim objStream, content
        Set objStream = CreateObject("ADODB.Stream")
        objStream.CharSet = encoding
        objStream.Open
        objStream.LoadFromFile path
        content = objStream.ReadText()
        objStream.Close
        Set objStream = Nothing
        readFile = content
    End If
End Function

Function writeFile(path, encoding, content)
    path = shell.expandEnvironmentStrings(path)

    If Not fso.FileExists(path) Then
        MsgBox "File D.N.E!" & vbNewLine & "(檔案不存在)"
        readFile = False
    Else
        Dim objStream
        Set objStream = CreateObject("ADODB.Stream")
        objStream.CharSet = encoding
        objStream.Open
        objStream.WriteText content
        objStream.SaveToFile path, 2
        objStream.Close
        Set objStream = Nothing
    End If
End Function

Const zhconvertURL = "https://api.zhconvert.org/convert?converter={targetType}&text={text}"
Dim xmlhttp
Set xmlhttp = CreateObject("MSXML2.serverXMLHTTP")
Dim zhc_htmlfile
Set zhc_htmlfile = CreateObject("htmlfile") 
Set zhc_window = zhc_htmlfile.parentWindow
zhc_window.execScript "var res = {}; var text = ''", "JScript"
Function zhConvert(text, targetType)
    Dim url : url = zhconvertURL
    url = Replace(url, "{targetType}", targetType)
    url = Replace(url, "{text}", text)
    xmlhttp.Open "GET", url, False
    xmlhttp.Send
    Dim res : res = xmlhttp.responseText
    Dim js : js = "res = " & res & "; text = res['data']['text']"
    zhc_window.execScript js, "JScript"
    zhConvert = zhc_window.text
End Function

Function parseDraftTimedText(json) 
    Dim htmlfile
    Set htmlfile = CreateObject("htmlfile") 
    Set window = htmlfile.parentWindow
    Dim js : js = "" & _ 
    "var jsonObject = " & json & ";" & _ 
    "var texts = jsonObject['materials']['texts'];" & _ 
    "var contentList = [];" & _ 
    "for(var i = 0; i < texts.length; i++){" & _ 
    "   var data = texts[i]['content'].split(/\[|\]/g);" & _ 
    "   if(data.length == 3){contentList[i] = data[1];}" & _ 
    "}"
    window.execScript js, "JScript"
    Dim jsContentList, contentList, item
    Set jsContentList = window.contentList
    Set contentList = CreateObject("System.Collections.ArrayList")
    For i = 0 To jsContentList.length-1
        exestring = "item = jsContentList.[" & i & "]"
        execute(exestring)
        contentList.Add item
    Next
    Set parseDraftTimedText = contentList
End Function

Function updateDraftTimedText(json, contentList)
    Dim htmlfile
    Set htmlfile = CreateObject("htmlfile") 
    Set window = htmlfile.parentWindow
    Dim js : js = "" & _ 
    "function encodeURI(text){" & _ 
    "   text = text.replace(/""/g, '\\""');" & _ 
    "   return text;" & _ 
    "}" & _ 
    "function jsonStringify(jsObject){" & _ 
    "    if(jsObject === undefined){return 'undefined';}" & _ 
    "    else if(jsObject === null){return 'null';}" & _ 
    "    else if(jsObject === true){return 'true';}" & _ 
    "    else if(jsObject === false){return 'false';}" & _ 
    "    else if(typeof jsObject == 'string'){return '""' + encodeURI(jsObject) + '""';}" & _ 
    "    else if(typeof jsObject == 'number'){return jsObject;}" & _ 
    "    else if('length' in jsObject){" & _ 
    "        var innerJsonList = [];" & _ 
    "        for(var i = 0; i < jsObject.length; i++){" & _ 
    "            innerJsonList.push(jsonStringify(jsObject[i]));" & _ 
    "        }" & _ 
    "        return '[' + innerJsonList.join(',') + ']';" & _ 
    "    }" & _ 
    "    else{" & _ 
    "        var innerJsonList = [];" & _ 
    "        for(var key in jsObject){" & _ 
    "            innerJsonList.push('""' + key + '"":' + jsonStringify(jsObject[key]));" & _ 
    "        }" & _ 
    "        return '{' + innerJsonList.join(',') + '}';" & _ 
    "    }" & _ 
    "}" & _ 
    "var jsonObject = " & json & ";" & _ 
    "var texts = jsonObject['materials']['texts'];" & _ 
    "var contentList = ['" & Join(contentList.ToArray(), "', '") & "'];" & _ 
    "for(var i = 0; i < texts.length; i++){" & _ 
    "   var data = texts[i]['content'].replace(']', '[').split('[');" & _ 
    "   if(data.length == 3){" & _ 
    "       data[1] = contentList[i];" & _ 
    "       texts[i]['content'] = data[0] + '[' + data[1] + ']' + data[2];" & _ 
    "   }" & _ 
    "}" & _ 
    "var jsonString = jsonStringify(jsonObject);"
    window.execScript js, "JScript"
    updateDraftTimedText = window.jsonString
End Function

Sub main()
    changeDirectory("%AppData%\..\Local\JianyingPro\User Data\Projects\com.lveditor.draft")
    Dim projectNames, projectName
    Dim projectNameList
    Set projectNameList = CreateObject("System.Collections.ArrayList")
    listDirectory "./", False, projectNameList
    projectNames = Join(projectNameList.ToArray(), vbNewLine)
    projectName = InputBox("Please Enter The Project Name!" & vbNewLine & "(請輸入專案名稱)" & vbNewLine & "-----" & vbNewLine & projectNames, "Choose Project")
    Dim content
    content = readFile("./" & projectName & "/draft_content.json", "utf-8")
    If content = False Then
        Exit Sub
    End If
    Dim timedTextContentList
    Set timedTextContentList = parseDraftTimedText(content)
    Dim targetTypeOptionsList
    Set targetTypeOptionsList = CreateObject("System.Collections.ArrayList")
    targetTypeOptionsList.Add "Traditional(繁體化)"
    targetTypeOptionsList.Add "Simplified(簡體化)"
    targetTypeOptionsList.Add "Taiwan(台灣化)"
    targetTypeOptionsList.Add "Hongkong(香港化)"
    targetTypeOptionsList.Add "China(中國化)"
    targetTypeOptionsList.Add "Bopomofo(注音化)"
    targetTypeOptionsList.Add "Pinyin(拼音化)"
    targetTypeOptionsList.Add "Mars(火星化)"
    targetTypeOptionsList.Add "Exit(退出程式)"
    targetTypeOptionsList.Add "Restoration(恢復備份)"
    Dim targetTypeOptions : targetTypeOptions = ""
    For i = 0 To targetTypeOptionsList.Count-1
        targetTypeOptions = targetTypeOptions & vbNewLine & _
        "" & i+1 & ". " & targetTypeOptionsList.Item(i)
    Next
    Dim targetTypeIndex, targetType
    targetTypeIndex = InputBox("Please Enter The Index Of Target Type!" & vbNewLine & "(請輸入目標模式的索引值)" & vbNewLine & "-----" & vbNewLine & targetTypeOptions, "Choose Project")
    targetTypeIndex = CInt(targetTypeIndex)-1
    If targetTypeIndex >= targetTypeOptionsList.Count Then
        MsgBox "The Index Must Be Less Than " & targetTypeOptionsList.Count & "!" & vbNewLine & "(索引值必須小於" & targetTypeOptionsList.Count & ")"
        Exit Sub
    End If
    targetType = targetTypeOptionsList.Item(targetTypeIndex)
    targetType = Split(targetType, "(")(0)
    If targetType = "Exit" Then
        Exit Sub
    Elseif targetType = "Restoration" Then
        content = readFile("./" & projectName & "/draft_content.json.bak", "utf-8")
        If content = False Then
            Exit Sub
        End If
        writeFile "./" & projectName & "/draft_content.json", "utf-8", content
        MsgBox "Done!" & vbNewLine & "(完成)"
        Exit Sub
    End If
    For i = 0 To timedTextContentList.Count-1
        timedTextContentList.Item(i) = zhConvert(timedTextContentList.Item(i), targetType)
    Next
    Dim newContent : newContent = updateDraftTimedText(content, timedTextContentList)
    ' MsgBox newContent
    MsgBox "Done!" & vbNewLine & "(完成)"
    writeFile "./" & projectName & "/draft_content.json", "utf-8", newContent
End Sub

main()