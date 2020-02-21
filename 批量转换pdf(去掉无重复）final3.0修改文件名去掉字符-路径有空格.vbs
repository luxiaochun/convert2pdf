'WScript.CreateObject("WScript.Shell") 找不到指定文件的错误修正
'#若是路径用空格，两边再加chr(34) 就可以运行ws.Run Chr(34) & b1 & Chr(34), 1
Dim filetype
filetype = ".pdf" '修改这个 改变转换的格式

'filetype = ".docx"
Dim pub_oFso, pub_oFolder, pub_oSubFolders, pub_oFiles, sPath, final_path, fso, myfile
Dim this_Script_Name, mf, s
'求得本脚本的路径名字
sPath = wscript.ScriptFullName
'求得本脚本的名字
this_Script_Name = wscript.ScriptName
'求得本脚本所在文件夹的路径
final_path = Left(sPath, (Len(sPath) - Len(this_Script_Name)))
'###########第一步
'定义词典，用来去除重复的文件
'[复杂].epub 和 [复杂].azw3 实际上是一个东西，你只需要转换一次成pdf就行
'思路就是 定义词典，然后每次myDic.add key，item 其中key取"[复杂].epub 和 [复杂].azw3"中的共用文件名 [复杂]
'这样就过滤了重复，然后item只储存完整路径，比如只储存了 C:\Users\Administrator\Desktop\电子书整理2\[复杂].epub
'这样去重复了之后的 就是在字典里key是唯一的。
'###########第二步生成bat文件
'【for /l %%i in (1,1,1) do ("G:\Calibre Portable\Calibre\ebook-convert.exe" "C:\Users\Administrator\Desktop\电子书3\[复杂].azw3" "[复杂].pdf")】
'这一条语句就执行了转换
'###########第三步运行bat文件
'###########第四步运行bat文件结束后，销毁（删除bat）
Dim myDic, resultDic
Set myDic = CreateObject("Scripting.Dictionary")
Set resultDic = CreateObject("Scripting.Dictionary")
'递归运行，找出文件（去重复）
FilesTree (final_path)
'生成bat文件
Set fso = CreateObject("scripting.filesystemobject")
Set myfile=fso.CreateTextFile(final_path&"批量转换pdf.bat",True)
Dim a1, a2, a3, a4, a5, a6, a7
'拼接bat语句
a1 = "for /l %%i in (1,1,1) do ( "
a5 = " )"
a2 = "G:\Calibre Portable\Calibre\ebook-convert.exe"
a3 = " pause "
Set myDic2 = CreateObject("Scripting.Dictionary")
'拼接中间步骤，由于单独用拼接字符串，字符串的大小有限制，所以用字典的key来，保存每一条记录
For Each Key In myDic.Keys
    myDic2.Add Chr(34) & a2 & Chr(34) & " " & Chr(34) & myDic(Key) & Chr(34) & " " & Chr(34) & Key & filetype & Chr(34), ""
Next
'写入头一句
myfile.WriteLine a1
'写入中间部分
For Each Key In myDic2.Keys
    myfile.WriteLine Key
Next
'写入最后一个括号
myfile.WriteLine a5
'myfile.WriteLine a3
myfile.Close
Dim b1
b1 = final_path & "批量转换pdf.bat"
'convertct b1,"UTF-8" 转换也没有用 bat不支持 Unicode和utf-8
Set ws = wscript.CreateObject("wscript.shell")
'b1 = Chr(34) & b1 & Chr(34)
'b1 = Chr(34) & "start "& Chr(34) & b1 & Chr(34)
'MsgBox b1
ws.Run Chr(34) & b1 & Chr(34), 1

'###################完成模板，留作复用，2s后自动消失的对话框，代替MsgBox"#################################
'Set temp_VAR = CreateObject("WScript.Shell")
'temp_STR = "mshta.exe vbscript:close(CreateObject(""WScript.Shell"").Popup(""转换完成"",2,""标题""))"
'temp_VAR.Run temp_STR
'MsgBox 1
'Set myfile = fso.GetFile(final_path&"批量转换pdf.bat")
'myfile.Delete
'将读取的文件内容以指定编码写入文件
Function convertct(filepath, charset)
    Dim FileName, FileContents, dFileContents
    FileName = filepath
    FileContents = LoadFile(FileName)
    Set savefile = CreateObject("adodb.stream")
    savefile.Type = 2  '这里1为二进制，2为文本型
    savefile.Mode = 3
    savefile.Open()
    savefile.charset = charset
    savefile.Position = savefile.Size
    savefile.Writetext (FileContents) 'write写二进制,writetext写文本型
    savefile.SaveToFile filepath, 2
    savefile.Close()
    Set savefile = Nothing
End Function
 
'以文件本身编码读取文件
Function LoadFile(path)
    Dim Stm2
    Set Stm2 = CreateObject("ADODB.Stream")
    Stm2.Type = 2
    Stm2.Mode = 3
    Stm2.Open
    Stm2.charset = CheckCode(path)
    'Stm2.Charset = "UTF-8"
    'Stm2.Charset = "Unicode"
    'Stm2.Charset = "GB2312"
    Stm2.Position = Stm2.Size
    Stm2.LoadFromFile path
    LoadFile = Stm2.ReadText
    Stm2.Close
    Set Stm2 = Nothing
End Function
 
'该函数检查并返回文件的编码类型
Function CheckCode(file)
    Dim slz
    Set slz = CreateObject("Adodb.Stream")
    slz.Type = 1
    slz.Mode = 3
    slz.Open
    slz.Position = 0
    slz.LoadFromFile file
    Bin = slz.read(2)
    If is_valid_utf8(read(file)) Then
        Codes = "UTF-8"
    ElseIf AscB(MidB(Bin, 1, 1)) = &HFF And AscB(MidB(Bin, 2, 1)) = &HFE Then
        Codes = "Unicode"
    Else
        Codes = "GB2312"
    End If
    slz.Close
    Set slz = Nothing
    CheckCode = Codes
End Function
 
'将Byte()数组转成String字符串
Function read(path)
    Dim ado, a(), i, n
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 1: ado.Open
    ado.LoadFromFile path
    n = ado.Size - 1
    ReDim a(n)
    For i = 0 To n
        a(i) = ChrW(AscB(ado.read(1)))
    Next
    read = Join(a, "")
End Function
 
'准确验证文件是否为utf-8（能验证无BOM头的uft-8文件）
Function is_valid_utf8(ByRef input) 'ByRef以提高效率
    Dim s, re
    Set re = New Regexp
    s = "[\xC0-\xDF]([^\x80-\xBF]|$)"
    s = s & "|[\xE0-\xEF].{0,1}([^\x80-\xBF]|$)"
    s = s & "|[\xF0-\xF7].{0,2}([^\x80-\xBF]|$)"
    s = s & "|[\xF8-\xFB].{0,3}([^\x80-\xBF]|$)"
    s = s & "|[\xFC-\xFD].{0,4}([^\x80-\xBF]|$)"
    s = s & "|[\xFE-\xFE].{0,5}([^\x80-\xBF]|$)"
    s = s & "|[\x00-\x7F][\x80-\xBF]"
    s = s & "|[\xC0-\xDF].[\x80-\xBF]"
    s = s & "|[\xE0-\xEF]..[\x80-\xBF]"
    s = s & "|[\xF0-\xF7]...[\x80-\xBF]"
    s = s & "|[\xF8-\xFB]....[\x80-\xBF]"
    s = s & "|[\xFC-\xFD].....[\x80-\xBF]"
    s = s & "|[\xFE-\xFE]......[\x80-\xBF]"
    s = s & "|^[\x80-\xBF]"
    re.Pattern = s
    is_valid_utf8 = (Not re.Test(input))
End Function
 
'核心转换代码
Function FilesTree(sPath)
    Dim f1, f2, f3, f4, f5, f6, f7, f8
    Dim p1, p2, p3, p4, p5, p6, p7, p8
    '遍历一个文件夹下的所有文件夹文件夹
    Set oFso = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFso.GetFolder(sPath)
    Set oSubFolders = oFolder.SubFolders
    Set oFiles = oFolder.Files
    On Error Resume Next
    For Each ofile In oFiles
        ' WScript.Echo oFile.Path
        ' 拷贝资料pdf，Excel，Word，PPT
        f4 = Right(ofile.Name, 4) '后缀名为".pdf"，".xls"，".ppt"，".doc"
        f5 = Right(ofile.Name, 5) '后缀名为".xlsx"，".pptx"，".docx"
        f1 = Left(ofile.Name, 1) '前缀为$等临时文件
        If f5 = ".azw3" Or f5 = ".epub" Or f5 = ".mobi" Then
            mf = ofile.Name
            s = ""
            For i = 1 To Len(mf)
                b = Mid(mf, i, 1)
                If Asc(b) = 63 Then b = " "
                s = s & b
            Next
            ofile.Name = s
            myDic.Add Left(ofile.Name, Len(ofile.Name) - 5), ofile.path
            'MsgBox ofile.Name
            'MsgBox ofile.Path
        ElseIf f4 = ".azw" Then
            mf = ofile.Name
            s = ""
            For i = 1 To Len(mf)
                b = Mid(mf, i, 1)
                If Asc(b) = 63 Then b = " "
                s = s & b
            Next
            ofile.Name = s
            myDic.Add Left(ofile.Name, Len(ofile.Name) - 5), ofile.path
            'MsgBox ofile.Name
        End If
    Next
    '递归
    For Each oSubFolder In oSubFolders
        FilesTree (oSubFolder.path) '递归
    Next
    '释放对象
    Set oFolder = Nothing
    Set oSubFolders = Nothing
    Set oFso = Nothing
    Set resultDic = myDic
End Function
