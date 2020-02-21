Const filetype = ".pdf" '修改这个 改变转换的格式

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
Dim myDic, resultDic
Set myDic = CreateObject("Scripting.Dictionary")
Set resultDic = CreateObject("Scripting.Dictionary")
'递归运行，找出文件（去重复）
FilesTree (final_path)

'###########第二步调用ebook-convert.exe转换
'"G:\Calibre Portable\Calibre\ebook-convert.exe" "C:\Users\Administrator\Desktop\电子书3\[复杂].azw3" "[复杂].pdf"
'这一条语句就执行了转换
a2 = "G:\Calibre Portable\Calibre\ebook-convert.exe"
Set ws = wscript.CreateObject("wscript.shell")
For Each Key In myDic.Keys
	strCommand = Chr(34) & a2 & Chr(34) & " " & Chr(34) & myDic(Key) & Chr(34) & " " & Chr(34) & Key & filetype & Chr(34)
	'WScript.Echo strCommand
	ws.Run strCommand, 0, True
Next
 
'核心转换代码
Function FilesTree(sPath)
	Dim ext
    Dim f1, f2, f3, f4, f5, f6, f7, f8
    Dim p1, p2, p3, p4, p5, p6, p7, p8
    '遍历一个文件夹下的所有文件夹文件夹
    Set oFso = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFso.GetFolder(sPath)
    Set oSubFolders = oFolder.SubFolders
    Set oFiles = oFolder.Files
    'On Error Resume Next
    For Each ofile In oFiles
        ' WScript.Echo oFile.Path
        ext = ofso.GetExtensionName(ofile.Path) '获取后缀名
        If ext = "azw3" Or ext = "epub" Or ext = "mobi" Or ext = "azw"  Then
        	ofile.Name = Replace(ofile.Name, "?", " ") '将问号替换为空格
        	myDic.Add Left(ofile.Name, Len(ofile.Name) - Len(ext)), ofile.path
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
