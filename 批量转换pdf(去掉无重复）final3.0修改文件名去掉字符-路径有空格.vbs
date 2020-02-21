Const filetype = ".pdf" '�޸���� �ı�ת���ĸ�ʽ

Dim pub_oFso, pub_oFolder, pub_oSubFolders, pub_oFiles, sPath, final_path, fso, myfile
Dim this_Script_Name, mf, s
'��ñ��ű���·������
sPath = wscript.ScriptFullName
'��ñ��ű�������
this_Script_Name = wscript.ScriptName
'��ñ��ű������ļ��е�·��
final_path = Left(sPath, (Len(sPath) - Len(this_Script_Name)))

'###########��һ��
'����ʵ䣬����ȥ���ظ����ļ�
'[����].epub �� [����].azw3 ʵ������һ����������ֻ��Ҫת��һ�γ�pdf����
'˼·���� ����ʵ䣬Ȼ��ÿ��myDic.add key��item ����keyȡ"[����].epub �� [����].azw3"�еĹ����ļ��� [����]
'�����͹������ظ���Ȼ��itemֻ��������·��������ֻ������ C:\Users\Administrator\Desktop\����������2\[����].epub
'����ȥ�ظ���֮��� �������ֵ���key��Ψһ�ġ�
Dim myDic, resultDic
Set myDic = CreateObject("Scripting.Dictionary")
Set resultDic = CreateObject("Scripting.Dictionary")
'�ݹ����У��ҳ��ļ���ȥ�ظ���
FilesTree (final_path)

'###########�ڶ�������ebook-convert.exeת��
'"G:\Calibre Portable\Calibre\ebook-convert.exe" "C:\Users\Administrator\Desktop\������3\[����].azw3" "[����].pdf"
'��һ������ִ����ת��
a2 = "G:\Calibre Portable\Calibre\ebook-convert.exe"
Set ws = wscript.CreateObject("wscript.shell")
For Each Key In myDic.Keys
	strCommand = Chr(34) & a2 & Chr(34) & " " & Chr(34) & myDic(Key) & Chr(34) & " " & Chr(34) & Key & filetype & Chr(34)
	'WScript.Echo strCommand
	ws.Run strCommand, 0, True
Next
 
'����ת������
Function FilesTree(sPath)
	Dim ext
    Dim f1, f2, f3, f4, f5, f6, f7, f8
    Dim p1, p2, p3, p4, p5, p6, p7, p8
    '����һ���ļ����µ������ļ����ļ���
    Set oFso = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFso.GetFolder(sPath)
    Set oSubFolders = oFolder.SubFolders
    Set oFiles = oFolder.Files
    'On Error Resume Next
    For Each ofile In oFiles
        ' WScript.Echo oFile.Path
        ext = ofso.GetExtensionName(ofile.Path) '��ȡ��׺��
        If ext = "azw3" Or ext = "epub" Or ext = "mobi" Or ext = "azw"  Then
        	ofile.Name = Replace(ofile.Name, "?", " ") '���ʺ��滻Ϊ�ո�
        	myDic.Add Left(ofile.Name, Len(ofile.Name) - Len(ext)), ofile.path
        End If
    Next
    '�ݹ�
    For Each oSubFolder In oSubFolders
        FilesTree (oSubFolder.path) '�ݹ�
    Next
    '�ͷŶ���
    Set oFolder = Nothing
    Set oSubFolders = Nothing
    Set oFso = Nothing
    Set resultDic = myDic
End Function
