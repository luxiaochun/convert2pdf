'WScript.CreateObject("WScript.Shell") �Ҳ���ָ���ļ��Ĵ�������
'#����·���ÿո������ټ�chr(34) �Ϳ�������ws.Run Chr(34) & b1 & Chr(34), 1
Dim filetype
filetype = ".pdf" '�޸���� �ı�ת���ĸ�ʽ

'filetype = ".docx"
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
'###########�ڶ�������bat�ļ�
'��for /l %%i in (1,1,1) do ("G:\Calibre Portable\Calibre\ebook-convert.exe" "C:\Users\Administrator\Desktop\������3\[����].azw3" "[����].pdf")��
'��һ������ִ����ת��
'###########����������bat�ļ�
'###########���Ĳ�����bat�ļ����������٣�ɾ��bat��
Dim myDic, resultDic
Set myDic = CreateObject("Scripting.Dictionary")
Set resultDic = CreateObject("Scripting.Dictionary")
'�ݹ����У��ҳ��ļ���ȥ�ظ���
FilesTree (final_path)
'����bat�ļ�
Set fso = CreateObject("scripting.filesystemobject")
Set myfile=fso.CreateTextFile(final_path&"����ת��pdf.bat",True)
Dim a1, a2, a3, a4, a5, a6, a7
'ƴ��bat���
a1 = "for /l %%i in (1,1,1) do ( "
a5 = " )"
a2 = "G:\Calibre Portable\Calibre\ebook-convert.exe"
a3 = " pause "
Set myDic2 = CreateObject("Scripting.Dictionary")
'ƴ���м䲽�裬���ڵ�����ƴ���ַ������ַ����Ĵ�С�����ƣ��������ֵ��key��������ÿһ����¼
For Each Key In myDic.Keys
    myDic2.Add Chr(34) & a2 & Chr(34) & " " & Chr(34) & myDic(Key) & Chr(34) & " " & Chr(34) & Key & filetype & Chr(34), ""
Next
'д��ͷһ��
myfile.WriteLine a1
'д���м䲿��
For Each Key In myDic2.Keys
    myfile.WriteLine Key
Next
'д�����һ������
myfile.WriteLine a5
'myfile.WriteLine a3
myfile.Close
Dim b1
b1 = final_path & "����ת��pdf.bat"
'convertct b1,"UTF-8" ת��Ҳû���� bat��֧�� Unicode��utf-8
Set ws = wscript.CreateObject("wscript.shell")
'b1 = Chr(34) & b1 & Chr(34)
'b1 = Chr(34) & "start "& Chr(34) & b1 & Chr(34)
'MsgBox b1
ws.Run Chr(34) & b1 & Chr(34), 1

'###################���ģ�壬�������ã�2s���Զ���ʧ�ĶԻ��򣬴���MsgBox"#################################
'Set temp_VAR = CreateObject("WScript.Shell")
'temp_STR = "mshta.exe vbscript:close(CreateObject(""WScript.Shell"").Popup(""ת�����"",2,""����""))"
'temp_VAR.Run temp_STR
'MsgBox 1
'Set myfile = fso.GetFile(final_path&"����ת��pdf.bat")
'myfile.Delete
'����ȡ���ļ�������ָ������д���ļ�
Function convertct(filepath, charset)
    Dim FileName, FileContents, dFileContents
    FileName = filepath
    FileContents = LoadFile(FileName)
    Set savefile = CreateObject("adodb.stream")
    savefile.Type = 2  '����1Ϊ�����ƣ�2Ϊ�ı���
    savefile.Mode = 3
    savefile.Open()
    savefile.charset = charset
    savefile.Position = savefile.Size
    savefile.Writetext (FileContents) 'writeд������,writetextд�ı���
    savefile.SaveToFile filepath, 2
    savefile.Close()
    Set savefile = Nothing
End Function
 
'���ļ���������ȡ�ļ�
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
 
'�ú�����鲢�����ļ��ı�������
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
 
'��Byte()����ת��String�ַ���
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
 
'׼ȷ��֤�ļ��Ƿ�Ϊutf-8������֤��BOMͷ��uft-8�ļ���
Function is_valid_utf8(ByRef input) 'ByRef�����Ч��
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
 
'����ת������
Function FilesTree(sPath)
    Dim f1, f2, f3, f4, f5, f6, f7, f8
    Dim p1, p2, p3, p4, p5, p6, p7, p8
    '����һ���ļ����µ������ļ����ļ���
    Set oFso = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFso.GetFolder(sPath)
    Set oSubFolders = oFolder.SubFolders
    Set oFiles = oFolder.Files
    On Error Resume Next
    For Each ofile In oFiles
        ' WScript.Echo oFile.Path
        ' ��������pdf��Excel��Word��PPT
        f4 = Right(ofile.Name, 4) '��׺��Ϊ".pdf"��".xls"��".ppt"��".doc"
        f5 = Right(ofile.Name, 5) '��׺��Ϊ".xlsx"��".pptx"��".docx"
        f1 = Left(ofile.Name, 1) 'ǰ׺Ϊ$����ʱ�ļ�
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
