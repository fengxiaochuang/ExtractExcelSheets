Dim fso,fld,Path
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Path = "D:\project\node\mail-sender-tool\test\new" '获取脚本所在文件夹字符串
Set fld=fso.GetFolder(Path) '通过路径字符串获取文件夹对象

Dim Sum,IsChooseDelete,ThisTime
Sum = 0
Dim LogFile
Set LogFile= fso.opentextFile("log.txt",8,true)

Dim List
Set List= fso.opentextFile("ConvertFileList.txt",2,true)

Call LogOut("开始遍历文件")
Call TreatSubFolder(fld) '调用该过程进行递归遍历该文件夹对象下的所有文件对象及子文件夹对象

Sub LogOut(msg)
    ThisTime=Now
    LogFile.WriteLine(year(ThisTime) & "-" & Month(ThisTime) & "-" & day(ThisTime) & " " & Hour(ThisTime) & ":" & Minute(ThisTime) & ":" & Second(ThisTime) & ": " & msg)
End Sub

Sub TreatSubFolder(fld) 
    Dim File
    Dim ts
    For Each File In fld.Files '遍历该文件夹对象下的所有文件对象
        If UCase(fso.GetExtensionName(File)) ="XLSX" Then
            List.WriteLine(File.Path)
            Sum = Sum + 1
        End If
    Next
    'Dim subfld
    'For Each subfld In fld.SubFolders '递归遍历子文件夹对象
        'TreatSubFolder subfld
    'Next
End Sub

List.close

Call LogOut("文件遍历已完成，已找到" & Sum & "个excel文档")

On Error Resume Next
Set ExcelApp = CreateObject("Excel.Application")
On Error Goto 0

ExcelApp.Visible=false '设置视图不可见


Sum = 0
Dim FilePath,FileLine
Set List= fso.opentextFile("ConvertFileList.txt",1,true)
Do While List.AtEndOfLine <> True 
    FileLine=List.ReadLine
    If FileLine <> "" and Mid(FileLine,1,2) <> "~$" Then
        Sum = Sum + 1 '获取用户修改后的文件列表行数
    End If
loop
List.close

'MsgBox "现在开始转换，若是在运行过程中弹出Word窗口"&vbCrlf&"请直接最小化Word窗口，不要关闭!"&vbCrlf&"请直接最小化Word窗口，不要关闭!"&vbCrlf&"请直接最小化Word窗口，不要关闭!"&vbCrlf&"重要的事情说三遍！关闭会导致脚本退出", vbOKOnly + vbExclamation, "警告"

Dim Finished,filename
Finished = 0
Set List= fso.opentextFile("ConvertFileList.txt",1,true)
Do While List.AtEndOfLine <> True 
    FilePath=List.ReadLine
    If Mid(FilePath,1,2) <> "~$" Then '不处理临时文件
        Set objExcel = ExcelApp.Workbooks.Open(FilePath)
        If ExcelApp.Visible = true Then
            ExcelApp.ActiveWorkbook.ActiveWindow.WindowState = 2 'wdWindowStateMinimize
        End If
        Set Sheets = objExcel.Sheets
        For i = 1 To Sheets.Count 
        	If Sheets(i).Name <> "统计" Then
        'filename = objExcel.Path & "\" & objExcel.Name & "-" & Sheets(i).Name & ".xlsx"
        'LogOut("文档" & FilePath & "已转换完成。(" & filename & ")")
            Sheets(i).Copy
            Sheets(i).Cells.Copy
            Sheets(i).Cells.PasteSpecial Paste = xlPasteValues
            Sheets(i).Cells.PasteSpecial Paste = xlPasteFormats
            OutputPath = objExcel.Path & "/out/"
            If Not fso.FolderExists(OutputPath) Then
                fso.CreateFolder OutputPath
            End If
            SavePath = OutputPath & Replace(objExcel.Name,".xlsx", "") & "-" & Sheets(i).Name & ".xlsx"
            ExcelApp.ActiveWorkbook.SaveAs SavePath
            ExcelApp.ActiveWorkbook.Close savechanges = False
            End If
        Next
        LogOut("文档" & FilePath & "已转换完成。(" & Finished & "/" & Sum & ")")
        ExcelApp.ActiveWorkbook.Close savechanges = False
        Finished = Finished + 1
    End If
loop

'扫尾处理开始
List.close
LogOut("转换已完成")
LogFile.close 

Dim Msg
Msg = "已成功转换" & Finished & "个文件"
MsgBox Msg & vbCrlf & "日志文件在" & fso.GetFolder(Path).Path & "\log.txt"
Set fso = nothing
ExcelApp.Quit
Wscript.Quit
