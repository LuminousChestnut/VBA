
Dim n%
For n = 2 To 19
If Cells(n,2) < 60 Then
	Cells(n,2).Interior.ColorIndex = 3
End If
Next

For n = 1 to 3
	For y = 1 to 10
	'..
	Next y
Next n

inputbox("提示语")

If n > x Then 
	Msgbox ""
Else
	End if

Workbooks.Add() '新增工作簿
Workbooks.Open("路径") '打开工作簿 
ActiveWorkbook.Close()  '关闭活动工作簿
Worksheets.Add() '新增工作表

Range("").Active '活跃单元格
Range("").Copy '复制单元格
Range("").Delete '删除单元格
Range("").Clear '清除单元格
Range("").Select '选择单元格
Selection.ClearContents  '清除内容
Range("A1").Cut [A13] '剪切单元格
ActiveCell '活动单元格
Sheets().Delete   '删除工作表
Msgbox ActiveSheet.* '活动工作表

```
Path 路径
Eg. [B1] = Workbooks('*.xlsm').Path （工作簿路径）
Count 工作表数量
Address 地址（绝对）
Interior.ColorIndex 颜色
```
