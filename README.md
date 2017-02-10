# vbaexcel

练习vba写法
excel2007

# download
修改完的文件名字为result.xlsm
点击result.xlsm文件，然后点击download按钮

# 介绍下基本用法
文件格式修改为xlsm: excel2007 xlxs格式不支持宏命令只能修改表数据
我是日文系统中文乱码: 修改了sheet页名字请勿动
在sheet页method中生成了一个按钮click点击即可会自动跳的sheet页result

我试了下基本符合要求 

# vba 代码
```ruby
Sub btn_click()
 Application.ScreenUpdating = False
 Dim i As Integer
 Dim j As Integer
 Dim k As Integer
 Dim y As Integer
 Dim arr()
 Dim brr()
 
 i = 1
 j = 1
 k = 2
 
 Sheets("result").Rows("2:" & Rows.Count).Clear
 
 Do Until Sheets("method").Cells(j, 1).Value = ""
 
  Do Until Sheets("data").Cells(i, 4).Value = ""
 
   If Sheets("method").Cells(j, 1).Value = Sheets("data").Cells(i, 4).Value Then
   
    Sheets("result").Range(Replace("A" & Str(k) & ":" & "V" & Str(k), " ", "")).Value = Sheets("data").Range(Replace("A" & Str(i) & ":" & "V" & Str(i), " ", "")).Value
    k = k + 1
   End If
   i = i + 1
  
  Loop
  i = 1
  j = j + 1
  
 Loop
 
 Sheets("result").Activate
 Application.ScreenUpdating = True

End Sub

```

或者点击alt+F8 查看

