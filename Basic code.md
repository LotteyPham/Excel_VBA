--Sưu tầm --

## Add Serial Numbers (đánh số thự tự tự động)

Code macro này sẽ giúp bạn bổ sung số serial tự động trên trang Excel.
Sau khi bạn chạy mã macro này, màn hình sẽ hiển thị input box để bạn nhập tối đa số serial và sau đó, nó sẽ chèn các số vào cột theo thứ tự.

```
Sub AddSerialNumbers()
Dim i As Integer
On Error GoTo Last
i = InputBox("Enter Value", "Enter Serial Numbers")
For i = 1 To i
ActiveCell.Value = i
ActiveCell.Offset(1, 0).Activate
Next i
Last:Exit Sub
End Sub
```

## Add Multiple Columns (chèn cột)

Sau khi chạy mã macro, màn hình sẽ hiển thị một input box và bạn phải nhập số cột mà bạn muốn chèn.

```
Sub InsertMultipleColumns()
Dim i As Integer
Dim j As Integer
ActiveCell.EntireColumn.Select
On Error GoTo Last
i = InputBox("Enter number of columns to insert", "Insert         Columns")
For j = 1 To i
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightorAbove
Next j
Last:Exit Sub
End Sub
```

## Add Multiple Rows (chèn dòng)

Sau khi chạy mã macro, màn hình sẽ hiển thị một input box và bạn phải nhập số hàng mà bạn muốn chèn.

```
Sub InsertMultipleRows()
Dim i As Integer
Dim j As Integer
ActiveCell.EntireRow.Select
On Error GoTo Last
i = InputBox("Enter number of columns to insert", "Insert Columns")
For j = 1 To i
Selection.Insert Shift:=xlToDown,
CopyOrigin:=xlFormatFromRightorAbove
Next j
Last:Exit Sub
End Sub                 
```

## Auto Fit Columns (tự động canh các cột)

Nhanh chóng tự động khớp tất cả các hàng trong worksheet của bạn.
Mã macro này sẽ chọn tất cả các ô trong worksheet và tự động khớp ngay lập tức các cột.

```
Sub AutoFitColumns()
Cells.Select
Cells.EntireColumn.AutoFit
End Sub
```

## Auto Fit Rows (tự động canh các dòng)

Bạn có thể sử dụng mã code này để tự động khớp tất cả các hàng trong worksheet.
Khi bạn chạy mã này, nó sẽ chọn tất cả các ô trong worksheet và tự động khớp ngay lập tức các hàng.

```
Sub AutoFitRows()
Cells.Select
Cells.EntireRow.AutoFit
End Sub
```

## Remove Text Wrap (bỏ chế độ wrap text)

Mã code này sẽ giúp bạn xóa text wrap khỏi toàn bộ worksheet với một cái nhấp chuột. Đầu tiên nó sẽ chọn tất cả các cột và sau đó xóa text wrap và tự động khớp các hàng và cột.

```
Sub RemoveWrapText()
Cells.Select
Selection.WrapText = False
Cells.EntireRow.AutoFit
Cells.EntireColumn.AutoFit
End Sub
```

## Unmerge Cells (không kết nối các ô)

Chọn các ô và chạy mã này, nó sẽ không sát nhập tất cả các ô vừa chọn với dữ liệu bị mất của bạn.                                            

```
Sub UnmergeCells()
Selection.UnMerge
End Sub
```

## Open Calculator (mở máy tính trên excel)

Trong cửa sổ có một máy tính cụ thể và sử dụng mã macro này, bạn có thể mở máy tính trực tiếp từ Excel cho việc tính toán.

```
Sub OpenCalculator()
Application.ActivateMicrosoftApp Index:=0
End Sub
```

## Add Header/Footer Date (thêm ngày ở chân trang/đầu trang)

Sử dụng mã này để bổ sung ngày vào phần header và footer trong worksheet.
Bạn có thể điều chỉnh mã này để đổi từ header sang footer.

```
Sub dateInHeader()
With ActiveSheet.PageSetup
.LeftHeader = ""
.CenterHeader = "&D"
.RightHeader = ""
.LeftFooter = ""
.CenterFooter = ""
.RightFooter = ""
End With
ActiveWindow.View = xlNormalView
End Sub
```

## Custom Header/Footer (chèn đầu trang/chân trang theo ý bạn)

Nếu bạn muốn chèn header tùy chỉnh thì đây là một mã dành cho bạn.
Chạy mã này, nhập giá trị tùy chỉnh vào input box. Để thay đổi liên kết của header hoặc footer, bạn có thể điều chỉnh mã.

```
Sub customHeader()
Dim myText As Stringmy
Text = InputBox("Enter your text here", "Enter Text")
With ActiveSheet.PageSetup
.LeftHeader = ""
.CenterHeader = myText
.RightHeader = ""
.LeftFooter = ""
.CenterFooter = ""
.RightFooter = ""
End With
End Sub 
```
