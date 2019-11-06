REM  *****  BASIC  *****

Sub Shimashima


Dim oActiveSheet As Object
Dim objRange As Object
Dim oZeile As Object
Dim objCursor As Object
Dim objRange2 As Object

Dim max_row As Integer
'Dim max_column As Integer
Dim x As Integer
Dim y As Integer

oActiveSheet = ThisComponent.CurrentController.getActiveSheet()
objRange = oActiveSheet.getCellRangeByName("A1")
objCursor = oActiveSheet.createCursorByRange(objRange)
objCursor.gotoEndOfUsedArea(True)


max_row = objCursor.Rows.Count
'max_column =  objCursor.Columns.Count
objRange2 = oActiveSheet.getCellRangeByPosition(0, 0, 10, max_row)

Dim objBorderLineBlack As New com.sun.star.table.BorderLine
Dim objTableBorder As New com.sun.star.table.TableBorder

With objBorderLineBlack
              .Color = RGB(0, 0, 0)
              .OuterLineWidth = 30
       End With
       
With objTableBorder
              '
              '表の外枠の設定
              .LeftLine = objBorderLineBlack
              .RightLine = objBorderLineBlack
              .TopLine = objBorderLineBlack
              .BottomLine = objBorderLineBlack
              .IsLeftLineValid = True
              .IsRightLineValid = True
              .IsTopLineValid = True
              .IsBottomLineValid = True
              '
              '表内の横罫線の設定
              .HorizontalLine = objBorderLineBlack
              .IsHorizontalLineValid = True
              '
              '表内の縦罫線の設定
              .VerticalLine = objBorderLineBlack
              .IsVerticalLineValid = True
       End With
       '
objRange2.TableBorder = objTableBorder

For x=2 To max_row Step 2

oZeile = oActiveSheet.getCellRangeByPosition(0, x, 10, x)
oZeile.CellBackColor = RGB(242, 242, 242)

Next x

For y=1 To max_row Step 2

oZeile = oActiveSheet.getCellRangeByPosition(0, y, 10, y)
oZeile.CellBackColor = RGB(255, 255, 255)

Next y


End Sub


