REM  *****  BASIC  *****

Sub ZoteroCS_to_orderTable

Dim oActiveSheet As Object
Dim objRange As Object
Dim oZeile As Object
Dim objCursor As Object

Dim max_row As Integer


oActiveSheet = ThisComponent.CurrentController.getActiveSheet()

'いらない列削除
oActiveSheet.getColumns.removeByIndex(0,2)
oActiveSheet.getColumns.removeByIndex(3,1)
oActiveSheet.getColumns.removeByIndex(4,15)
oActiveSheet.getColumns.removeByIndex(5,3)
oActiveSheet.getColumns.removeByIndex(6,8)
oActiveSheet.getColumns.removeByIndex(7,3)
oActiveSheet.getColumns.removeByIndex(8,20)
oActiveSheet.getColumns.removeByIndex(9,26)

'列の順番を整理
'まず最終行の数を獲得
objRange = oActiveSheet.getCellRangeByName("A1")
objCursor = oActiveSheet.createCursorByRange(objRange)
objCursor.gotoEndOfUsedArea(True)
max_row = objCursor.Rows.Count

'挿入先の列作成
oActiveSheet.getColumns.insertByIndex(1,1)


Dim srcRangeAdr1 As Object
Dim destCellAdr1 As Object
Dim x1 As Integer

For x1=0 To max_row step 1

srcRangeAdr1 = oActiveSheet.getCellByPosition(3,x1,3,max_row).getRangeAddress() 'コピー元の情報"Title"
destCellAdr1 = oActiveSheet.getCellByPosition(1,x1).getCellAddress() 'コピー先のアドレス
oActiveSheet.copyRange(destCellAdr1,srcRangeAdr1) 'コピー


Next x1


oActiveSheet.getColumns.insertByIndex(3,1)


Dim srcRangeAdr2 As Object
Dim destCellAdr2 As Object
Dim x2 As Integer

For x2=0 To max_row step 1
srcRangeAdr2 = oActiveSheet.getCellByPosition(10,x2,10,max_row).getRangeAddress() 'コピー元の情報"Edition"
destCellAdr2 = oActiveSheet.getCellByPosition(3,x2).getCellAddress() 'コピー先のアドレス
oActiveSheet.copyRange(destCellAdr2,srcRangeAdr2) 'コピー

Next x2

Dim srcRangeAdr3 As Object
Dim destCellAdr3 As Object
Dim x3 As Integer

For x3=0 To max_row step 1
srcRangeAdr3 = oActiveSheet.getCellByPosition(0,x3,0,max_row).getRangeAddress() 'コピー元の情報"publication year"
destCellAdr3 = oActiveSheet.getCellByPosition(4,x3).getCellAddress() 'コピー先のアドレス
oActiveSheet.copyRange(destCellAdr3,srcRangeAdr3) 'コピー

Next x3

oActiveSheet.getColumns.insertByIndex(5,1)

Dim srcRangeAdr4 As Object
Dim destCellAdr4 As Object
Dim x4 As Integer

For x4=0 To max_row step 1
srcRangeAdr4 = oActiveSheet.getCellByPosition(8,x4,8,max_row).getRangeAddress() 'コピー元の情報"Publisher"
destCellAdr4 = oActiveSheet.getCellByPosition(5,x4).getCellAddress() 'コピー先のアドレス
oActiveSheet.copyRange(destCellAdr4,srcRangeAdr4) 'コピー

Next x4 

oActiveSheet.getColumns.insertByIndex(4,1)

Dim srcRangeAdr6 As Object
Dim destCellAdr6 As Object
Dim x6 As Integer

For x6=0 To max_row step 1
srcRangeAdr6 = oActiveSheet.getCellByPosition(8,x6,8,max_row).getRangeAddress() 'コピー元の情報"Series"
destCellAdr6 = oActiveSheet.getCellByPosition(4,x6).getCellAddress() 'コピー先のアドレス
oActiveSheet.copyRange(destCellAdr6,srcRangeAdr6) 'コピー

Next x6 

oActiveSheet.getColumns.removeByIndex(8,2)

Dim srcRangeAdr5 As Object
Dim destCellAdr5 As Object
Dim x5 As Integer

For x5=0 To max_row step 1
srcRangeAdr5 = oActiveSheet.getCellByPosition(8,x5,8,max_row).getRangeAddress() 'コピー元の情報"Extra"
destCellAdr5 = oActiveSheet.getCellByPosition(10,x5).getCellAddress() 'コピー先のアドレス
oActiveSheet.copyRange(destCellAdr5,srcRangeAdr5) 'コピー

Next x5

oActiveSheet.getColumns.removeByIndex(8,1)

'A列の内容をクリア
Dim objRange2 As Object
objRange2 = oActiveSheet.getCellRangeByPosition(0, 1, 0, max_row)
objRange2.clearContents(511)

'注文番号の列作成
oActiveSheet.getCellRangeByName("A1").String = "注文番号"
oActiveSheet.getCellRangeByName("B1").String = "タイトル"
oActiveSheet.getCellRangeByName("C1").String = "著者"
oActiveSheet.getCellRangeByName("D1").String = "版"
oActiveSheet.getCellRangeByName("E1").String = "シリーズ名"
oActiveSheet.getCellRangeByName("F1").String = "出版年"
oActiveSheet.getCellRangeByName("G1").String = "出版社"
oActiveSheet.getCellRangeByName("I1").String = "備考"

'値段の列作成
oActiveSheet.getCellRangeByName("J1").String = "値段"
oActiveSheet.getCellRangeByName("K1").String = "送料込み値段"


End Sub

