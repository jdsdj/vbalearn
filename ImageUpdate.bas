Function ImageUpdate(rng As Range, value2 As String, imgSuffix As String, value4 As Integer, value5 As Integer)
  x = rng.Row
  y = rng.Column
  On Error Resume Next
  Application.ScreenUpdatinating = False '关闭屏幕更新
  Dim img As Shape
  For Each img In In ActiveSheet.Shapes  '删除原有图片
    If Not Application.Intersect(img.TopLeftCell, Cells(x, y).Offset(value5, value4)) Is Nothing Then
    img.Delete
    End If
  Next
  If Not IsEmpty(Cells(x, y)) And Dir(value2 & "\" & Cells(x, y).Value & "." & imgSuffix) <> "" Then
    Cells(x, y).Select
    ML = Cells(x, y).Offset(value5, value4).Left
    MT = Cells(x, y).Offset(value5, value4).Top
    MW = Cells(x, y).Offset(value5, value4).Width
    MH = Cells(x, y).Offset(value5, value4).Height
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, ML, MT, MW, MH).Select '添加图片图片
    Selection.ShapeRange.Fill.UserPicture value2 & "\" & Cells(x, y).Value & "." & value3 '当前文件所在目录下以当前单元内容为名称的.jpg图片
  End If
  Application.ScreenUpdatinating = True '开启屏幕更新
  tupian = "OK"
End Function
'ImageUpdate(A2,"e:\12","gif",4,1)  QQ:16846067
'1、单元格内容对应的图片名 2、图片地址 3、图片格式 4、左右偏移几个单元格 5 上下偏移位置
