VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub 繰り返しボタン_Click()
  Worksheets("Sheet2").Activate
  ''Worksheets("Sheet2").Range(Cells(1, 1), Cells(1000, 1000)).Clear
  Do While Sheet2.Shapes.Count > 0
    Sheet2.Shapes(0).Delete
  Loop
  
  Dim oShape1 As Shape 'オートシェイプ保存用変数は、Shape型にする
  Dim oShape2 As Shape 'オートシェイプ保存用変数は、Shape型にする
  Dim oShape3 As Shape 'オートシェイプ保存用変数は、Shape型にする
  'Dim oShape4 As Shape 'オートシェイプ保存用変数は、Shape型にする
  Set oShape1 = ActiveSheet.Shapes.AddShape(3, 100, 300, 120, 50)
  Set oShape2 = ActiveSheet.Shapes.AddShape(3, 100, 700, 120, 50)
  Set oShape3 = ActiveSheet.Shapes.AddShape(1, 230, 310, 320, 20)
  
  oShape2.Rotation = 180
  
  oShape1.Fill.ForeColor.RGB = RGB(255, 255, 255)
  oShape2.Fill.ForeColor.RGB = RGB(255, 255, 255)
  oShape3.Fill.ForeColor.RGB = RGB(255, 255, 255)
  
  oShape1.Line.ForeColor.RGB = RGB(0, 0, 0)
  oShape2.Line.ForeColor.RGB = RGB(0, 0, 0)
  oShape3.Line.ForeColor.RGB = RGB(0, 0, 0)
  
  oShape1.Line.Weight = 1
  oShape2.Line.Weight = 1
  oShape3.Line.Weight = 1
  
  oShape3.Line.Visible = True
  oShape3.Line.Transparency = 1
  oShape3.Fill.Transparency = 1
  oShape3.TextFrame.Characters.Text = "ここにFor文の条件を書く"
  oShape3.TextFrame.Characters.Font.Color = 1
  oShape3.TextFrame.HorizontalAlignment = xlLeft
  
  
End Sub
