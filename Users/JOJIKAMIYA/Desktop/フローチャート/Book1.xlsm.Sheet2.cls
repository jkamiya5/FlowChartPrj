VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub フロー図生成_Click()
  
    Dim delIdx As Long
    delIdx = 1
    '全てのシェイプが処理対象
    Do Until delIdx > ActiveSheet.Shapes.Count
        With ActiveSheet
            '四角形のときは削除する
            If .Shapes(delIdx).Name <> "フロー図生成" Then
               .Shapes(delIdx).Delete
               '要素番号は進めない
            Else
               '要素番号を次に進める
               delIdx = delIdx + 1
            End If
        End With
    Loop
    
    
    Dim i As Long
    Dim LateralFrameSize As Long
    i = Range("AG1").Row    ' 1
    LateralFrameSize = Range("AG1").Column ' 1
    Debug.Print (i)
    Debug.Print (LateralFrameSize)
     
    Dim s As String
    s = Range("A1").Address ' "$A$1"
    Debug.Print (s)
  
    Dim oShape1 As Shape 'オートシェイプ保存用変数は、Shape型にする
    Dim oShape2 As Shape 'オートシェイプ保存用変数は、Shape型にする
    Dim commentShape As Shape 'オートシェイプ保存用変数は、Shape型にする
    'Dim oShape4 As Shape 'オートシェイプ保存用変数は、Shape型にする
    Set oShape1 = ActiveSheet.Shapes.AddShape(3, 100, 300, 120, 50)
    Set oShape2 = ActiveSheet.Shapes.AddShape(3, 100, 700, 120, 50)
    Set commentShape = ActiveSheet.Shapes.AddShape(1, 230, 310, 320, 20)
    commentShape.Name = "commentShape"
    
    oShape2.Rotation = 180
    
    oShape1.Fill.ForeColor.RGB = RGB(255, 255, 255)
    oShape2.Fill.ForeColor.RGB = RGB(255, 255, 255)
    commentShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    oShape1.Line.ForeColor.RGB = RGB(0, 0, 0)
    oShape2.Line.ForeColor.RGB = RGB(0, 0, 0)
    commentShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    
    oShape1.Line.Weight = 1
    oShape2.Line.Weight = 1
    commentShape.Line.Weight = 1
    
    commentShape.Line.Visible = True
    commentShape.Line.Transparency = 1
    commentShape.Fill.Transparency = 1
    commentShape.TextFrame.Characters.Text = "ここにFor文の条件を書く"
    commentShape.TextFrame.Characters.Font.Color = 1
    commentShape.TextFrame.HorizontalAlignment = xlLeft
    
    
    Dim right As Integer
    right = ActiveSheet.Shapes("commentShape").Left + ActiveSheet.Shapes("commentShape").Width
    If right = LateralFrameSize Then
        MsgBox "右端を超えています"
    End If
    
End Sub
