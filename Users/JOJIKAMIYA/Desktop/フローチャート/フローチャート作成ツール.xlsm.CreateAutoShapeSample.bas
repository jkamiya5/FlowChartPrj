Attribute VB_Name = "CreateAutoShapeSample"
Public Sub CreateForLoopObj(ForLoopCondition As String)

    
    '右端の列位置取得
    Dim LateralFrameSize As Long ' 1
    LateralFrameSize = Range("AG1").Column ' 1
        
    'オートシェイプ保存用変数は、Shape型にする
    Dim ForStartShape As Shape
    Dim ForEndShape As Shape
    Dim CommentShape As Shape
    
    Set ForStartShape = ActiveSheet.Shapes.AddShape(3, 100, 300, 120, 50)
    Set ForEndShape = ActiveSheet.Shapes.AddShape(3, 100, 700, 120, 50)
    Set CommentShape = ActiveSheet.Shapes.AddShape(1, 230, 310, 320, 20)
    CommentShape.Name = "CommentShape"
        
        
    'オートシェイプの初期値設定
    Call CommonFunctions.InitialValueAutoShape(ForStartShape)
    Call CommonFunctions.InitialValueAutoShape(ForEndShape)
    Call CommonFunctions.InitialValueAutoShape(CommentShape)
    
    '終端オートシェイプを反転
    ForEndShape.Rotation = 180
    'コメント用オートシェイプの設定
    CommentShape.Line.Visible = True
    CommentShape.Line.Transparency = 1
    CommentShape.Fill.Transparency = 1
    CommentShape.TextFrame.Characters.Font.Color = 1
    CommentShape.TextFrame.HorizontalAlignment = xlLeft
    CommentShape.TextFrame.Characters.Text = ForLoopCondition
    
    Call CommonFunctions.図形の位置取得(ForStartShape)
    
    '始端と終端を線で結ぶ
    Dim rngStart As Range, rngEnd As Range
    Dim BX As Single, BY As Single, EX As Single, EY As Single
    'Shapeを配置するための基準となるセル
    Set rngStart = Range("B2")
    Set rngEnd = Range("J2")
    'セルのLeft、Top、Widthプロパティを利用して位置決め
    BX = rngStart.Left
    BY = rngStart.Top
    EX = rngEnd.Left + rngEnd.Width
    EY = rngEnd.Top
    '直線
    ActiveSheet.Shapes.AddLine BX, BY, EX, EY
    '赤色・太さ1.5ポイントの矢印線
    With ActiveSheet.Shapes.AddLine(BX, BY + 10, EX, EY + 10).Line
        .ForeColor.RGB = vbRed
        .Weight = 1.5
        .EndArrowheadStyle = msoArrowheadTriangle
    End With
End Sub

