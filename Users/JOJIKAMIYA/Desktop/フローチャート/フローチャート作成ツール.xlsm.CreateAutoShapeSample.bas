Attribute VB_Name = "CreateAutoShapeSample"
'Forループ文を表すオートシェイプの生成
Public Sub CreateForLoopObj(ForLoopCondition As String)

    
    '右端の列位置取得
    Dim LateralFrameSize As Long ' 1
    LateralFrameSize = Range("AG1").Column ' 1
        
    'オートシェイプ保存用変数は、Shape型にする
    Dim ForStartShape As Shape
    Dim ForEndShape As Shape
    Dim CommentShape As Shape
    Dim LineShape As Shape
    
    Set ForStartShape = ActiveSheet.Shapes.AddShape(3, 100, 300, 120, 50)
    Set ForEndShape = ActiveSheet.Shapes.AddShape(3, 100, 700, 120, 50)
    Set CommentShape = ActiveSheet.Shapes.AddShape(1, 230, 310, 320, 20)
    ForStartShape.Name = "ForStartShape"
    ForEndShape.Name = "ForEndShape"
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
    
    
    '始端と終端を線で結ぶ
    Dim rngStart As Range, rngEnd As Range
    Dim BX As Single, BY As Single, EX As Single, EY As Single
    'セルのLeft、Top、Widthプロパティを利用して位置決め
    BX = ForStartShape.Left + (ForStartShape.Width / 2)
    BY = ForStartShape.Top + ForStartShape.Height
    EX = ForEndShape.Left + (ForEndShape.Width / 2)
    EY = ForEndShape.Top
    '直線
    With ActiveSheet.Shapes.AddLine(BX, BY, EX, EY).Line
        .ForeColor.RGB = vbBlack
        .Weight = 1
        .Parent.Name = "Line"
    End With
End Sub

'処理を表すオートシェイプの生成
Public Sub CreateProcObj(ProcDescription As String)
        
    'オートシェイプ保存用変数は、Shape型にする
    Dim ProcShape As Shape
    Set ProcShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 100, 300, 120, 50)
    ProcShape.Name = "ProcShape"
    
    'オートシェイプの初期値設定
    Call CommonFunctions.InitialValueAutoShape(ProcShape)
    '処理オートシェイプ用に図形のオプションを定義する
    ProcShape.Line.Visible = True
    ProcShape.TextFrame.Characters.Font.Color = 1
    ProcShape.TextFrame.HorizontalAlignment = xlLeft
    ProcShape.TextFrame.Characters.Text = ProcDescription
    ProcShape.TextFrame.AutoSize = True
        
End Sub

