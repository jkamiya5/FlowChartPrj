VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim objName(1) As String
Dim ForLoopCondition(10) As String
Dim ProcDescription(10) As String


Private Sub フロー図生成_Click()

    objName(0) = "フロー図生成"
    '上記のオブジェクト以外削除
    Call CommonFunctions.ClearSpecificShapes(objName)
    
    ProcDescription(0) = "●●という処理を行う"
    ProcDescription(1) = "cccという処理を行う"
    ProcDescription(2) = "dddddddddddddddddddddddddddddddddddddddddddddddddddddddという処理を行う"
    ProcDescription(3) = "dddddddddddddddddddddddddddddddddddddddddddddddddddddddという処理を行う"
    ProcDescription(4) = "dddddddddddddddddddddddddddddddddddddddddddddddddddddddという処理を行う"
    ProcDescription(5) = "dddddddddddddddddddddddddddddddddddddddddddddddddddddddという処理を行う"
    ProcDescription(6) = "dddddddddddddddddddddddddddddddddddddddddddddddddddddddという処理を行う"
    ProcDescription(7) = "dddddddddddddddddddddddddddddddddddddddddddddddddddddddという処理を行う"
    ProcDescription(8) = "dddddddddddddddddddddddddddddddddddddddddddddddddddddddという処理を行う"
    '処理1を生成
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(0))
    '処理2を生成
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(1))
    '処理3を生成
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(2))
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(3))
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(4))
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(5))
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(6))
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(7))
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(8))
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(2))
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(3))
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(4))
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(5))
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(6))
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(7))
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(8))
    
    
    ForLoopCondition(0) = "○○が～の間"
    'Forループを作成
    Call CreateAutoShapeSample.CreateForLoopObj(ForLoopCondition(0))
        
    '横方向のレイアウトを整える
    Call CommonFunctions.CenteringForGivenColumn("B5")
    '縦方向のレイアウトを整える
    Call CommonFunctions.EquallySpacedAutoShapeForStatement
    
    ActiveSheet.Shapes.SelectAll
    '始端から終端まで線で結ぶ
    Call CommonFunctions.ConnecteStraightLineStartEnd(Selection.ShapeRange("ForStartShape"), Selection.ShapeRange("ForEndShape"))
    
    

    
End Sub


