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
    
    ForLoopCondition(0) = "○○が～の間"
    Call CreateAutoShapeSample.CreateForLoopObj(ForLoopCondition(0))
    ProcDescription(0) = "●●という処理を行う"
    ProcDescription(1) = "cccという処理を行う"
    ProcDescription(2) = "dddddddddddddddddddddddddddddddddddddddddddddddddddddddという処理を行う"
    
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(0))
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(1))
    Call CreateAutoShapeSample.CreateProcObj(ProcDescription(2))
        
    Call CommonFunctions.CenteringForGivenColumn("B5")
    Call CommonFunctions.CenteringForGivenRow
    

    
End Sub


