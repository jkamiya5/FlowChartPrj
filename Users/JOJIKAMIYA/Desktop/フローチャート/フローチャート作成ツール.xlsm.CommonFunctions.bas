Attribute VB_Name = "CommonFunctions"
Option Explicit


Public Type Position
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
    Width As Integer
    Height As Integer
End Type

'すべての図形を削除する
Public Sub ClearAllShapes()
    Dim delIdx As Long
    delIdx = 1
    'すべての図形が処理対象
    Do Until delIdx > ActiveSheet.Shapes.Count
        With ActiveSheet
            .Shapes(delIdx).Delete
        End With
        delIdx = delIdx + 1
    Loop
End Sub
'指定した図形以外削除する
Public Sub ClearSpecificShapes(objName() As String)
    Dim delIdx As Long
    delIdx = 1
    'すべての図形が処理対象
    Do Until delIdx > ActiveSheet.Shapes.Count
        With ActiveSheet
            '｢フロー図生成｣というオブジェクト名は対象外
            Dim i As Integer
            For i = LBound(objName) To UBound(objName)
                If .Shapes(delIdx).Name = objName(i) Then
                   delIdx = delIdx + 1
                   Exit For
                Else
                   .Shapes(delIdx).Delete
                End If
                i = i + 1
            Next
        End With
    Loop
End Sub

'オートシェイプの初期値設定
Public Sub InitialValueAutoShape(targetShape As Shape)
    targetShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
    targetShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    targetShape.Line.Weight = 1
End Sub

'オートシェイプを指定のセルを軸にしてセンタリングする
Public Sub CenteringForGivenColumn(CellName As String)
    'オートシェイプの位置設定
    'グループ化してD10列にセンタリング
    Dim myShRange As ShapeRange
    '---全ての図形を選択
    ActiveSheet.Shapes.SelectAll
    '---選択図形範囲を変数に格納
    Set myShRange = Selection.ShapeRange
    With myShRange
        .Group                         '---グループ化
        .Top = Range(CellName).Top     '---図形範囲の上端位置をB2の位置へ
        .Left = Range(CellName).Left   '---図形範囲の左端位置をB2の位置へ
        .Ungroup                       '---グループ化解除
    End With
End Sub


'オートシェイプを指定のセルを軸にしてセンタリングする
Public Sub CenteringForGivenRow()
    
    
    Dim Top As Integer
    Dim Bottom As Integer
    Dim TargetCount As Integer
    Top = 0
    Bottom = 0
    TargetCount = 0
    
    'アクティブなオートシェイプを取得
    ActiveSheet.Shapes.SelectAll
    Dim Myshp As Variant
    For Myshp = 1 To Selection.ShapeRange.Count
        If Selection.ShapeRange(Myshp).Name <> "CommentShape" _
            And Selection.ShapeRange(Myshp).Name <> "Line" Then
            If Selection.ShapeRange(Myshp).Name = "ForStartShape" Then
                '始端の位置を取得
                Top = Selection.ShapeRange(Myshp).Top
            ElseIf Selection.ShapeRange(Myshp).Name = "ForEndShape" Then
                '終端の位置を取得
                Bottom = Selection.ShapeRange(Myshp).Top
            End If
            If Selection.ShapeRange(Myshp).Name = "ProcShape" Then
                '｢処理｣シェイプの数取得
                TargetCount = TargetCount + 1
            End If
        End If
    Next
    
    Dim i As Integer
    i = 1
    For Myshp = 1 To Mycount
        '｢処理｣シェイプを等間隔に配置
        If Selection.ShapeRange(Myshp).Name = "ProcShape" Then
            Selection.ShapeRange(Myshp).Top = Top + (Bottom - Top) / (TargetCount + 1) * i
            i = i + 1
        End If
    Next
End Sub


