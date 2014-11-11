Attribute VB_Name = "CommonFunctions"
Option Explicit


'始端と終端を線で結ぶ
Public Sub ConnecteStraightLineStartEnd(ForStartShape As Shape, ForEndShape As Shape)
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
        .Parent.Name = "Line1"
    End With
    ActiveSheet.Shapes("Line1").ZOrder msoSendToBack
End Sub

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


'オートシェイプをFor文の中で等間隔に配置する
Public Sub EquallySpacedAutoShapeForStatement()
        
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
                Top = Selection.ShapeRange(Myshp).Top + Selection.ShapeRange(Myshp).Height
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
        
    
    '全処理シェイプの長さの和が｢始端の下｣から｢終端の上｣までの長さを超える場合
    '始端から終端を合計の長さより多く取る
    Dim j As Integer
    Dim TotalLength As Integer
    For j = 1 To Selection.ShapeRange.Count
        '｢処理｣シェイプを等間隔に配置
        If Selection.ShapeRange(j).Name = "ProcShape" Then
            TotalLength = TotalLength + (Selection.ShapeRange(j).Top + Selection.ShapeRange(j).Height)
        End If
    Next
    If TotalLength >= Top + Bottom Then
        Selection.ShapeRange("ForEndShape").Top = Selection.ShapeRange("ForEndShape").Top + TotalLength
        Bottom = Selection.ShapeRange("ForEndShape").Top
    End If
    
    '等間隔に配置する
    Dim i As Integer
    Dim ProcShapeIndex As Integer
    ProcShapeIndex = 1
    For i = 1 To Selection.ShapeRange.Count
        '｢処理｣シェイプを等間隔に配置
        If Selection.ShapeRange(i).Name = "ProcShape" Then
            Selection.ShapeRange(i).Top = Top + (Bottom - Top) / (TargetCount + 1) * ProcShapeIndex
            ProcShapeIndex = ProcShapeIndex + 1
        End If
    Next

    
End Sub


