Attribute VB_Name = "CommonFunctions"
Option Explicit
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


Public Sub 図形の位置取得(targetShape As Shape)
  Dim 左 As Double, 上 As Double, 横幅 As Double, 縦幅 As Double
  Dim 右 As Double, 下 As Double
  If TypeName(targetShape) = "Range" Then
    MsgBox "図形が選択されていません", vbCritical
    Exit Sub
  End If
  With targetShape.ShapeRange
    If .Count > 1 Then
      MsgBox "図形は複数選択しないでください", vbCritical
      Exit Sub
    End If
    左 = .Left                               'セル の左位置
    上 = .Top                               'セル の上位置
    横幅 = .Width                          'セル の横幅
    縦幅 = .Height                         'セル の縦幅
    右 = 左 + 横幅                        'セル の右位置
    下 = 上 + 縦幅                        'セル の下位置
    MsgBox "左位置は　" & Format(左, "0.00") & "　です" & vbCrLf & _
    "上位置は　" & Format(上, "0.00") & "　です" & vbCrLf & _
    "右位置は　" & Format(右, "0.00") & "　です" & vbCrLf & _
    "下位置は　" & Format(下, "0.00") & "　です", Title:="図形の位置（ポイント）"
  End With
End Sub
