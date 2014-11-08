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

'���ׂĂ̐}�`���폜����
Public Sub ClearAllShapes()
    Dim delIdx As Long
    delIdx = 1
    '���ׂĂ̐}�`�������Ώ�
    Do Until delIdx > ActiveSheet.Shapes.Count
        With ActiveSheet
            .Shapes(delIdx).Delete
        End With
        delIdx = delIdx + 1
    Loop
End Sub
'�w�肵���}�`�ȊO�폜����
Public Sub ClearSpecificShapes(objName() As String)
    Dim delIdx As Long
    delIdx = 1
    '���ׂĂ̐}�`�������Ώ�
    Do Until delIdx > ActiveSheet.Shapes.Count
        With ActiveSheet
            '��t���[�}������Ƃ����I�u�W�F�N�g���͑ΏۊO
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

'�I�[�g�V�F�C�v�̏����l�ݒ�
Public Sub InitialValueAutoShape(targetShape As Shape)
    targetShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
    targetShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    targetShape.Line.Weight = 1
End Sub

'�I�[�g�V�F�C�v���w��̃Z�������ɂ��ăZ���^�����O����
Public Sub CenteringForGivenColumn(CellName As String)
    '�I�[�g�V�F�C�v�̈ʒu�ݒ�
    '�O���[�v������D10��ɃZ���^�����O
    Dim myShRange As ShapeRange
    '---�S�Ă̐}�`��I��
    ActiveSheet.Shapes.SelectAll
    '---�I��}�`�͈͂�ϐ��Ɋi�[
    Set myShRange = Selection.ShapeRange
    With myShRange
        .Group                         '---�O���[�v��
        .Top = Range(CellName).Top     '---�}�`�͈͂̏�[�ʒu��B2�̈ʒu��
        .Left = Range(CellName).Left   '---�}�`�͈͂̍��[�ʒu��B2�̈ʒu��
        .Ungroup                       '---�O���[�v������
    End With
End Sub


'�I�[�g�V�F�C�v���w��̃Z�������ɂ��ăZ���^�����O����
Public Sub CenteringForGivenRow()
    
    
    Dim Top As Integer
    Dim Bottom As Integer
    Dim TargetCount As Integer
    Top = 0
    Bottom = 0
    TargetCount = 0
    
    '�A�N�e�B�u�ȃI�[�g�V�F�C�v���擾
    ActiveSheet.Shapes.SelectAll
    Dim Myshp As Variant
    For Myshp = 1 To Selection.ShapeRange.Count
        If Selection.ShapeRange(Myshp).Name <> "CommentShape" _
            And Selection.ShapeRange(Myshp).Name <> "Line" Then
            If Selection.ShapeRange(Myshp).Name = "ForStartShape" Then
                '�n�[�̈ʒu���擾
                Top = Selection.ShapeRange(Myshp).Top
            ElseIf Selection.ShapeRange(Myshp).Name = "ForEndShape" Then
                '�I�[�̈ʒu���擾
                Bottom = Selection.ShapeRange(Myshp).Top
            End If
            If Selection.ShapeRange(Myshp).Name = "ProcShape" Then
                '�������V�F�C�v�̐��擾
                TargetCount = TargetCount + 1
            End If
        End If
    Next
    
    Dim i As Integer
    i = 1
    For Myshp = 1 To Mycount
        '�������V�F�C�v�𓙊Ԋu�ɔz�u
        If Selection.ShapeRange(Myshp).Name = "ProcShape" Then
            Selection.ShapeRange(Myshp).Top = Top + (Bottom - Top) / (TargetCount + 1) * i
            i = i + 1
        End If
    Next
End Sub


