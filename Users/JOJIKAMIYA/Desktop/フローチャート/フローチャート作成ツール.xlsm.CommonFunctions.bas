Attribute VB_Name = "CommonFunctions"
Option Explicit


'�n�[�ƏI�[����Ō���
Public Sub ConnecteStraightLineStartEnd(ForStartShape As Shape, ForEndShape As Shape)
    Dim rngStart As Range, rngEnd As Range
    Dim BX As Single, BY As Single, EX As Single, EY As Single
    '�Z����Left�ATop�AWidth�v���p�e�B�𗘗p���Ĉʒu����
    BX = ForStartShape.Left + (ForStartShape.Width / 2)
    BY = ForStartShape.Top + ForStartShape.Height
    EX = ForEndShape.Left + (ForEndShape.Width / 2)
    EY = ForEndShape.Top
    '����
    With ActiveSheet.Shapes.AddLine(BX, BY, EX, EY).Line
        .ForeColor.RGB = vbBlack
        .Weight = 1
        .Parent.Name = "Line1"
    End With
    ActiveSheet.Shapes("Line1").ZOrder msoSendToBack
End Sub

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


'�I�[�g�V�F�C�v��For���̒��œ��Ԋu�ɔz�u����
Public Sub EquallySpacedAutoShapeForStatement()
        
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
                Top = Selection.ShapeRange(Myshp).Top + Selection.ShapeRange(Myshp).Height
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
        
    
    '�S�����V�F�C�v�̒����̘a����n�[�̉�����碏I�[�̏㣂܂ł̒����𒴂���ꍇ
    '�n�[����I�[�����v�̒�����葽�����
    Dim j As Integer
    Dim TotalLength As Integer
    For j = 1 To Selection.ShapeRange.Count
        '�������V�F�C�v�𓙊Ԋu�ɔz�u
        If Selection.ShapeRange(j).Name = "ProcShape" Then
            TotalLength = TotalLength + (Selection.ShapeRange(j).Top + Selection.ShapeRange(j).Height)
        End If
    Next
    If TotalLength >= Top + Bottom Then
        Selection.ShapeRange("ForEndShape").Top = Selection.ShapeRange("ForEndShape").Top + TotalLength
        Bottom = Selection.ShapeRange("ForEndShape").Top
    End If
    
    '���Ԋu�ɔz�u����
    Dim i As Integer
    Dim ProcShapeIndex As Integer
    ProcShapeIndex = 1
    For i = 1 To Selection.ShapeRange.Count
        '�������V�F�C�v�𓙊Ԋu�ɔz�u
        If Selection.ShapeRange(i).Name = "ProcShape" Then
            Selection.ShapeRange(i).Top = Top + (Bottom - Top) / (TargetCount + 1) * ProcShapeIndex
            ProcShapeIndex = ProcShapeIndex + 1
        End If
    Next

    
End Sub


