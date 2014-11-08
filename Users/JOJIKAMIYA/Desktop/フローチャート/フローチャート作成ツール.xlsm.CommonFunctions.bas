Attribute VB_Name = "CommonFunctions"
Option Explicit
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


Public Sub �}�`�̈ʒu�擾(targetShape As Shape)
  Dim �� As Double, �� As Double, ���� As Double, �c�� As Double
  Dim �E As Double, �� As Double
  If TypeName(targetShape) = "Range" Then
    MsgBox "�}�`���I������Ă��܂���", vbCritical
    Exit Sub
  End If
  With targetShape.ShapeRange
    If .Count > 1 Then
      MsgBox "�}�`�͕����I�����Ȃ��ł�������", vbCritical
      Exit Sub
    End If
    �� = .Left                               '�Z�� �̍��ʒu
    �� = .Top                               '�Z�� �̏�ʒu
    ���� = .Width                          '�Z�� �̉���
    �c�� = .Height                         '�Z�� �̏c��
    �E = �� + ����                        '�Z�� �̉E�ʒu
    �� = �� + �c��                        '�Z�� �̉��ʒu
    MsgBox "���ʒu�́@" & Format(��, "0.00") & "�@�ł�" & vbCrLf & _
    "��ʒu�́@" & Format(��, "0.00") & "�@�ł�" & vbCrLf & _
    "�E�ʒu�́@" & Format(�E, "0.00") & "�@�ł�" & vbCrLf & _
    "���ʒu�́@" & Format(��, "0.00") & "�@�ł�", Title:="�}�`�̈ʒu�i�|�C���g�j"
  End With
End Sub
