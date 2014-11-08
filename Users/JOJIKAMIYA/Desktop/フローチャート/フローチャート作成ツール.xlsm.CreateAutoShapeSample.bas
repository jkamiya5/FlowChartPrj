Attribute VB_Name = "CreateAutoShapeSample"
Public Sub CreateForLoopObj(ForLoopCondition As String)

    
    '�E�[�̗�ʒu�擾
    Dim LateralFrameSize As Long ' 1
    LateralFrameSize = Range("AG1").Column ' 1
        
    '�I�[�g�V�F�C�v�ۑ��p�ϐ��́AShape�^�ɂ���
    Dim ForStartShape As Shape
    Dim ForEndShape As Shape
    Dim CommentShape As Shape
    
    Set ForStartShape = ActiveSheet.Shapes.AddShape(3, 100, 300, 120, 50)
    Set ForEndShape = ActiveSheet.Shapes.AddShape(3, 100, 700, 120, 50)
    Set CommentShape = ActiveSheet.Shapes.AddShape(1, 230, 310, 320, 20)
    CommentShape.Name = "CommentShape"
        
        
    '�I�[�g�V�F�C�v�̏����l�ݒ�
    Call CommonFunctions.InitialValueAutoShape(ForStartShape)
    Call CommonFunctions.InitialValueAutoShape(ForEndShape)
    Call CommonFunctions.InitialValueAutoShape(CommentShape)
    
    '�I�[�I�[�g�V�F�C�v�𔽓]
    ForEndShape.Rotation = 180
    '�R�����g�p�I�[�g�V�F�C�v�̐ݒ�
    CommentShape.Line.Visible = True
    CommentShape.Line.Transparency = 1
    CommentShape.Fill.Transparency = 1
    CommentShape.TextFrame.Characters.Font.Color = 1
    CommentShape.TextFrame.HorizontalAlignment = xlLeft
    CommentShape.TextFrame.Characters.Text = ForLoopCondition
    
    Call CommonFunctions.�}�`�̈ʒu�擾(ForStartShape)
    
    '�n�[�ƏI�[����Ō���
    Dim rngStart As Range, rngEnd As Range
    Dim BX As Single, BY As Single, EX As Single, EY As Single
    'Shape��z�u���邽�߂̊�ƂȂ�Z��
    Set rngStart = Range("B2")
    Set rngEnd = Range("J2")
    '�Z����Left�ATop�AWidth�v���p�e�B�𗘗p���Ĉʒu����
    BX = rngStart.Left
    BY = rngStart.Top
    EX = rngEnd.Left + rngEnd.Width
    EY = rngEnd.Top
    '����
    ActiveSheet.Shapes.AddLine BX, BY, EX, EY
    '�ԐF�E����1.5�|�C���g�̖���
    With ActiveSheet.Shapes.AddLine(BX, BY + 10, EX, EY + 10).Line
        .ForeColor.RGB = vbRed
        .Weight = 1.5
        .EndArrowheadStyle = msoArrowheadTriangle
    End With
End Sub

