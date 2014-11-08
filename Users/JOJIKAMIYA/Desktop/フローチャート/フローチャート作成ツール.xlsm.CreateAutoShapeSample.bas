Attribute VB_Name = "CreateAutoShapeSample"
'For���[�v����\���I�[�g�V�F�C�v�̐���
Public Sub CreateForLoopObj(ForLoopCondition As String)

    
    '�E�[�̗�ʒu�擾
    Dim LateralFrameSize As Long ' 1
    LateralFrameSize = Range("AG1").Column ' 1
        
    '�I�[�g�V�F�C�v�ۑ��p�ϐ��́AShape�^�ɂ���
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
    
    
    '�n�[�ƏI�[����Ō���
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
        .Parent.Name = "Line"
    End With
End Sub

'������\���I�[�g�V�F�C�v�̐���
Public Sub CreateProcObj(ProcDescription As String)
        
    '�I�[�g�V�F�C�v�ۑ��p�ϐ��́AShape�^�ɂ���
    Dim ProcShape As Shape
    Set ProcShape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 100, 300, 120, 50)
    ProcShape.Name = "ProcShape"
    
    '�I�[�g�V�F�C�v�̏����l�ݒ�
    Call CommonFunctions.InitialValueAutoShape(ProcShape)
    '�����I�[�g�V�F�C�v�p�ɐ}�`�̃I�v�V�������`����
    ProcShape.Line.Visible = True
    ProcShape.TextFrame.Characters.Font.Color = 1
    ProcShape.TextFrame.HorizontalAlignment = xlLeft
    ProcShape.TextFrame.Characters.Text = ProcDescription
    ProcShape.TextFrame.AutoSize = True
        
End Sub

