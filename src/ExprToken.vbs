class ExprToken
    ''' �����̃g�[�N����\���N���X
    ''' �v���C�x�[�g�ϐ� typ �́A�g�[�N���̌^��\��
    ''' �v���C�x�[�g�ϐ� lex �́A�g�[�N���̌�b�f��\��
    private typ
    private lex
    private head
    private tail
    
    public function init(byval t, byval l, byval hd, byval tl)
        ''' �g�[�N��������������
        ''' ������ t �́A�g�[�N���̌^ (String) ���󂯎��
        ''' ������ l �́A�g�[�N���̌�b�f (String) ���󂯎��B
        ''' �߂�l�Ƃ��āA�������g�ւ̎Q�� (ExprToken) ��Ԃ��B
        typ = t
        lex = l
        head = hd
        tail = tl
        set init = me
    end function
    
    public function getPos()
        getPos = array(head, tail)
    end function
    
    public function getLex()
        ''' �g�[�N���̌�b�f��Ԃ��B
        ''' �߂�l�Ƃ��āA�g�[�N���̌�b�f (String) ��Ԃ��B
        getLex = lex
    end function
    
    public function getTyp()
        ''' �g�[�N���̌^��Ԃ��B
        ''' �߂�l�Ƃ��āA�g�[�N���̌^ (String) ��Ԃ��B
        getTyp = typ
    end function
    
    public function setLex(byval l)
        ''' �g�[�N���̌�b�f��ݒ肷��B
        ''' �߂�l��Ԃ��Ȃ��B
        setLex = l
    end function
    
    public function setTyp(byval t)
        ''' �g�[�N���̌^��ݒ肷��B
        ''' �߂�l��Ԃ��Ȃ��B
        typ = t
    end function
    
    public function toString()
        ''' ������ŕ\�������g�[�N����Ԃ��B
        ''' �߂�l�Ƃ��āA������ŕ\�������g�[�N�� (String) ��Ԃ��B
        toString = "<" & typ & "," & lex & ">"
    end function
end class
