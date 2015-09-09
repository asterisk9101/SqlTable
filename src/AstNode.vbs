class AstNode
    ''' ���ۍ\���؂̃m�[�h��\���N���X�B
    private children
    private token
    ''' �v���C�x�[�g�ϐ� children �́A���̃m�[�h�̎q���m�[�h�̃��X�g (ArrayList) ���i�[����B
    ''' �v���C�x�[�g�ϐ� token �́A���̃m�[�h�����g�[�N�� (ExprToken) ���i�[����B
    
    private sub Class_Initialize
        set children = new ArrayList
        set token = nothing
    end sub
    
    public function init(byval t)
        set token = t
        set init = me
    end function
    
    public function getToken()
        set getToken = token
    end function
    
    public function getChildren()
        set getChildren = children
    end function
    
    public function push(byval node)
        call children.push(node)
    end function
    
    public function getLastChild()
        set getLastChild = children.item(children.length() - 1)
    end function
    
    public function pop()
        set pop = children.item(children.length() - 1)
        call children.removeAt(children.length() - 1)
    end function
    
    private function isNil()
        isNil = token is nothing
    end function
    
    public function toStringTree()
        ''' �؍\�����|�[�����h�L�@�ŕ\������������ŕԂ��B
        ''' �߂�l�Ƃ��āA�؍\�����|�[�����h�L�@�ŕ\������������ (String) ��Ԃ��B
        if children.length() = 0 then
            toStringTree = me.toString()
            exit function
        end if
        
        dim buf, i, t
        
        if not isNil() then buf = buf & "(" & me.toString() & " "
        
        i = 0
        do while i < children.length()
            set t = children.item(i)
            if i > 0 then buf = buf & " "
            buf = buf & t.toStringTree()
            i = i + 1
        loop
        if not isNil() then buf = buf & ")"
        toStringTree = buf
    end function
    
    public function getPos()
        dim pos, item, childpos
        pos = token.getPos()
        
        for each item in children.toArray()
            childpos = item.getPos()
            if childpos(0) < pos(0) then pos(0) = childpos(0)
            if childpos(1) > pos(1) then pos(1) = childpos(1)
        next
        
        getPos = pos
    end function
    
    public function toString()
        dim ret
        if not token is nothing then
            ret = token.toString()
        else
            ret = "nil"
        end if
        toString = ret
    end function
end class
