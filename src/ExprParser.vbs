'!require ExprLexer.vbs
'!require AstNode.vbs

class ExprParser
    ''' �\����͊�
    ''' ExprLexer ����͂����g�[�N�����󂯎��A���߂�ꂽ�\���ɉ����Ă��邩�m�F���Ȃ���\���؂��\�z����B
    private lex
    private token
    ''' �v���C�x�[�g�ϐ� lex �́A���������ꂽ�����͊� (ExprLexer) ���i�[����B
    ''' �v���C�x�[�g�ϐ� token �́A��͂���g�[�N�� (ExprToken) ���i�[����B
    
    public function init(byval lexer)
        set lex = lexer
        set token = lex.nextToken()
        set init = me
    end function
    
    private function genVirtualNode(byval name)
        set genVirtualNode = (new AstNode).init((new ExprToken).init(name, "", lex.getAt(), 0))
    end function
    
'    public function program()
'        dim node
'        set node = genVirtualNode("PROGRAM")
'        
'        call node.push(statements())
'        
'        set program = node
'    end function
'    
'    public function statements()
'        dim node
'        set node = genVirtualNode("STATEMENTS")
'        
'        dim t
'        t = token.getTyp()
'        do while t <> "EOF"
'            call node.push(statement())
'            t = token.getTyp()
'        loop
'        
'        set statements = node
'    end function
'    
'    public function statement()
'        dim node
'        set node = genVirtualNode("STATEMENTS")
'        
'        dim t
'        t = token.getTyp()
'        do while t <> "EOF" and t <> "RBRACE" and t <> "TERM"
'            if token.getTyp() = "WORD" then
'                select case token.getLex()
'                case "if"   call node.push(stat_if())
'                case "while"    call node.push(stat_while())
'                case "for"  call node.push(stat_for())
'                case "func" call node.push(stat_func())
'                case "return"   call node.push(stat_return())
'                case "var"  call node.push(stat_var())
'                case else   call node.push(expression())
'                end select
'            else
'                call node.push()
'            end if
'            t = token.getTyp()
'        loop
'        
'        set statement = node
'    end function
'    
'    public function stat_if()
'        dim node
'        set node = genVirtualNode("STAT_IF")
'        
'        call match("WORD") ' if
'        call match("LBRA")
'        call node.push(expr())
'        call match("RBRA")
'        call match("LBRACE")
'        call node.push(statements())
'        call match("RBRACE")
'        
'        set stat_if = node
'    end function
'    
'    public function stat_for()
'        dim node
'        set node = genVirtualNode("STAT_FOR")
'        
'        call match("WORD") ' for
'        
'        call match("LPAR")
'        call node.push(expression()) ' init
'        call match("TERM")
'        call node.push(expression()) ' cond
'        call match("TERM")
'        call node.push(expression()) ' iter
'        call match("TERM")
'        call match("RPAR")
'        
'        call match("LBRACE")
'        call node.push(statements())
'        call match("RBRACE")
'        
'        set stat_for = node
'    end function
'    
'    public function stat_while()
'        dim node
'        set node = genVirtualNode("STAT_WHILE")
'        
'        call match("WORD") ' while
'        call match("LPAR")
'        call node.push(expression())
'        call match("RPAR")
'        call match("RBRACE")
'        call node.push(statements())
'        call match("RBRACE")
'        
'        set stat_while = node
'    end function
'    
'    public function stat_func()
'        dim node
'        set node = genVirtualNode("STAT_FUNC")
'        
'        call match("WORD") ' func
'        call node.push(identifier())
'        call match("LPAR")
'        call node.push(arguments())
'        call match("RPAR")
'        call match("LBRACE")
'        call node.push(statements())
'        call match("RBRACE")
'        
'        set stat_func = node
'    end function
'    
'    public function stat_return()
'        dim node
'        set node = genVirtualNode("STAT_RETURN")
'        
'        call match("WORD") ' return
'        call node.push(expression())
'        if check("TERM") or check("EOF") or check("RBRACE") then
'            set token = lex.nextToken()
'            set stat_return = node
'        else
'            call err.raise(12345, TypeName(me), "")
'        end if
'        
'    end function
'    
'    public function stat_var()
'        dim node
'        set node = genVirtualNode("STAT_VAR")
'        
'        call match("WORD") ' var
'        call match("ASSIGN")
'        call node.push(expression())
'        if check("TERM") or check("EOF") or check("RBRACE") then
'            set token = lex.nextToken()
'            set stat_return = node
'        else
'            call err.raise(12345, TypeName(me), "")
'        end if
'        
'    end function
'    
'    public function stat_class()
'        ' 
'    end function
'    
'    public function expression()
'        dim node
'        set node = genVirtualNode("EXPRESSION")
'    end function
'    
'    public function arguments()
'    end function
'    
'    public function identifier()
'    end function
    
    public function expr()
        dim root, parent, child
        set root = genVirtualNode("EXPR")
        set expr = root ' �߂�l
        
        ' ��̎�
        if check("EOF") then exit function
        
        ' �l�����̎�
        call root.push(val())
        set parent = root
        set child = (new AstNode).init(token) ' ���Z�q�g�[�N�������҂����m�[�h
        
        do while isOP2(child.getToken())
            set parent = compare_priority(root, child)
            call child.push(parent.pop())
            call parent.push(child)
            
            set token = lex.nextToken()
            call child.push(val())
            set child = (new AstNode).init(token)
        loop
        
        set expr = root
    end function
    
    private function ary()
        dim node
        set node = genVirtualNode("ARRAY")
        set ary = node
        
        call match("LBRA")' ���J�b�R���X�L�b�v����
        
        ' ��̔z��
        if check("RBRA") then exit function
        
        ' �v�f 1 �ȏ�̔z��
        set node = expr_list(node)
        call match("RBRA")
        
        set ary = node
    end function
    
    private function expr_list(byval node)
        call node.push(expr())
        do while check("COMMA")
            set token = lex.nextToken()
            call node.push(expr())
        loop
        set expr_list = node
    end function
    
    private function word()
        dim node
        set node = (new AstNode).init(token)
        
        ' word(...) �Ȃ�֐��Ƃ��ď�������
        set token = lex.nextToken()
        if check("LPAR") then set node = func(node)
        
        set word = node
    end function
    
    private function func(byval node)
        call match("LPAR")
        set func = node ' �߂�l
        ' �����Ȃ��̊֐�
        if check("RPAR") then exit function
        set node = expr_list(node)
        call match("RPAR")
        set func = node
    end function
    
    private function val()
        dim node
        select case token.getTyp()
        case "LPAR"
            set node = paren_expr()
        case "LBRA"
            set node = ary()
        case "NOT"
            set node = (new AstNode).init(token)
            call node.push(val())
        case "NUMBER", "STRING", "REGEX", "DATE", "TRUE", "FALSE"
            set node = (new AstNode).init(token)
            set token = lex.nextToken()
        case "WORD"
            set node = WORD()
        case else
            call err.raise(12345, TypeName(me), "token is not value: " & token.toString())
        end select
        
        set val = node
    end function
    
    private function paren_expr()
        dim node
        
        call match("LPAR")
        set node = expr()
        call match("RPAR")
        
        set paren_expr = node
    end function
    
    private function compare_priority(byval parent, byval newChild)
        ''' �g�[�N���̗D�揇�ʂ��r����B
        dim oldChild, oldChild_p, newChild_p
        set oldChild = parent.getLastChild()
        oldChild_p = priority(oldChild.getToken())
        newChild_p = priority(newChild.getToken())
        
        if oldChild_p < newChild_p then
            set parent = compare_priority(oldChild, newChild)
        end if
        
        set compare_priority = parent
    end function
    
    private function priority(byval token)
        select case token.getLex()
        case "(", ")"
            priority = 100
        case "!"
            priority = 90
        case "*", "/", "%"
            priority = 80
        case "+", "-"
            priority = 70
        case ">", "<", ">=", "<=", "==", "!=", "~", "!~"
            priority = 60
        case "&"
            priority = 50
        case "|"
            priority = 40
        case ","
            priority = 30
        case else
            priority = 1000 ' �퉉�Z�q�inumber, string, regex, word, date, expr���z����, ary���z����j
        end select
    end function
    
    private function isOP2(byval token)
        isOP2 = false
        select case token.getLex()
        case "+", "-", "*", "/", "%", _
             ">", "<", ">=", "<=", "==", "!=", "~", "!~", _
             "&", "|"
            isOP2 = true
        end select
    end function
    
    private function match(byval t)
        if token.getTyp() <> t then
            call err.raise(12345, TypeName(me), "expected " & t & " token, but " & token.getTyp() & " found.")
        end if
        set token = lex.nextToken()
    end function
    
    private function check(byval t)
        if token.getTyp() = t then
            check = true
        else
            check = false
        end if
    end function
end class
