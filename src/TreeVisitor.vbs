'!require ArrayList.vbs
'!require ObjectString.vbs

class TreeVisitor
    ''' ���ۍ\���؂̖K���
    ''' ���ۍ\���؂ɂ��������Čv�Z���ʂ�Ԃ��B
    private dic
    ''' �v���C�x�[�g�ϐ� dic �́ASQL �ϐ����i�[����B
    
    private sub Class_Initialize
        call init()
    end sub
    
    public function init()
        ''' �g�ݍ��ݕϐ����`����B
        set dic = createObject("Scripting.Dictionary")
        call dic.add("true", true)
        call dic.add("false", false)
        call dic.add("null", null)
        call dic.add("empty", empty)
        set init = me
    end function
    
    public function initAssign(byval ary)
        ''' �z��Ƃ��ė^����ꂽ�l�� $n �̕ϐ��ɑ������
        ''' ������ ary �Ƃ��āA�ϐ��Ɋi�[����l���܂ޔz�� (array) ���󂯎��B
        ''' �߂�l�Ƃ��āA�������g�ւ̎Q�� (TreeVisitor) ��Ԃ��B
        call init()
        dim e, i
        i = 1
        for each e in ary
            call assign("$" & i, e)
            i = i + 1
        next
        call assign("$0", join(ary, ""))
        set initAssign = me
    end function
    
    public function assign(byval name, byval value)
        ''' �ϐ����`����B
        ''' ������ name �Ƃ��āA�ϐ��� (String) ���󂯎��B
        ''' ������ value �Ƃ��āA�ϐ��l (Variant) ���󂯎��B
        ''' �߂�l�Ƃ��āA�������g�ւ̎Q�� (TreeVisitor) ��Ԃ��B
        if isObject(value) then
            set dic.item(name) = value
        else
            dic.item(name) = value
        end if
        set assign = me
    end function
    
    public function evalate(byval tree)
        ''' ���ۍ\���؂��ċA�I�ɕ]������B
        ''' ������ tree �Ƃ��āA���ۍ\���� (AstTree) ���󂯎��B
        ''' �߂�l�Ƃ��āA�]������ (Variant) ��Ԃ��B
        dim result, child, children, args, t
        set args = new ArrayList
        
        for each child in tree.getChildren.toArray()
            call args.push(evalate(child))
        next
        set t = tree.getToken()
        select case t.getTyp
        case "ADD"  result = add_(args)
        case "SUB"  result = sub_(args)
        case "MUL"  result = mul(args)
        case "DIV"  result = div(args)
        case "SUR"  result = sur(args)
        case "GT"   result = gt(args)
        case "LT"   result = lt(args)
        case "GTE"  result = gte(args)
        case "LTE"  result = lte(args)
        case "EQL"  result = eql(args)
        case "NEQ"  result = neq(args)
        case "AND"  result = and_(args)
        case "OR"   result = or_(args)
        case "NOT"  result = not_(args)
        case "MATCH"    result = mch(args)
        case "UNMATCH"  result = umh(args)
        case "COMMA"    result = comma(args)
        case "NUMBER"   result = cdbl(t.getLex())
        case "STRING"   result = t.getLex()
        case "EXPR"     call bind(result, args.item(0))
        case "ARRAY"    result = args.toArray()
        case "OBJECT"   set result = obj(args)
        case "PAIR"     result = pair(args)
        case "DATE"     result = cdate(replace(t.getLex(), "#", ""))
        case "REGEX"    set result = parse_regex(t.getLex())
        case "WORD"
            if dic.exists(t.getLex) then
                result = dic(t.getLex)
            else
                result = Empty
            end if
        case "FUNC"
            result = func(t, args)
        case else
            call Err.raise(1000, TypeName(me), "illigal token")
        end select
        
        if isObject(result) then
            set evalate = result
        else
            evalate = result
        end if
    end function
    
    private function bind(byref a, byval b)
        if isObject(b) then
            set a = b
        else
            a = b
        end if
    end function
    
    private function obj(byval args)
        set obj = createObject("Scripting.Dictionary")
        dim item
        for each item in args.toArray()
            call obj.add(item(0), item(1))
        next
        set obj = obj
    end function
    
    private function pair(byval args)
        pair = array(args.item(0), args.item(1))
    end function
    
    private function func(byval token, byval args)
        dim o
        select case token.getLex()
        case "replace"
            func = replace(args.item(0), args.item(1), args.item(2))
        case "slice"
            set o = (new ObjectString).init(args.item(0))
            select case args.length()
            case 2 func = o.sliceTail(args.item(1))
            case 3 func = o.slice(args.item(1), args.item(2))
            end select
        case "indexOf"
            func = inStr(args.item(0), args.item(1)) - 1
        case "length"
            func = len(args.item(0))
        case else
            call err.raise(12345, TypeName(me), "function?")
        end select
    end function
    
    private function parse_regex(byval lex)
        dim re
        set re = createObject("VBScript.RegExp")
        
        ' Lexer �Ɠ������Ƃ����Ă���B
        ' Lexer �ł̓g�[�N���̐؂�o���������l����������
        ' �I�u�W�F�N�g�ւ̕ϊ��͂��Ă��Ȃ������B
        
        ' �����p�^�[���̏���
        dim at, ch, pattern
        at = 2  ' �擪�� / �̓X�L�b�v����
        ch = mid(lex, at, 1)
        pattern = ""
        do while ch <> "/"
            if ch = "\" then
                at = at + 1
                ch = mid(lex, at, 1)
                select case ch
                case "n" pattern = pattern & vbLf
                case "r" pattern = pattern & vbCr
                case "t" pattern = pattern & vbTab
                case "/" pattern = pattern & ch
                end select
            end if
            pattern = pattern & ch
            at = at + 1
            ch = mid(lex, at, 1)
        loop
        re.pattern = pattern
        
        ' �����t���O�̏���
        at = at + 1
        ch = mid(lex, at, 1)
        do while ch <> ""
            select case ch
            case "i" re.ignoreCase = true
            case "g" re.global = true
            case "m" re.multiline = true
            end select
            at = at + 1
            ch = mid(lex, at, 1)
        loop
        set parse_regex = re
    end function
    
    private function add_(byval args)
        ' + ���P�����Z�q�Ƃ��Ďg�p����Ă����ꍇ�͂��̂܂܂̒l��Ԃ��B
        if args.length() = 1 then
            add_ = args.item(0)
            exit function
        end if
        
        dim result, arg
        for each arg in args.toArray()
            result = result + arg
        next
        add_ = result
    end function
    
    private function sub_(byval args)
        ' - ���P�����Z�q�Ƃ��Ďg�p����Ă���ꍇ�́A�����𔽓]���ĕԂ��B
        if args.length() = 1 then
            sub_ = -args.item(0)
            exit function
        end if
        
        dim result, arg
        result = args.item(0)
        args.removeAt(0)
        for each arg in args.toArray()
            result = result - arg
        next
        sub_ = result
    end function
    
    private function mul(byval args)
        dim result, arg
        result = 1
        for each arg in args.toArray()
            result = result * arg
        next
        mul = result
    end function
    
    private function div(byval args)
        dim result, arg
        result = args.item(0)
        args.removeAt(0)
        for each arg in args.toArray()
            result = result / arg
        next
        div = result
    end function
    
    private function sur(byval args)
        sur = args.item(0) mod args.item(1)
    end function
    
    private function and_(byval args)
        and_ = args.item(0) and args.item(1)
    end function
    
    private function or_(byval args)
         or_ = args.item(0) or args.item(1)
    end function
    
    private function gt(byval args)
         gt = args.item(0) > args.item(1)
    end function
    
    private function lt(byval args)
         lt = args.item(0) < args.item(1)
    end function
    
    private function gte(byval args)
        gte = args.item(0) >= args.item(1)
    end function
    
    private function lte(byval args)
        lte = args.item(0) <= args.item(1)
    end function
    
    private function mch(byval args)
        dim result
        result = false
        if TypeName(args.item(0)) = "IRegExp2" then
            result = args.item(0).test(args.item(1))
        elseif TypeName(args.item(1)) = "IRegExp2" then
            result = args.item(1).test(args.item(0))
        else
            call Err.raise(1000, TypeName(me), "illigal expression: RegExp position")
        end if
        mch = result
    end function
    
    private function umh(byval args)
        umh = not mch(args)
    end function
    
    private function eql(byval args)
        eql = (args.item(0) = args.item(1))
    end function
    
    private function neq(byval args)
        neq = not eql(args)
    end function
    
    private function comma(byval args)
        dim arg, result, ary, i
        set ary = new ArrayList
        for each arg in args.toArray()
            ary.push(arg)
        next
        comma = ary.toArray()
    end function
    
    private function not_(byval args)
        dim arg, result
        result = false
        for each arg in args
            result = not arg
        next
        not_ = result
    end function
end class
