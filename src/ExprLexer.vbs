'!require ExprToken.vbs

class ExprLexer
    ''' �����͊��\���N���X�B
    ''' �����Ƃ��ĕ�������󂯎��g�[�N���ɕ�������B
    ''' �v���C�x�[�g�ϐ� expr �́A�����ƂȂ镶���� (String) ���i�[����B
    ''' �v���C�x�[�g�ϐ� at �́Aexpr �̒��ŉ������ڂ���͂��Ă��邩�Ƃ������ (Number) ���i�[����B
    ''' �v���C�x�[�g�ϐ� ch �́Aepxr �̒��� at �����ڂɑ��݂���ꕶ�� (String) ���i�[����B
    ''' �v���C�x�[�g�ϐ� token �́A���O�ɉ�͂����g�[�N�� (ExprToken) ���i�[����B
    private expr
    private at
    private ch
    private token
    
    public function getAt()
        getAt = at
    end function
    
    public function currentToken()
        set currentToken = token
    end function
    
    private sub Class_Initialize
        ''' ExprLexer �̃R���X�g���N�^
        set token = nothing
    end sub
    
    private sub Class_Terminate
        ''' ExprLexer �̃f�X�g���N�^
        expr = null
        set token = nothing
    end sub
    
    public function init(byval str)
        ''' ExprLexer ������������B
        ''' ���ϐ� str �Ƃ��āA��͂��鎮 (String) ���󂯎��B
        ''' �߂�l�Ƃ��āA�������g�ւ̎Q�� (ExprLexer) ��Ԃ��B
        at = 1
        expr = str
        ch = mid(expr, 1, 1)
        set token = nothing
        set init = me
    end function
    
    private function error_(byval m)
        call err.raise(1000, TypeName(me), m)
    end function
    
    private function next_()
        ''' ��͈ʒu���ꕶ���i�߂�B
        ''' �߂�l�Ƃ��āA�ǂݐi�߂���̕�����Ԃ��B
        at = at + 1
        ch = mid(expr, at, 1)
        next_ = ch
    end function
    
    private function white()
        ''' ���̒��Ɍ����󔒕����𖳎����ĉ�͈ʒu��i�߂�B
        ''' �߂�l��Ԃ��Ȃ��B
        do while ch <> ""
            if asc(ch) < 0 or 32 < asc(ch) then exit do ' �󔒕����ȊO�����ꂽ�� break ����B
            next_()
        loop
    end function
    
    public function nextToken()
        ''' ���̉�͂�i�߂�B
        ''' �߂�l�Ƃ��āA��͂��ꂽ�g�[�N�� (ExprToken) ��Ԃ��B
        dim pos
        white()
        pos = at
        select case ch
        case ""           set token = (new ExprToken).init("EOF", "", pos, pos + 1) ' �Ō�܂œ��B�����ꍇ�� next_() ���Ȃ��B
        case "(" next_(): set token = (new ExprToken).init("LPAR", "(", pos, at)
        case ")" next_(): set token = (new ExprToken).init("RPAR", ")", pos, at)
        case "[" next_(): set token = (new ExprToken).init("LBRA", "[", pos, at)
        case "]" next_(): set token = (new ExprToken).init("RBRA", "]", pos, at)
        case "{" next_(): set token = (new ExprToken).init("LBRACE", "{", pos, at)
        case "}" next_(): set token = (new ExprToken).init("RBRACE", "}", pos, at)
        case "," next_(): set token = (new ExprToken).init("COMMA", ",", pos, at)
        case "%" next_(): set token = (new ExprToken).init("SUR", "%", pos, at)
        case "&" next_(): set token = (new ExprToken).init("AND", "&", pos, at)
        case "|" next_(): set token = (new ExprToken).init("OR", "|", pos, at)
        case ";" next_(): set token = (new ExprToken).init("TERM", ";", pos, at)
        case "~" next_(): set token = (new ExprToken).init("MATCH", "~", pos, at)
        case "#" set token = yyyymmdd()
        case "+" next_(): set token = (new ExprToken).init("ADD", "+", pos, at)
        case "-" next_(): set token = (new ExprToken).init("SUB", "-", pos, at)
        case "*" next_(): set token = (new ExprToken).init("MUL", "*", pos, at)
        case "/"
            if token is nothing then
                ' ���̐擪�� / �����ꂽ�ꍇ�͏��Z�L���łȂ����Ƃ��m�肷��̂�
                ' REGEX �^�Ƃ��ăp�[�X����B
                set token = regex()
            elseif token.gettyp = "NUMBER" or token.getTyp = "WORD" then
                ' ��O�̃g�[�N���������܂��͎��ʎq�ł���Ώ��Z�L���Ƃ���
                next_()
                set token = (new ExprToken).init("DIV", "/", pos, at)
            else
                ' �����ɂ͍X�ɏꍇ�������K�v�����A
                ' �����ł͍\����͂����Ȃ��̂Ő��K�\�����e�����Ƃ��ăp�[�X����B
                set token = regex()
            end if
        case "="
            next_()
            if ch = "=" then
                next_()
                set token = (new ExprToken).init("EQL", "==", pos, at)
            else
                set token = (new ExprToken).init("ASSIGN", "=", pos, at)
            end if
        case "!"
            next_()
            if ch = "=" then
                next_()
                set token = (new ExprToken).init("NEQ", "!=", pos, at)
            elseif ch = "~" then
                next_()
                set token = (new ExprToken).init("UNMATCH", "!~", pos, at)
            else
                set token = (new ExprToken).init("NOT", "!", pos, at)
            end if
        case ">"
            if next_() = "=" then
                next_()
                set token = (new ExprToken).init("GTE", ">=", pos, at)
            else
                set token = (new ExprToken).init("GT", ">", pos, at)
            end if
        case "<"
            if next_() = "=" then
                next_()
                set token = (new ExprToken).init("LTE", "<=", pos, at)
            else
                set token = (new ExprToken).init("LT", "<", pos, at)
            end if
        case """" set token = str_("""")
        case "'" set token = str_("'")
        case else
            if 48 <= asc(ch) and asc(ch) <= 57 then
                set token = number("")
            else
                set token = word()
            end if
        end select
        set nextToken = token
    end function
    
    private function number(byval sign)
        ''' ���l���e��������͂���B
        ''' ������ sign �Ƃ��āA�������� (String) ���󂯎��B
        ''' �߂�l�Ƃ��āAdouble �^�̐��l��Ԃ��B
        dim ret, pos
        ret = sign
        if sign <> "" then
            pos = at - 1
        else
            pos = at
        end if
        
        do while isNumeric(ch)
            ret = ret & ch
            call next_()
        loop
        
        if ch = "." then
            ret = ret & ch
            call next_()
            do while isNumeric(ch)
                ret = ret & ch
                call next_()
            loop
        end if
        set number = (new ExprToken).init("NUMBER", cdbl(ret), pos, at)
    end function
    
    private function yyyymmdd()
        ''' ���t���e��������͂���B
        ''' �߂�l�Ƃ��āA���t���e������\�������� (String) ��Ԃ��B
        dim ret, pos
        ret = ""
        pos = at
        do while next_() <> "#"
            if ch = "" then error_("date format end is not found: " & pos)
            ret = ret & ch
        loop
        call next_() ' �I���� # ���X�L�b�v����
        if not isDate(ret) then error_("format is not date: " & pos)
        set yyyymmdd = (new ExprToken).init("DATE", ret, pos, at)
    end function
    
    private function word()
        ''' �P�����͂���B
        dim w, pos
        w = ch
        pos = at
        do
            select case next_()
            case "", "+", "-", "/", "*", "%", "!", """", "&", "(", ")", "[", "]", "=", "|", ">", "<", ","
                exit do
            case else
                if 0 <= asc(ch) and asc(ch) <= 32 then
                    exit do
                end if
            end select
            w = w & ch
        loop
        set word = (new ExprToken).init("WORD", w, pos, at)
    end function
    
    private function str_(byval q)
        ''' �����񃊃e�������p�[�X����B
        ''' ������ q �Ƃ��āA���p�� (String) ���󂯎��B
        ''' �߂�l�Ƃ��āA�����񃊃e������\�������� (String) ��Ԃ��B
        dim pos, s
        pos = at
        s = "" ' ���p�����X�L�b�v
        do while next_() <> q
            if ch = "\" then
                select case next_()
                case ""  error_("string end not found: " & pos)
                case "t" s = s & chr(9)   ' TAB
                case "n" s = s & chr(10)  ' LF
                case "r" s = s & chr(13)  ' CR
                case "\" s = s & "\"      ' \
                case q   s = s & q        ' q
                case else s = s & " "
                end select
            else
                if ch = "" then error_("string end not found: " & pos)
                s = s & ch
            end if
        loop
        next_() ' �I���̈��p�����X�L�b�v
        if q = "'" then
            set str_ = (new ExprToken).init("STRING", s, pos, at)
        elseif q = """" then
            set str_ = (new ExprToken).init("WORD", s, pos, at)
        end if
    end function
    
    private function regex()
        ''' ���K�\�����e��������͂���B
        ''' �߂�l�Ƃ��āA���K�\�����e������\�������� (String) ��Ԃ��B
        dim re, pos
        pos = at
        re = "/"
        do while next_() <> "/"
            if ch = "\" then
                next_()
                select case ch
                case "/" re = re & "/"
                end select
            else
                if ch = "" then error_("regex end not found: " & pos)
                re = re & ch
            end if
        loop
        if re = "/" then error_("invalid regex: " & pos)
        re = re & "/"
        do
            select case next_()
            case "i" re = re & "i"
            case "m" re = re & "m"
            case "g" re = re & "g"
            case else exit do
            end select
        loop
        set regex = (new ExprToken).init("REGEX", re, pos, at)
    end function
end class
