'!require ExprToken.vbs

class ExprLexer
    ''' 字句解析器を表すクラス。
    ''' 数式として文字列を受け取りトークンに分解する。
    ''' プライベート変数 expr は、数式となる文字列 (String) を格納する。
    ''' プライベート変数 at は、expr の中で何文字目を解析しているかという情報 (Number) を格納する。
    ''' プライベート変数 ch は、epxr の中で at 文字目に存在する一文字 (String) を格納する。
    ''' プライベート変数 token は、直前に解析したトークン (ExprToken) を格納する。
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
        ''' ExprLexer のコンストラクタ
        set token = nothing
    end sub
    
    private sub Class_Terminate
        ''' ExprLexer のデストラクタ
        expr = null
        set token = nothing
    end sub
    
    public function init(byval str)
        ''' ExprLexer を初期化する。
        ''' 第一変数 str として、解析する式 (String) を受け取る。
        ''' 戻り値として、自分自身への参照 (ExprLexer) を返す。
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
        ''' 解析位置を一文字進める。
        ''' 戻り値として、読み進めた先の文字を返す。
        at = at + 1
        ch = mid(expr, at, 1)
        next_ = ch
    end function
    
    private function white()
        ''' 式の中に現れる空白文字を無視して解析位置を進める。
        ''' 戻り値を返さない。
        do while ch <> ""
            if asc(ch) < 0 or 32 < asc(ch) then exit do ' 空白文字以外が現れたら break する。
            next_()
        loop
    end function
    
    public function nextToken()
        ''' 式の解析を進める。
        ''' 戻り値として、解析されたトークン (ExprToken) を返す。
        dim pos
        white()
        pos = at
        select case ch
        case ""           set token = (new ExprToken).init("EOF", "", pos, pos + 1) ' 最後まで到達した場合は next_() しない。
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
                ' 式の先頭に / が現れた場合は除算記号でないことが確定するので
                ' REGEX 型としてパースする。
                set token = regex()
            elseif token.gettyp = "NUMBER" or token.getTyp = "WORD" then
                ' 一つ前のトークンが数字または識別子であれば除算記号とする
                next_()
                set token = (new ExprToken).init("DIV", "/", pos, at)
            else
                ' 厳密には更に場合分けが必要だが、
                ' ここでは構文解析をしないので正規表現リテラルとしてパースする。
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
        ''' 数値リテラルを解析する。
        ''' 第一引数 sign として、正負符号 (String) を受け取る。
        ''' 戻り値として、double 型の数値を返す。
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
        ''' 日付リテラルを解析する。
        ''' 戻り値として、日付リテラルを表す文字列 (String) を返す。
        dim ret, pos
        ret = ""
        pos = at
        do while next_() <> "#"
            if ch = "" then error_("date format end is not found: " & pos)
            ret = ret & ch
        loop
        call next_() ' 終わりの # をスキップする
        if not isDate(ret) then error_("format is not date: " & pos)
        set yyyymmdd = (new ExprToken).init("DATE", ret, pos, at)
    end function
    
    private function word()
        ''' 単語を解析する。
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
        ''' 文字列リテラルをパースする。
        ''' 第一引数 q として、引用符 (String) を受け取る。
        ''' 戻り値として、文字列リテラルを表す文字列 (String) を返す。
        dim pos, s
        pos = at
        s = "" ' 引用符をスキップ
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
        next_() ' 終了の引用符をスキップ
        if q = "'" then
            set str_ = (new ExprToken).init("STRING", s, pos, at)
        elseif q = """" then
            set str_ = (new ExprToken).init("WORD", s, pos, at)
        end if
    end function
    
    private function regex()
        ''' 正規表現リテラルを解析する。
        ''' 戻り値として、正規表現リテラルを表す文字列 (String) を返す。
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
