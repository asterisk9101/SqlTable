class ExprToken
    ''' 数式のトークンを表すクラス
    ''' プライベート変数 typ は、トークンの型を表す
    ''' プレイベート変数 lex は、トークンの語彙素を表す
    private typ
    private lex
    private head
    private tail
    
    public function init(byval t, byval l, byval hd, byval tl)
        ''' トークンを初期化する
        ''' 第一引数 t は、トークンの型 (String) を受け取る
        ''' 第二引数 l は、トークンの語彙素 (String) を受け取る。
        ''' 戻り値として、自分自身への参照 (ExprToken) を返す。
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
        ''' トークンの語彙素を返す。
        ''' 戻り値として、トークンの語彙素 (String) を返す。
        getLex = lex
    end function
    
    public function getTyp()
        ''' トークンの型を返す。
        ''' 戻り値として、トークンの型 (String) を返す。
        getTyp = typ
    end function
    
    public function setLex(byval l)
        ''' トークンの語彙素を設定する。
        ''' 戻り値を返さない。
        setLex = l
    end function
    
    public function setTyp(byval t)
        ''' トークンの型を設定する。
        ''' 戻り値を返さない。
        typ = t
    end function
    
    public function toString()
        ''' 文字列で表現したトークンを返す。
        ''' 戻り値として、文字列で表現したトークン (String) を返す。
        toString = "<" & typ & "," & lex & ">"
    end function
end class
