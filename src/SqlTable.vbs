'!require ArrayList.vbs
'!require ExprLexer.vbs
'!require ExprParser.vbs
'!require TreeVisitor.vbs

class SqlTable
    ''' SQL 的なインターフェースを持つ二次元配列を表す
    private head
    private body
    private lexer
    private parser
    private visitor
    ''' クラス変数 head は、テーブルのヘッダ (ArrayList) を格納する。
    ''' クラス変数 body は、テーブル状の連結リスト (ArrayList) を格納する。
    ''' クラス変数 lexer は、字句解析器 (ExprLexer) を格納する。
    ''' クラス変数 parser は、構文解析器 (ExprLexer) を格納する。
    ''' クラス変数 visitor は、訪問器 (TreeVisitor) を格納する。
    
    private sub Class_Initialize
        set head = new ArrayList
        set body = new ArrayList
        set lexer = new ExprLexer
        set parser = new ExprParser
        set visitor = new TreeVisitor
    end sub
    
    private function error_(byval m)
        call Err.raise(1000, TypeName(me), m)
    end function
    
    public function from(byval ary, byval withHeader)
        ''' 二次元配列を受け取ってテーブル状のリストを構築する。
        ''' 第一引数 ary として、二次元配列 (Array(,)) を受け取る。
        ''' 第二引数 withHeader として、第一引数の二次元配列にヘッダを含むかを示す真偽値 (Boolean) を受け取る。
        ''' 戻り値として、自分自身への参照 (SqlTable) を返す。
        if isEmpty(ary) then
            call fromEmpty(ary, withHeader)
        elseif isArray(ary) then
            call fromArrayTable(ary, withHeader)
        elseif TypeName(ary) = "ArrayList" then
            call fromArrayList(ary, withHeader)
        else
            call error_("Type Error: expected Array or ArrayList instead of " & TypeName(ary))
        end if
        set from = me
    end function
    
    private function fromEmpty(byval ary, byval withHeader)
        set body = new ArrayList
    end function
    
    private function fromArrayTable(byval ary, byval withHeader)
        dim r
        r = rank(ary)
        if r = 1 then
            ' 一次元の配列
            set body = new ArrayList
            call body.push((new ArrayList).init(ary))
        elseif r = 2 then
            ' 二次元の配列
            set body = toListTable(ary)
        else
            call error_("TypeError: fromArrayTable")
        end if
        
        if withHeader then
            set head = fixHeader(body.shift())
        else
            set head = getPseudoHeader(body.item(0).length())
        end if
        
        call err.clear()
    end function
    
    private function rank(byval ary)
        ''' 配列の次元を調べる。
        ''' 第一引数 ary として、配列 (Array) を受け取る。
        ''' 戻り値として、配列の次元数 (Number) を返す。
        on error resume next
        call err.clear()
        dim i
        i = 0
        do while err.number = 0
            i = i + 1
            call ubound(ary, i)
        loop
        rank = i - 1
        call err.clear()
        on error goto 0
    end function
    
    private function fromArrayList(byval list, byval withHeader)
        if list.length() = 0 then call error_("Argument Error: arraylist's length is 0.")
        if list.item(0).length() = 0 then call error_("Argument Error: arraylist's length is 0.")
        
        dim i, len, row, width
        i = 0
        len = list.length()
        width = checkLength(list)
        do while i < len
            call rightPadding(list.item(i), width)
            i = i + 1
        loop
        
        if withHeader then
            set head = fixHeader(list.shift())
            set body = list
        else
            set head = getPseudoHeader(body.item(0).length())
            set body = list
        end if
        
    end function
    
    private function checkLength(byval list)
        dim i, len, row
        len = list.length()
        i = 0
        do while i < len
            set row = list.item(i)
            if TypeName(row) <> "ArrayList" then call error_("TypeError: checkLength")
            if row.length() > checkLength then
                checkLength = row.length()
            end if
            i = i + 1
        loop
        checkLength = checkLength
    end function
    
    public function update(byval col, byval value)
        set update = me
        if body.length() = 0 then exit function
        
        dim index
        index = indexof(head, col)
        if index = -1 then call error_("column not found: " & col)
        
        dim iter
        for each iter in body.toArray()
            call iter.setItem(index, value)
        next
        
        set update = me
    end function
    
    public function concat(byval table)
        dim listTable
        set listTable = toListTable(table.toArrayTable(false))
        
        dim iter, width
        width = head.length()
        for each iter in listTable.toArray()
            set iter = rightPadding(iter, width)
            set iter = iter.slice(0, width)
            call body.push(iter)
        next
        
        set concat = me
    end function
    
    public function insert(byval ary)
        set insert = me
        if isEmpty(head) then exit function
        
        dim list
        if isArray(ary) then set list = toList(ary)
        set list = rightPadding(list, head.length())
        set list = list.slice(0, head.length())
        
        call body.push(list)
        
        set insert = me
    end function
    
    public function insertAll(byval table)
        set insertAll = new SqlTable
        if TypeName(table) <> "SqlTable" then error_("TypeError: argument type is " & TypeName(table))
        if table.count = 0 then
            call body.unshift(head)
            set insertAll = (new SqlTable).from(body, true)
            call body.pop()
        end if
        
        dim listTable
        set listTable = toListTable(table.toArrayTable(true))
        
        dim header1, header2, newheader
        set header1 = head
        set header2 = listTable.shift()
        ' header1 と header2 を連結して重複を削除する
        set newheader = header1.clone().concat(header2.clone())
        set newheader = newheader.uniq()
        
        ' body のレコードを全てコピーする
        ' その際にフィールド数を newheader に合わせる。
        dim newbody, width
        set newbody = new ArrayList
        width = newheader.length()
        for each row in body.toArray()
            call newbody.push(rightPadding(row.clone(), width))
        next
        
        ' 追加するレコードを newheader に合わせて並べ替えながらコピーする。
        dim keys, row, key, index
        keys = newheader.toArray()
        for each row in listTable.toArray()
            call newbody.push(new ArrayList)
            for each key in keys
                index = indexof(header2, key)
                if index = -1 then
                    call newbody.peek().push(Empty)
                else
                    call newbody.peek().push(row.item(index))
                end if
            next
        next
        
        call newbody.unshift(newheader)
        call insertAll.from(newbody, true)
    end function
    
    ' 現在のヘッダを取得する
    public function describe()
        describe = array()
        if isEmpty(head) then exit function
        
        describe = head.toArray()
    end function
    
    ' 引数を元にテーブルのヘッダを設定する
    public function setHeader(byval ary)
        set setHeader = me
        if body.length() = 0 then exit function
        
        dim header
        set header = toList(ary)
        set header = rightPadding(header, head.length())
        set header = header.slice(0, head.length())
        set header = fixHeader(header)
        
        set head = header
        
        set setHeader = me
    end function
    
    private function fixHeader(byval list)
        ' 空欄(Empty)の処理
        set list = fixEmpty(list)
        
        ' 数字から始まるなど参照不可の列名の処理
        set list = fixInvalidName(list)
        
        ' 同名の列名の処理
        set list = fixSameName(list)
        
        set fixHeader = list
    end function
    
    private function fixEmpty(byval list)
        dim i, len
        i = 0
        len = list.length()
        do while i < len
            if isEmpty(list.item(i)) then
                call list.setItem(i, "$" & i + 1)
            end if
            i = i + 1
        loop
        set fixEmpty = list
    end function
    
    private function fixInvalidName(byval list)
        dim i, len
        i = 0
        len = list.length()
        do while i < len
            select case left(list.item(i), 1)
            case "", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
                 "(", ")", "[", "]", "+", "-", "*", "/", "%", "&", "|", _
                 "!", "=", "<", ">", ",", "~", "#", """", "'"
                call list.setItem(i, "$" & list.item(i))
            end select
            i = i + 1
        loop
        set fixInvalidName = list
    end function
    
    private function fixSameName(byval list)
        dim i, len, val, j, index
        i = 0
        len = list.length()
        do while i < len
            val = list.item(i)
            call list.setItem(i, empty)
            
            index = indexOf(list, val)
            j = 2
            suffix = "_" & j
            do while index <> -1
                ' 同じ要素全てに個別のサフィックスを付ける
                if indexOf(list, val & suffix) = -1 then
                    call list.setItem(index, val & suffix)
                    index = indexOf(list, val)
                end if
                j = j + 1
                suffix = "_" & j
            loop
            
            call list.setItem(i, val)
            i = i + 1
        loop
        set fixSameName = list
    end function
    
    public function delHeader()
        delHeader = empty
        if isEmpty(head) then exit function
        
        delHeader = head.toArray()
        set haed = getPseudoHeader(head.length())
    end function
    
    ' ヘッダを生成して返す
    ' $1, $2, $3 ... $n
    private function getPseudoHeader(byval cols)
        dim list, i
        set list = new ArrayList
        for i = 1 to cols ' cols は数値
            call list.push("$" & i)
        next
        set getPseudoHeader = list
    end function
    
    ' テーブルのレコード数を返す
    public function count()
        count = body.length()
    end function
    
    public function height()
        height = count()
    end function
    
    ' テーブルのフィールド数を返す
    public function width()
        width = head.length()
    end function
    
    private function assign(byval header, byval values)
        dim i
        i = 0
        do while i < header.length()
            ' ヘッダの変数に従って値を格納する
            call visitor.assign(header.item(i), values.item(i))
            ' $n 形式のヘッダとして値を格納する
            call visitor.assign("$" & i + 1, values.item(i))
            i = i + 1
        loop
        call visitor.assign("$0", values.toString())
    end function
    
    public function range(byval offset, byval limit)
        ''' 指定した範囲のレコードを返す。
        ''' 第一引数 offset として、レコードの選択を始める行番号 (Number) を受け取る。
        ''' 第二引数 limit として、取り出すレコードの最大数 (Number) を受け取る。
        ''' 戻り値として、指定範囲のレコードを含むテーブル (SqlTable) を返す。
        set range = new SqlTable
        if isEmpty(head) then exit function
        
        dim listTable, i, length
        set listTable = new ArrayList
        i = offset
        length = body.length()
        do while i < length and i < limit + offset
            call listTable.push(body.item(i))
            i = i + 1
        loop
        
        call listTable.unshift(head.clone())
        call range.from(listTable, true)
    end function
    
    public function take(byval count)
        ''' テーブルの最初のレコードから順番に指定した数のレコードを抽出して返す。
        ''' テーブルが保持するレコード数が、指定したレコード数に満たない場合は全てのレコードを返す。
        ''' 第一引数 count として、取り出すレコードの数 (Number) を受け取る。
        ''' 戻り値として、抽出したレコードを含むテーブル (SqlTable) を返す。
        set take = me.range(0, count)
    end function
    
    public function takeWhile(byval expr)
        ''' 指定された条件が満たされる限りテーブルの最初のレコードから順にレコードを抽出して返す。
        ''' 第一引数 expr として、抽出条件 (String) を受け取る。
        ''' 戻り値として、抽出されたレコードを含むテーブル (SqlTable) を返す。
        set takeWhile = new SqlTable
        if isEmpty(head) then exit function
        
        dim listTable
        set listTable = new ArrayList
        
        if body.length() = 0 then
            call listTable.push(head.clone())
            call takeWhile.from(listTable, true)
            exit function
        end if
        
        dim tree, row
        call lexer.init(expr)
        call parser.init(lexer)
        set tree = parser.expr()
        
        for each row in body.toArray()
            call assign(head, row)
            if visitor.evalate(tree) then call listTable.push(row)
        next
        
        call listTable.unshift(head.clone())
        call takeWhile.from(listTable, true)
    end function
    
    public function skip(byval count)
        ''' テーブルの最初のレコードから指定した数だけレコードをスキップして残りのレコードを返す。
        ''' 第一引数 count として、スキップするレコードの数 (Number) を受け取る。
        ''' 戻り値として、残りのレコードを含むテーブル (SqlTable) を返す。
        set skip = me.range(count, me.count())
    end function
    
    public function skipWhile(byval expr)
        ''' 指定された条件が満たされる限りテーブルの最初のレコードから順にスキップして残りのレコードを返す。
        ''' 第一引数 expr として、スキップする条件 (String) を受け取る。
        ''' 戻り値として、残りのレコードを含むテーブル (SqlTable) を返す。
        set skipWhile = new SqlTable
        if isEmpty(head) then exit function
        
        dim listTable
        set listTable = new ArrayList
        if body.length() = 0 then
            call listTable.push(head.clone())
            call skipWhile.from(listTable, true)
            exit function
        end if
        
        dim tree
        call lexer.init(expr)
        call parser.init(lexer)
        set tree = parser.expr()
        
        dim i, length
        i = 0
        length = body.length()
        do while i < length
            call assign(head, body.item(i))
            if not visitor.evalate(tree) then
                call listTable.push(body.item(i))
            end if
            i = i + 1
        loop
        
        call listTable.unshift(head.clone())
        call skipWhile.from(listTable, true)
    end function
    
    public function distinct()
        ''' 行の重複を排除する
        ''' 戻り値として、重複行を排除した新しいテーブル (SqlTable) を返す。
        
        set distinct = new SqlTable
        if isEmpty(head) then exit function
        
        dim after
        set after = new ArrayList
        if body.length() = 0 then
            call after.push(head.clone())
            call distinct.from(after, true)
            exit function
        end if
        
        dim rec1, rec2, agree
        for each rec1 in body.toArray()
            
            ' after に一致する行が有るか調べる。
            agree = false
            for each rec2 in after.toArray()
                if rec1.compare(rec2) then
                    agree = true
                    exit for
                end if
            next
            
            ' after に一致する行が無い場合
            if agree = false then
                call after.push(rec1.clone())
            end if
        next
        
        call after.unshift(head.clone())
        call distinct.from(after, true)
    end function
    
    public function sum(byval key)
        sum = 0
        if body.length() = 0 then exit function
        
        dim index
        index = indexof(head, key)
        if index = -1 then error_("ArgumentError: " & key)
        
        dim iter, typ, result
        result = 0
        for each iter in body.toArray()
            typ = vartype(iter.item(index))
            if 5 >= typ and typ >= 3 then ' 数値なら
                result = result + iter.item(index)
            end if
        next
        
        sum = result
    end function
    
    public function max(byval key)
        max = 0
        if body.length() = 0 then exit function
        
        dim index
        index = indexof(head, key)
        if index = -1 then error_("ArgumentError: " & key)
        
        dim iter, result
        result = 0
        for each iter in body.toArray()
            typ = vartype(iter.item(index))
            if 5 >= typ and typ >= 3 then ' 数値なら
                if iter.item(index) > result then
                    result = iter.item(index)
                end if
            end if
        next
        max = result
    end function
    
    public function min(byval key)
        min = 0
        if body.length() = 0 then exit function
        
        dim index
        index = indexof(head, key)
        if index = -1 then error_("ArgumentError: " & key)
        
        dim iter, result
        result = 0
        for each iter in body.toArray()
            typ = vartype(iter.item(index))
            if 5 >= typ and typ >= 3 then ' 数値なら
                if result > iter.item(index) then
                    result = iter.item(index)
                end if
            end if
        next
        min = result
    end function
    
    public function ave(byval key)
        ave = 0
        if body.length() = 0 then exit function
        
        dim index
        index = indexof(head, key)
        if index = -1 then error_("ArgumentError: " & key)
        
        dim iter, count, typ, sum
        sum = 0
        count = 0
        for each iter in body.toArray()
            typ = vartype(iter.item(index))
            if 5 >= typ and typ >= 3 then ' 数値なら
                sum = sum + iter.item(index)
                count = count + 1
            end if
        next
        
        if count <> 0 then
            ave = sum/count
        else
            ave = null
        end if
    end function
    
    ' filter の変種。破壊的に動作する。
    ' cond が true となるレコードだけで新しい SqlTable オブジェクトを生成して返す。
    ' false となったレコードは、extract を実行したインスタンスに残る。
    public function extract(byval cond)
        set extract = new SqlTable
        if isEmpty(head) then exit function
        
        dim trueTable, falseTable
        set trueTable = new ArrayList
        set falseTable = new ArrayList
        if body.length() = 0 then
            call trueTable.push(head.clone())
            call extract.from(trueTable, true)
            exit function
        end if
        
        dim tree, row
        call lexer.init(cond)
        call parser.init(lexer)
        set tree = parser.expr()
        
        for each row in body.toArray()
            call assign(head, row)
            if visitor.evalate(tree) then
                call trueTable.push(row)
            else
                call falseTable.push(row)
            end if
        next
        
        set body = falseTable
        call trueTable.unshift(head.clone())
        call extract.from(trueTable, true)
    end function
    
    public function innerJoin(byval key1, byval table, byval key2)
        set innerJoin = new SqlTable
        if body.length() = 0 then exit function
        
        dim index1
        index1 = indexof(head, key1)
        if index1 < 0 then error_("field not found: " & key1)
        
        dim head2, listTable, index2
        set listTable = toListTable(table.toArrayTable(true))
        set head2 = listTable.shift()
        index2 = indexof(head2, key2)
        if index2 < 0 then error_("field not found: " & key2)
        
        dim newbody, iter1, iter2, row
        set newbody = new ArrayList
        for each iter1 in body.toArray()
            for each iter2 in listTable.toArray()
                if iter1.item(index1) = iter2.item(index2) then
                    call newbody.push(iter1.clone().concat(iter2.clone()))
                end if
            next
        next
        
        call newbody.unshift(head.clone().concat(head2.clone()))
        call innerJoin.from(newbody, true)
    end function
    
    public function leftOuterJoin(byval key1, byval table, byval key2)
        set leftOuterJoin = new SqlTable
        if body.length() = 0 then exit function
        
        dim listTable, head2
        set listTable = toListTable(table.toArrayTable(true))
        set head2 = listTable.shift()
        
        dim index1, index2
        index1 = indexof(head, key1)
        index2 = indexof(head2, key2)
        if index1 < 0 then error_("field not found: " & key1)
        if index2 < 0 then error_("field not found: " & key2)
        
        dim newheader, newbody
        set newheader = header1.clone().concat(header2.clone())
        set newbody = new ArrayList
        call newbody.push(newheader)
        
        dim iter1, iter2, matchflag
        for each iter1 in body.toArray()
            matchflag = false
            for each iter2 in listTable.toArray()
                if iter1.item(index1) = iter2.item(index2) then
                    call newbody.push(iter1.clone().concat(iter2.clone()))
                    matchflag = true
                end if
            next
            ' 一度もマッチしなかった行は、空白で列の数を合わせて newbody に追加する。
            if not matchflag then newbody.push(rightPadding(iter1.clone(), newheader.length()))
        next
        
        call leftOuterJoin.from(newbody, true)
    end function
    
    public function rightOuterJoin(byval key1, byval table, byval key2)
        set rightouterJoin = new SqlTable
        if body.length() = 0 then exit function
        
        dim listTable, head2
        set listTable = toListTable(table.toArrayTable(true))
        set head2 = listTable.shift()
        
        dim newheader, newbody
        set newheader = head.clone().concat(head2.clone())
        set newbody = new ArrayList
        call newbody.push(newheader)
        
        dim index1, index2
        index1 = indexof(head, key1)
        index2 = indexof(head2, key2)
        if index1 < 0 then error_("field not found: " & key1)
        if index2 < 0 then error_("field not found: " & key2)
        
        dim iter1, iter2
        for each iter1 in body.toArray()
            for each iter2 in listTable.toArray()
                if iter1.item(index1) = iter2.item(index2) then
                    call newbody.push(iter1.clone().concat(iter2.clone()))
                end if
            next
        next
        
        ' 一度もマッチしなかった行は、空白で列の数を合わせて newbody に追加する。
        dim matchflag
        for each iter1 in listTable.toArray()
            matchflag = false
            for each iter2 in body.toArray()
                if iter1.item(index2) = iter2.item(index1) then
                    matchflag = true
                    exit for
                end if
            next
            if not matchflag then newbody.push(leftPadding(iter1.clone(), newheader.length()))
        next
        
        call rightOuterJoin.from(newbody, true)
    end function
    
    public function fullOuterJoin(byval key1, byval table, byval key2)
        set fullOuterJoin = new SqlTable
        if body.length() = 0 then exit function
        
        dim listTable, head2
        set listTable = toListTable(table.toArrayTable(true))
        set head2 = listTable.shift()
        
        dim newheader, newbody
        set newheader = head.clone().concat(head2.clone())
        set newbody = new ArrayList
        call newbody.push(newheader)
        
        dim index1, index2
        index1 = indexof(head, key1)
        index2 = indexof(head2, key2)
        if index1 < 0 then error_("field not found: " & key1)
        if index2 < 0 then error_("field not found: " & key2)
        
        ' leftOuterJoin と同じ
        dim iter1, iter2, matchflag
        for each iter1 in body.toArray()
            matchflag = false
            for each iter2 in listTable.toArray()
                if iter1.item(index1) = iter2.item(index2) then
                    call newbody.push(iter1.clone().concat(iter2.clone()))
                    matchflag = true
                end if
            next
            if not matchflag then newbody.push(rightPadding(iter1.clone(), newheader.length()))
        next
        
        ' 右側のテーブルの中で一度も選択されていないレコードを探して、
        ' 新しいテーブルの最後に insert する。
        for each iter1 in listTable.toArray()
            matchflag = false
            for each iter2 in body.toArray()
                if iter1.item(index2) = iter2.item(index1) then
                    matchflag = true
                    exit for
                end if
            next
            if not matchflag then newbody.push(leftPadding(iter1.clone(), newheader.length()))
        next
        
        call fullOuterJoin.from(newbody, true)
    end function
    
    private function leftPadding(byval list, byval length)
        ''' 指定された長さに達するまでリストに Empty をプッシュする。
        ''' 第一引数 list として、リスト (ArrayList) を受け取る。
        ''' 第二引数 length として、長さ (Number) を受け取る。
        ''' 戻り値として、指定された長さのリスト (ArrayList) を返す。
        do while list.length() < length
            call list.unshift(Empty)
        loop
        set leftPadding = list
    end function
    
    private function rightPadding(byval list, byval length)
        ''' 指定された長さに達するまでリストに Empty をプッシュする。
        ''' 第一引数 list として、リスト (ArrayList) を受け取る。
        ''' 第二引数 length として、長さ (Number) を受け取る。
        ''' 戻り値として、指定された長さのリスト (ArrayList) を返す。
        do while list.length() < length
            call list.push(Empty)
        loop
        set rightPadding = list
    end function
    
    private function indexof(byval list, byval key)
        ''' リストのメンバーを検索し、発見した位置を返す。
        ''' 第一引数 list として、検索されたリスト (ArrayList) を受け取る。
        ''' 第二引数 key として、検索する値 (String) を受け取る
        ''' 戻り値として、最初に key を発見した位置 (Number) を返す。発見できなければ -1 を返す。
        dim iter, i
        i = -1
        for each iter in list.toArray()
            i = i + 1
            if iter = key then
                indexof = i
                exit function
            end if
        next
        indexof = -1
    end function
    
    public function map(byval cols)
        ''' 列を生成する。
        ''' 第一引数 cols として、列を表す式の配列 (Array) を受け取る。
        ''' 戻り値として、新しいテーブル (SqlTable) を返す。
        set map = new SqlTable
        if isEmpty(head) then exit function
        
        cols = "[" & cols & "]"
        dim node
        call lexer.init(cols)
        call parser.init(lexer)
        set node = parser.expr()
        
        dim table, list, item, pos
        set table = new ArrayList
        set list = new ArrayList
        for each item in node.getChildren().item(0).getChildren().toArray()
            pos = item.getPos()
            call list.push(mid(cols, pos(0), pos(1) - pos(0)))
        next
        call table.push(list)
        
        dim col, row
        ' 新しい行の挿入
        for each row in body.toArray()
            set list = new ArrayList
            call assign(head, row) ' visitor の内部変数を書き換える
            for each col in visitor.evalate(node)
                call list.push(col)
            next
            call table.push(list)
        next
        
        set map = map.from(table, true)
    end function
    
    public function filter(byval cond)
        set filter = new SqlTable
        if isEmpty(head) then exit function
        
        dim table, row
        set table = new ArrayList
        call table.push(head.clone())
        
        if body.length() = 0 then
            call filter.from(table, true)
            exit function
        end if
        
        dim tree
        call lexer.init(cond)
        call parser.init(lexer)
        set tree = parser.expr()
        
        for each row in body.toArray()
            call assign(head, row)
            if visitor.evalate(tree) then call table.push(row)
        next
        call filter.from(table, true)
    end function
    
    public function orderby(byval querysString)
        set orderby = new SqlTable
        querysString = "[" & querysString & "]"
        
        dim querys, key
        call lexer.init(querysString)
        call parser.init(lexer)
        call visitor.init()
        for each key in head.toArray()
            call visitor.assign(key, key)
        next
        querys = visitor.evalate(parser.expr())
        
        dim listTable
        set listTable = body.clone()
        call listTable.unshift(head.clone())
        set listTable = sort(listTable, (new ArrayList).init(querys))
        
        call orderby.from(listTable, true)
    end function
    
    private function sort(byval listTable, byval querys)
        
        dim header
        set header = listTable.shift()
        
        dim query, key, index
        set query = querys.shift()
        key = query.keys()(0)
        index = indexof(header, key)
        if index = -1 then call error_("header not found")
        
        dim table
        select case lcase(query.item(key))
        case "asc"      set table = mergesort(listTable, index)
        case "desc"     set table = rev_mergesort(listTable, index)
        case "floatup"  set table = floatupsort(listTable, index)
        case else       set table = mergesort(listTable, index)
        end select
        
        call table.unshift(header)
        
        dim tables
        set tables = new ArrayList
        for each table in splitTable(table, index).toArray()
            if querys.length() <> 0 then
                set table = sort(table, querys.clone())
            end if
            call tables.push(table)
        next
        
        set sort = reduce(tables)
    end function
    
    private function reduce(byval tables)
        dim list
        set list = tables.shift()
        for each table in tables.toArray()
            call table.shift() ' drop header
            for each row in table.toArray()
                call list.push(row)
            next
        next
        set reduce = list
    end function
    
    private function splitTable(byval ListTable, byval index)
        ''' 指定された列に同じ値を持つ行を一つのグループとして分割したテーブルのリストを返す
        ''' この関数は第一引数 ListTable を破壊的する。
        ''' 第一引数 ListTable として、リストのリスト (ArrayList<ArrayList>) を受け取る。
        ''' 第二引数 index として、グループ化する基準となる列番号 (Number) を受け取る。
        ''' 戻り値として、テーブルの集合 (ArrayList<ArrayList<ArrayList>>) を返す。
        if index < 0 then call error_("split")
        
        dim header
        set header = ListTable.shift()
        
        dim tables, table
        set tables = new ArrayList
        do while ListTable.length() > 0
            set table = new ArrayList
            call table.push(ListTable.shift())
            value = table.item(0).item(index)
            do while ListTable.length() > 0
                if ListTable.item(0).item(index) <> value then exit do
                call table.push(ListTable.shift())
            loop
            call table.unshift(header.clone())
            call tables.push(table)
        loop
        set splitTable = tables
    end function
    
    private function mergesort(byval seq, byval key)
        ''' 昇順にソートする。
        if seq.length() <= 1 then
            set mergesort = seq
            exit function
        end if
        
        dim half, ary1, ary2
        half = seq.length() \ 2 ' 少数を切り捨てた商を返す
        set ary1 = seq.getRange(0, half)
        set ary2 = seq.getRange(half, seq.length() - half)
        
        set ary1 = mergesort(ary1, key)
        set ary2 = mergesort(ary2, key)
        
        dim result
        set result = new ArrayList
        do while ary1.length() > 0 and ary2.length() > 0
            if ary1.item(0).item(key) < ary2.item(0).item(key) then
                call result.push(ary1.shift())
            else
                call result.push(ary2.shift())
            end if
        loop
        
        if ary1.length() > 0 then result.concat(ary1)
        if ary2.length() > 0 then result.concat(ary2)
        
        set mergesort = result
    end function
    
    private function rev_mergesort(byval seq, byval key)
        ''' 降順にソートする。
        if seq.length() <= 1 then
            set rev_mergesort = seq
            exit function
        end if
        
        dim half, ary1, ary2
        half = seq.length() \ 2 ' 少数を切り捨てた商を返す
        set ary1 = seq.getRange(0, half)
        set ary2 = seq.getRange(half, seq.length() - half)
        
        set ary1 = rev_mergesort(ary1, key)
        set ary2 = rev_mergesort(ary2, key)
        
        dim result
        set result = new ArrayList
        do while ary1.length() > 0 and ary2.length() > 0
            if ary1.item(0).item(key) > ary2.item(0).item(key) then
                call result.push(ary1.shift())
            else
                call result.push(ary2.shift())
            end if
        loop
        
        if ary1.length() > 0 then result.concat(ary1)
        if ary2.length() > 0 then result.concat(ary2)
        
        set rev_mergesort = result
    end function
    
    ' 業務都合により実装
    ' 1    1
    ' 3    1
    ' 1 -> 3
    ' 3    3
    ' 2    2
    private function floatupsort(byval seq, byval key)
        dim after, i
        set after = new ArrayList
        do while seq.length <> 0
            call after.push(seq.shift())
            i = 0
            do while i < seq.length
                if after.peek.item(key) = seq.item(i).item(key) then
                    call after.push(seq.removeAt(i))
                else
                    i = i + 1
                end if
            loop
        loop
        set floatupsort = after
    end function
    
    private function toList(byval ary)
        ''' 配列をリストに変換する
        ''' 第一引数 ary として、配列 (Array) を受け取る。
        ''' 戻り値として、リスト (ArrayList) を返す。
        dim list, iter
        set list = new ArrayList
        for each iter in ary
            call list.push(iter)
        next
        set toList = list
    end function
    
    private function toListTable(byval ary2d)
        ''' 二次元配列をテーブル状のリスト（リストのリスト）にして返す
        ''' 第一引数 ary2d として、二次元配列 (Array) を受け取る。
        ''' 戻り値として、テーブル上のリスト (ArrayList) を返す。
        if not isArray(ary2d) then call error_("not 2d array: " & TypeName(ary2d))
        
        dim table, row, i, j
        set table = new ArrayList
        for i = lbound(ary2d) to ubound(ary2d)
            set row = new ArrayList
            for j = lbound(ary2d, 2) to ubound(ary2d, 2)
                call row.push(ary2d(i, j))
            next
            call table.push(row)
        next
        set toListTable = table
    end function
    
    public function toArrayTable(byval withHeader)
        if withHeader then
            toArrayTable = toArrayTableWithHeader()
        else
            toArrayTable = toArrayTableWithoutHeader()
        end if
    end function
    
    public function toArrayTableWithHeader()
        toArrayTableWithHeader = Array()
        if isEmpty(head) then exit function
        if body.length() = 0 then
            toArrayTableWithHeader = oneRowArrayTable(head)
            exit function
        end if
        call body.unshift(head)
        toArrayTableWithHeader = convertToArrayTable(body)
        call body.shift()
    end function
    
    public function toArrayTableWithoutHeader()
        toArrayTableWithoutHeader = Array()
        if isEmpty(head) then exit function
        if body.length() = 0 then exit function
        if body.length() = 1 then
            toArrayTableWithoutHeader = oneRowArrayTable(body.item(1))
            exit function
        end if
        toArrayTableWithoutHeader = convertToArrayTable(body)
    end function
    
    private function oneRowArrayTable(byval list)
        ''' リストを一行だけの二次元配列に変換する
        ''' 第一引数 list として、リスト (ArrayList) を受け取る。
        ''' 戻り値として、二次元配列 (Array) を返す。
        dim i, len, ret()
        i = 0
        len = list.length()
        redim ret(0, list.length() - 1)
        do while i < len
            ret(0, i) = list.item(i)
            i = i + 1
        loop
        oneRowArrayTable = ret
    end function
    
    private function convertToArrayTable(byval list)
        ''' リストのリストを二次元配列に格納する
        ''' 第一引数 list として、リストのリスト (ArrayList) を受け取る。
        ''' 戻り値として、二次元配列 (Array(,)) を返す。
        dim row, col, ary()
        row = list.length() - 1
        col = list.item(0).length() - 1
        redim ary(row, col)
        
        dim i, j
        i = 0
        do while i <= row
            j = 0
            do while j <= col
                if list.item(i).length() < j then exit do
                ary(i, j) = list.item(i).item(j)
                j = j + 1
            loop
            i = i + 1
        loop
        convertToArrayTable = ary
    end function
    
    public function toString(byval withHeader)
        ''' テーブル全体を文字列にして返す。
        ''' テーブルが初期化されていない場合は、空の文字列を返す。
        ''' 戻り値として、テーブル全体を表す文字列 (String) を返す。
        
        if withHeader then
            toString = toStringWithHeader()
        else
            toString = toStringWithoutHeader()
        end if
        
    end function
    
    public function toStringWithHeader()
        toStringWithHeader = ""
        if isEmpty(head) then exit function
        if body.length() = 0 then exit function
        
        dim sb
        set sb = createObject("System.Text.StringBuilder")
        
        ' ヘッダの設定
        call sb.append_3(join(head.toArray(), vbTab) & vbCrLf)
        dim col
        for each col in head.toArray()
            call sb.append_3(string(len(col), "-"))
            call sb.append_3(vbTab)
        next
        call sb.append_3(vbCrLf)
        
        ' 行の設定
        dim row
        for each row in body.toArray()
            call sb.append_3(join(row.toArray(), vbTab))
            call sb.append_3(vbCrLf)
        next
        
        toStringWithHeader = sb.toString()
    end function
    
    public function toStringWithoutHeader()
        toStringWithoutHeader = ""
        if isEmpty(head) then exit function
        if body.length() = 0 then exit function
        
        dim sb
        set sb = createObject("System.Text.StringBuilder")
        
        ' 行の設定
        dim row
        for each row in body.toArray()
            call sb.append_3(join(row.toArray(), vbTab))
            call sb.append_3(vbCrLf)
        next
        
        toStringWithoutHeader = sb.toString()
    end function
    
    public function setValue(byval r, byval c, byval v)
        ''' 指定されたアドレスの値を返す。アドレスが存在しなければ Empty を返す。
        ''' 第一引数 r として、行番号 (Number) を受け取る。
        ''' 第二引数 c として、列番号 (Number) を受け取る。
        ''' 戻り値として、指定されたアドレスの値 (Variant) を返す。
        getValue = empty
        if isEmpty(head) then exit function
        if c < 0 or head.length() - 1 < c then exit function
        if r < 0 or body.length() - 1 < r then exit function
        
        call body.item(r).setItem(c, v)
    end function
    
    public function getValue(byval r, byval c)
        ''' 指定されたアドレスの値を返す。アドレスが存在しなければ Empty を返す。
        ''' 第一引数 r として、行番号 (Number) を受け取る。
        ''' 第二引数 c として、列番号 (Number) を受け取る。
        ''' 戻り値として、指定されたアドレスの値 (Variant) を返す。
        getValue = empty
        if isEmpty(head) = 0 then exit function
        if c < 0 or head.length() - 1 < c then exit function
        if r < 0 or body.length() - 1 < r then exit function
        
        getValue = body.item(r).item(c)
    end function
    
    public function getRow(byval r)
        ''' 指定された行を配列にして返す。行が無ければ Array() を返す。
        ''' 行番号は 0 から始まる。ヘッダを取り出すには describe メソッドを使用する。
        ''' 第一引数 r として、取得する行番号 (Number) を受け取る。
        ''' 戻り値として、指定された行の配列 (Array) を返す。
        getRow = array()
        if isEmpty(head) then exit function
        if body.length() = 0 then exit function
        if index < 0 or body.length() - 1 < r then exit function
        
        getRow = body.item(r).toArray()
    end function
    
    public function getColumn(byval col)
        ''' 指定された列を配列にして返す。テーブルの行がなければ Array() を返す。
        ''' 第一引数 col として、取得するヘッダ名 (String) を受け取る。
        ''' 戻り値として、指定された列の配列 (Array) を返す。
        
        getColumn = array()
        if isEmpty(head) then exit function
        if body.length() = 0 then exit function
        
        dim ary(), index
        redim ary(body.length() - 1)
        index = indexof(head, col)
        if index = -1 then error_("header not found: " & col)
        
        dim i, len
        i = 0
        len = body.length()
        do while i < len
            ary(i) = body.item(i).item(index)
            i = i + 1
        loop
        getColumn = ary
    end function
    
    public function addColumn(byval name, byval defaultValue)
        ''' テーブルにフィールドを追加する
        ''' 第一引数 name として、追加するフィールドの名前 (String) を受け取る。
        ''' 第二引数 defaultValue として、追加したフィールドの初期値 (Variant) を受け取る。
        ''' 戻り値は返さない。
        
        call head.push(name) ' ToDo 名前のチェック
        
        dim i, len
        i = 0
        len = body.length()
        do while i < len
            body.item(i).push(defaultValue)
            i = i + 1
        loop
    end function
    
    public function modColumn(byval before, byval after)
        ''' フィールド名を変更する。
        ''' 第一引数 before として、変更前の名前 (String) を受け取る。
        ''' 第二引数 after として、変更後の名前 (String) を受け取る。
        ''' 戻り値は返さない。
        
        if isEmpty(head) then call err.raise(12345, TypeName(me), "header is empty")
        
        dim index
        index = indexof(head, before)
        if index = -1 then call err.raise(12345, TypeName(me), "header not found: " & before)
        if indexof(head, after) <> -1 then call err.raise(12345, TypeName(me), "header already added: " & after)
        call head.setItem(index, after)
    end function
end class
