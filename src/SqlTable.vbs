'!require ArrayList.vbs
'!require ExprLexer.vbs
'!require ExprParser.vbs
'!require TreeVisitor.vbs

class SqlTable
    ''' SQL �I�ȃC���^�[�t�F�[�X�����񎟌��z���\��
    private head
    private body
    private lexer
    private parser
    private visitor
    ''' �N���X�ϐ� head �́A�e�[�u���̃w�b�_ (ArrayList) ���i�[����B
    ''' �N���X�ϐ� body �́A�e�[�u����̘A�����X�g (ArrayList) ���i�[����B
    ''' �N���X�ϐ� lexer �́A�����͊� (ExprLexer) ���i�[����B
    ''' �N���X�ϐ� parser �́A�\����͊� (ExprLexer) ���i�[����B
    ''' �N���X�ϐ� visitor �́A�K��� (TreeVisitor) ���i�[����B
    
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
        ''' �񎟌��z����󂯎���ăe�[�u����̃��X�g���\�z����B
        ''' ������ ary �Ƃ��āA�񎟌��z�� (Array(,)) ���󂯎��B
        ''' ������ withHeader �Ƃ��āA�������̓񎟌��z��Ƀw�b�_���܂ނ��������^�U�l (Boolean) ���󂯎��B
        ''' �߂�l�Ƃ��āA�������g�ւ̎Q�� (SqlTable) ��Ԃ��B
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
            ' �ꎟ���̔z��
            set body = new ArrayList
            call body.push((new ArrayList).init(ary))
        elseif r = 2 then
            ' �񎟌��̔z��
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
        ''' �z��̎����𒲂ׂ�B
        ''' ������ ary �Ƃ��āA�z�� (Array) ���󂯎��B
        ''' �߂�l�Ƃ��āA�z��̎����� (Number) ��Ԃ��B
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
        ' header1 �� header2 ��A�����ďd�����폜����
        set newheader = header1.clone().concat(header2.clone())
        set newheader = newheader.uniq()
        
        ' body �̃��R�[�h��S�ăR�s�[����
        ' ���̍ۂɃt�B�[���h���� newheader �ɍ��킹��B
        dim newbody, width
        set newbody = new ArrayList
        width = newheader.length()
        for each row in body.toArray()
            call newbody.push(rightPadding(row.clone(), width))
        next
        
        ' �ǉ����郌�R�[�h�� newheader �ɍ��킹�ĕ��בւ��Ȃ���R�s�[����B
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
    
    ' ���݂̃w�b�_���擾����
    public function describe()
        describe = array()
        if isEmpty(head) then exit function
        
        describe = head.toArray()
    end function
    
    ' ���������Ƀe�[�u���̃w�b�_��ݒ肷��
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
        ' ��(Empty)�̏���
        set list = fixEmpty(list)
        
        ' ��������n�܂�ȂǎQ�ƕs�̗񖼂̏���
        set list = fixInvalidName(list)
        
        ' �����̗񖼂̏���
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
                ' �����v�f�S�ĂɌʂ̃T�t�B�b�N�X��t����
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
    
    ' �w�b�_�𐶐����ĕԂ�
    ' $1, $2, $3 ... $n
    private function getPseudoHeader(byval cols)
        dim list, i
        set list = new ArrayList
        for i = 1 to cols ' cols �͐��l
            call list.push("$" & i)
        next
        set getPseudoHeader = list
    end function
    
    ' �e�[�u���̃��R�[�h����Ԃ�
    public function count()
        count = body.length()
    end function
    
    public function height()
        height = count()
    end function
    
    ' �e�[�u���̃t�B�[���h����Ԃ�
    public function width()
        width = head.length()
    end function
    
    private function assign(byval header, byval values)
        dim i
        i = 0
        do while i < header.length()
            ' �w�b�_�̕ϐ��ɏ]���Ēl���i�[����
            call visitor.assign(header.item(i), values.item(i))
            ' $n �`���̃w�b�_�Ƃ��Ēl���i�[����
            call visitor.assign("$" & i + 1, values.item(i))
            i = i + 1
        loop
        call visitor.assign("$0", values.toString())
    end function
    
    public function range(byval offset, byval limit)
        ''' �w�肵���͈͂̃��R�[�h��Ԃ��B
        ''' ������ offset �Ƃ��āA���R�[�h�̑I�����n�߂�s�ԍ� (Number) ���󂯎��B
        ''' ������ limit �Ƃ��āA���o�����R�[�h�̍ő吔 (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āA�w��͈͂̃��R�[�h���܂ރe�[�u�� (SqlTable) ��Ԃ��B
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
        ''' �e�[�u���̍ŏ��̃��R�[�h���珇�ԂɎw�肵�����̃��R�[�h�𒊏o���ĕԂ��B
        ''' �e�[�u�����ێ����郌�R�[�h�����A�w�肵�����R�[�h���ɖ����Ȃ��ꍇ�͑S�Ẵ��R�[�h��Ԃ��B
        ''' ������ count �Ƃ��āA���o�����R�[�h�̐� (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āA���o�������R�[�h���܂ރe�[�u�� (SqlTable) ��Ԃ��B
        set take = me.range(0, count)
    end function
    
    public function takeWhile(byval expr)
        ''' �w�肳�ꂽ������������������e�[�u���̍ŏ��̃��R�[�h���珇�Ƀ��R�[�h�𒊏o���ĕԂ��B
        ''' ������ expr �Ƃ��āA���o���� (String) ���󂯎��B
        ''' �߂�l�Ƃ��āA���o���ꂽ���R�[�h���܂ރe�[�u�� (SqlTable) ��Ԃ��B
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
        ''' �e�[�u���̍ŏ��̃��R�[�h����w�肵�����������R�[�h���X�L�b�v���Ďc��̃��R�[�h��Ԃ��B
        ''' ������ count �Ƃ��āA�X�L�b�v���郌�R�[�h�̐� (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āA�c��̃��R�[�h���܂ރe�[�u�� (SqlTable) ��Ԃ��B
        set skip = me.range(count, me.count())
    end function
    
    public function skipWhile(byval expr)
        ''' �w�肳�ꂽ������������������e�[�u���̍ŏ��̃��R�[�h���珇�ɃX�L�b�v���Ďc��̃��R�[�h��Ԃ��B
        ''' ������ expr �Ƃ��āA�X�L�b�v������� (String) ���󂯎��B
        ''' �߂�l�Ƃ��āA�c��̃��R�[�h���܂ރe�[�u�� (SqlTable) ��Ԃ��B
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
        ''' �s�̏d����r������
        ''' �߂�l�Ƃ��āA�d���s��r�������V�����e�[�u�� (SqlTable) ��Ԃ��B
        
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
            
            ' after �Ɉ�v����s���L�邩���ׂ�B
            agree = false
            for each rec2 in after.toArray()
                if rec1.compare(rec2) then
                    agree = true
                    exit for
                end if
            next
            
            ' after �Ɉ�v����s�������ꍇ
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
            if 5 >= typ and typ >= 3 then ' ���l�Ȃ�
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
            if 5 >= typ and typ >= 3 then ' ���l�Ȃ�
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
            if 5 >= typ and typ >= 3 then ' ���l�Ȃ�
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
            if 5 >= typ and typ >= 3 then ' ���l�Ȃ�
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
    
    ' filter �̕ώ�B�j��I�ɓ��삷��B
    ' cond �� true �ƂȂ郌�R�[�h�����ŐV���� SqlTable �I�u�W�F�N�g�𐶐����ĕԂ��B
    ' false �ƂȂ������R�[�h�́Aextract �����s�����C���X�^���X�Ɏc��B
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
            ' ��x���}�b�`���Ȃ������s�́A�󔒂ŗ�̐������킹�� newbody �ɒǉ�����B
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
        
        ' ��x���}�b�`���Ȃ������s�́A�󔒂ŗ�̐������킹�� newbody �ɒǉ�����B
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
        
        ' leftOuterJoin �Ɠ���
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
        
        ' �E���̃e�[�u���̒��ň�x���I������Ă��Ȃ����R�[�h��T���āA
        ' �V�����e�[�u���̍Ō�� insert ����B
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
        ''' �w�肳�ꂽ�����ɒB����܂Ń��X�g�� Empty ���v�b�V������B
        ''' ������ list �Ƃ��āA���X�g (ArrayList) ���󂯎��B
        ''' ������ length �Ƃ��āA���� (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āA�w�肳�ꂽ�����̃��X�g (ArrayList) ��Ԃ��B
        do while list.length() < length
            call list.unshift(Empty)
        loop
        set leftPadding = list
    end function
    
    private function rightPadding(byval list, byval length)
        ''' �w�肳�ꂽ�����ɒB����܂Ń��X�g�� Empty ���v�b�V������B
        ''' ������ list �Ƃ��āA���X�g (ArrayList) ���󂯎��B
        ''' ������ length �Ƃ��āA���� (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āA�w�肳�ꂽ�����̃��X�g (ArrayList) ��Ԃ��B
        do while list.length() < length
            call list.push(Empty)
        loop
        set rightPadding = list
    end function
    
    private function indexof(byval list, byval key)
        ''' ���X�g�̃����o�[���������A���������ʒu��Ԃ��B
        ''' ������ list �Ƃ��āA�������ꂽ���X�g (ArrayList) ���󂯎��B
        ''' ������ key �Ƃ��āA��������l (String) ���󂯎��
        ''' �߂�l�Ƃ��āA�ŏ��� key �𔭌������ʒu (Number) ��Ԃ��B�����ł��Ȃ���� -1 ��Ԃ��B
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
        ''' ��𐶐�����B
        ''' ������ cols �Ƃ��āA���\�����̔z�� (Array) ���󂯎��B
        ''' �߂�l�Ƃ��āA�V�����e�[�u�� (SqlTable) ��Ԃ��B
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
        ' �V�����s�̑}��
        for each row in body.toArray()
            set list = new ArrayList
            call assign(head, row) ' visitor �̓����ϐ�������������
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
        ''' �w�肳�ꂽ��ɓ����l�����s����̃O���[�v�Ƃ��ĕ��������e�[�u���̃��X�g��Ԃ�
        ''' ���̊֐��͑����� ListTable ��j��I����B
        ''' ������ ListTable �Ƃ��āA���X�g�̃��X�g (ArrayList<ArrayList>) ���󂯎��B
        ''' ������ index �Ƃ��āA�O���[�v�������ƂȂ��ԍ� (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āA�e�[�u���̏W�� (ArrayList<ArrayList<ArrayList>>) ��Ԃ��B
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
        ''' �����Ƀ\�[�g����B
        if seq.length() <= 1 then
            set mergesort = seq
            exit function
        end if
        
        dim half, ary1, ary2
        half = seq.length() \ 2 ' ������؂�̂Ă�����Ԃ�
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
        ''' �~���Ƀ\�[�g����B
        if seq.length() <= 1 then
            set rev_mergesort = seq
            exit function
        end if
        
        dim half, ary1, ary2
        half = seq.length() \ 2 ' ������؂�̂Ă�����Ԃ�
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
    
    ' �Ɩ��s���ɂ�����
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
        ''' �z������X�g�ɕϊ�����
        ''' ������ ary �Ƃ��āA�z�� (Array) ���󂯎��B
        ''' �߂�l�Ƃ��āA���X�g (ArrayList) ��Ԃ��B
        dim list, iter
        set list = new ArrayList
        for each iter in ary
            call list.push(iter)
        next
        set toList = list
    end function
    
    private function toListTable(byval ary2d)
        ''' �񎟌��z����e�[�u����̃��X�g�i���X�g�̃��X�g�j�ɂ��ĕԂ�
        ''' ������ ary2d �Ƃ��āA�񎟌��z�� (Array) ���󂯎��B
        ''' �߂�l�Ƃ��āA�e�[�u����̃��X�g (ArrayList) ��Ԃ��B
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
        ''' ���X�g����s�����̓񎟌��z��ɕϊ�����
        ''' ������ list �Ƃ��āA���X�g (ArrayList) ���󂯎��B
        ''' �߂�l�Ƃ��āA�񎟌��z�� (Array) ��Ԃ��B
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
        ''' ���X�g�̃��X�g��񎟌��z��Ɋi�[����
        ''' ������ list �Ƃ��āA���X�g�̃��X�g (ArrayList) ���󂯎��B
        ''' �߂�l�Ƃ��āA�񎟌��z�� (Array(,)) ��Ԃ��B
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
        ''' �e�[�u���S�̂𕶎���ɂ��ĕԂ��B
        ''' �e�[�u��������������Ă��Ȃ��ꍇ�́A��̕������Ԃ��B
        ''' �߂�l�Ƃ��āA�e�[�u���S�̂�\�������� (String) ��Ԃ��B
        
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
        
        ' �w�b�_�̐ݒ�
        call sb.append_3(join(head.toArray(), vbTab) & vbCrLf)
        dim col
        for each col in head.toArray()
            call sb.append_3(string(len(col), "-"))
            call sb.append_3(vbTab)
        next
        call sb.append_3(vbCrLf)
        
        ' �s�̐ݒ�
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
        
        ' �s�̐ݒ�
        dim row
        for each row in body.toArray()
            call sb.append_3(join(row.toArray(), vbTab))
            call sb.append_3(vbCrLf)
        next
        
        toStringWithoutHeader = sb.toString()
    end function
    
    public function setValue(byval r, byval c, byval v)
        ''' �w�肳�ꂽ�A�h���X�̒l��Ԃ��B�A�h���X�����݂��Ȃ���� Empty ��Ԃ��B
        ''' ������ r �Ƃ��āA�s�ԍ� (Number) ���󂯎��B
        ''' ������ c �Ƃ��āA��ԍ� (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āA�w�肳�ꂽ�A�h���X�̒l (Variant) ��Ԃ��B
        getValue = empty
        if isEmpty(head) then exit function
        if c < 0 or head.length() - 1 < c then exit function
        if r < 0 or body.length() - 1 < r then exit function
        
        call body.item(r).setItem(c, v)
    end function
    
    public function getValue(byval r, byval c)
        ''' �w�肳�ꂽ�A�h���X�̒l��Ԃ��B�A�h���X�����݂��Ȃ���� Empty ��Ԃ��B
        ''' ������ r �Ƃ��āA�s�ԍ� (Number) ���󂯎��B
        ''' ������ c �Ƃ��āA��ԍ� (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āA�w�肳�ꂽ�A�h���X�̒l (Variant) ��Ԃ��B
        getValue = empty
        if isEmpty(head) = 0 then exit function
        if c < 0 or head.length() - 1 < c then exit function
        if r < 0 or body.length() - 1 < r then exit function
        
        getValue = body.item(r).item(c)
    end function
    
    public function getRow(byval r)
        ''' �w�肳�ꂽ�s��z��ɂ��ĕԂ��B�s��������� Array() ��Ԃ��B
        ''' �s�ԍ��� 0 ����n�܂�B�w�b�_�����o���ɂ� describe ���\�b�h���g�p����B
        ''' ������ r �Ƃ��āA�擾����s�ԍ� (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āA�w�肳�ꂽ�s�̔z�� (Array) ��Ԃ��B
        getRow = array()
        if isEmpty(head) then exit function
        if body.length() = 0 then exit function
        if index < 0 or body.length() - 1 < r then exit function
        
        getRow = body.item(r).toArray()
    end function
    
    public function getColumn(byval col)
        ''' �w�肳�ꂽ���z��ɂ��ĕԂ��B�e�[�u���̍s���Ȃ���� Array() ��Ԃ��B
        ''' ������ col �Ƃ��āA�擾����w�b�_�� (String) ���󂯎��B
        ''' �߂�l�Ƃ��āA�w�肳�ꂽ��̔z�� (Array) ��Ԃ��B
        
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
        ''' �e�[�u���Ƀt�B�[���h��ǉ�����
        ''' ������ name �Ƃ��āA�ǉ�����t�B�[���h�̖��O (String) ���󂯎��B
        ''' ������ defaultValue �Ƃ��āA�ǉ������t�B�[���h�̏����l (Variant) ���󂯎��B
        ''' �߂�l�͕Ԃ��Ȃ��B
        
        call head.push(name) ' ToDo ���O�̃`�F�b�N
        
        dim i, len
        i = 0
        len = body.length()
        do while i < len
            body.item(i).push(defaultValue)
            i = i + 1
        loop
    end function
    
    public function modColumn(byval before, byval after)
        ''' �t�B�[���h����ύX����B
        ''' ������ before �Ƃ��āA�ύX�O�̖��O (String) ���󂯎��B
        ''' ������ after �Ƃ��āA�ύX��̖��O (String) ���󂯎��B
        ''' �߂�l�͕Ԃ��Ȃ��B
        
        if isEmpty(head) then call err.raise(12345, TypeName(me), "header is empty")
        
        dim index
        index = indexof(head, before)
        if index = -1 then call err.raise(12345, TypeName(me), "header not found: " & before)
        if indexof(head, after) <> -1 then call err.raise(12345, TypeName(me), "header already added: " & after)
        call head.setItem(index, after)
    end function
end class
