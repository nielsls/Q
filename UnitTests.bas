Option Explicit

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private lasttick As Long
Private testcounter As Long

Sub TestAll()
    testcounter = 0

    Dim a, b, c, d, e, arr
    a = [{ 1, 4, 7, 10, 13; 2, 5, 8, 11, 14; 3, 6, 9, 12, 15 }]
    b = [{ 1, 4, 7, 10, 13 }]
    c = [{ 1; 2; 3 }]
    d = 17
    e = Empty
    arr = Array(a, b, c, d, e)
    
    Tic
    Dim i As Long
    For i = 1 To 10
    
    Debug.Assert test("A", "true", True)
    Debug.Assert test("A", "false", False)
    Debug.Assert test("A", "42", 42)
    Debug.Assert test("A", """test string""", "test string")
    Debug.Assert test("A", "[]", Empty)
    Debug.Assert test("A", "[1 4 7 10 13]", b)
    Debug.Assert test("A", "[1; 2; 3]", c)
    Debug.Assert test("ans", "[]")
    Debug.Assert test("2+2;ans+1", "5")
    
    Debug.Assert test("#[]", "0")
    Debug.Assert test("[]==[]", "[]")
    Debug.Assert test("[]>=[]", "[]")
    Debug.Assert test("[]<=[]", "[]")
    Debug.Assert test("[]>[]", "[]")
    Debug.Assert test("[]<[]", "[]")
    Debug.Assert test("[]<>[]", "[]")
    Debug.Assert test("[]+3", "[]")
    Debug.Assert test("3-[]", "[]")
    Debug.Assert test("[]*3", "[]")
    Debug.Assert test("any([])", "false")
    Debug.Assert test("all([])", "true")
    Debug.Assert test("sum([])", "0")
    Debug.Assert test("prod([])", "1")
    
    Debug.Assert test("17e3", "17 * 1000")
    Debug.Assert test("17e-3+2", "17 * 0.001 + 2")
    Debug.Assert test("17.2e2", "1720")
    
    Debug.Assert test("1(1)", "1")                                     ' Scalar operators
    Debug.Assert test("---1+(+2)", "1")
    Debug.Assert test("16/2/2*3", "16./2./2.*3")
    Debug.Assert test("2^2^3", "2.^2.^3")
    
    Debug.Assert test("17*true", "-17")
    Debug.Assert test("rows(42)", "1")
    Debug.Assert test("cols(42)", "1")
    Debug.Assert test("numel(42)", "1")
    Debug.Assert test("size(42)", "[1 1]")
    Debug.Assert test("3", "rows(A)", a)                                     ' Size, dimension...
    Debug.Assert test("5", "cols(A)", a)
    Debug.Assert test("15", "numel(A)", a)
    Debug.Assert test("[3 5]", "size(A)", a)
    
    Debug.Assert test("1", "A(1)", a)                                          ' Indexing
    Debug.Assert test("1", "A(1,1)", a)
    Debug.Assert test("14", "A(2,end)", a)
    Debug.Assert test("15", "A(end)", a)
    Debug.Assert test("1:15", "A(:)'", a)
    Debug.Assert test("1:3:13", "A(1,:)", a)
    Debug.Assert test("(1:15)'", "A(:)", a)                            ' Colon operator
    Debug.Assert test("(1:15)", "A(1:end)", a)
    Debug.Assert test("(1:3:15)", "A(1:3:end)", a)
    Debug.Assert test("32", "2+#A*2", a)                                  ' Count operator
    Debug.Assert test("1", "inv(2)*2")
    
    Debug.Assert test("1&&2", "true")
    Debug.Assert test("1&&0", "false")
    Debug.Assert test("1||2", "true")
    Debug.Assert test("0||false", "false")
    Debug.Assert test("true||eye(4)", "true")
    
    Debug.Assert test("1>2", "false")                                        ' Comparison operators
    Debug.Assert test("1>=2", "false")
    Debug.Assert test("1<2", "true")
    Debug.Assert test("1<=2", "true")
    Debug.Assert test("1~=2", "true")
    Debug.Assert test("1<>2", "true")
    Debug.Assert test("1=2", "false")
    Debug.Assert test("1==2", "false")
    
    Debug.Assert test("2+3", 5)
    Debug.Assert test("2-3", -1)
    Debug.Assert test("2*3", 6)
    Debug.Assert test("8/4", 2)
    
    Debug.Assert test("sort(A)", "sort(A, ""descend"")(end:-1:1,:)", a)
    Debug.Assert test("1", "((7*A)./(A*7))(1,1)", a)  ' Arithmetic operators
    Debug.Assert test("0", "((7+A)-(A+7))(1,1)", a)
    Debug.Assert test("A", "A*eye(5)", a)
    Debug.Assert test("A", "eye(3)*A", a)
    Debug.Assert test("eye(3)", "round(A*inv(A))", Q("randn(3,3)"))     ' Matrix functions
    Debug.Assert test("7", "log(exp(A))", 7)
    Debug.Assert test("7", "log(exp(A))", 7)
    Debug.Assert test("7", "sqrt(A.^2)", 7)
    Debug.Assert test("[ 6, 15, 24, 33, 42 ]", "sum(A)", a)
    Debug.Assert test("true", "isempty([])")
    Debug.Assert test("false", "isempty(ones(3))")
    Debug.Assert test("true", "islogical(true)")
    Debug.Assert test("false", "islogical(17)")
    Debug.Assert test("true", "islogical(A>10)", a)
    Debug.Assert test("false", "islogical(A)", a)
    
    Dim fun
    For Each fun In Split("zeros,ones,eye,true,false,rand,randn", ",")
        Debug.Assert test("size(" + fun + ")", "[1,1]")
        Debug.Assert test("size(" + fun + "(7))", "[7,7]")
        Debug.Assert test("size(" + fun + "(3,4))", "[3,4]")
        Debug.Assert test("size(" + fun + "([3,4]))", "[3,4]")
        Debug.Assert test(fun + "(0)", "[]")
        Debug.Assert test(fun + "(0,0)", "[]")
        Debug.Assert test(fun + "([0,0])", "[]")
        Debug.Assert test("all(all(isnum(17*" + fun + "(3,4))))", "true")
    Next fun
    
    For Each fun In Split("zeros,ones,true,false", ",")
        Debug.Assert test("unique(" + fun + "(2,3))", CStr(fun))
    Next fun
    Debug.Assert test("unique(repmat(""hello"", 4,5))", """hello""")
    
    ' Error testing - expression must result in error
    Debug.Assert testerror("[](1)")
    Debug.Assert testerror("1(2)")
    Debug.Assert testerror("zeros(2)+ones(3)")
    Debug.Assert testerror("false(4)-true(5)")
    Debug.Assert testerror("2(")
    Debug.Assert testerror("2+(")
    Debug.Assert testerror("true&&eye(2)")
    Debug.Assert testerror("false||ones(2)")
    
    Debug.Assert test("true", "sum(sum((A-tick2ret(ret2tick(A))).^2))<0.00001", Q("randn(5,4)"))
    Debug.Assert test("true", "sum(sum((A-tick2ret(ret2tick(A,2),2)).^2))<0.00001", Q("randn(2,7)"))
    
    Dim arrItem: For Each arrItem In arr
        Debug.Assert test("A([])", "[]", arrItem)
        Debug.Assert test("A(:)", "reshape(A(1:end),#A,,)", arrItem)
        Debug.Assert test("A(:,:)", "A", arrItem)
        Debug.Assert test("A(1:end,1:end)", "A", arrItem)
        
        Debug.Assert test("rows(A)*cols(A)", "#A", arrItem)
        Debug.Assert test("size(A)", "[rows(A) cols(A)]", arrItem)
        Debug.Assert test("prod(size(A))", "#A", arrItem)
        Debug.Assert test("size(A')(end:-1:1)", "size(A)", arrItem)
        
        Debug.Assert test("A==A", "true(size(A))", arrItem)
        Debug.Assert test("A<>A", "false(size(A))", arrItem)
        Debug.Assert test("A>=A", "true(size(A))", arrItem)
        Debug.Assert test("A<=A", "true(size(A))", arrItem)
        Debug.Assert test("A>A", "false(size(A))", arrItem)
        Debug.Assert test("A<A", "false(size(A))", arrItem)
            
        Debug.Assert test("A>10", "~(A<=10)", arrItem)
        Debug.Assert test("A<10", "~(A>=10)", arrItem)
        Debug.Assert test("A==10", "~(A~=10)", arrItem)
        Debug.Assert test("A=10", "~(A<>10)", arrItem)
        Debug.Assert test("A>10|A=10", "A>=10", arrItem)
        Debug.Assert test("A<10|A=10", "A<=10", arrItem)
        Debug.Assert test("A>=10&A<=10", "A==10", arrItem)
        Debug.Assert test("xor(A>10,A<10)", "A<>10", arrItem)
        
        Debug.Assert test("all(A(A>10)>10)", "true", arrItem)
        Debug.Assert test("all(A(A>=10)>=10)", "true", arrItem)
        Debug.Assert test("all(A(A<10)<10)", "true", arrItem)
        Debug.Assert test("all(A(A<=10)<=10)", "true", arrItem)
        Debug.Assert test("all(A(A=10)=10)", "true", arrItem)
        Debug.Assert test("all(A(A<>10)<>10)", "true", arrItem)
        Debug.Assert test("all(A(find(A>10))>10)", "true", arrItem)

        Debug.Assert test("A''", "A", arrItem)
        Debug.Assert test("log(exp(A))", "A", arrItem)
        Debug.Assert test("sqrt(A.^2)", "A", arrItem)
        Debug.Assert test("islogical(A>10)", "true", a)
        Debug.Assert test("round(A+0.1)", "fix(A+0.2)", a)
        Debug.Assert test("round(-A+0.1)", "fix(-A-0.2)", a)
        Debug.Assert test("floor(A(A>0))", "fix(A(A>0))", a)
        Debug.Assert test("ceil(A(A<0))", "fix(A(A<0))", a)
        
        Debug.Assert test("A+A", "2*A", arrItem)
        Debug.Assert test("A-A", "0*A", arrItem)
        Debug.Assert test("A-A", "zeros(size(A))", arrItem)
        Debug.Assert test("A.*A.*A", "A.^3", arrItem)
        Debug.Assert test("A./A", "ones(size(A))", arrItem)
        
        Debug.Assert test("A*eye(cols(A))", "A", arrItem)
        Debug.Assert test("eye(rows(A))*A", "A", arrItem)
        
        Debug.Assert test("cumsum(A,1)(end,:)", "sum(A,1)", arrItem)
        Debug.Assert test("cumsum(A,2)(:,end)", "sum(A,2)", arrItem)
        Debug.Assert test("cumprod(A,1)(end,:)", "prod(A,1)", arrItem)
        Debug.Assert test("cumprod(A,2)(:,end)", "prod(A,2)", arrItem)
        
        Debug.Assert test("reshape(A,#A,[])", "A(:)", arrItem)
        Debug.Assert test("reshape(A,#A,,)", "A(:)", arrItem)
        Debug.Assert test("reshape(A,[],#A)", "A(:)'", arrItem)
        Debug.Assert test("reshape(A,,#A)", "A(:)'", arrItem)
        
        Debug.Assert test("rows(repmat(A,7,8))", "rows(A)*7", arrItem)
        Debug.Assert test("cols(repmat(A,7,8))", "cols(A)*8", arrItem)
        Debug.Assert test("sum(sum(repmat(A,7,8)))", "sum(sum(A))*7*8", arrItem)
        Debug.Assert test("repmat(A,0,3)", "[]", arrItem)
        Debug.Assert test("repmat(A,3,0)", "[]", arrItem)
        
        Debug.Assert test("repmat(A,1,3)", "[A A A]", arrItem)
        Debug.Assert test("repmat(A,3,1)", "[A; A; A]", arrItem)
        Debug.Assert test("repmat(A,2,3)", "[A A A; A A A]", arrItem)
        
        Debug.Assert test("sum(sum(A))", "sum(A(:))", arrItem)
        Debug.Assert test("prod(prod(A))", "prod(A(:))", arrItem)
        Debug.Assert test("max(max(A))", "max(A(:))", arrItem)
        Debug.Assert test("min(min(A))", "min(A(:))", arrItem)
        
        Debug.Assert test("sum(A,1)", "sum(A',2)'", arrItem)
        Debug.Assert test("prod(A,1)", "prod(A',2)'", arrItem)
        Debug.Assert test("max(A,,1)", "max(A',,2)'", arrItem)
        Debug.Assert test("min(A,,1)", "min(A',,2)'", arrItem)
        Debug.Assert test("cumsum(A,1)", "cumsum(A',2)'", arrItem)
        Debug.Assert test("cumprod(A,1)", "cumprod(A',2)'", arrItem)
        Debug.Assert test("cummax(A,1)", "cummax(A',2)'", arrItem)
        Debug.Assert test("cummin(A,1)", "cummin(A',2)'", arrItem)
        
        Debug.Assert test("cummax(A)", "-cummin(-A)", arrItem)
        Debug.Assert test("cummax(A)", "-cummin(-A,1)", arrItem)
        Debug.Assert test("cummax(A,2)", "-cummin(-A,2)", arrItem)
        Debug.Assert test("max(A)", "-min(-A)", arrItem)
        Debug.Assert test("max(A,,1)", "-min(-A,,1)", arrItem)
        Debug.Assert test("max(A,,2)", "-min(-A,,2)", arrItem)
        
        Debug.Assert test("#cov(A)", "cols(A)^2", arrItem)
        
        
        ' Error testing - expression must result in error
        Debug.Assert testerror("A(#A+1)", arrItem)
        Debug.Assert testerror("A(0)", arrItem)
        Debug.Assert testerror("A(-1)", arrItem)
        Debug.Assert testerror("A(999999)", arrItem)
        Debug.Assert testerror("A(1,0)", arrItem)
        Debug.Assert testerror("A(-1,1)", arrItem)
        Debug.Assert testerror("A")
        Debug.Assert testerror("A+B", 3)
        
    Next arrItem
    Next i
    Toc
    Debug.Print "#tests = " & testcounter / (i - 1)
End Sub

Function test(code1 As String, code2 As String, Optional item As Variant) As Boolean
    testcounter = testcounter + 1
    Dim r1 As Variant, r2 As Variant
    If IsMissing(item) Then
        r1 = Q(code1)
        r2 = Q(code2)
    Else
        r1 = Q(code1, item)
        r2 = Q(code2, item)
    End If
    test = Q("isequal(A,B)", r1, r2)
End Function

Function testerror(code, Optional item As Variant) As Boolean
    testcounter = testcounter + 1
    On Error GoTo FoundError
    If IsMissing(item) Then
        Q code
    Else
        Q code, item
    End If
    Exit Function
FoundError:
    testerror = True
End Function

Sub Tic()
    lasttick = GetTickCount
End Sub

Sub Toc()
    Debug.Print (GetTickCount - lasttick) / 1000# & " seconds elapsed"
End Sub
