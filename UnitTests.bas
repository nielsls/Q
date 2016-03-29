Option Explicit

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private lasttick As Long

Sub TestAll()

    Dim emptymat As Variant

    Dim a, b, c, d, e, arr
    a = [{ 1, 4, 7, 10, 13; 2, 5, 8, 11, 14; 3, 6, 9, 12, 15 }]
    b = [{ 1, 4, 7, 10, 13 }]
    c = [{ 1; 2; 3 }]
    d = 17
    e = Empty
    arr = Array(a, b, c, d, e)
    
    Tic
    
    Dim i: For i = 1 To 5

    Debug.Assert Q("isequal(A,B)", Q("true"), True)
    Debug.Assert Q("isequal(A,B)", Q("false"), False)
    Debug.Assert Q("isequal(A,B)", Q("42"), 42)
    Debug.Assert Q("isequal(A,B)", Q("""test string"""), "test string")
    
    Debug.Assert Q("isequal(A,B)", Q("[]"), emptymat)
    Debug.Assert Q("isequal(A,B)", Q("#[]"), 0)
    Debug.Assert Q("isequal(A,B)", Q("[]==[]"), emptymat)
    Debug.Assert Q("isequal(A,B)", Q("[]>=[]"), emptymat)
    Debug.Assert Q("isequal(A,B)", Q("[]<=[]"), emptymat)
    Debug.Assert Q("isequal(A,B)", Q("[]>[]"), emptymat)
    Debug.Assert Q("isequal(A,B)", Q("[]<[]"), emptymat)
    Debug.Assert Q("isequal(A,B)", Q("[]<>[]"), emptymat)
    Debug.Assert Q("isequal(A,B)", Q("any([])"), False)
    Debug.Assert Q("isequal(A,B)", Q("all([])"), True)
    Debug.Assert Q("isequal(A,B)", Q("sum([])"), 0)
    Debug.Assert Q("isequal(A,B)", Q("prod([])"), 1)
    
    Debug.Assert Q("isequal(A,B)", Q("17e3"), 17 * 1000)
    Debug.Assert Q("isequal(A,B)", Q("17e-3+2"), 17 * 0.001 + 2)
    Debug.Assert Q("isequal(A,B)", Q("17.2e2"), 1720)
    Debug.Assert Q("isequal(A,B)", Q("A+B+C+D", 1, 2, 3, 4), 10)

    Debug.Assert Q("isequal(A,B)", Q("1(1)"), 1)                                      ' Scalar operators
    Debug.Assert Q("isequal(A,B)", Q("---1+(+2)"), 1)
    Debug.Assert Q("isequal(A,B)", Q("16/2/2*3"), Q("16./2./2.*3"))
    Debug.Assert Q("isequal(A,B)", Q("2^2^3"), Q("2.^2.^3"))
    
    Debug.Assert Q("isequal(A,B)", Q("17*true"), -17)
    Debug.Assert Q("isequal(A,B)", Q("rows(A)", a), 3)                                      ' Size, dimension...
    Debug.Assert Q("isequal(A,B)", Q("rows(A)", a), 3)
    Debug.Assert Q("isequal(A,B)", Q("rows(42)"), 1)
    Debug.Assert Q("isequal(A,B)", Q("cols(A)", a), 5)
    Debug.Assert Q("isequal(A,B)", Q("cols(42)"), 1)
    Debug.Assert Q("isequal(A,B)", Q("numel(A)", a), 15)
    Debug.Assert Q("isequal(A,B)", Q("numel(42)"), 1)
    Debug.Assert Q("isequal(A,B)", Q("size(A)", a), [{ 3, 5 }])
    Debug.Assert Q("isequal(A,B)", Q("size(42)"), [{ 1, 1 }])
    Debug.Assert Q("isequal(A,B)", Q("A(1)", a), 1)                                           ' Indexing
    Debug.Assert Q("isequal(A,B)", Q("A(1,1)", a), 1)
    Debug.Assert Q("isequal(A,B)", Q("A(2,end)", a), 14)
    Debug.Assert Q("isequal(A,B)", Q("A(end)", a), 15)
    Debug.Assert Q("isequal(A,B)", Q("A(:)'", a), Q("1:15"))
    Debug.Assert Q("isequal(A,B)", Q("A(1,:)", a), Q("1:3:13"))
    Debug.Assert Q("isequal(A,B)", Q("A(:)", a), Q("(1:15)'"))                              ' Colon operator
    Debug.Assert Q("isequal(A,B)", Q("A(1:end)", a), Q("(1:15)"))
    Debug.Assert Q("isequal(A,B)", Q("A(1:3:end)", a), Q("(1:3:15)"))
    Debug.Assert Q("isequal(A,B)", Q("2+#A*2", a), 32)                                    ' Count operator
    
    Debug.Assert Q("isequal(A,B)", Q("inv(2)*2", a), 1)
    
    Debug.Assert Q("isequal(A,B)", Q("1>2"), False)                                         ' Comparison operators
    Debug.Assert Q("isequal(A,B)", Q("1>=2"), False)
    Debug.Assert Q("isequal(A,B)", Q("1<2"), True)
    Debug.Assert Q("isequal(A,B)", Q("1<=2"), True)
    Debug.Assert Q("isequal(A,B)", Q("1~=2"), True)
    Debug.Assert Q("isequal(A,B)", Q("1<>2"), True)
    Debug.Assert Q("isequal(A,B)", Q("1=2"), False)
    Debug.Assert Q("isequal(A,B)", Q("1==2"), False)
    
    Debug.Assert Q("isequal(A,B)", Q("sort(A)", a), Q("sort(A, ""descend"")(end:-1:1,:)", a))
    Debug.Assert Q("isequal(A,B)", Q("((7*A)./(A*7))(1,1)", a), 1)   ' Arithmetic operators
    Debug.Assert Q("isequal(A,B)", Q("((7+A)-(A+7))(1,1)", a), 0)
    Debug.Assert Q("isequal(A,B)", Q("A*eye(5)", a), a)
    Debug.Assert Q("isequal(A,B)", Q("eye(3)*A", a), a)
    Debug.Assert Q("isequal(A,B)", Q("round(A*inv(A))", Q("randn(3,3)")), Q("eye(3)"))      ' Matrix functions
    Debug.Assert Q("isequal(A,B)", Q("log(exp(A))", 7), 7)
    Debug.Assert Q("isequal(A,B)", Q("log(exp(A))", 7), 7)
    Debug.Assert Q("isequal(A,B)", Q("sqrt(A.^2)", 7), 7)
    Debug.Assert Q("isequal(A,B)", Q("sum(A)", a), [{ 6, 15, 24, 33, 42 }])
    Debug.Assert Q("isequal(A,B)", Q("isempty(A)", Empty), True)
    Debug.Assert Q("isequal(A,B)", Q("isempty(A)", Q("ones(3)")), False)
    Debug.Assert Q("isequal(A,B)", Q("islogical(true)"), True)
    Debug.Assert Q("isequal(A,B)", Q("islogical(17)"), False)
    Debug.Assert Q("isequal(A,B)", Q("islogical(A>10)", a), True)
    Debug.Assert Q("isequal(A,B)", Q("islogical(A)", a), False)
    
    Dim f: For Each f In Split("zeros,ones,eye,true,false,rand,randn", ",")
        Debug.Assert test("size(" + f + ")", "[1,1]")
        Debug.Assert test("size(" + f + "(7))", "[7,7]")
        Debug.Assert test("size(" + f + "(3,4))", "[3,4]")
        Debug.Assert test("size(" + f + "([3,4]))", "[3,4]")
        Debug.Assert test(f + "(0)", "[]")
        Debug.Assert test(f + "(0,0)", "[]")
        Debug.Assert test(f + "([0,0])", "[]")
        Debug.Assert test("all(all(isnum(17*" + f + "(3,4))))", "true")
    Next f
    
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
    Next arrItem
    
    Next i
    
    Toc
    
End Sub

Function test(code1 As String, code2 As String, Optional item As Variant) As Boolean
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

Sub Tic()
    lasttick = GetTickCount
End Sub

Sub Toc()
    Debug.Print (GetTickCount - lasttick) / 1000# & " seconds elapsed"
End Sub
