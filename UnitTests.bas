Option Explicit

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private lasttick As Long

Sub TestOne()
    Dim a, b, c, d, arr
    a = [{ 1, 4, 7, 10, 13; 2, 5, 8, 11, 14; 3, 6, 9, 12, 15 }]
    b = [{ 1, 4, 7, 10, 13 }]
    c = [{ 1; 2; 3 }]
    d = 17
    
    Dim x: x = Q("a+b", a, b)
End Sub

Sub TestAll()

    Dim a, b, c, d, e, arr
    a = [{ 1, 4, 7, 10, 13; 2, 5, 8, 11, 14; 3, 6, 9, 12, 15 }]
    b = [{ 1, 4, 7, 10, 13 }]
    c = [{ 1; 2; 3 }]
    d = 17
    e = Empty
    arr = Array(a, b, c, d, e)
    
    Tic
    
    Dim i: For i = 1 To 5

    Debug.Assert Q("isequal(a,b)", Q("true"), True)
    Debug.Assert Q("isequal(a,b)", Q("false"), False)
    Debug.Assert Q("isequal(a,b)", Q("42"), 42)
    Debug.Assert Q("isequal(a,b)", Q("""test string"""), "test string")
    Debug.Assert Q("isequal(a,b)", Q("[]"), [NA()])
    Debug.Assert Q("isequal(a,b)", Q("#[]"), 0)
    Debug.Assert Q("isequal(a,b)", Q("[]==[]"), [NA()])
    Debug.Assert Q("isequal(a,b)", Q("[]>=[]"), [NA()])
    Debug.Assert Q("isequal(a,b)", Q("[]<=[]"), [NA()])
    Debug.Assert Q("isequal(a,b)", Q("[]>[]"), [NA()])
    Debug.Assert Q("isequal(a,b)", Q("[]<[]"), [NA()])
    Debug.Assert Q("isequal(a,b)", Q("[]<>[]"), [NA()])
    Debug.Assert Q("isequal(a,b)", Q("17e3"), 17 * 1000)
    Debug.Assert Q("isequal(a,b)", Q("17e-3+2"), 17 * 0.001 + 2)
    Debug.Assert Q("isequal(a,b)", Q("17.2e2"), 1720)
    Debug.Assert Q("isequal(a,b)", Q("a+b+c+d", 1, 2, 3, 4), 10)

    Debug.Assert Q("isequal(a,b)", Q("1(1)"), 1)                                      ' Scalar operators
    Debug.Assert Q("isequal(a,b)", Q("---1+(+2)"), 1)
    Debug.Assert Q("isequal(a,b)", Q("16/2/2*3"), Q("16./2./2.*3"))
    Debug.Assert Q("isequal(a,b)", Q("2^2^3"), Q("2.^2.^3"))
    
    Debug.Assert Q("isequal(a,b)", Q("17*true"), -17)
    Debug.Assert Q("isequal(a,b)", Q("rows(a)", a), 3)                                      ' Size, dimension...
    Debug.Assert Q("isequal(a,b)", Q("rows(a)", a), 3)
    Debug.Assert Q("isequal(a,b)", Q("rows(42)"), 1)
    Debug.Assert Q("isequal(a,b)", Q("cols(a)", a), 5)
    Debug.Assert Q("isequal(a,b)", Q("cols(42)"), 1)
    Debug.Assert Q("isequal(a,b)", Q("numel(a)", a), 15)
    Debug.Assert Q("isequal(a,b)", Q("numel(42)"), 1)
    Debug.Assert Q("isequal(a,b)", Q("size(a)", a), [{ 3, 5 }])
    Debug.Assert Q("isequal(a,b)", Q("size(42)"), [{ 1, 1 }])
    Debug.Assert Q("isequal(a,b)", Q("a(1)", a), 1)                                           ' Indexing
    Debug.Assert Q("isequal(a,b)", Q("a(1,1)", a), 1)
    Debug.Assert Q("isequal(a,b)", Q("a(2,end)", a), 14)
    Debug.Assert Q("isequal(a,b)", Q("a(end)", a), 15)
    Debug.Assert Q("isequal(a,b)", Q("a(:)'", a), Q("1:15"))
    Debug.Assert Q("isequal(a,b)", Q("a(1,:)", a), Q("1:3:13"))
    Debug.Assert Q("isequal(a,b)", Q("a(:)", a), Q("(1:15)'"))                              ' Colon operator
    Debug.Assert Q("isequal(a,b)", Q("a(1:end)", a), Q("(1:15)"))
    Debug.Assert Q("isequal(a,b)", Q("a(1:3:end)", a), Q("(1:3:15)"))
    Debug.Assert Q("isequal(a,b)", Q("2+#a*2", a), 32)                                    ' Count operator
    
    Debug.Assert Q("isequal(a,b)", Q("inv(2)*2", a), 1)
    
    Debug.Assert Q("isequal(a,b)", Q("1>2"), False)                                         ' Comparison operators
    Debug.Assert Q("isequal(a,b)", Q("1>=2"), False)
    Debug.Assert Q("isequal(a,b)", Q("1<2"), True)
    Debug.Assert Q("isequal(a,b)", Q("1<=2"), True)
    Debug.Assert Q("isequal(a,b)", Q("1~=2"), True)
    Debug.Assert Q("isequal(a,b)", Q("1<>2"), True)
    Debug.Assert Q("isequal(a,b)", Q("1=2"), False)
    Debug.Assert Q("isequal(a,b)", Q("1==2"), False)
    
    Debug.Assert Q("isequal(a,b)", Q("any([])"), False)
    Debug.Assert Q("isequal(a,b)", Q("all([])"), True)
    Debug.Assert Q("isequal(a,b)", Q("sum([])"), 0)
    Debug.Assert Q("isequal(a,b)", Q("prod([])"), 1)
    
    Debug.Assert Q("isequal(a,b)", Q("sort(a)", a), Q("sort(a, ""descend"")(end:-1:1,:)", a))
    Debug.Assert Q("isequal(a,b)", Q("((7*a)./(a*7))(1,1)", a), 1)   ' Arithmetic operators
    Debug.Assert Q("isequal(a,b)", Q("((7+a)-(a+7))(1,1)", a), 0)
    Debug.Assert Q("isequal(a,b)", Q("a*eye(5)", a), a)
    Debug.Assert Q("isequal(a,b)", Q("eye(3)*a", a), a)
    Debug.Assert Q("isequal(a,b)", Q("round(a*inv(a))", Q("randn(3,3)")), Q("eye(3)"))      ' Matrix functions
    Debug.Assert Q("isequal(a,b)", Q("log(exp(a))", 7), 7)
    Debug.Assert Q("isequal(a,b)", Q("log(exp(a))", 7), 7)
    Debug.Assert Q("isequal(a,b)", Q("sqrt(a.^2)", 7), 7)
    Debug.Assert Q("isequal(a,b)", Q("sum(a)", a), [{ 6, 15, 24, 33, 42 }])
    Debug.Assert Q("isequal(a,b)", Q("isempty(a)", Empty), True)
    Debug.Assert Q("isequal(a,b)", Q("isempty(a)", Q("ones(3)")), False)
    Debug.Assert Q("isequal(a,b)", Q("islogical(true)"), True)
    Debug.Assert Q("isequal(a,b)", Q("islogical(17)"), False)
    Debug.Assert Q("isequal(a,b)", Q("islogical(a>10)", a), True)
    Debug.Assert Q("isequal(a,b)", Q("islogical(a)", a), False)
    
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
        Debug.Assert test("a([])", "[]", arrItem)
        
        Debug.Assert test("rows(a)*cols(a)", "#a", arrItem)
        Debug.Assert test("size(a)", "[rows(a) cols(a)]", arrItem)
        Debug.Assert test("size(a')(end:-1:1)", "size(a)", arrItem)
    
        Debug.Assert test("a>10", "~(a<=10)", arrItem)
        Debug.Assert test("a<10", "~(a>=10)", arrItem)
        Debug.Assert test("a==10", "~(a~=10)", arrItem)
        Debug.Assert test("a=10", "~(a<>10)", arrItem)
        Debug.Assert test("a>10|a=10", "a>=10", arrItem)
        Debug.Assert test("a<10|a=10", "a<=10", arrItem)
        Debug.Assert test("a>=10&a<=10", "a==10", arrItem)
        Debug.Assert test("xor(a>10,a<10)", "a<>10", arrItem)
        
        Debug.Assert test("all(a(a>10)>10)", "true", arrItem)
        Debug.Assert test("all(a(a>=10)>=10)", "true", arrItem)
        Debug.Assert test("all(a(a<10)<10)", "true", arrItem)
        Debug.Assert test("all(a(a<=10)<=10)", "true", arrItem)
        Debug.Assert test("all(a(a=10)=10)", "true", arrItem)
        Debug.Assert test("all(a(a<>10)<>10)", "true", arrItem)
        Debug.Assert test("all(a(find(a>10))>10)", "true", arrItem)

        Debug.Assert test("a''", "a", arrItem)
        Debug.Assert test("log(exp(a))", "a", arrItem)
        Debug.Assert test("sqrt(a.^2)", "a", arrItem)
        Debug.Assert test("islogical(a>10)", "true", a)
        Debug.Assert test("round(a+0.1)", "fix(a+0.2)", a)
        Debug.Assert test("floor(a(a>0))", "fix(a(a>0))", a)
        Debug.Assert test("ceil(a(a<0))", "fix(a(a<0))", a)
        
        Debug.Assert test("a+a", "2*a", arrItem)
        Debug.Assert test("a-a", "0*a", arrItem)
        Debug.Assert test("a-a", "zeros(size(a))", arrItem)
        Debug.Assert test("a.*a.*a", "a.^3", arrItem)
        Debug.Assert test("a./a", "ones(size(a))", arrItem)
        
        Debug.Assert test("cumsum(a,1)(end,:)", "sum(a,1)", arrItem)
        Debug.Assert test("cumsum(a,2)(:,end)", "sum(a,2)", arrItem)
        Debug.Assert test("cumprod(a,1)(end,:)", "prod(a,1)", arrItem)
        Debug.Assert test("cumprod(a,2)(:,end)", "prod(a,2)", arrItem)
        
        Debug.Assert test("reshape(a,#a,[])", "a(:)", arrItem)
        Debug.Assert test("reshape(a,#a,,)", "a(:)", arrItem)
        Debug.Assert test("reshape(a,[],#a)", "a(:)'", arrItem)
        Debug.Assert test("reshape(a,,#a)", "a(:)'", arrItem)
        
        Debug.Assert test("rows(repmat(a,7,8))", "rows(a)*7", arrItem)
        Debug.Assert test("cols(repmat(a,7,8))", "cols(a)*8", arrItem)
        Debug.Assert test("sum(sum(repmat(a,7,8)))", "sum(sum(a))*7*8", arrItem)
        Debug.Assert test("repmat(a,0,3)", "[]", arrItem)
        Debug.Assert test("repmat(a,3,0)", "[]", arrItem)
        
        Debug.Assert test("cummax(a)", "-cummin(-a)", arrItem)
        Debug.Assert test("cummax(a)", "-cummin(-a,1)", arrItem)
        Debug.Assert test("cummax(a,2)", "-cummin(-a,2)", arrItem)
        
        Debug.Assert test("max(a)", "-min(-a)", arrItem)
        Debug.Assert test("max(a,,1)", "-min(-a,,1)", arrItem)
        Debug.Assert test("max(a,,2)", "-min(-a,,2)", arrItem)
        
        Debug.Assert test("corr(a)", "corr(a,,1)", arrItem)
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
    test = Q("isequal(a,b)", r1, r2)
End Function

Sub Tic()
    lasttick = GetTickCount
End Sub

Sub Toc()
    Debug.Print (GetTickCount - lasttick) / 1000# & " seconds elapsed"
End Sub
