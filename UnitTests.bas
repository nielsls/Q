Sub TestAll()

    Dim a, b, c, d, arr, arrItem
    a = [{ 1, 4, 7, 10, 13; 2, 5, 8, 11, 14; 3, 6, 9, 12, 15 }]
    b = [{ 1, 4, 7, 10, 13 }]
    c = [{ 1; 2; 3 }]
    d = 17
    arr = Array(a, b, c, d)

    Debug.Assert Q("isequal(a,b)", Q("[]"), [NA()])

    Debug.Assert Q("isequal(a,b)", Q("-1.2(1)"), -1.2)                                      ' Scalar operators
    Debug.Assert Q("isequal(a,b)", Q("---1+(+2)"), 1)
    Debug.Assert Q("isequal(a,b)", Q("16/2/2*3"), Q("16./2./2.*3"))
    Debug.Assert Q("isequal(a,b)", Q("2^2^3"), Q("2.^2.^3"))
    Debug.Assert Q("isequal(a,b)", Q("true"), True)
    Debug.Assert Q("isequal(a,b)", Q("false"), False)
    
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
    Debug.Assert Q("isequal(a,b)", Q("#[]"), 0)
    
    Debug.Assert Q("isequal(a,b)", Q("1>2"), False)                                         ' Comparison operators
    Debug.Assert Q("isequal(a,b)", Q("1>=2"), False)
    Debug.Assert Q("isequal(a,b)", Q("1<2"), True)
    Debug.Assert Q("isequal(a,b)", Q("1<=2"), True)
    Debug.Assert Q("isequal(a,b)", Q("1~=2"), True)
    Debug.Assert Q("isequal(a,b)", Q("1<>2"), True)
    Debug.Assert Q("isequal(a,b)", Q("1=2"), False)
    Debug.Assert Q("isequal(a,b)", Q("1==2"), False)
    
    For Each arrItem In arr
        Debug.Assert Q("isequal(a,b)", Q("rows(a)*cols(a)", arrItem), Q("#a", arrItem))
    
        Debug.Assert Q("isequal(a,b)", Q("a>10", arrItem), Q("~(a<=10)", arrItem))
        Debug.Assert Q("isequal(a,b)", Q("a<10", arrItem), Q("~(a>=10)", arrItem))
        Debug.Assert Q("isequal(a,b)", Q("a==10", arrItem), Q("~(a~=10)", arrItem))
        Debug.Assert Q("isequal(a,b)", Q("a=10", arrItem), Q("~(a<>10)", arrItem))
        Debug.Assert Q("isequal(a,b)", Q("a>10|a=10", arrItem), Q("a>=10", arrItem))
        Debug.Assert Q("isequal(a,b)", Q("a<10|a=10", arrItem), Q("a<=10", arrItem))

        Debug.Assert Q("isequal(a,b)", Q("a''", arrItem), arrItem)
        Debug.Assert Q("isequal(a,b)", Q("log(exp(a))", arrItem), arrItem)
        Debug.Assert Q("isequal(a,b)", Q("sqrt(a.^2)", arrItem), arrItem)
        
        Debug.Assert Q("isequal(a,b)", Q("a+a", arrItem), Q("2*a", arrItem))
        Debug.Assert Q("isequal(a,b)", Q("a-a", arrItem), Q("0*a", arrItem))
        Debug.Assert Q("isequal(a,b)", Q("a-a", arrItem), Q("zeros(size(a))", arrItem))
        Debug.Assert Q("isequal(a,b)", Q("a./a", arrItem), Q("ones(size(a))", arrItem))
        Debug.Assert Q("isequal(a,b)", Q("a.*a.*a", arrItem), Q("a.^3", arrItem))
        
        Debug.Assert Q("isequal(a,b)", Q("cumsum(a,1)(end,:)", arrItem), Q("sum(a,1)", arrItem))
        Debug.Assert Q("isequal(a,b)", Q("cumsum(a,2)(:,end)", arrItem), Q("sum(a,2)", arrItem))
        Debug.Assert Q("isequal(a,b)", Q("cumprod(a,1)(end,:)", arrItem), Q("prod(a,1)", arrItem))
        Debug.Assert Q("isequal(a,b)", Q("cumprod(a,2)(:,end)", arrItem), Q("prod(a,2)", arrItem))
        
        Debug.Assert Q("isequal(a,b)", Q("reshape(a,#a,[])", arrItem), Q("a(:)", arrItem))
        Debug.Assert Q("isequal(a,b)", Q("reshape(a,[],#a)", arrItem), Q("a(:)'", arrItem))
        
        Debug.Assert Q("isequal(a,b)", Q("rows(repmat(a,7,8))", arrItem), Q("rows(a)*7", arrItem))
        Debug.Assert Q("isequal(a,b)", Q("cols(repmat(a,7,8))", arrItem), Q("cols(a)*8", arrItem))
        Debug.Assert Q("isequal(a,b)", Q("sum(sum(repmat(a,7,8)))", arrItem), Q("sum(sum(a))*7*8", arrItem))
        
        Debug.Assert Q("isequal(a,b)", Q("cummax(a)", a), Q("-cummin(-a,1)", a))
        Debug.Assert Q("isequal(a,b)", Q("cummax(a,2)", a), Q("-cummin(-a,2)", a))
    Next arrItem

    Debug.Assert Q("isequal(a,b)", Q("sort(a)", a), Q("sort(a, ""descend"")(end:-1:1,:)", a))

    Debug.Assert Q("isequal(a,b)", Q("7''"), 7)                                            ' Arithmetic operators
    Debug.Assert Q("isequal(a,b)", Q("((7*a)./(a*7))(1,1)", a), 1)
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
    

End Sub
