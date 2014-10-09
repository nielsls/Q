'*******************************************************************
' Q - A MATLAB-like matrix parser for Microsoft Excel
' Version 1.0
'
' Q features a single public function, Q(), containing an expression parser.
' Q() is able to parse and evaluate a subset of the MATLAB programming language.
' It features almost all MATLAB operators, selected standard functions
' and has complete support for submatrices, '()', and concatenation, '[]'.
'
' Example usage:
'   - =Q("2+2")                -> 4
'   - =Q("a+b+c",3,4,5)        -> 12
'   - =Q("eye(3)")             -> the 3x3 identity matrix
'   - =Q("mean(a)",A1:D5)      -> row vector with the mean of each column in cells A1:D5
'   - =Q("a.*b",A1:D5,F1:I5)   -> element wise multiplication of cells A1:D5 and F1:I5
'   - =Q("a([1 3],end)",A1:D5) -> 2x1 matrix with the last entries in row 1 and 3 of cells A1:D5
'
' Features:
'   - All standard MATLAB operators: :,::,+,-,*,/,.*,./,^,.^,||,&&,|,&,<,<=,>,>=,==,~=,~,'
'   - Most used MATLAB functions: eye,zeros,ones,sum,cumsum,cumprod,prod,
'     mean,median,prctile,std,isequal,fix,rand,randn,repmat,find,sqrt,exp,inv...
'   - Indexing via a(2,:) or a(5,3:end)
'   - Concatenate matrices with '[]', i.e. [ a b; c d]
'   - Excel functions: if,iferror
'   - Prefix function calls with ! to call external VBA functions not found in Q.
'
' For the newest version, go to:
' http://github.com/nielsls/Q
'
' 2014, Niels Lykke Sørensen

Option Explicit
Option Base 1

Private Const REGEXPATTERN = _
    " |\""[^\""]*\""|\(|\)|\[|\]|\,|\;" & _
    "|[a-zA-Z][a-zA-Z0-9_]*|[0-9]+\.?[0-9]*" & _
    "|\|\||&&|\||&|<>|~=|<=|>=|<|>|==|=" & _
    "|\:|\+|\-|\*|\.\*|\/|\.\/|\^|\.\^|\'|~|#|!"
Private arguments As Variant
Private tokens As Object
Private tokenIndex As Long
Private endValues As Variant
Private errorMsg As String

' Entry point - the only public function in the library
Public Function Q(expr As Variant, ParamArray args() As Variant) As Variant
    On Error GoTo ErrorHandler

    arguments = args
    tokenIndex = 0
    endValues = Empty
    errorMsg = ""

    With CreateObject("VBScript.RegExp")
        'Split input string into valid tokens
        .Global = True
        .IgnoreCase = False
        .pattern = REGEXPATTERN
        Set tokens = .Execute(expr)
        
        'Check valid tokens constitute the full input string
        Dim i As Long, token As Variant
        For Each token In tokens
            Assert i = token.FirstIndex, "Illegal token: " & Mid(expr, i + 1, 1)
            i = i + token.Length
        Next token
        Assert i = Len(expr), "Illegal token: " & Mid(expr, i + 1, 1)
    End With

    Dim root As Variant
    root = Parse_Binary()

    'Utils_DumpTree root    'Uncomment for debugging

    Q = eval_tree(root)

    Assert Tokens_Current() = "", "'" & Tokens_Current & "' not expected here."
    Exit Function
    
ErrorHandler:
    ' Cannot rely on usual error message transfer only as
    ' this does not work when using Application.Run()
    If errorMsg = "" Then errorMsg = err.Description
    Q = "ERROR - " & errorMsg
End Function

'*******************************************************************
' Evaluator invariants:
'   - All variables must be either strings, scalars or 2d arrays.
'     A variable cannot be a 1x1 array; then it must be a scalar.
'     Use Utils_Conform() to correctly shape all variables.
'   - Variable names: [a-z]
'   - Function names: [a-zA-Z][a-zA-Z0-9_]*
'   - The empty matrix/scalar [] has per definition 0 rows, 0 cols,
'     dimension 0 and is internally represented by value #NA / NA()
'*******************************************************************

'***************************
'*** SUPPORTED OPERATORS ***
'***************************

'Returns true if token is a suitable operator
'op = Array( <function name>, <precedence level>, <left associative> )
Private Function Parse_FindOp(token As String, opType As String, ByRef op As Variant) As Boolean
    op = Null
    
    Select Case opType
        Case "binary"
            Select Case token
                Case "||": op = Array("orshortcircuit", 1, True)
                Case "&&": op = Array("andshortcircuit", 2, True)
                Case "|": op = Array("or", 3, True)
                Case "&": op = Array("and", 4, True)
                Case "<": op = Array("lt", 5, True)
                Case "<=": op = Array("lte", 5, True)
                Case ">": op = Array("gt", 5, True)
                Case ">=": op = Array("gte", 5, True)
                Case "==": op = Array("eq", 5, True)
                Case "=": op = Array("eq", 5, True)
                Case "~=": op = Array("ne", 5, True)
                Case "<>": op = Array("ne", 5, True)
                Case ":": op = Array("colon", 6, False)
                Case "+": op = Array("plus", 7, True)
                Case "-": op = Array("minus", 7, True)
                Case "*": op = Array("mtimes", 8, True)
                Case ".*": op = Array("times", 8, True)
                Case "/": op = Array("mdivide", 8, True)
                Case "./": op = Array("divide", 8, True)
                Case "^": op = Array("mpower", 9, True)
                Case ".^": op = Array("power", 9, True)
            End Select
            
        Case "unaryprefix"
            Select Case token
                Case "+": op = "uplus"
                Case "-": op = "uminus"
                Case "~": op = "negate"
                Case "#": op = "numel"
                Case "!": op = "extern"
            End Select
            
        Case "unarypostfix"
            Select Case token
                Case "'": op = "transpose"
            End Select
    End Select
    
    Parse_FindOp = Not IsNull(op)
End Function

'******************
'*** BUILD TREE ***
'******************
Private Function Parse_Matrix() As Variant
    Do While Tokens_Current() <> "]"
        Utils_Stack_Push Parse_List(True), Parse_Matrix
        If Tokens_Current() = ";" Then Tokens_Advance
    Loop
End Function

Private Function Parse_List(Optional isSpaceSeparator As Boolean = False) As Variant
    Do
        Utils_Stack_Push Parse_Binary(), Parse_List
        Select Case Tokens_Current()
            Case ";", ")", "]": Exit Do
            Case ",": Tokens_Advance
            Case Else: If Tokens_Previous() <> " " Or Not isSpaceSeparator Then Exit Do
        End Select
    Loop While True
End Function

Private Function Parse_Binary(Optional lastPrec As Long = -999) As Variant
    Parse_Binary = Parse_Prefix()
    Dim op: Do While Parse_FindOp(Tokens_Current(), "binary", op)
        If op(2) + CLng(op(3)) < lastPrec Then Exit Do
        Tokens_Advance
        Parse_Binary = Array("op_" & op(1), Array(Parse_Binary, Parse_Binary(CLng(op(2)))))
    Loop
End Function

Private Function Parse_Prefix() As Variant
    Dim op
    If Not Parse_FindOp(Tokens_Current(), "unaryprefix", op) Then
        Parse_Prefix = Parse_Postfix()
    Else
        Tokens_Advance
        Parse_Prefix = Array("op_" & op, Array(Parse_Prefix()))
    End If
End Function

Private Function Parse_Postfix() As Variant
    Parse_Postfix = Parse_Atomic
    Dim op: Do
        If Parse_FindOp(Tokens_Current(), "unarypostfix", op) Then
            Parse_Postfix = Array("op_" & op, Array(Parse_Postfix))
            Tokens_Advance
        ElseIf Tokens_Current() = "(" Then
            Tokens_Advance
            Parse_Postfix = Array("eval_index", Array(Parse_Postfix, Parse_List()))
            Tokens_AssertAndAdvance ")"
        Else
            Exit Do
        End If
    Loop While True
End Function

Private Function Parse_Atomic() As Variant
    Dim token: token = Tokens_Current()
    Select Case token
        Case ""
            Assert False, "Missing argument"
            
        Case "true"
            Parse_Atomic = Array("eval_constant", Array(True))
            Tokens_Advance
            
        Case "false"
            Parse_Atomic = Array("eval_constant", Array(False))
            Tokens_Advance
            
        Case "end"
            Parse_Atomic = Array("eval_end", Array())
            Tokens_Advance
            
        Case ":"
            Parse_Atomic = Array("eval_colon", Array())
            Tokens_Advance
            
        Case "("
            Tokens_Advance
            Parse_Atomic = Parse_Binary()
            Tokens_AssertAndAdvance ")"
            
        Case "["
            Tokens_Advance
            Parse_Atomic = Array("eval_concat", Parse_Matrix())
            Tokens_AssertAndAdvance "]"
        
        Case Else
            Select Case Asc(token) 'filter on first char of token
                Case Asc("""")
                    Parse_Atomic = Array("eval_constant", Array(Mid(token, 2, Len(token) - 2)))
                    Tokens_Advance
                    
                Case Asc("0") To Asc("9")
                    Parse_Atomic = Array("eval_constant", Array(Val(token)))
                    Tokens_Advance
                    
                Case Asc("a") To Asc("z"), Asc("A") To Asc("Z")
                    If Len(token) = 1 Then
                        Parse_Atomic = Array("eval_arg", Array(Asc(token) - Asc("a")))
                        Tokens_Advance
                    Else
                        Tokens_Advance
                        Tokens_AssertAndAdvance "("
                        Parse_Atomic = Array("fn_" & token, Parse_List())
                        Tokens_AssertAndAdvance ")"
                    End If
                    
                Case Else
                    Assert False, "Unexpected token: " & token
            End Select
    End Select
End Function

'*********************
'*** TOKEN CONTROL ***
'*********************
Private Function Tokens_Advance() As Boolean
    Do
        tokenIndex = tokenIndex + 1
    Loop While Tokens_Current() = " "
    Tokens_Advance = tokenIndex < tokens.count
End Function

Private Function Tokens_Current() As Variant
    If tokenIndex < tokens.count Then
        Tokens_Current = tokens(tokenIndex)
    Else
        Tokens_Current = ""
    End If
End Function

Private Function Tokens_Previous() As Variant
    If tokenIndex - 1 >= 0 And tokenIndex - 1 < tokens.count Then
        Tokens_Previous = tokens(tokenIndex - 1)
    Else
        Tokens_Previous = ""
    End If
End Function

Private Sub Tokens_AssertAndAdvance(token As String)
    Assert token = Tokens_Current(), "Missing token: " & token
    Tokens_Advance
End Sub

'*************
'*** UTILS ***
'*************
Private Sub Utils_DumpTree(tree As Variant, Optional spacer As String = "")
    If Utils_Dimensions(tree) > 0 Then
        Dim leaf: For Each leaf In tree
            Utils_DumpTree leaf, spacer & "  "
        Next leaf
    Else
        Debug.Print spacer & tree
    End If
End Sub

Private Function Utils_Dimensions(v As Variant) As Long
    Dim dimnum As Long, errorCheck As Integer
    On Error GoTo FinalDimension
    For dimnum = 1 To 60000
        errorCheck = LBound(v, dimnum)
    Next
FinalDimension:
    Utils_Dimensions = dimnum - 1
End Function

Private Function Utils_Numel(v As Variant) As Long
    Select Case Utils_Dimensions(v)
        Case 0: If WorksheetFunction.IsNA(v) Then Utils_Numel = 0 Else Utils_Numel = 1
        Case 1: Utils_Numel = UBound(v)
        Case 2: Utils_Numel = UBound(v, 1) * UBound(v, 2)
        Case Else: Assert False, "Dimension > 2"
    End Select
End Function

Private Function Utils_IsNA(v As Variant) As Boolean
    Utils_IsNA = Utils_Numel(v) = 0
End Function

Private Sub Utils_Conform(ByRef v As Variant)
    Dim r As Variant, i As Long
    Select Case Utils_Dimensions(v)
        Case 1:
            If UBound(v) = 1 Then
                v = v(1)
            Else
                ReDim r(1 To 1, 1 To UBound(v)) As Variant
                For i = 1 To UBound(r, 2)
                    r(1, i) = v(i)
                Next i
                v = r
            End If
            
        Case 2:
            If UBound(v, 1) = 1 And UBound(v, 2) = 1 Then v = v(1, 1)
            
        Case Is > 2:
            Assert False, "Dimension > 2"
    End Select
End Sub

Private Sub Utils_ForceMatrix(v As Variant)
    If Utils_Dimensions(v) = 0 Then
        Dim r: ReDim r(1 To 1, 1 To 1) As Variant
        r(1, 1) = v
        v = r
    End If
End Sub

Private Function Utils_Rows(ByRef v As Variant) As Long
    Select Case Utils_Numel(v)
        Case 0: Utils_Rows = 0
        Case 1: Utils_Rows = 1
        Case Else: Utils_Rows = UBound(v, 1)
    End Select
End Function

Private Function Utils_Cols(ByRef v As Variant) As Long
    Select Case Utils_Numel(v)
        Case 0: Utils_Cols = 0
        Case 1: Utils_Cols = 1
        Case Else: Utils_Cols = UBound(v, 2)
    End Select
End Function

Private Sub Utils_Size(v As Variant, ByRef r As Long, ByRef c As Long)
    r = 0: c = 0
    Select Case Utils_Dimensions(v)
        Case 0: If Not WorksheetFunction.IsNA(v) Then r = 1: c = 1
        Case 1: r = UBound(v): c = 1
        Case 2: r = UBound(v, 1): c = UBound(v, 2)
        Case Else: Assert False, "Dimension > 2"
    End Select
End Sub

Private Sub Utils_Ind2Sub(rows As Long, k As Long, ByRef i As Long, ByRef j As Long)
    j = (k - 1) \ rows + 1
    i = k - rows * (j - 1)
End Sub

Private Sub Utils_Stack_Push(item As Variant, stack As Variant)
    On Error GoTo NotInitiated
    ReDim Preserve stack(LBound(stack) To UBound(stack) + 1)
    stack(UBound(stack)) = item
    Exit Sub
NotInitiated:
    stack = Array(item)
End Sub

Private Function Utils_Stack_Pop(stack As Variant) As Variant
    Dim ub As Long: ub = UBound(stack)
    Dim lb As Long: lb = LBound(stack)
    Utils_Stack_Pop = stack(ub)
    If ub > lb Then ReDim Preserve stack(lb To ub - 1) Else stack = Null
End Function

Private Function Utils_Stack_Peek(stack As Variant) As Variant
    Utils_Stack_Peek = stack(UBound(stack))
End Function

Private Function Utils_Stack_Size(stack As Variant) As Long
    On Error Resume Next
    Utils_Stack_Size = UBound(stack)
End Function

Private Function MAX(a As Variant, b As Variant) As Variant
    If a > b Then MAX = a Else MAX = b
End Function

Private Function MIN(a As Variant, b As Variant) As Variant
    If a < b Then MIN = a Else MIN = b
End Function

Private Sub Utils_CalcArgs(args As Variant)
    Dim i As Long
    For i = 1 To UBound(args)
        args(i) = eval_tree(args(i))
    Next i
End Sub

'do cols -> return 0, do rows -> return 1
Private Function Utils_CalcDimDirection(args As Variant, Optional dimIndex As Long = 2)
    If UBound(args) >= dimIndex Then
        Utils_CalcDimDirection = args(dimIndex) - 1
    ElseIf Utils_Rows(args(1)) = 1 Then
        Utils_CalcDimDirection = 1
    Else
        Utils_CalcDimDirection = 0
    End If
End Function

Private Sub Utils_AssertArgsCount(args As Variant, lb As Long, ub As Long)
    Assert _
        LBound(args) >= lb And UBound(args) <= ub, _
        "Number of arguments must be between " & lb & " and " & ub
End Sub

Private Sub Assert(expr As Boolean, Optional msg As String = "Unknown error")
    If expr Then Exit Sub
    errorMsg = msg
    err.Raise vbObjectError + 1
End Sub

'**********************
'*** EVAL FUNCTIONS ***
'**********************

Private Function eval_tree(root As Variant) As Variant
    ' Precalculate argument trees for all ordinary functions except if() and iferror()
    If left(root(1), 3) = "fn_" And root(1) <> "fn_if" And root(1) <> "fn_iferror" Then
        Utils_CalcArgs root(2)
    End If
    Select Case root(1)
        ' This is ugly, but much faster than just naively calling Application.Run()
        ' Just hardcode the most used functions
        Case "eval_constant": eval_tree = eval_constant(root(2))
        Case "eval_arg": eval_tree = eval_arg(root(2))
        Case "eval_index": eval_tree = eval_index(root(2))
        Case "eval_end": eval_tree = eval_end(root(2))
        Case "eval_colon": eval_tree = eval_colon(root(2))
        Case "eval_concat": eval_tree = eval_concat(root(2))
        Case "op_eq": eval_tree = op_eq(root(2))
        Case "op_plus": eval_tree = op_plus(root(2))
        Case "op_minus": eval_tree = op_minus(root(2))
        Case "op_mtimes": eval_tree = op_mtimes(root(2))
        Case "op_colon": eval_tree = op_colon(root(2))
        Case "fn_sum": eval_tree = fn_sum(root(2))
        Case "fn_repmat": eval_tree = fn_repmat(root(2))
        Case Else
            eval_tree = Application.Run(root(1), root(2))
    End Select
End Function

Private Function eval_constant(args As Variant) As Variant
    eval_constant = args(1)
End Function

Private Function eval_arg(args As Variant) As Variant
    If args(1) > UBound(arguments) Then
        Assert False, "Argument '" & Chr(Asc("a") + args(1)) & "' not found."
    End If
    eval_arg = CVar(arguments(args(1)))
    Utils_Conform eval_arg
End Function

Private Function eval_end(args As Variant) As Variant
    If Utils_Stack_Size(endValues) > 0 Then
        eval_end = Utils_Stack_Peek(endValues)
    Else
        err.Raise vbObjectError + 1
    End If
End Function

Private Function eval_colon(args As Variant) As Variant
    Assert False, "colon not allowed here..."
End Function

Private Function Utils_IsVector(r As Long, c As Long) As Boolean
    Utils_IsVector = (r = 1 And c > 1) Or (r > 1 And c = 1)
End Function

Private Function eval_indexarg(root As Variant, endValue As Long) As Variant
    Dim r As Variant
    If root(1) = "eval_colon" Then
        ReDim r(endValue, 1)
        Dim idx As Long: For idx = 1 To endValue
            r(idx, 1) = idx
        Next idx
    Else
        Utils_Stack_Push endValue, endValues
        r = eval_tree(root)
        Utils_Stack_Pop endValues
        If fn_islogical(Array(r)) Then
            r = fn_find(Array(r))
        End If
    End If
    eval_indexarg = r
End Function

' Evaluates matrix indexing/subsetting
Private Function eval_index(args As Variant) As Variant
    Dim r As Variant, i1 As Variant, i2 As Variant
    Dim matrows As Long, matcols As Long, r1 As Long, c1 As Long, r2 As Long, c2 As Long
    Dim idx As Long, r_i As Long, r_j As Long, arg_i As Long, arg_j As Long, i1_i As Long, i1_j As Long, i2_i As Long, i2_j As Long
    
    args(1) = eval_tree(args(1))
    Utils_ForceMatrix args(1)
    Utils_Size args(1), matrows, matcols
    
    On Error GoTo ErrorHandler
    Select Case UBound(args(2))
        
        Case 1:
            i1 = eval_indexarg(args(2)(1), matrows * matcols)
            Utils_Size i1, r1, c1
            If r1 = 0 Or c1 = 0 Then
                r = [NA()]
            Else
                Utils_ForceMatrix i1
                If Utils_IsVector(r1, c1) And Utils_IsVector(matrows, matcols) Then
                    If matrows > 1 Or args(2)(1)(1) = "eval_colon" Then
                        ReDim r(r1 * c1, 1)
                    Else
                        ReDim r(1, r1 * c1)
                    End If
                Else
                    ReDim r(r1, c1)
                End If
                For idx = 1 To UBound(r, 1) * UBound(r, 2)
                    Utils_Ind2Sub r1, idx, i1_i, i1_j
                    Utils_Ind2Sub matrows, CLng(i1(i1_i, i1_j)), arg_i, arg_j
                    Utils_Ind2Sub UBound(r, 1), idx, r_i, r_j
                    r(r_i, r_j) = args(1)(arg_i, arg_j)
                Next idx
            End If
            
        Case 2:
            i1 = eval_indexarg(args(2)(1), matrows)
            i2 = eval_indexarg(args(2)(2), matcols)
            Utils_Size i1, r1, c1
            Utils_Size i2, r2, c2
            If r1 = 0 Or c1 = 0 Or r2 = 0 Or c2 = 0 Then
                r = [NA()]
            Else
                Utils_ForceMatrix i1
                Utils_ForceMatrix i2
                ReDim r(r1 * c1, r2 * c2)
                For r_i = 1 To UBound(r, 1)
                    For r_j = 1 To UBound(r, 2)
                        Utils_Ind2Sub r1, r_i, i1_i, i1_j
                        Utils_Ind2Sub r2, r_j, i2_i, i2_j
                        r(r_i, r_j) = args(1)(i1(i1_i, i1_j), i2(i2_i, i2_j))
                    Next r_j
                Next r_i
            End If
            
        Case Else:
            Debug.Print "Too many index arguments..."
    
    End Select

    Utils_Conform r
    eval_index = r
    Exit Function
    
ErrorHandler:
    Assert False, err.Description
End Function

' Evaluates matrix concatenation []
Private Function eval_concat(args As Variant) As Variant

    ' Get matrices and check their sizes are compatible for concatenation
    Dim totalRows As Long, totalCols As Long
    Dim requiredRows As Long, requiredCols As Long
    Dim rows As Long, cols As Long, i As Long, j As Long
    For i = 1 To Utils_Stack_Size(args)
        totalCols = 0
        For j = 1 To Utils_Stack_Size(args(i))
            args(i)(j) = eval_tree(args(i)(j))
            Utils_Size args(i)(j), rows, cols
            If j = 1 Then
                requiredRows = rows
            Else
                Assert requiredRows = rows, "Concatenation: Different row counts"
            End If
            totalCols = totalCols + cols
        Next j
        If i = 1 Then
            requiredCols = totalCols
        Else
            Assert requiredCols = totalCols, "Concatenation: Different column counts"
        End If
        totalRows = totalRows + rows
    Next i
    
    ' Perform the actual concatenation by copying input matrices
    If totalRows = 0 Or totalCols = 0 Then
        eval_concat = [NA()]
        Exit Function
    End If
    Dim r: ReDim r(totalRows, totalCols) As Variant
    Dim x As Long, y As Long
    totalRows = 0
    For i = 1 To Utils_Stack_Size(args)
        totalCols = 0
        For j = 1 To Utils_Stack_Size(args(i))
            Utils_ForceMatrix args(i)(j)
            Utils_Size args(i)(j), rows, cols
            For x = 1 To rows
                For y = 1 To cols
                    r(totalRows + x, totalCols + y) = args(i)(j)(x, y)
                Next y
            Next x
            totalCols = totalCols + cols
        Next j
        totalRows = totalRows + rows
    Next i
    Utils_Conform r
    eval_concat = r

End Function

'*****************
'*** OPERATORS ***
'*****************

' Matches operator !
Private Function op_extern(args As Variant) As Variant
    args(1)(1) = Mid(args(1)(1), 4)
    Dim a As Variant: a = args(1)(2)
    Utils_CalcArgs a
    Select Case UBound(a)
        Case 0: op_extern = Application.Run(args(1)(1))
        Case 1: op_extern = Application.Run(args(1)(1), a(1))
        Case 2: op_extern = Application.Run(args(1)(1), a(1), a(2))
        Case 3: op_extern = Application.Run(args(1)(1), a(1), a(2), a(3))
        Case 4: op_extern = Application.Run(args(1)(1), a(1), a(2), a(3), a(4))
        Case 5: op_extern = Application.Run(args(1)(1), a(1), a(2), a(3), a(4), a(5))
        Case 6: op_extern = Application.Run(args(1)(1), a(1), a(2), a(3), a(4), a(5), a(6))
        Case 7: op_extern = Application.Run(args(1)(1), a(1), a(2), a(3), a(4), a(5), a(6), a(7))
        Case 8: op_extern = Application.Run(args(1)(1), a(1), a(2), a(3), a(4), a(5), a(6), a(7), a(8))
        Case 9: op_extern = Application.Run(args(1)(1), a(1), a(2), a(3), a(4), a(5), a(6), a(7), a(8), a(9))
        Case 10: op_extern = Application.Run(args(1)(1), a(1), a(2), a(3), a(4), a(5), a(6), a(7), a(8), a(9), a(10))
        Case Else: Assert False, "Cannot evaluate " & args(1)(1) & ": Too many arguments"
    End Select
    Utils_Conform op_extern
End Function

' Matches operator ||
Private Function op_orshortcircuit(args As Variant) As Variant
    On Error GoTo ErrorHandler
    If CBool(eval_tree(args(1))) Then
        op_orshortcircuit = True
    Else
        op_orshortcircuit = CBool(eval_tree(args(2)))
    End If
    Exit Function
ErrorHandler:
    Assert False, "Operator ||: Could not convert argument to boolean value"
End Function

' Matches operator &&
Private Function op_andshortcircuit(args As Variant) As Variant
    On Error GoTo ErrorHandler
    If Not CBool(eval_tree(args(1))) Then
        op_andshortcircuit = False
    Else
        op_andshortcircuit = CBool(eval_tree(args(2)))
    End If
    Exit Function
ErrorHandler:
    Assert False, "Operator &&: Could not convert argument to boolean value"
End Function

' Matches operator &
Private Function op_and(args As Variant) As Variant
    Utils_CalcArgs args
    Utils_ForceMatrix args(1): Utils_ForceMatrix args(2)
    Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
    Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
    Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2)
    Dim r: ReDim r(MAX(r1, r2), MAX(c1, c2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = CBool(args(1)(MIN(x, r1), MIN(y, c1))) And CBool(args(2)(MIN(x, r2), MIN(y, c2)))
        Next y
    Next x
    Utils_Conform r
    op_and = r
End Function

' Matches operator |
Private Function op_or(args As Variant) As Variant
    Utils_CalcArgs args
    Utils_ForceMatrix args(1): Utils_ForceMatrix args(2)
    Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
    Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
    Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2)
    Dim r: ReDim r(MAX(r1, r2), MAX(c1, c2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = CBool(args(1)(MIN(x, r1), MIN(y, c1))) Or CBool(args(2)(MIN(x, r2), MIN(y, c2)))
        Next y
    Next x
    Utils_Conform r
    op_or = r
End Function

' Matches operator <
Private Function op_lt(args As Variant) As Variant
    Utils_CalcArgs args
    Utils_ForceMatrix args(1): Utils_ForceMatrix args(2)
    Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
    Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
    Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2)
    Dim r: ReDim r(MAX(r1, r2), MAX(c1, c2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = args(1)(MIN(x, r1), MIN(y, c1)) < args(2)(MIN(x, r2), MIN(y, c2))
        Next y
    Next x
    Utils_Conform r
    op_lt = r
End Function

' Matches operator <=
Private Function op_lte(args As Variant) As Variant
    Utils_CalcArgs args
    Utils_ForceMatrix args(1): Utils_ForceMatrix args(2)
    Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
    Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
    Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2)
    Dim r: ReDim r(MAX(r1, r2), MAX(c1, c2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = args(1)(MIN(x, r1), MIN(y, c1)) <= args(2)(MIN(x, r2), MIN(y, c2))
        Next y
    Next x
    Utils_Conform r
    op_lte = r
End Function

' Matches operator >
Private Function op_gt(args As Variant) As Variant
    Utils_CalcArgs args
    Utils_ForceMatrix args(1): Utils_ForceMatrix args(2)
    Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
    Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
    Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2)
    Dim r: ReDim r(MAX(r1, r2), MAX(c1, c2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = args(1)(MIN(x, r1), MIN(y, c1)) > args(2)(MIN(x, r2), MIN(y, c2))
        Next y
    Next x
    Utils_Conform r
    op_gt = r
End Function

' Matches operator >=
Private Function op_gte(args As Variant) As Variant
    Utils_CalcArgs args
    Utils_ForceMatrix args(1): Utils_ForceMatrix args(2)
    Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
    Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
    Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2)
    Dim r: ReDim r(MAX(r1, r2), MAX(c1, c2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = args(1)(MIN(x, r1), MIN(y, c1)) >= args(2)(MIN(x, r2), MIN(y, c2))
        Next y
    Next x
    Utils_Conform r
    op_gte = r
End Function

' Matches operator ==
Private Function op_eq(args As Variant) As Variant
    Utils_CalcArgs args
    Utils_ForceMatrix args(1): Utils_ForceMatrix args(2)
    Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
    Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
    Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2)
    Dim r: ReDim r(MAX(r1, r2), MAX(c1, c2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = args(1)(MIN(x, r1), MIN(y, c1)) = args(2)(MIN(x, r2), MIN(y, c2))
        Next y
    Next x
    Utils_Conform r
    op_eq = r
End Function

' Matches operator ~=
Private Function op_ne(args As Variant) As Variant
    Utils_CalcArgs args
    Utils_ForceMatrix args(1): Utils_ForceMatrix args(2)
    Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
    Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
    Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2)
    Dim r: ReDim r(MAX(r1, r2), MAX(c1, c2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = args(1)(MIN(x, r1), MIN(y, c1)) <> args(2)(MIN(x, r2), MIN(y, c2))
        Next y
    Next x
    Utils_Conform r
    op_ne = r
End Function

' Matches operator ~
Private Function op_negate(args As Variant) As Variant
    Utils_CalcArgs args
    If Utils_Dimensions(args(1)) = 0 Then
        op_negate = Not CBool(args(1))
    Else
        Dim i As Long, j As Long
        Dim r: ReDim r(UBound(args(1), 1), UBound(args(1), 2))
        For i = 1 To UBound(r, 1)
            For j = 1 To UBound(r, 2)
                r(i, j) = Not CBool(args(1)(i, j))
            Next j
        Next i
        op_negate = r
    End If
End Function

' Matches operator : with one or two arguments
Private Function op_colon(args As Variant) As Variant
    Dim m As Long, i As Long, step As Double, start As Double
    start = eval_tree(args(1))
    If args(2)(1) <> "op_colon" Then
        ' x:y
        step = 1
        m = fn_fix(Array(eval_tree(args(2)) - start))
    Else
        ' x:y:z
        step = eval_tree(args(2)(2)(1))
        m = fn_fix(Array((eval_tree(args(2)(2)(2)) - start) / step))
    End If
    If m < 0 Then
        op_colon = [NA()]
        Exit Function
    End If
    Dim r: ReDim r(1, 1 + m) As Variant
    For i = 0 To m
        r(1, 1 + i) = start + i * step
    Next i
    Utils_Conform r
    op_colon = r
End Function

' Matches operator +
Private Function op_plus(args As Variant) As Variant
    Utils_CalcArgs args
    Utils_ForceMatrix args(1): Utils_ForceMatrix args(2)
    Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
    Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
    Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2)
    Dim a1 As Variant, a2 As Variant
    Dim r: ReDim r(MAX(r1, r2), MAX(c1, c2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = args(1)(MIN(x, r1), MIN(y, c1)) + args(2)(MIN(x, r2), MIN(y, c2))
        Next y
    Next x
    Utils_Conform r
    op_plus = r
End Function

' Matches unary operator +
Private Function op_uplus(args As Variant) As Variant
    op_uplus = eval_tree(args(1))
End Function

' Matches binary operator -
Private Function op_minus(args As Variant) As Variant
    Utils_CalcArgs args
    Utils_ForceMatrix args(1): Utils_ForceMatrix args(2)
    Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
    Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
    Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2)
    Dim r: ReDim r(MAX(r1, r2), MAX(c1, c2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = args(1)(MIN(x, r1), MIN(y, c1)) - args(2)(MIN(x, r2), MIN(y, c2))
        Next y
    Next x
    Utils_Conform r
    op_minus = r
End Function

' Matches prefix unary operator -
Private Function op_uminus(args As Variant) As Variant
    Dim rows As Long, cols As Long
    Utils_CalcArgs args
    Utils_Size args(1), rows, cols
    If rows <= 1 And cols <= 1 Then
        op_uminus = -args(1)
    Else
        Dim i As Long, j As Long
        ReDim r(rows, cols)
        For i = 1 To rows
            For j = 1 To cols
                r(i, j) = -args(1)(i, j)
            Next j
        Next i
        op_uminus = r
    End If
End Function

' Matches operator *
Private Function op_mtimes(args As Variant) As Variant
    Utils_CalcArgs args
    If Utils_Dimensions(args(1)) = 2 And Utils_Dimensions(args(2)) = 2 Then
        Assert UBound(args(1), 2) = UBound(args(2), 1), "mtimes(): Matrix sizes not compatible"
        op_mtimes = WorksheetFunction.MMult(args(1), args(2))
    Else
        Utils_ForceMatrix args(1): Utils_ForceMatrix args(2)
        Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
        Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
        Dim r: ReDim r(MAX(r1, r2), MAX(c1, c2))
        Dim x As Long, y As Long
        For x = 1 To UBound(r, 1)
            For y = 1 To UBound(r, 2)
                r(x, y) = args(1)(MIN(x, r1), MIN(y, c1)) * args(2)(MIN(x, r2), MIN(y, c2))
            Next y
        Next x
        op_mtimes = r
    End If
    Utils_Conform op_mtimes
End Function

' Matches operator .*
Private Function op_times(args As Variant) As Variant
    Utils_CalcArgs args
    Utils_ForceMatrix args(1): Utils_ForceMatrix args(2)
    Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
    Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
    Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2)
    Dim r: ReDim r(MAX(r1, r2), MAX(c1, c2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = args(1)(MIN(x, r1), MIN(y, c1)) * args(2)(MIN(x, r2), MIN(y, c2))
        Next y
    Next x
    Utils_Conform r
    op_times = r
End Function

' Matches operator /
Private Function op_mdivide(args As Variant) As Variant
    Utils_CalcArgs args
    Utils_ForceMatrix args(1): Utils_ForceMatrix args(2)
    Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
    Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
    Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1)
    Dim r: ReDim r(MAX(r1, r2), MAX(c1, c2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = args(1)(MIN(x, r1), MIN(y, c1)) / args(2)(MIN(x, r2), MIN(y, c2))
        Next y
    Next x
    Utils_Conform r
    op_mdivide = r
End Function

' Matches operator ./
Private Function op_divide(args As Variant) As Variant
    Utils_CalcArgs args
    Utils_ForceMatrix args(1): Utils_ForceMatrix args(2)
    Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
    Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
    Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2)
    Dim r: ReDim r(MAX(r1, r2), MAX(c1, c2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = args(1)(MIN(x, r1), MIN(y, c1)) / args(2)(MIN(x, r2), MIN(y, c2))
        Next y
    Next x
    Utils_Conform r
    op_divide = r
End Function

' Matches operator ^
Private Function op_mpower(args As Variant) As Variant
    Utils_CalcArgs args
    Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
    Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
    If r1 = 1 And c1 = 1 And r2 = 1 And c2 = 1 Then
        op_mpower = args(1) ^ args(2)
    Else
        Assert False, "mpower: Input must be scalars"
    End If
End Function

' Matches operator .^
Private Function op_power(args As Variant) As Variant
    Utils_CalcArgs args
    Utils_ForceMatrix args(1): Utils_ForceMatrix args(2)
    Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
    Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
    Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2)
    Dim r: ReDim r(MAX(r1, r2), MAX(c1, c2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = args(1)(MIN(x, r1), MIN(y, c1)) ^ args(2)(MIN(x, r2), MIN(y, c2))
        Next y
    Next x
    Utils_Conform r
    op_power = r
End Function

' Matches postfix unary operator '
Private Function op_transpose(args As Variant) As Variant
    op_transpose = Application.WorksheetFunction.Transpose(eval_tree(args(1)))
    Utils_Conform op_transpose
End Function

' Matches operator #
Private Function op_numel(args As Variant) As Variant
    op_numel = Utils_Numel(eval_tree(args(1)))
End Function

'*****************
'*** FUNCTIONS ***
'*****************

' b = islogical(A)
'
' b = islogical(A) returns true if all elements of A are boolean values.
Private Function fn_islogical(args As Variant) As Variant
    Utils_ForceMatrix args(1)
    Dim i As Long, j As Long
    For i = 1 To UBound(args(1), 1)
        For j = 1 To UBound(args(1), 2)
            If Not WorksheetFunction.IsLogical(args(1)(i, j)) Then
                fn_islogical = False
                Exit Function
            End If
        Next j
    Next i
    fn_islogical = True
End Function

' I = find(X)
'
' I = find(A) locates all nonzero elements of array X, and returns the linear indices
' of those elements in vector I. If X is a row vector, then I is a row vector;
' otherwise, I is a column vector. If X contains no nonzero elements or is an empty array,
' then I is an empty array.
Private Function fn_find(args As Variant) As Variant
    Dim rows As Long, cols As Long
    Utils_Size args(1), rows, cols
    If rows <= 1 And cols <= 1 Then
        If CDbl(args(1)) = 0 Then fn_find = [NA()] Else fn_find = 1
    Else
        Dim counter As Long, i As Long, j As Long
        For i = 1 To rows
            For j = 1 To cols
                If CDbl(args(1)(i, j)) <> 0 Then counter = counter + 1
            Next j
        Next i
        If counter = 0 Then fn_find = [NA()]: Exit Function
        Dim isRowVec As Long: isRowVec = -(rows = 1)
        Dim r: ReDim r(isRowVec + (1 - isRowVec) * counter, 1 - isRowVec + isRowVec * counter)
        counter = 0
        For j = 1 To cols
            For i = 1 To rows
                If CDbl(args(1)(i, j)) <> 0 Then
                    counter = counter + 1
                    r(isRowVec + (1 - isRowVec) * counter, 1 - isRowVec + isRowVec * counter) _
                        = (j - 1) * rows + i
                End If
            Next i
        Next j
        Utils_Conform r
        fn_find = r
    End If
End Function

' X = fix(A)
'
' X = fix(A) rounds the elements of A towards 0.
Private Function fn_fix(args As Variant) As Variant
    If Utils_Dimensions(args(1)) = 0 Then
        fn_fix = WorksheetFunction.RoundDown(args(1), 0)
    Else
        Dim x As Long, y As Long
        Dim r: ReDim r(UBound(args(1), 1), UBound(args(1), 2)) As Variant
        For x = 1 To UBound(r, 1)
            For y = 1 To UBound(r, 2)
                r(x, y) = WorksheetFunction.RoundDown(args(1)(x, y), 0)
            Next y
        Next x
        fn_fix = r
    End If
End Function

' X = round(A)
'
' X = round(A) rounds the elements of A towards the nearest integer.
Private Function fn_round(args As Variant) As Variant
    If Utils_Dimensions(args(1)) = 0 Then
        fn_round = WorksheetFunction.Round(args(1), 0)
    Else
        Dim x As Long, y As Long
        Dim r: ReDim r(UBound(args(1), 1), UBound(args(1), 2)) As Variant
        For x = 1 To UBound(r, 1)
            For y = 1 To UBound(r, 2)
                r(x, y) = WorksheetFunction.Round(args(1)(x, y), 0)
            Next y
        Next x
        fn_round = r
    End If
End Function

Private Function fn_inv(args As Variant) As Variant
    If Utils_Dimensions(args(1)) = 0 Then
        fn_inv = 1# / args(1)
    Else
        Assert UBound(args(1), 1) = UBound(args(1), 2), "inv: matrix not quadratic."
        fn_inv = WorksheetFunction.MInverse(args(1))
    End If
End Function

' X = exp(A)
'
' X = exp(A) returns the exponential for each element of A.
Private Function fn_exp(args As Variant) As Variant
    If Utils_Dimensions(args(1)) = 0 Then
        fn_exp = Exp(args(1))
    Else
        Dim i As Long, j As Long
        Dim r: ReDim r(UBound(args(1), 1), UBound(args(1), 2))
        For i = 1 To UBound(r, 1)
            For j = 1 To UBound(r, 2)
                r(i, j) = Exp(args(1)(i, j))
            Next j
        Next i
        fn_exp = r
    End If
End Function

' X = log(A)
'
' X = log(A) returns the natural logarithm of the elements of A.
Private Function fn_log(args As Variant) As Variant
    If Utils_Dimensions(args(1)) = 0 Then
        fn_log = Log(args(1))
    Else
        Dim i As Long, j As Long
        Dim r: ReDim r(UBound(args(1), 1), UBound(args(1), 2))
        For i = 1 To UBound(r, 1)
            For j = 1 To UBound(r, 2)
                r(i, j) = Log(args(1)(i, j))
            Next j
        Next i
        fn_log = r
    End If
End Function

' X = sqrt(A)
'
' X = sqrt(A) returns the square root of the elements of A.
Private Function fn_sqrt(args As Variant) As Variant
    If Utils_Dimensions(args(1)) = 0 Then
        fn_sqrt = Sqr(args(1))
    Else
        Dim i As Long, j As Long
        Dim r: ReDim r(UBound(args(1), 1), UBound(args(1), 2))
        For i = 1 To UBound(r, 1)
            For j = 1 To UBound(r, 2)
                r(i, j) = Sqr(args(1)(i, j))
            Next j
        Next i
        fn_sqrt = r
    End If
End Function

' r = rows(A)
'
' r = rows(A) returns the number of rows in A.
Private Function fn_rows(args As Variant) As Variant
    fn_rows = Utils_Rows(args(1))
End Function

' c = cols(A)
'
' c = cols(A) returns the number of columns in A.
Private Function fn_cols(args As Variant) As Variant
    fn_cols = Utils_Cols(args(1))
End Function

' n = numel(A)
'
' n = numel(A) returns the number of elements in A.
Private Function fn_numel(args As Variant) As Variant
    fn_numel = Utils_Numel(args(1))
End Function

' X = zeros(n)
' X = zeros(n,m)
'
' X = zeros(n) returns an n-by-n matrix of zeros.
'
' X = zeros(n,m) returns an n-by-m matrix of zeros.
Private Function fn_zeros(args As Variant) As Variant
    fn_zeros = fn_repmat(Array(0, CLng(args(1)), CLng(args(UBound(args)))))
    Utils_Conform fn_zeros
End Function

' X = ones(n)
' X = ones(n,m)
'
' X = ones(n) returns an n-by-n matrix of ones.
'
' X = ones(n,m) returns an n-by-m matrix of ones.
Private Function fn_ones(args As Variant) As Variant
    fn_ones = fn_repmat(Array(1, CLng(args(1)), CLng(args(UBound(args)))))
    Utils_Conform fn_ones
End Function

' X = ones(n)
'
' X = ones(n) returns the n-by-n identity matrix.
Private Function fn_eye(args As Variant) As Variant
    Dim i As Long, r As Variant
    r = fn_repmat(Array(0, CLng(args(1)), CLng(args(1))))
    For i = 1 To UBound(r, 1)
        r(i, i) = 1
    Next i
    Utils_Conform r
    fn_eye = r
End Function

' X = tick2ret(A)
' X = tick2ret(A,method)
' X = tick2ret(A,method,dim)
Private Function fn_tick2ret(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 3
    Dim i As Long, j As Long, x As Long, r As Variant
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args, 3)
    Dim simple As Boolean: simple = True
    If UBound(args) > 1 Then
        If LCase(args(2)) = "continuous" Then
            simple = False
        End If
    End If
    ReDim r(UBound(args(1), 1) - (1 - x), UBound(args(1), 2) - x)
    For i = 1 To UBound(r, 1)
        For j = 1 To UBound(r, 2)
            If simple Then
                r(i, j) = args(1)(i + (1 - x), j + x) / args(1)(i, j) - 1
            Else
                r(i, j) = Log(args(1)(i + (1 - x), j + x) / args(1)(i, j))
            End If
        Next j
    Next i
    Utils_Conform r
    fn_tick2ret = r
End Function

' B = cumsum(A)
' B = cumsum(A,dim)
'
' B = cumsum(A) returns the cumulative sum along different dimensions of
' an array. If A is a vector, cumsum(A) returns a vector containing the
' cumulative sum of the elements of A. If A is a matrix, cumsum(A)
' returns a matrix the same size as A containing the cumulative sums for
' each column of A.
'
' B = cumsum(A,dim) returns the cumulative sum of the elements along the
' dimension of A specified by scalar dim. For example, cumsum(A,1) works
' along the first dimension (the columns); cumsum(A,2) works along the '
' second dimension (the rows).
Private Function fn_cumsum(args As Variant) As Variant
    Dim i As Long, j As Long, x As Long
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args)
    For i = 2 - x To UBound(args(1), 1)
        For j = 1 + x To UBound(args(1), 2)
            args(1)(i, j) = args(1)(i, j) + args(1)(i - (1 - x), j - x)
        Next j
    Next i
    Utils_Conform args(1)
    fn_cumsum = args(1)
End Function

' B = cumprod(A)
' B = cumprod(A,dim)
'
' B = cumprod(A) returns the cumulative product along different dimensions of
' an array. If A is a vector, cumsum(A) returns a vector containing the
' cumulative product of the elements of A. If A is a matrix, cumprod(A)
' returns a matrix the same size as A containing the cumulative products for
' each column of A.
'
' B = cumprod(A,dim) returns the cumulative product of the elements along the
' dimension of A specified by scalar dim. For example, cumprod(A,1) works
' along the first dimension (the columns); cumprod(A,2) works along the '
' second dimension (the rows).
Private Function fn_cumprod(args As Variant) As Variant
    Dim i As Long, j As Long, x As Long
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args)
    For i = 2 - x To UBound(args(1), 1)
        For j = 1 + x To UBound(args(1), 2)
            args(1)(i, j) = args(1)(i, j) * args(1)(i - (1 - x), j - x)
        Next j
    Next i
    Utils_Conform args(1)
    fn_cumprod = args(1)
End Function

' X = std(A)
' X = std(A,dim)
'
' X = std(A) returns the standard deviation of the elements of A
' along the first array dimension whose size does not equal 1.
'
' X = std(A,dim) sums the elements of A along dimension dim.
' The dim input is a positive integer scalar.
Private Function fn_std(args As Variant) As Variant
    Dim i As Long, x As Long, r As Variant
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args)
    ReDim r(x * Utils_Rows(args(1)) + (1 - x), (1 - x) * Utils_Cols(args(1)) + x)
    For i = 1 To UBound(r, -x + 2)
        r(x * i + (1 - x), (1 - x) * i + x) _
            = WorksheetFunction.StDev(WorksheetFunction.index(args(1), x * i, (1 - x) * i))
    Next i
    Utils_Conform r
    fn_std = r
End Function

' X = sum(A)
' X = sum(A,dim)
'
' X = sum(A) returns the sum of the elements of A along the
' first array dimension whose size does not equal 1.
'
' X = sum(A,dim) sums the elements of A along dimension dim.
' The dim input is a positive integer scalar.
Private Function fn_sum(args As Variant) As Variant
    Dim x As Long, i As Long, r As Variant
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args)
    ReDim r(x * Utils_Rows(args(1)) + (1 - x), (1 - x) * Utils_Cols(args(1)) + x)
    For i = 1 To UBound(r, -x + 2)
        r(x * i + (1 - x), (1 - x) * i + x) _
            = WorksheetFunction.sum(WorksheetFunction.index(args(1), x * i, (1 - x) * i))
    Next i
    Utils_Conform r
    fn_sum = r
End Function

' X = prod(A)
' X = prod(A,dim)
'
' X = prod(A) returns the product of the elements of A
' along the first array dimension whose size does not equal 1.
'
' X = prod(A,dim) multiplies the elements of A along dimension dim.
' The dim input is a positive integer scalar.
Private Function fn_prod(args As Variant) As Variant
    Dim i As Long, x As Long, r As Variant
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args)
    ReDim r(x * Utils_Rows(args(1)) + (1 - x), (1 - x) * Utils_Cols(args(1)) + x)
    For i = 1 To UBound(r, -x + 2)
        r(x * i + (1 - x), (1 - x) * i + x) _
            = WorksheetFunction.Product(WorksheetFunction.index(args(1), x * i, (1 - x) * i))
    Next i
    Utils_Conform r
    fn_prod = r
End Function

' X = mean(A)
' X = mean(A,dim)
Private Function fn_mean(args As Variant) As Variant
    Dim i As Long, x As Long, r As Variant
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args)
    ReDim r(x * Utils_Rows(args(1)) + (1 - x), (1 - x) * Utils_Cols(args(1)) + x)
    For i = 1 To UBound(r, -x + 2)
        r(x * i + (1 - x), (1 - x) * i + x) _
            = WorksheetFunction.Average(WorksheetFunction.index(args(1), x * i, (1 - x) * i))
    Next i
    Utils_Conform r
    fn_mean = r
End Function

' X = median(A)
' X = median(A,dim)
Private Function fn_median(args As Variant) As Variant
    Dim i As Long, x As Long, r As Variant
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args)
    ReDim r(x * Utils_Rows(args(1)) + (1 - x), (1 - x) * Utils_Cols(args(1)) + x)
    For i = 1 To UBound(r, -x + 2)
        r(x * i + (1 - x), (1 - x) * i + x) _
            = WorksheetFunction.Median(WorksheetFunction.index(args(1), x * i, (1 - x) * i))
    Next i
    Utils_Conform r
    fn_median = r
End Function

' X = prctile(A)
' X = prctile(A,p)
' X = prctile(A,p,dim)
Private Function fn_prctile(args As Variant) As Variant
    Dim i As Long, x As Long, r As Variant
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args, 3)
    ReDim r(x * Utils_Rows(args(1)) + (1 - x), (1 - x) * Utils_Cols(args(1)) + x)
    For i = 1 To UBound(r, -x + 2)
        r(x * i + (1 - x), (1 - x) * i + x) _
            = WorksheetFunction.Percentile(WorksheetFunction.index(args(1), x * i, (1 - x) * i), args(2))
    Next i
    Utils_Conform r
    fn_prctile = r
End Function

' b = isequal(A,B)
Private Function fn_isequal(args As Variant) As Variant
    fn_isequal = False
    Dim dim1 As Long: dim1 = Utils_Dimensions(args(1))
    Dim dim2 As Long: dim2 = Utils_Dimensions(args(2))
    If dim1 = 0 And dim2 = 0 Then
        fn_isequal = (args(1) = args(2))
    ElseIf dim1 = 2 And dim2 = 2 Then
        Dim size1 As Variant: size1 = fn_size(Array(args(1)))
        Dim size2 As Variant: size2 = fn_size(Array(args(2)))
        If size1(1, 1) <> size2(1, 1) Or size1(1, 2) <> size2(1, 2) Then Exit Function
        Dim i As Long, j As Long
        For i = 1 To size1(1, 1)
            For j = 1 To size1(1, 2)
                If args(1)(i, j) <> args(2)(i, j) Then Exit Function
            Next j
        Next i
        fn_isequal = True
    End If
End Function

' X = isempty(A)
Private Function fn_isempty(args As Variant) As Variant
    fn_isempty = False
    If Utils_Dimensions(args(1)) = 0 Then
        fn_isempty = WorksheetFunction.IsNA(args(1))
    End If
End Function

' X = size(A)
'
' X = size(A) returns a 1-by-2 vector with the number of rows and columns in A.
Private Function fn_size(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim r: ReDim r(1, 2)
    r(1, 1) = Utils_Rows(args(1))
    r(1, 2) = Utils_Cols(args(1))
    fn_size = r
End Function

' X = rand(n)
' X = rand(m,n)
'
' X = rand(n) returns an n-by-n matrix containing pseudorandom values
' drawn from the standard uniform distribution on the open interval (0,1).
'
' X = rand(m,n) returns an m-by-n matrix containing pseudorandom values
' drawn from the standard uniform distribution on the open interval (0,1).
Private Function fn_rand(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    Dim r: ReDim r(args(1), args(UBound(args))) As Variant
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = Rnd
        Next y
    Next
    Utils_Conform r
    fn_rand = r
End Function

' X = randn(n)
' X = randn(m,n)
'
' X = randn(n) returns an n-by-n matrix containing pseudorandom values
' drawn from the standard normal distribution with mean 0 and variance 1.
'
' X = randn(m,n) returns an m-by-n matrix containing pseudorandom values
' drawn from the standard normal distribution with mean 0 and variance 1.
Private Function fn_randn(args As Variant) As Variant
    Dim r: ReDim r(args(1), args(UBound(args))) As Variant
    Dim x As Long, y As Long
    Dim c As Long: c = 3
    Dim n(2) As Double, tmp As Double
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            If c > 2 Then
                Do
                    n(1) = 2 * Rnd - 1
                    n(2) = 2 * Rnd - 1
                    tmp = n(1) * n(1) + n(2) * n(2)
                Loop Until tmp <= 1
                tmp = Sqr(-2 * Log(tmp) / tmp)
                n(1) = n(1) * tmp
                n(2) = n(2) * tmp
                c = 1
            End If
            r(x, y) = n(c)
            c = c + 1
        Next y
    Next x
    Utils_Conform r
    fn_randn = r
End Function

' X = repmat(A,n)
' X = repmat(A,m,n)
' X = repmat(A,[m n])
'
' X = repmat(A,n) creates a large matrix X consisting of an n-by-n tiling of A.
' X = repmat(A,m,n) creates a large matrix X consisting of an m-by-n tiling of A.
' X = repmat(A,[m n]) creates a large matrix X consisting of an m-by-n tiling of A.
Private Function fn_repmat(args As Variant) As Variant
    Dim r As Variant, matrows As Long, matcols As Long, m As Long, n As Long, i As Long, j As Long
    Select Case UBound(args)
        Case 2
            Select Case Utils_Numel(args(2))
                Case 1
                    m = args(2)
                    n = args(2)
                Case 2
                    m = args(2)(1, 1)
                    n = args(2)(MAX(2, UBound(args(2), 1)), MAX(2, UBound(args(2), 2)))
                Case Else
                    Assert False, "repmat: Wrong argument"
            End Select
        Case 3
            m = args(2)
            n = args(3)
        Case Else
            Assert False, "repmat: Wrong number of input arguments."
    End Select
    Utils_ForceMatrix args(1)
    Utils_Size args(1), matrows, matcols
    ReDim r(matrows * m, matcols * n)
    For m = 0 To m - 1
        For n = 0 To n - 1
            For i = 1 To matrows
                For j = 1 To matcols
                    r(m * matrows + i, n * matcols + j) = args(1)(i, j)
                Next j
            Next i
        Next n
    Next m
    Utils_Conform r
    fn_repmat = r
End Function

' B = reshape(A,m,n)
' B = reshape(A,[],n)
' B = reshape(A,m,[])
'
' B = reshape(A,m,n) returns the m-by-n matrix B whose elements are taken
' column-wise from A. An error results if A does not have m*n elements.
' Either m or n can be the empty matrix [] in which case the length of the
' dimension is calculated automatically.
Private Function fn_reshape(args As Variant) As Variant
    Dim r As Variant, rows As Long, cols As Long, idx As Long
    Dim r_i As Long, r_j As Long, arg_i As Long, arg_j As Long
    Utils_ForceMatrix args(1)
    Utils_Size args(1), rows, cols
    If WorksheetFunction.IsNA(args(2)) Then args(2) = rows * cols / args(3)
    If WorksheetFunction.IsNA(args(3)) Then args(3) = rows * cols / args(2)
    ReDim r(args(2), args(3))
    For idx = 1 To rows * cols
        Utils_Ind2Sub rows, idx, arg_i, arg_j
        Utils_Ind2Sub CLng(args(2)), idx, r_i, r_j
        r(r_i, r_j) = args(1)(arg_i, arg_j)
    Next idx
    Utils_Conform r
    fn_reshape = r
End Function

' X = tostring(A)
'
Private Function fn_tostring(args As Variant) As Variant
    Utils_ForceMatrix args(1)
    Dim i As Long, j As Long
    ReDim r(UBound(args(1), 1), UBound(args(1), 2))
    For i = 1 To UBound(r, 1)
        For j = 1 To UBound(r, 2)
            r(i, j) = args(1)(i, j) & ""
        Next j
    Next i
    Utils_Conform r
    fn_tostring = r
End Function

' X = if(a,B,C)
'
' X = if(a,B,C) returns B if a evaluates to true; otherwise C.
' if() functions as an operator implementing short circuiting.
' I.e. if C is an expression it is not evaluated unless a is true and vice versa.
Private Function fn_if(args As Variant) As Variant
    If CBool(eval_tree(args(1))) Then
        fn_if = eval_tree(args(2))
    Else
        fn_if = eval_tree(args(3))
    End If
End Function

' X = iferror(A,B)
'
' X = iferror(A,B) returns A if the evaluation of A does not result in a error; then B is returned instead.
Private Function fn_iferror(args As Variant) As Variant
    On Error GoTo ErrorHandler:
    fn_iferror = eval_tree(args(1))
    Exit Function
ErrorHandler:
    errorMsg = ""
    fn_iferror = eval_tree(args(2))
End Function

' X = count(A)
' X = count(A,dim)
'
' X = count(A) counts the number of elements in A which do not evaluate to false
Private Function fn_count(args As Variant) As Variant
    If Utils_IsNA(args(1)) Then fn_count = 0: Exit Function
    Dim x As Long, i As Long, j As Long, r As Variant
    Dim rows As Long, cols As Long
    Utils_ForceMatrix args(1)
    Utils_Size args(1), rows, cols
    x = Utils_CalcDimDirection(args)
    ReDim r(x * rows + (1 - x), (1 - x) * cols + x)
    For i = 1 To rows
        For j = 1 To cols
            r(i * x + (1 - x), j * (1 - x) + x) _
                = r(i * x + (1 - x), j * (1 - x) + x) - CBool(args(1)(i, j))
        Next j
    Next i
    Utils_Conform r
    fn_count = r
End Function

' X = diff(A)
' X = diff(A,dim)
' X = diff(A,dim,n)
Private Function fn_diff(args As Variant) As Variant
    If Utils_IsNA(args(1)) Then fn_diff = [NA()]: Exit Function
    Utils_ForceMatrix args(1)
    Dim x As Long: x = Utils_CalcDimDirection(args)
    If UBound(args(1), 1 + x) < 2 Then fn_diff = [NA()]: Exit Function
    Dim i As Long, j As Long, r As Variant
    ReDim r(UBound(args(1), 1) - (1 - x), UBound(args(1), 2) - x)
    For i = 2 - x To UBound(args(1), 1)
        For j = 1 + x To UBound(args(1), 2)
            r(i - (1 - x), j - x) = args(1)(i, j) - args(1)(i - (1 - x), j - x)
        Next j
    Next i
    Utils_Conform r
    fn_diff = r
    If UBound(args) > 2 Then
        If args(3) > 1 Then
            fn_diff = fn_diff(Array(r, 1 + x, args(3) - 1))
        End If
    End If
End Function

' B = sort(A)
' B = sort(A,dim)
' B = sort(A,dim,mode)
Private Function fn_sort(args As Variant) As Variant
    Dim x As Long: x = Utils_CalcDimDirection(args)
    Dim ascend As Boolean: ascend = True
    If UBound(args) > 2 Then
        Assert args(3) = "ascend" Or args(3) = "descend", _
            "sort(): Parameter mode must be either ""ascend"" or ""descend""."
        ascend = args(3) = "ascend"
    End If
    Utils_ForceMatrix args(1)
    Dim i As Long
    For i = 1 To UBound(args(1), 2 - x)
        If x = 0 Then
            Utils_QuickSortCol args(1), 1, UBound(args(1), 1), i, ascend
        Else
            Utils_QuickSortRow args(1), 1, UBound(args(1), 2), i, ascend
        End If
    Next i
    fn_sort = args(1)
End Function

' Implementation of quick-sort - is a helper for fn_sort()
' Sorts on columns
Private Function Utils_QuickSortCol(arr As Variant, first As Long, last As Long, col As Long, ascend As Boolean)
    If first >= last Then Exit Function
    Dim tmp As Variant
    Dim pivot As Variant: pivot = arr(first, col)
    Dim left As Long: left = first
    Dim right As Long: right = last
    While left <= right
        While Utils_QuickSortCompare(arr(left, col), pivot, ascend) > 0
            left = left + 1
        Wend
        While Utils_QuickSortCompare(arr(right, col), pivot, ascend) < 0
            right = right - 1
        Wend
        If left <= right Then
            tmp = arr(left, col)
            arr(left, col) = arr(right, col)
            arr(right, col) = tmp
            left = left + 1
            right = right - 1
        End If
    Wend
    Utils_QuickSortCol arr, first, right, col, ascend
    Utils_QuickSortCol arr, left, last, col, ascend
End Function

' Implementation of quick-sort - is a helper for fn_sort()
' Sorts on rows
Private Function Utils_QuickSortRow(arr As Variant, first As Long, last As Long, row As Long, ascend As Boolean)
    If first >= last Then Exit Function
    Dim tmp As Variant
    Dim pivot As Variant: pivot = arr(row, first)
    Dim left As Long: left = first
    Dim right As Long: right = last
    While left <= right
        While Utils_QuickSortCompare(arr(row, left), pivot, ascend) > 0
            left = left + 1
        Wend
        While Utils_QuickSortCompare(arr(row, right), pivot, ascend) < 0
            right = right - 1
        Wend
        If left <= right Then
            tmp = arr(row, left)
            arr(row, left) = arr(row, right)
            arr(row, right) = tmp
            left = left + 1
            right = right - 1
        End If
    Wend
    Utils_QuickSortRow arr, first, right, row, ascend
    Utils_QuickSortRow arr, left, last, row, ascend
End Function

' Called from Utils_QuickSortRow and Utils_QuickSortCol. Compares numerics and strings.
Public Function Utils_QuickSortCompare(arg1 As Variant, arg2 As Variant, ascend As Boolean) As Long
    If IsNumeric(arg1) Then
        If IsNumeric(arg2) Then
            Utils_QuickSortCompare = (1 + 2 * CLng(ascend)) * (arg1 - arg2)
        Else
            Utils_QuickSortCompare = -1 - 2 * CLng(ascend)
        End If
    Else
        If IsNumeric(arg2) Then
            Utils_QuickSortCompare = 1 + 2 * CLng(ascend)
        Else
            Utils_QuickSortCompare = (1 + 2 * CLng(ascend)) * StrComp(CStr(arg1), CStr(arg2))
        End If
    End If
End Function
