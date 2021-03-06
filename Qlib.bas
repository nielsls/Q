'*******************************************************************
' Q - A MATLAB-like matrix parser for Microsoft Excel
'
' Q features a single public function, Q(), containing an expression parser.
' Q() is able to parse and evaluate a subset of the MATLAB programming language.
' It features almost all MATLAB operators, selected standard functions
' and has complete support for submatrices, '()', and concatenation, '[]'.
'
' Example usage:
'   - =Q("2+2")                -> 4
'   - =Q("A+B+C",3,4,5)        -> 12
'   - =Q("eye(3)")             -> the 3x3 identity matrix
'   - =Q("mean(A)",A1:D5)      -> vector with the mean of each column in cells A1:D5
'   - =Q("A.*B",A1:D5,F1:I5)   -> element wise multiplication of cells A1:D5 and F1:I5
'   - =Q("A([1 3],end)",A1:D5) -> get the last entries in row 1 and 3 of cells A1:D5
'   - =Q("sort(A)",A1:D5)      -> sort each column of cells A1:D5
'   - =Q("3+4;ans^2")          -> 49
'                                 Multiple expressions are separated by ";" or linebreak.
'                                 "ans" returns the last result and the very last
'                                 result is then returned by Q().
'
' Features:
'   - All standard MATLAB operators: :,::,+,-,*,/,.*,./,^,.^,||,&&,|,&,<,<=,>,>=,==,~=,~,'
'   - Most used MATLAB functions: eye,zeros,ones,sum,cumsum,cumprod,prod,mean,median,corr,
'     cov,prctile,std,isequal,fix,rand,randn,repmat,find,sqrt,exp,sort and many more...
'   - Indexing via A(2,:) or A(5,3:end)
'   - Concatenate matrices with '[]', i.e. [ A B; C D]
'   - Multiple expressions separated by ";" or a linebreak.
'     The variable ans contains the result of the previous expression
'   - Excel functions: if,iferror
'   - Prefix function calls with ! to call external VBA functions not found in Q.
'
' For the newest version, go to:
' http://github.com/nielsls/Q
'
' 2017, Niels Lykke Sørensen

Option Explicit
Option Base 1

Private Const version = "2.21"
    
Private Const NUMERICS = "0123456789"
Private Const ALPHAS = "abcdefghijklmnopqrstuvwxyz"
Private Const SINGLE_OPS = "()[],;:+-#'"
Private Const COMBO_OPS = ".|&<>~=*/^!"
Private Const DOUBLE_MAX = 1.79769313486231E+308
Private Const DOUBLE_MIN = -1.79769313486231E+308

Private expression As String
Private expressionIndex As Long
Private currentToken As String
Private previousTokenIsSpace As Boolean
Private arguments As Variant
Private endValues As Variant ' A stack of numbers providing the right value of the "end" constant
Private ans As Variant       ' Result of last answer when multiple expressions are used as input

' Entry point - the only public function in the library
Public Function Q(expr As Variant, ParamArray args() As Variant) As Variant
    On Error GoTo ErrorHandler

    expression = expr
    arguments = args
    endValues = Empty
    ans = Empty
    expressionIndex = 1
    Tokens_Next ' Find first token in input string
    
    Dim root As Variant
    While currentToken <> ""
        If currentToken = ";" Or currentToken = vbLf Then
            Tokens_Next
        Else
            root = Parse_Binary()
            'Utils_DumpTree root    'Uncomment for debugging
            ans = calc_tree(root)
            Utils_Assert _
                currentToken = "" Or currentToken = ";" Or currentToken = vbLf, _
                "'" & currentToken & "' not allowed here"
        End If
    Wend
    
    Select Case Utils_Numel(ans)
        Case 0
            'If an empty matrix is returned from Q() and
            'used in a cell it must be converted to #N/A
            'to avoid being converted to 0.
            If Utils_WasCalledFromCell() Then Q = [NA()]
        Case 1
            Q = ans(1, 1)
        Case Else
            Q = ans
    End Select
    Exit Function
    
ErrorHandler:
    ' If Q was called from cell, fail silently with an error msg.
    ' Else, raise a new error.
    Utils_Assert _
        Utils_WasCalledFromCell(), _
        Err.Description & " in """ & expression & """"
    Q = "ERROR - " & Err.Description
End Function

Private Function Utils_WasCalledFromCell() As Boolean
    Utils_WasCalledFromCell = TypeName(Application.Caller) <> "Error"
End Function

'*******************************************************************
' Evaluator invariants:
'   - All variables, including scalars, are represented internally
'     as 2-dimensional matrices
'   - All matrices are 1-based just as in MATLAB.
'   - Variable names: [A-Z]
'   - Function names: [a-z][a-z0-9_]*
'   - Numbers support e/E exponent
'   - The empty matrix/scalar [] has per definition 0 rows, 0 cols,
'     dimension 0 and is internally represented by the default
'     Variant value Empty
'     Missing parameters in function calls are given the value Empty
'*******************************************************************

'*********************
'*** TOKEN CONTROL ***
'*********************

Private Sub Tokens_Next()
    previousTokenIsSpace = Tokens_MoveCharPointer(" ")
    If expressionIndex > Len(expression) Then currentToken = "": Exit Sub
    
    Dim startIndex As Long: startIndex = expressionIndex
    Select Case Asc(Mid(expression, expressionIndex, 1))
    
        Case Asc("A") To Asc("Z")
            expressionIndex = expressionIndex + 1
            
        Case Asc("a") To Asc("z")
            Tokens_MoveCharPointer NUMERICS & ALPHAS & "_"
            
        Case Asc("0") To Asc("9")
            Tokens_MoveCharPointer NUMERICS
            If Tokens_MoveCharPointer(".", False, True) Then
                Tokens_MoveCharPointer NUMERICS
            End If
            If Tokens_MoveCharPointer("eE", False, True) Then
                Tokens_MoveCharPointer NUMERICS & "-", False, True
                Tokens_MoveCharPointer NUMERICS
            End If
            
        Case Asc("""")
            expressionIndex = expressionIndex + 1
            Tokens_MoveCharPointer """", True
            Utils_Assert expressionIndex <= Len(expression), "missing '""'"
            expressionIndex = expressionIndex + 1
            
        Case Asc(vbLf) 'New line
            expressionIndex = expressionIndex + 1
                
        Case Else
            If Not Tokens_MoveCharPointer(SINGLE_OPS, False, True) Then
                Tokens_MoveCharPointer COMBO_OPS
            End If
         
    End Select
    
    currentToken = Mid(expression, startIndex, expressionIndex - startIndex)
    Utils_Assert expressionIndex > startIndex Or expressionIndex > Len(expression), _
        "Illegal char: " & Mid(expression, expressionIndex, 1)
End Sub

Private Sub Tokens_AssertAndNext(Token As String)
    Utils_Assert Token = currentToken, "missing token: " & Token
    Tokens_Next
End Sub

Private Function Tokens_MoveCharPointer(str As String, _
    Optional stopWhenFound As Boolean = False, _
    Optional singleCharOnly As Boolean = False) As Boolean
    While expressionIndex <= Len(expression) _
        And stopWhenFound <> (InStr(str, Mid(expression, expressionIndex, 1)) > 0)
        expressionIndex = expressionIndex + 1
        Tokens_MoveCharPointer = True
        If singleCharOnly Then Exit Function
    Wend
End Function

'Returns true if token is a suitable operator
Private Function Parse_FindBinaryOp(Token As String, ByRef op As Variant) As Boolean
    Parse_FindBinaryOp = True
    Select Case Token
        'op = Array( <function name>, <precedence level>, <left associative> )
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
        Case Else: Parse_FindBinaryOp = False
    End Select
End Function

Private Function Parse_FindUnaryPrefixOp(Token As String, ByRef op As Variant) As Boolean
    Parse_FindUnaryPrefixOp = True
    Select Case Token
        Case "+": op = "uplus"
        Case "-": op = "uminus"
        Case "~": op = "negate"
        Case "#": op = "numel"
        Case "!": op = "extern"
        Case Else: Parse_FindUnaryPrefixOp = False
    End Select
End Function

Private Function Parse_FindUnaryPostfixOp(Token As String, ByRef op As Variant) As Boolean
    Parse_FindUnaryPostfixOp = True
    Select Case Token
        Case "'": op = "transpose"
        Case Else: Parse_FindUnaryPostfixOp = False
    End Select
End Function

Private Function Parse_Matrix() As Variant
    While currentToken <> "]"
        Utils_Stack_Push Array("fn_hcat", Parse_List(True)), Parse_Matrix
        If currentToken = ";" Then Tokens_Next
        Utils_Assert currentToken <> "", "Missing ']'"
    Wend
End Function

Private Function Parse_List(Optional isSpaceSeparator As Boolean = False) As Variant
    Do While InStr(";)]", currentToken) = 0
        If currentToken = "," Then
            Utils_Stack_Push Array("eval_constant", Empty), Parse_List
        Else
            Utils_Stack_Push Parse_Binary(), Parse_List
        End If
        If currentToken = "," Then
            Tokens_Next
        ElseIf Not (previousTokenIsSpace And isSpaceSeparator) Then
            Exit Do
        End If
    Loop
End Function

Private Function Parse_Binary(Optional lastPrec As Long = -999) As Variant
    Parse_Binary = Parse_Prefix()
    Dim op: Do While Parse_FindBinaryOp(currentToken, op)
        If op(2) + CLng(op(3)) < lastPrec Then Exit Do
        Tokens_Next
        Parse_Binary = Array("op_" & op(1), Array(Parse_Binary, Parse_Binary(CLng(op(2)))))
    Loop
End Function

Private Function Parse_Prefix() As Variant
    Dim op
    If Parse_FindUnaryPrefixOp(currentToken, op) Then
        Tokens_Next
        Parse_Prefix = Array("op_" & op, Array(Parse_Prefix()))
    Else
        Parse_Prefix = Parse_Postfix()
    End If
End Function

Private Function Parse_Postfix() As Variant
    Parse_Postfix = Parse_Atomic()
    Dim op: Do
        If Parse_FindUnaryPostfixOp(currentToken, op) Then
            Tokens_Next
            Parse_Postfix = Array("op_" & op, Array(Parse_Postfix))
        ElseIf currentToken = "(" Then
            Tokens_Next
            Parse_Postfix = Array("op_index", Array(Parse_Postfix, Parse_List()))
            Tokens_AssertAndNext ")"
        Else
            Exit Do
        End If
    Loop While True
End Function

Private Function Parse_Atomic() As Variant
    Utils_Assert currentToken <> "", "missing argument"
    Select Case Asc(currentToken) ' Filter on first char of token
            
        Case Asc("(")
            Tokens_Next
            Parse_Atomic = Parse_Binary()
            Tokens_AssertAndNext ")"
            
        Case Asc(":")
            Parse_Atomic = Array("eval_colon")
            Tokens_Next
            
        Case Asc("0") To Asc("9") ' Found a numeric constant
            Parse_Atomic = Array("eval_constant", Utils_ToMatrix(val(currentToken)))
            Tokens_Next
            
        Case Asc("A") To Asc("Z")  ' Found an input variable
            Parse_Atomic = Array("eval_variable", Asc(currentToken) - Asc("A"))
            Tokens_Next
            
        Case Asc("a") To Asc("z")
            If currentToken = "end" Then
                Parse_Atomic = Array("eval_end")
                Tokens_Next
            ElseIf currentToken = "ans" Then
                Parse_Atomic = Array("eval_ans")
                Tokens_Next
            Else                   ' Found a function call
                Parse_Atomic = "fn_" & currentToken
                Tokens_Next
                If currentToken = "(" Then
                    Tokens_AssertAndNext "("
                    Parse_Atomic = Array(Parse_Atomic, Parse_List())
                    Tokens_AssertAndNext ")"
                Else
                    Parse_Atomic = Array(Parse_Atomic, Empty)
                End If
            End If
            
        Case Asc("[")  ' Found a matrix concatenation
            Tokens_Next
            Parse_Atomic = Array("fn_vcat", Parse_Matrix())
            Tokens_AssertAndNext "]"
            
        Case Asc("""") ' Found a string constant
            Parse_Atomic = Array("eval_constant", Utils_ToMatrix(Mid(currentToken, 2, Len(currentToken) - 2)))
            Tokens_Next
            
        Case Else
            Utils_Fail "unexpected token: " & currentToken
            
    End Select
End Function

'*************
'*** UTILS ***
'*************
Private Function MAX(a As Variant, b As Variant) As Variant
    If a > b Then MAX = a Else MAX = b
End Function

Private Function MIN(a As Variant, b As Variant) As Variant
    If a < b Then MIN = a Else MIN = b
End Function

Private Function IFF(expr As Boolean, alt1 As Variant, alt2 As Variant) As Variant
    If expr Then IFF = alt1 Else IFF = alt2
End Function

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
    If Not IsEmpty(v) Then Utils_Numel = UBound(v, 1) * UBound(v, 2)
End Function

Private Function Utils_Rows(ByRef v As Variant) As Long
    If Not IsEmpty(v) Then Utils_Rows = UBound(v, 1)
End Function

Private Function Utils_Cols(ByRef v As Variant) As Long
    If Not IsEmpty(v) Then Utils_Cols = UBound(v, 2)
End Function

Private Sub Utils_Size(v As Variant, ByRef r As Variant, ByRef c As Variant)
    If IsEmpty(v) Then
        r = 0
        c = 0
    Else
        r = UBound(v, 1)
        c = UBound(v, 2)
    End If
End Sub

' From linear index to subscripts in a column-major matrix
Private Sub Utils_Ind2Sub(rows As Long, ind As Long, ByRef i As Long, ByRef j As Long)
    j = (ind - 1) \ rows + 1
    i = ind - rows * (j - 1)
End Sub

Private Function Utils_ToMatrix(val As Variant) As Variant
    Utils_ToMatrix = val
    Utils_Conform Utils_ToMatrix
End Function

' Makes sure that a scalar is transformed to a 1x1 matrix
' and a 1-dim vector is transformed to a 2-dim vector of size 1xN
Private Sub Utils_Conform(ByRef v As Variant)
    If IsEmpty(v) Then Exit Sub
    Dim r
    Select Case Utils_Dimensions(v)
        Case 0:
            ReDim r(1, 1)
            r(1, 1) = v
            v = r
        Case 1:
            ReDim r(1, UBound(v))
            Dim i As Long
            For i = 1 To UBound(r, 2)
                r(1, i) = v(i)
            Next i
            v = r
        Case Is > 2:
            Utils_Fail "dimension > 2"
    End Select
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

Private Function Utils_RepScalar(ByRef out As Variant, val As Variant, n As Long, m As Long) As Variant
    Utils_Assert n >= 0 And m >= 0, "cannot create matrix with negative size"
    If n = 0 Or m = 0 Then Exit Function
    ReDim out(n, m)
    For n = 1 To UBound(out, 1)
        For m = 1 To UBound(out, 2)
            out(n, m) = val
        Next m
    Next n
End Function

' Transforms all entries in the vector from trees to values
Private Sub Utils_CalcArgs(args As Variant)
    Dim i As Long: For i = 1 To Utils_Stack_Size(args)
        args(i) = calc_tree(args(i))
    Next i
End Sub

' Test if a flag was supplied in the args, i.e. such as "descend" in the sort function
Private Function Utils_IsFlagSet(args As Variant, flag As String) As Boolean
    Dim i As Long
    For i = UBound(args) To 1 Step -1
        If StrComp(TypeName(args(i)(1, 1)), "String") = 0 Then
            If StrComp(args(i)(1, 1), flag, vbTextCompare) = 0 Then
                Utils_IsFlagSet = True
                Exit Function
            End If
        End If
    Next i
End Function

'do cols -> return 0, do rows -> return 1
Private Function Utils_CalcDimDirection(args As Variant, Optional dimIndex As Long = 2) As Long
    If UBound(args) >= dimIndex Then
        If IsNumeric(args(dimIndex)(1, 1)) Then
            Utils_CalcDimDirection = args(dimIndex)(1, 1) - 1
            Exit Function
        End If
    End If
    Utils_CalcDimDirection = -(Utils_Rows(args(1)) = 1)
End Function

' Returns the size of the return matrix in functions like zeros, rand, repmat, ...
' Size must be last in the args and can be either nothing, 1 scalar, 2 scalars or a vector with two scalars
Private Function Utils_GetSizeFromArgs(args As Variant, ByRef n As Long, ByRef m As Long, Optional index As Long = 2)
    Select Case Utils_Stack_Size(args)
        Case Is < index
            n = 1: m = 1
        Case Is = index
            Select Case Utils_Numel(args(index))
                Case 1
                    n = args(index)(1, 1)
                    m = n
                Case 2
                    n = args(index)(1, 1)
                    m = args(index)(MIN(2, UBound(args(index), 1)), MIN(2, UBound(args(index), 2)))
                Case Else
                    Utils_Fail "bad size input"
            End Select
        Case Is = index + 1
            n = args(index)(1, 1)
            m = args(index + 1)(1, 1)
        Case Else
            Utils_Fail "bad size input"
    End Select
End Function

' Easy way of obtaining the value of an optional argument
Private Function Utils_GetOptionalScalarArg(args As Variant, index As Long, defaultValue As Variant) As Variant
    If index <= Utils_Stack_Size(args) Then
        If IsEmpty(args(index)) Then
            Utils_GetOptionalScalarArg = Empty
        Else
            Utils_GetOptionalScalarArg = args(index)(1, 1)
        End If
    Else
        Utils_GetOptionalScalarArg = defaultValue
    End If
End Function

' Do the initial calculations that are the same for every binary operation
' Broadcasting is done, st. a matrix with size (n,m) will match all
' matrices with size (1,1), (1,m), (n,1) or (n,m).
Private Function Utils_SetupBinaryOperation( _
        args As Variant, out As Variant, _
        ByRef r As Long, ByRef c As Long, _
        ByRef arg1_r As Long, ByRef arg1_c As Long, _
        ByRef arg2_r As Long, ByRef arg2_c As Long, _
        Optional preCalcArgs As Boolean = True) As Variant
    If preCalcArgs Then Utils_CalcArgs args
    Utils_Size args(1), arg1_r, arg1_c
    Utils_Size args(2), arg2_r, arg2_c
    Utils_Assert _
        (arg1_r = 1 Or arg2_r = 1 Or arg1_r = arg2_r) And (arg1_c = 1 Or arg2_c = 1 Or arg1_c = arg2_c), _
        "dimension mismatch"
    If arg1_r > 0 And arg2_r > 0 Then
        r = MAX(arg1_r, arg2_r)
        c = MAX(arg1_c, arg2_c)
        ReDim out(r, c)
    Else
        r = 0: c = 0
    End If
End Function

' Do the initial processing for operations which reduce the dimension of a matrix.
' These includes sum, prod, mean, any, all, count etc...
Private Sub Utils_SetupReducedDimOperation(ByRef args As Variant, ByRef mat As Variant, ByRef x As Long, Optional dimIndex As Long = 2)
    x = Utils_CalcDimDirection(args, dimIndex)
    ReDim mat(x * UBound(args(1), 1) + (1 - x), (1 - x) * UBound(args(1), 2) + x)
End Sub

' Is called from each Q function to ensure the given function has
' has been called with the right number of arguments
Private Sub Utils_AssertArgsCount(args As Variant, lb As Long, ub As Long)
    Dim size As Long: size = Utils_Stack_Size(args)
    Utils_Assert size >= lb, "too few arguments"
    Utils_Assert size <= ub, "too many arguments"
End Sub

' Allows each Q function to fail gracefully with a nice error message.
Private Sub Utils_Assert(expr As Boolean, Optional msg As String = "unknown error")
    If expr Then Exit Sub
    Err.Raise 999, , msg
End Sub

Private Sub Utils_Fail(msg As String)
    Utils_Assert False, msg
End Sub

'*****************
'*** CALC ROOT ***
'*****************

' calc_tree is responsible for turning the
' abstract syntax tree (AST) into an actual
' result.
' It does this by recursively calling the
' different functions and operators needed.
'
' Branches of "select case" statements are
' used to determine which function to call.
' A more elegant solution would be to just
' invoke Application.Run().
' However, Application.Run is not feasible
' as it is both slow and fucks up error
' propagation.
Private Function calc_tree(root As Variant) As Variant
    Dim prefix As String, name As String
    root(1) = Split(root(1), "_")
    prefix = root(1)(0)
    name = root(1)(1)
    
    Select Case prefix
    
    ' Non-functions
    Case "eval":
        Select Case name
        Case "constant":
            calc_tree = root(2)
        Case "variable":
            Utils_Assert root(2) <= UBound(arguments), "variable '" & Chr(Asc("A") + root(2)) & "' not found."
            calc_tree = CVar(arguments(root(2)))
            Utils_Conform calc_tree
        Case "end":
            Utils_Assert Utils_Stack_Size(endValues) > 0, """end"" not allowed here."
            calc_tree = Utils_ToMatrix(Utils_Stack_Peek(endValues))
        Case "ans":
            calc_tree = ans
        Case "colon":
            Utils_Fail "colon not allowed here"
        End Select
       
    ' Operators
    Case "op"
        Select Case name
        Case "and": calc_tree = op_and(root(2))
        Case "andshortcircuit": calc_tree = op_andshortcircuit(root(2))
        Case "colon": calc_tree = op_colon(root(2))
        Case "divide": calc_tree = op_divide(root(2))
        Case "eq": calc_tree = op_eq(root(2))
        Case "extern": calc_tree = op_extern(root(2))
        Case "gt": calc_tree = op_gt(root(2))
        Case "gte": calc_tree = op_gte(root(2))
        Case "index": calc_tree = op_index(root(2))
        Case "lt": calc_tree = op_lt(root(2))
        Case "lte": calc_tree = op_lte(root(2))
        Case "minus": calc_tree = op_minus(root(2))
        Case "mdivide": calc_tree = op_mdivide(root(2))
        Case "mpower": calc_tree = op_mpower(root(2))
        Case "mtimes": calc_tree = op_mtimes(root(2))
        Case "ne": calc_tree = op_ne(root(2))
        Case "negate": calc_tree = op_negate(root(2))
        Case "numel": calc_tree = op_numel(root(2))
        Case "or": calc_tree = op_or(root(2))
        Case "orshortcircuit": calc_tree = op_orshortcircuit(root(2))
        Case "plus": calc_tree = op_plus(root(2))
        Case "power": calc_tree = op_power(root(2))
        Case "times": calc_tree = op_times(root(2))
        Case "transpose": calc_tree = op_transpose(root(2))
        Case "uplus": calc_tree = op_uplus(root(2))
        Case "uminus": calc_tree = op_uminus(root(2))
        End Select
      
    ' Functions
    Case "fn"
        If name <> "if" And name <> "iferror" And name <> "expand" Then Utils_CalcArgs root(2)
        Select Case name
        Case "all": calc_tree = fn_all(root(2))
        Case "any": calc_tree = fn_any(root(2))
        Case "arrayfun": calc_tree = fn_arrayfun(root(2))
        Case "binom": calc_tree = fn_binom(root(2))
        Case "ceil": calc_tree = fn_ceil(root(2))
        Case "cols": calc_tree = fn_cols(root(2))
        Case "corr": calc_tree = fn_corr(root(2))
        Case "count": calc_tree = fn_count(root(2))
        Case "counta": calc_tree = fn_counta(root(2))
        Case "cov": calc_tree = fn_cov(root(2))
        Case "cummax": calc_tree = fn_cummax(root(2))
        Case "cummin": calc_tree = fn_cummin(root(2))
        Case "cumprod": calc_tree = fn_cumprod(root(2))
        Case "cumsum": calc_tree = fn_cumsum(root(2))
        Case "diag": calc_tree = fn_diag(root(2))
        Case "diff": calc_tree = fn_diff(root(2))
        Case "droperror": calc_tree = fn_droperror(root(2))
        Case "e": calc_tree = fn_e(root(2))
        Case "exp": calc_tree = fn_exp(root(2))
        Case "expand": calc_tree = fn_expand(root(2))
        Case "eye": calc_tree = fn_eye(root(2))
        Case "fact": calc_tree = fn_fact(root(2))
        Case "false": calc_tree = fn_false(root(2))
        Case "find": calc_tree = fn_find(root(2))
        Case "fix": calc_tree = fn_fix(root(2))
        Case "floor": calc_tree = fn_floor(root(2))
        Case "hcat": calc_tree = fn_hcat(root(2))
        Case "if": calc_tree = fn_if(root(2))
        Case "iferror": calc_tree = fn_iferror(root(2))
        Case "inv": calc_tree = fn_inv(root(2))
        Case "isempty": calc_tree = fn_isempty(root(2))
        Case "isequal": calc_tree = fn_isequal(root(2))
        Case "iserror": calc_tree = fn_iserror(root(2))
        Case "islogical": calc_tree = fn_islogical(root(2))
        Case "isnum": calc_tree = fn_isnum(root(2))
        Case "join": calc_tree = fn_join(root(2))
        Case "linspace": calc_tree = fn_linspace(root(2))
        Case "log": calc_tree = fn_log(root(2))
        Case "match": calc_tree = fn_match(root(2))
        Case "max": calc_tree = fn_max(root(2))
        Case "mean": calc_tree = fn_mean(root(2))
        Case "median": calc_tree = fn_median(root(2))
        Case "min": calc_tree = fn_min(root(2))
        Case "normcdf": calc_tree = fn_normcdf(root(2))
        Case "numel": calc_tree = fn_numel(root(2))
        Case "ones": calc_tree = fn_ones(root(2))
        Case "percentrank": calc_tree = fn_percentrank(root(2))
        Case "pi": calc_tree = fn_pi(root(2))
        Case "prctile": calc_tree = fn_prctile(root(2))
        Case "prod": calc_tree = fn_prod(root(2))
        Case "rand": calc_tree = fn_rand(root(2))
        Case "randi": calc_tree = fn_randi(root(2))
        Case "randn": calc_tree = fn_randn(root(2))
        Case "repmat": calc_tree = fn_repmat(root(2))
        Case "ret2tick": calc_tree = fn_ret2tick(root(2))
        Case "reshape": calc_tree = fn_reshape(root(2))
        Case "round": calc_tree = fn_round(root(2))
        Case "rows": calc_tree = fn_rows(root(2))
        Case "size": calc_tree = fn_size(root(2))
        Case "sort": calc_tree = fn_sort(root(2))
        Case "sorttable": calc_tree = fn_sorttable(root(2))
        Case "sqrt": calc_tree = fn_sqrt(root(2))
        Case "std": calc_tree = fn_std(root(2))
        Case "sum": calc_tree = fn_sum(root(2))
        Case "tick2ret": calc_tree = fn_tick2ret(root(2))
        Case "tostring": calc_tree = fn_tostring(root(2))
        Case "true": calc_tree = fn_true(root(2))
        Case "unique": calc_tree = fn_unique(root(2))
        Case "var": calc_tree = fn_var(root(2))
        Case "vcat": calc_tree = fn_vcat(root(2))
        Case "version": calc_tree = fn_version(root(2))
        Case "xor": calc_tree = fn_xor(root(2))
        Case "zeros": calc_tree = fn_zeros(root(2))
        Case Else:
            Utils_Assert _
                False, _
                "unknown function """ & name & """" _
                    & IFF(Len(name) = 1, "; did you mean variable " & UCase(name) & "?", "")
        End Select
            
    End Select
End Function

'*****************
'*** OPERATORS ***
'*****************

Private Function Utils_IsVectorShape(r As Long, c As Long) As Boolean
    Utils_IsVectorShape = (r = 1 And c > 1) Or (r > 1 And c = 1)
End Function

Private Sub op_indexarg(root As Variant, endValue As Long, ByRef idx As Variant, ByRef r As Long, ByRef c As Long, ByRef t As Long)
    If root(1) = "eval_colon" Then
        t = 1
        r = endValue
        c = 1
        
    ElseIf root(1) = "op_colon" Then
        t = 2
        ReDim idx(3)
        Utils_Stack_Push endValue, endValues
        idx(1) = calc_tree(root(2)(1))(1, 1)
        If root(2)(2)(1) <> "op_colon" Then
            idx(2) = 1
            idx(3) = calc_tree(root(2)(2))(1, 1)
        Else
            idx(2) = calc_tree(root(2)(2)(2)(1))(1, 1)
            idx(3) = calc_tree(root(2)(2)(2)(2))(1, 1)
        End If
        Utils_Stack_Pop endValues
        r = 1
        c = MAX(0, 1 + WorksheetFunction.RoundDown((idx(3) - idx(1)) / idx(2), 0))
        
    Else
        t = 3
        Utils_Stack_Push endValue, endValues
        idx = calc_tree(root)
        Utils_Stack_Pop endValues
        Utils_Size idx, r, c
        Dim i As Long, j As Long
        For i = 1 To r
            For j = 1 To c
                If Not WorksheetFunction.IsLogical(idx(i, j)) Then Exit Sub
            Next j
        Next i
        idx = fn_find(Array(idx))
        Utils_Size idx, r, c
        
    End If
End Sub

' Evaluates matrix indexing/subsetting
Private Function op_index(args As Variant) As Variant
    Dim out As Variant, out_i As Long, out_j As Long
    Dim arg_r As Long, arg_c As Long
    Dim arg_i As Long, arg_j As Long
    Dim idx1 As Variant, idx2 As Variant
    Dim idx1_i As Long, idx2_i As Long
    Dim idx1_j As Long, idx2_j As Long
    Dim idx1_r As Long, idx2_r As Long
    Dim idx1_c As Long, idx2_c As Long
    Dim t1 As Long, t2 As Long ' t=1 for colon only, t=2 for sequence, t=3 for vector/mat
    args(1) = calc_tree(args(1))
    Utils_Size args(1), arg_r, arg_c
    
    Select Case Utils_Stack_Size(args(2))
        
        Case 0:
            out = args(1)
        
        Case 1:
            op_indexarg args(2)(1), arg_r * arg_c, idx1, idx1_r, idx1_c, t1
            If idx1_r * idx1_c = 0 Then Exit Function
            If Utils_IsVectorShape(idx1_r, idx1_c) And Utils_IsVectorShape(arg_r, arg_c) Then
                If arg_r > 1 Or t1 = 1 Then ReDim out(idx1_r * idx1_c, 1) Else ReDim out(1, idx1_r * idx1_c)
            Else
                ReDim out(idx1_r, idx1_c)
            End If
            Dim k As Long
            For k = 1 To UBound(out, 1) * UBound(out, 2)
                Select Case t1
                    Case 1: Utils_Ind2Sub arg_r, k, arg_i, arg_j
                    Case 2: Utils_Ind2Sub arg_r, idx1(1) + (k - 1) * idx1(2), arg_i, arg_j
                    Case 3: Utils_Ind2Sub idx1_r, k, idx1_i, idx1_j
                            Utils_Ind2Sub arg_r, CLng(idx1(idx1_i, idx1_j)), arg_i, arg_j
                End Select
                Utils_Ind2Sub UBound(out, 1), k, out_i, out_j
                out(out_i, out_j) = args(1)(arg_i, arg_j)
            Next k
            
        Case 2:
            op_indexarg args(2)(1), arg_r, idx1, idx1_r, idx1_c, t1
            op_indexarg args(2)(2), arg_c, idx2, idx2_r, idx2_c, t2
            If idx1_r * idx1_c = 0 Or idx2_r * idx2_c = 0 Then Exit Function
            ReDim out(idx1_r * idx1_c, idx2_r * idx2_c)
            For out_i = 1 To UBound(out, 1)
                For out_j = 1 To UBound(out, 2)
                    Select Case t1
                        Case 1: arg_i = out_i
                        Case 2: arg_i = idx1(1) + (out_i - 1) * idx1(2)
                        Case 3: Utils_Ind2Sub idx1_r, out_i, idx1_i, idx1_j
                                arg_i = idx1(idx1_i, idx1_j)
                    End Select
                    Select Case t2
                        Case 1: arg_j = out_j
                        Case 2: arg_j = idx2(1) + (out_j - 1) * idx2(2)
                        Case 3: Utils_Ind2Sub idx2_r, out_j, idx2_i, idx2_j
                                arg_j = idx2(idx2_i, idx2_j)
                    End Select
                    out(out_i, out_j) = args(1)(arg_i, arg_j)
                Next out_j
            Next out_i
            
        Case Else:
            Utils_Fail "too many indices"
    
    End Select

    op_index = out
End Function

' Matches operator !
Private Function op_extern(args As Variant) As Variant
    args(1)(1) = Mid(args(1)(1), 4)
    Dim a: a = args(1)(2)
    Utils_CalcArgs a
    Select Case UBound(a)
        Case 0: op_extern = Run(args(1)(1))
        Case 1: op_extern = Run(args(1)(1), a(1))
        Case 2: op_extern = Run(args(1)(1), a(1), a(2))
        Case 3: op_extern = Run(args(1)(1), a(1), a(2), a(3))
        Case 4: op_extern = Run(args(1)(1), a(1), a(2), a(3), a(4))
        Case 5: op_extern = Run(args(1)(1), a(1), a(2), a(3), a(4), a(5))
        Case 6: op_extern = Run(args(1)(1), a(1), a(2), a(3), a(4), a(5), a(6))
        Case 7: op_extern = Run(args(1)(1), a(1), a(2), a(3), a(4), a(5), a(6), a(7))
        Case 8: op_extern = Run(args(1)(1), a(1), a(2), a(3), a(4), a(5), a(6), a(7), a(8))
        Case 9: op_extern = Run(args(1)(1), a(1), a(2), a(3), a(4), a(5), a(6), a(7), a(8), a(9))
        Case 10: op_extern = Run(args(1)(1), a(1), a(2), a(3), a(4), a(5), a(6), a(7), a(8), a(9), a(10))
        Case Else: Utils_Fail "cannot evaluate " & args(1)(1) & ": too many arguments"
    End Select
    Utils_Conform op_extern
End Function

' Matches operator ||
Private Function op_orshortcircuit(args As Variant) As Variant
    args(1) = calc_tree(args(1))
    Utils_Assert Utils_Numel(args(1)) = 1, "||: 1st argument must be scalar"
    On Error GoTo ErrorHandler
    If CBool(args(1)(1, 1)) Then
        op_orshortcircuit = Utils_ToMatrix(True)
    Else
        On Error GoTo 0
        args(2) = calc_tree(args(2))
        Utils_Assert Utils_Numel(args(2)) = 1, "||: 2nd argument must be scalar"
        On Error GoTo ErrorHandler
        op_orshortcircuit = Utils_ToMatrix(CBool(args(2)(1, 1)))
    End If
    Exit Function
ErrorHandler:
    Utils_Fail "||: could not convert argument to boolean value"
End Function

' Matches operator &&
Private Function op_andshortcircuit(args As Variant) As Variant
    args(1) = calc_tree(args(1))
    Utils_Assert Utils_Numel(args(1)) = 1, "&&: 1st argument must be scalar"
    On Error GoTo ErrorHandler
    If Not CBool(args(1)(1, 1)) Then
        op_andshortcircuit = Utils_ToMatrix(False)
    Else
        On Error GoTo 0
        args(2) = calc_tree(args(2))
        Utils_Assert Utils_Numel(args(2)) = 1, "&&: 2nd argument must be scalar"
        On Error GoTo ErrorHandler
        op_andshortcircuit = Utils_ToMatrix(CBool(args(2)(1, 1)))
    End If
    Exit Function
ErrorHandler:
    Utils_Fail "&&: could not convert argument to boolean value"
End Function

' Matches operator &
Private Function op_and(args As Variant) As Variant
    Dim out, r As Long, c As Long, i As Long, j As Long
    Dim arg1_r As Long, arg1_c As Long, arg2_r As Long, arg2_c As Long
    Utils_SetupBinaryOperation args, out, r, c, arg1_r, arg1_c, arg2_r, arg2_c
    For i = 1 To r
        For j = 1 To c
            out(i, j) = CBool(args(1)(MIN(i, arg1_r), MIN(j, arg1_c))) _
                    And CBool(args(2)(MIN(i, arg2_r), MIN(j, arg2_c)))
        Next j
    Next i
    op_and = out
End Function

' Matches operator |
Private Function op_or(args As Variant) As Variant
    Dim out, r As Long, c As Long, i As Long, j As Long
    Dim arg1_r As Long, arg1_c As Long, arg2_r As Long, arg2_c As Long
    Utils_SetupBinaryOperation args, out, r, c, arg1_r, arg1_c, arg2_r, arg2_c
    For i = 1 To r
        For j = 1 To c
            out(i, j) = CBool(args(1)(MIN(i, arg1_r), MIN(j, arg1_c))) _
                     Or CBool(args(2)(MIN(i, arg2_r), MIN(j, arg2_c)))
        Next j
    Next i
    op_or = out
End Function

' Matches operator <
Private Function op_lt(args As Variant) As Variant
    Dim out, r As Long, c As Long, i As Long, j As Long
    Dim arg1_r As Long, arg1_c As Long, arg2_r As Long, arg2_c As Long
    Utils_SetupBinaryOperation args, out, r, c, arg1_r, arg1_c, arg2_r, arg2_c
    For i = 1 To r
        For j = 1 To c
            out(i, j) = args(1)(MIN(i, arg1_r), MIN(j, arg1_c)) _
                      < args(2)(MIN(i, arg2_r), MIN(j, arg2_c))
        Next j
    Next i
    op_lt = out
End Function

' Matches operator <=
Private Function op_lte(args As Variant) As Variant
    Dim out, r As Long, c As Long, i As Long, j As Long
    Dim arg1_r As Long, arg1_c As Long, arg2_r As Long, arg2_c As Long
    Utils_SetupBinaryOperation args, out, r, c, arg1_r, arg1_c, arg2_r, arg2_c
    For i = 1 To r
        For j = 1 To c
            out(i, j) = args(1)(MIN(i, arg1_r), MIN(j, arg1_c)) _
                     <= args(2)(MIN(i, arg2_r), MIN(j, arg2_c))
        Next j
    Next i
    op_lte = out
End Function

' Matches operator >
Private Function op_gt(args As Variant) As Variant
    Dim out, r As Long, c As Long, i As Long, j As Long
    Dim arg1_r As Long, arg1_c As Long, arg2_r As Long, arg2_c As Long
    Utils_SetupBinaryOperation args, out, r, c, arg1_r, arg1_c, arg2_r, arg2_c
    For i = 1 To r
        For j = 1 To c
            out(i, j) = args(1)(MIN(i, arg1_r), MIN(j, arg1_c)) _
                      > args(2)(MIN(i, arg2_r), MIN(j, arg2_c))
        Next j
    Next i
    op_gt = out
End Function

' Matches operator >=
Private Function op_gte(args As Variant) As Variant
    Dim out, r As Long, c As Long, i As Long, j As Long
    Dim arg1_r As Long, arg1_c As Long, arg2_r As Long, arg2_c As Long
    Utils_SetupBinaryOperation args, out, r, c, arg1_r, arg1_c, arg2_r, arg2_c
    For i = 1 To r
        For j = 1 To c
            out(i, j) = args(1)(MIN(i, arg1_r), MIN(j, arg1_c)) _
                     >= args(2)(MIN(i, arg2_r), MIN(j, arg2_c))
        Next j
    Next i
    op_gte = out
End Function

' Matches operators = and ==
Private Function op_eq(args As Variant) As Variant
    Dim out, r As Long, c As Long, i As Long, j As Long
    Dim arg1_r As Long, arg1_c As Long, arg2_r As Long, arg2_c As Long
    Utils_SetupBinaryOperation args, out, r, c, arg1_r, arg1_c, arg2_r, arg2_c
    For i = 1 To r
        For j = 1 To c
            out(i, j) = args(1)(MIN(i, arg1_r), MIN(j, arg1_c)) _
                      = args(2)(MIN(i, arg2_r), MIN(j, arg2_c))
        Next j
    Next i
    op_eq = out
End Function

' Matches operator ~=
Private Function op_ne(args As Variant) As Variant
    Dim out, r As Long, c As Long, i As Long, j As Long
    Dim arg1_r As Long, arg1_c As Long, arg2_r As Long, arg2_c As Long
    Utils_SetupBinaryOperation args, out, r, c, arg1_r, arg1_c, arg2_r, arg2_c
    For i = 1 To r
        For j = 1 To c
            out(i, j) = args(1)(MIN(i, arg1_r), MIN(j, arg1_c)) _
                     <> args(2)(MIN(i, arg2_r), MIN(j, arg2_c))
        Next j
    Next i
    op_ne = out
End Function

' Matches operator ~
Private Function op_negate(args As Variant) As Variant
    Dim r As Long, c As Long, i As Long, j As Long
    Utils_CalcArgs args
    Utils_Size args(1), r, c
    For i = 1 To r
        For j = 1 To c
            args(1)(i, j) = Not CBool(args(1)(i, j))
        Next j
    Next i
    op_negate = args(1)
End Function

' Matches operator : with one or two arguments
Private Function op_colon(args As Variant) As Variant
    Dim m As Long, i As Long, step As Double, start As Double, last As Double
    start = calc_tree(args(1))(1, 1)
    If args(2)(1) <> "op_colon" Then
        ' x:y
        step = 1
        last = calc_tree(args(2))(1, 1)
        m = 1 + WorksheetFunction.RoundDown(last - start, 0)
    Else
        ' x:y:z
        step = calc_tree(args(2)(2)(1))(1, 1)
        last = calc_tree(args(2)(2)(2))(1, 1)
        m = 1 + WorksheetFunction.RoundDown((last - start) / step, 0)
    End If
    If m < 1 Then Exit Function
    Dim r: ReDim r(1, m)
    For i = 1 To m
        r(1, i) = start + (i - 1) * step
    Next i
    op_colon = r
End Function

' Matches operator +
Private Function op_plus(args As Variant) As Variant
    Dim out, r As Long, c As Long, i As Long, j As Long
    Dim arg1_r As Long, arg1_c As Long, arg2_r As Long, arg2_c As Long
    Utils_SetupBinaryOperation args, out, r, c, arg1_r, arg1_c, arg2_r, arg2_c
    For i = 1 To r
        For j = 1 To c
            out(i, j) = args(1)(MIN(i, arg1_r), MIN(j, arg1_c)) _
                      + args(2)(MIN(i, arg2_r), MIN(j, arg2_c))
        Next j
    Next i
    op_plus = out
End Function

' Matches unary operator +
Private Function op_uplus(args As Variant) As Variant
    op_uplus = calc_tree(args(1))
End Function

' Matches binary operator -
Private Function op_minus(args As Variant) As Variant
    Dim out, r As Long, c As Long, i As Long, j As Long
    Dim arg1_r As Long, arg1_c As Long, arg2_r As Long, arg2_c As Long
    Utils_SetupBinaryOperation args, out, r, c, arg1_r, arg1_c, arg2_r, arg2_c
    For i = 1 To r
        For j = 1 To c
            out(i, j) = args(1)(MIN(i, arg1_r), MIN(j, arg1_c)) _
                      - args(2)(MIN(i, arg2_r), MIN(j, arg2_c))
        Next j
    Next i
    op_minus = out
End Function

' Matches prefix unary operator -
Private Function op_uminus(args As Variant) As Variant
    Dim i As Long, j As Long, r As Long, c As Long
    Utils_CalcArgs args
    Utils_Size args(1), r, c
    For i = 1 To r
        For j = 1 To c
            args(1)(i, j) = -args(1)(i, j)
        Next j
    Next i
    op_uminus = args(1)
End Function

' Matches operator *
Private Function op_mtimes(args As Variant) As Variant
    Utils_CalcArgs args
    Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
    Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
    If r1 * c1 > 1 And r2 * c2 > 1 Then
        Utils_Assert UBound(args(1), 2) = UBound(args(2), 1), "operator *: matrix sizes not compatible"
        op_mtimes = WorksheetFunction.MMult(args(1), args(2))
        Utils_Conform op_mtimes
    Else
        If r1 = 0 Or r2 = 0 Then Exit Function
        Dim out: ReDim out(MAX(r1, r2), MAX(c1, c2))
        Dim i As Long, j As Long
        For i = 1 To UBound(out, 1)
            For j = 1 To UBound(out, 2)
                out(i, j) = args(1)(MIN(i, r1), MIN(j, c1)) _
                          * args(2)(MIN(i, r2), MIN(j, c2))
            Next j
        Next i
        op_mtimes = out
    End If
End Function

' Matches operator .*
Private Function op_times(args As Variant) As Variant
    Dim out, r As Long, c As Long, i As Long, j As Long
    Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long
    Utils_SetupBinaryOperation args, out, r, c, r1, c1, r2, c2
    For i = 1 To r
        For j = 1 To c
            out(i, j) = args(1)(MIN(i, r1), MIN(j, c1)) _
                      * args(2)(MIN(i, r2), MIN(j, c2))
        Next j
    Next i
    op_times = out
End Function

' Matches operator /
Private Function op_mdivide(args As Variant) As Variant
    Dim out, r As Long, c As Long, i As Long, j As Long
    Dim arg1_r As Long, arg1_c As Long, arg2_r As Long, arg2_c As Long
    Utils_SetupBinaryOperation args, out, r, c, arg1_r, arg1_c, arg2_r, arg2_c
    For i = 1 To r
        For j = 1 To c
            out(i, j) = args(1)(MIN(i, arg1_r), MIN(j, arg1_c)) _
                      / args(2)(MIN(i, arg2_r), MIN(j, arg2_c))
        Next j
    Next i
    op_mdivide = out
End Function

' Matches operator ./
Private Function op_divide(args As Variant) As Variant
    Dim out, r As Long, c As Long, i As Long, j As Long
    Dim arg1_r As Long, arg1_c As Long, arg2_r As Long, arg2_c As Long
    Utils_SetupBinaryOperation args, out, r, c, arg1_r, arg1_c, arg2_r, arg2_c
    For i = 1 To r
        For j = 1 To c
            out(i, j) = args(1)(MIN(i, arg1_r), MIN(j, arg1_c)) _
                      / args(2)(MIN(i, arg2_r), MIN(j, arg2_c))
        Next j
    Next i
    op_divide = out
End Function

' Matches operator ^
Private Function op_mpower(args As Variant) As Variant
    Utils_CalcArgs args
    Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
    Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
    If r1 = 1 And c1 = 1 And r2 = 1 And c2 = 1 Then
        op_mpower = Utils_ToMatrix(args(1)(1, 1) ^ args(2)(1, 1))
    Else
        Utils_Fail "operator ^: input must be scalars" & r1 & " " & c1 & " " & r2 & " " & c2 & " "
    End If
End Function

' Matches operator .^
Private Function op_power(args As Variant) As Variant
    Dim out, r As Long, c As Long, i As Long, j As Long
    Dim arg1_r As Long, arg1_c As Long, arg2_r As Long, arg2_c As Long
    Utils_SetupBinaryOperation args, out, r, c, arg1_r, arg1_c, arg2_r, arg2_c
    For i = 1 To r
        For j = 1 To c
            out(i, j) = args(1)(MIN(i, arg1_r), MIN(j, arg1_c)) _
                      ^ args(2)(MIN(i, arg2_r), MIN(j, arg2_c))
        Next j
    Next i
    op_power = out
End Function

' Matches postfix unary operator '
Private Function op_transpose(args As Variant) As Variant
    args(1) = calc_tree(args(1))
    If IsEmpty(args(1)) Then Exit Function
    op_transpose = WorksheetFunction.Transpose(args(1))
    Utils_Conform op_transpose
End Function

' Matches operator #
Private Function op_numel(args As Variant) As Variant
    op_numel = Utils_ToMatrix(Utils_Numel(calc_tree(args(1))))
End Function


'*****************
'*** FUNCTIONS ***
'*****************

' b = isempty(A)
'
' Returns true if A equals the empty matrix [].
Private Function fn_isempty(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    fn_isempty = Utils_ToMatrix(IsEmpty(args(1)))
End Function

' b = isequal(A,B)
'
' Returns true if A and B are equal
Private Function fn_isequal(args As Variant) As Variant
    Utils_AssertArgsCount args, 2, 2
    Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
    Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
    If r1 <> r2 Or c1 <> c2 Then
        fn_isequal = Utils_ToMatrix(False)
        Exit Function
    End If
    Dim i As Long, j As Long
    For i = 1 To r1
        For j = 1 To c1
            If args(1)(i, j) <> args(2)(i, j) Then
                fn_isequal = Utils_ToMatrix(False)
                Exit Function
            End If
        Next j
    Next i
    fn_isequal = Utils_ToMatrix(True)
End Function

' X = islogical(A)
'
' X = islogical(A) returns a matrix with the same size as A
' indicating what elements are numeric.
Private Function fn_islogical(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim r As Long, c As Long, i As Long, j As Long
    Utils_Size args(1), r, c
    For i = 1 To r
        For j = 1 To c
            args(1)(i, j) = WorksheetFunction.IsLogical(args(1)(i, j))
        Next j
    Next i
    fn_islogical = args(1)
End Function

' X = isnum(A)
'
' X = isnum(A) returns a matrix with the same size as A
' indicating what elements are numeric.
Private Function fn_isnum(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim r As Long, c As Long, i As Long, j As Long
    Utils_Size args(1), r, c
    For i = 1 To r
        For j = 1 To c
            args(1)(i, j) = IsNumeric(args(1)(i, j))
        Next j
    Next i
    fn_isnum = args(1)
End Function

' X = iserror(A)
'
' X = iserror(A) returns a matrix same size as A indicating
' whether what entries are errors.
Private Function fn_iserror(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim r As Long, c As Long, i As Long, j As Long
    Utils_Size args(1), r, c
    For i = 1 To r
        For j = 1 To c
            args(1)(i, j) = IsError(args(1)(i, j))
        Next j
    Next i
    fn_iserror = args(1)
End Function

' I = find(X)
' I = find(X,k)
' I = find(X,k,"first")
' I = find(X,k,"last")
'
' I = find(A) locates all nonzero elements of array X, and returns the linear indices
' of those elements in vector I. If X is a row vector, then I is a row vector;
' otherwise, I is a column vector. If X contains no nonzero elements or is an empty array,
' then I is an empty array.
Private Function fn_find(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 3
    If IsEmpty(args(1)) Then Exit Function
    Dim rows As Long, cols As Long
    Utils_Size args(1), rows, cols
    Dim counter As Long, i As Long, j As Long, stepsize As Long, numElements As Long
    For i = 1 To rows
        For j = 1 To cols
            If CDbl(args(1)(i, j)) <> 0 Then counter = counter + 1
        Next j
    Next i
    If UBound(args) >= 2 Then counter = MIN(counter, args(2))
    If counter <= 0 Then Exit Function
    stepsize = IFF(Utils_IsFlagSet(args, "last"), -1, 1)
    Dim isRowVec As Long: isRowVec = -(rows = 1)
    Dim r: ReDim r(isRowVec + (1 - isRowVec) * counter, 1 - isRowVec + isRowVec * counter)
    If stepsize > 0 Then counter = 1
    For j = IFF(stepsize > 0, 1, cols) To IFF(stepsize > 0, cols, 1) Step stepsize
        For i = IFF(stepsize > 0, 1, rows) To IFF(stepsize > 0, rows, 1) Step stepsize
            If CDbl(args(1)(i, j)) <> 0 Then
                r(isRowVec + (1 - isRowVec) * counter, 1 - isRowVec + isRowVec * counter) _
                    = (j - 1) * rows + i
                counter = counter + stepsize
                If counter < 1 Or counter > UBound(r, 1) * UBound(r, 2) Then GoTo found_all
            End If
        Next i
    Next j
found_all:
    fn_find = r
End Function

' X = fix(A)
'
' X = fix(A) rounds the elements of A towards 0.
Private Function fn_fix(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim r As Long, c As Long, i As Long, j As Long
    Utils_Size args(1), r, c
    With WorksheetFunction
        For i = 1 To r
            For j = 1 To c
                args(1)(i, j) = .RoundDown(args(1)(i, j), 0)
            Next j
        Next i
    End With
    fn_fix = args(1)
End Function

' X = round(A)
' X = round(A,k)
'
' X = round(A) rounds the elements of A.
' X = round(A,k) rounds the elements of A with k decimal places. Default is 0
Private Function fn_round(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    Dim k As Long: k = Utils_GetOptionalScalarArg(args, 2, 0)
    Dim r As Long, c As Long, i As Long, j As Long
    Utils_Size args(1), r, c
    With WorksheetFunction
        For i = 1 To r
            For j = 1 To c
                args(1)(i, j) = .Round(args(1)(i, j), k)
            Next j
        Next i
    End With
    fn_round = args(1)
End Function

' X = ceil(A)
'
' X = ceil(A) rounds each element of A to the nearest integer greater than
' or equal to that element.
Private Function fn_ceil(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim r As Long, c As Long, i As Long, j As Long
    Utils_Size args(1), r, c
    With WorksheetFunction
        For i = 1 To r
            For j = 1 To c
                args(1)(i, j) = .Ceiling(args(1)(i, j), 1)
            Next j
        Next i
    End With
    fn_ceil = args(1)
End Function

' X = floor(A)
'
' X = floor(A) rounds each element of A to the nearest integer smaller than
' or equal to that element.
Private Function fn_floor(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim r As Long, c As Long, i As Long, j As Long
    Utils_Size args(1), r, c
    With WorksheetFunction
        For i = 1 To r
            For j = 1 To c
                args(1)(i, j) = .Floor(args(1)(i, j), 1)
            Next j
        Next i
    End With
    fn_floor = args(1)
End Function

' X = inv(A)
'
' X = inv(A) inverts the matrix A.
' A must be a quadratic matrix
Private Function fn_inv(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Utils_Assert UBound(args(1), 1) = UBound(args(1), 2), "matrix not quadratic"
    fn_inv = WorksheetFunction.MInverse(args(1))
    Utils_Conform fn_inv
End Function

' X = exp(A)
'
' X = exp(A) returns the exponential of the elements of A.
Private Function fn_exp(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim r As Long, c As Long, i As Long, j As Long
    Utils_Size args(1), r, c
    For i = 1 To r
        For j = 1 To c
            args(1)(i, j) = Exp(args(1)(i, j))
        Next j
    Next i
    fn_exp = args(1)
End Function

' X = log(A)
'
' X = log(A) returns the natural logarithm of the elements of A.
Private Function fn_log(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim r As Long, c As Long, i As Long, j As Long
    Utils_Size args(1), r, c
    For i = 1 To r
        For j = 1 To c
            args(1)(i, j) = Log(args(1)(i, j))
        Next j
    Next i
    fn_log = args(1)
End Function

' X = sqrt(A)
'
' X = sqrt(A) returns the square root of the elements of A.
Private Function fn_sqrt(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim r As Long, c As Long, i As Long, j As Long
    Utils_Size args(1), r, c
    For i = 1 To r
        For j = 1 To c
            args(1)(i, j) = Sqr(args(1)(i, j))
        Next j
    Next i
    fn_sqrt = args(1)
End Function

' X = fact(A)
'
' X = fact(A) returns the factorial of the elements of A.
Private Function fn_fact(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim r As Long, c As Long, i As Long, j As Long
    Utils_Size args(1), r, c
    For i = 1 To r
        For j = 1 To c
            args(1)(i, j) = WorksheetFunction.Fact(args(1)(i, j))
        Next j
    Next i
    fn_fact = args(1)
End Function

' r = rows(A)
'
' r = rows(A) returns the number of rows in A.
Private Function fn_rows(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    fn_rows = Utils_ToMatrix(Utils_Rows(args(1)))
End Function

' c = cols(A)
'
' c = cols(A) returns the number of columns in A.
Private Function fn_cols(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    fn_cols = Utils_ToMatrix(Utils_Cols(args(1)))
End Function

' n = numel(A)
'
' n = numel(A) returns the number of elements in A.
Private Function fn_numel(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    fn_numel = Utils_ToMatrix(Utils_Numel(args(1)))
End Function

' X = zeros
' X = zeros(n)
' X = zeros(n,m)
' X = zeros([n m])
'
' zeros(...) returns a matrix of zeros.
Private Function fn_zeros(args As Variant) As Variant
    Utils_AssertArgsCount args, 0, 2
    Dim n As Long, m As Long
    Utils_GetSizeFromArgs args, n, m, 1
    Utils_RepScalar fn_zeros, 0, n, m
End Function

' X = ones
' X = ones(n)
' X = ones(m,n)
' X = ones([m n])
'
' ones(...) returns a matrix of ones.
Private Function fn_ones(args As Variant) As Variant
    Utils_AssertArgsCount args, 0, 2
    Dim n As Long, m As Long
    Utils_GetSizeFromArgs args, n, m, 1
    Utils_RepScalar fn_ones, 1, n, m
End Function

' X = eye
' X = eye(n)
' X = eye(n,m)
' X = eye([n m])
'
' eye(...) returns the identity matrix with ones on the diagonal and
' zeros elsewhere.
Private Function fn_eye(args As Variant) As Variant
    Utils_AssertArgsCount args, 0, 2
    Dim n As Long, m As Long
    Utils_GetSizeFromArgs args, n, m, 1
    If n < 1 Or m < 1 Then Exit Function
    Dim r: ReDim r(n, m)
    For n = 1 To UBound(r, 1)
        For m = 1 To UBound(r, 2)
            r(n, m) = -CLng(n = m)
        Next m
    Next n
    fn_eye = r
End Function

' X = true
' X = true(n)
' X = true(m,n)
' X = true([m n])
'
' true(...) returns a matrix of true's.
Private Function fn_true(args As Variant) As Variant
    Utils_AssertArgsCount args, 0, 2
    Dim n As Long, m As Long
    Utils_GetSizeFromArgs args, n, m, 1
    Utils_RepScalar fn_true, True, n, m
End Function

' X = false
' X = false(n)
' X = false(m,n)
' X = false([m n])
'
' false(...) returns a matrix of false's.
Private Function fn_false(args As Variant) As Variant
    Utils_AssertArgsCount args, 0, 2
    Dim n As Long, m As Long
    Utils_GetSizeFromArgs args, n, m, 1
    Utils_RepScalar fn_false, False, n, m
End Function

' X = linspace(x1,x2)
' X = linspace(x1,x2,n)
'
' linspace(x1,x2) returns a row vector of 100 evenly spaced points between x1 and x2
' linspace(x1,x2,n) returns n points with spacing (x2-x1)/(n-1).
Private Function fn_linspace(args As Variant) As Variant
    Utils_AssertArgsCount args, 2, 3
    Dim out, n As Long, i As Long, x1 As Double, x2 As Double
    x1 = args(1)(1, 1)
    x2 = args(2)(1, 1)
    n = Utils_GetOptionalScalarArg(args, 3, 100)
    ReDim out(1 To 1, 1 To n)
    For i = 1 To n
        out(1, i) = x1 + (x2 - x1) / (n - 1) * (i - 1)
    Next i
    fn_linspace = out
End Function

' X = xor(A,B)
Private Function fn_xor(args As Variant) As Variant
    Utils_AssertArgsCount args, 2, 2
    Dim out, r As Long, c As Long, i As Long, j As Long
    Dim arg1_r As Long, arg1_c As Long, arg2_r As Long, arg2_c As Long
    Utils_SetupBinaryOperation args, out, r, c, arg1_r, arg1_c, arg2_r, arg2_c, False
    For i = 1 To r
        For j = 1 To c
            out(i, j) = CBool(args(1)(MIN(i, arg1_r), MIN(j, arg1_c))) _
                    Xor CBool(args(2)(MIN(i, arg2_r), MIN(j, arg2_c)))
        Next j
    Next i
    fn_xor = out
End Function

' X = tick2ret(A)
' X = tick2ret(A,dim)
'
' Returns simple percentage returns
Private Function fn_tick2ret(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    Dim x As Long: x = Utils_CalcDimDirection(args, 2)
    Dim r: ReDim r(UBound(args(1), 1) - (1 - x), UBound(args(1), 2) - x)
    Dim i As Long, j As Long
    For i = 1 To UBound(r, 1)
        For j = 1 To UBound(r, 2)
            r(i, j) = args(1)(i + (1 - x), j + x) / args(1)(i, j) - 1
        Next j
    Next i
    fn_tick2ret = r
End Function

' X = ret2tick(A)
' X = ret2tick(A,dim)
'
' Creates timeseries from simple percentage returns.
' New base is 100
Private Function fn_ret2tick(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    Dim x As Long: x = Utils_CalcDimDirection(args, 2)
    Dim r: ReDim r(UBound(args(1), 1) + (1 - x), UBound(args(1), 2) + x)
    Dim i As Long, j As Long
    For j = 1 To UBound(r, 2 - x)
        r(1 * (1 - x) + j * x, j * (1 - x) + 1 * x) = 100
    Next j
    For i = 2 - x To UBound(r, 1)
        For j = 1 + x To UBound(r, 2)
            r(i, j) = r(i - (1 - x), j - x) * (1 + args(1)(i - (1 - x), j - x))
        Next j
    Next i
    fn_ret2tick = r
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
    Utils_AssertArgsCount args, 1, 2
    If IsEmpty(args(1)) Then fn_cumsum = Utils_ToMatrix(0): Exit Function
    Dim i As Long, j As Long, x As Long
    x = Utils_CalcDimDirection(args)
    For i = 2 - x To UBound(args(1), 1)
        For j = 1 + x To UBound(args(1), 2)
            args(1)(i, j) = args(1)(i, j) + args(1)(i - (1 - x), j - x)
        Next j
    Next i
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
' along the first dimension (the columns); cumprod(A,2) works along the
' second dimension (the rows).
Private Function fn_cumprod(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    If IsEmpty(args(1)) Then fn_cumprod = Utils_ToMatrix(1): Exit Function
    Dim i As Long, j As Long, x As Long
    x = Utils_CalcDimDirection(args)
    For i = 2 - x To UBound(args(1), 1)
        For j = 1 + x To UBound(args(1), 2)
            args(1)(i, j) = args(1)(i, j) * args(1)(i - (1 - x), j - x)
        Next j
    Next i
    fn_cumprod = args(1)
End Function

' B = cummax(A)
' B = cummax(A,dim)
Private Function fn_cummax(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    If IsEmpty(args(1)) Then fn_cummax = Utils_ToMatrix(DOUBLE_MIN): Exit Function
    Dim i As Long, j As Long, x As Long
    x = Utils_CalcDimDirection(args)
    For i = 2 - x To UBound(args(1), 1)
        For j = 1 + x To UBound(args(1), 2)
            args(1)(i, j) = MAX(args(1)(i, j), args(1)(i - (1 - x), j - x))
        Next j
    Next i
    fn_cummax = args(1)
End Function

' B = cummin(A)
' B = cummin(A,dim)
Private Function fn_cummin(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    If IsEmpty(args(1)) Then fn_cummin = Utils_ToMatrix(DOUBLE_MAX): Exit Function
    Dim i As Long, j As Long, x As Long
    x = Utils_CalcDimDirection(args)
    For i = 2 - x To UBound(args(1), 1)
        For j = 1 + x To UBound(args(1), 2)
            args(1)(i, j) = MIN(args(1)(i, j), args(1)(i - (1 - x), j - x))
        Next j
    Next i
    fn_cummin = args(1)
End Function

' X = std(A)
' X = std(A,dim)
'
' X = std(A) returns the standard deviation of the elements of A
' along the first array dimension whose size does not equal 1.
'
' X = std(A,dim) calculates the elements of A along dimension dim.
' The dim input is a positive integer scalar.
Private Function fn_std(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    If IsEmpty(args(1)) Then Exit Function
    Dim x As Long, i As Long, out As Variant
    Utils_SetupReducedDimOperation args, out, x
    With WorksheetFunction
        For i = 1 To UBound(out, 2 - x)
            out(x * i + (1 - x), (1 - x) * i + x) = .StDev(.index(args(1), x * i, (1 - x) * i))
        Next i
    End With
    fn_std = out
End Function

' X = var(A)
' X = var(A,dim)
'
' X = var(A) returns the variance of the elements of A
' along the first array dimension whose size does not equal 1.
'
' X = var(A,dim) calculates the elements of A along dimension dim.
' The dim input is a positive integer scalar.
Private Function fn_var(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    If IsEmpty(args(1)) Then Exit Function
    Dim x As Long, i As Long, out As Variant
    Utils_SetupReducedDimOperation args, out, x
    With WorksheetFunction
        For i = 1 To UBound(out, 2 - x)
            out(x * i + (1 - x), (1 - x) * i + x) = .Var(.index(args(1), x * i, (1 - x) * i))
        Next i
    End With
    fn_var = out
End Function

' X = corr(A)
'
' X = corr(A) returns a correlation matrix for the columns of A.
Private Function fn_corr(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    If IsEmpty(args(1)) Then Exit Function
    Utils_Assert UBound(args(1), 1) > 1, "too few rows"
    Dim c As Long: c = UBound(args(1), 2)
    Dim out: ReDim out(c, c)
    Dim i As Long, j As Long
    With WorksheetFunction
        For i = 1 To c
            For j = i To c
                out(i, j) = .Correl(.index(args(1), 0, i), .index(args(1), 0, j))
                out(j, i) = out(i, j)
            Next j
        Next i
    End With
    fn_corr = out
End Function

' X = cov(A)
'
' X = cov(A) returns a covariance matrix for the columns of A.
Private Function fn_cov(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    If IsEmpty(args(1)) Then Exit Function
    Dim c As Long: c = UBound(args(1), 2)
    Dim out: ReDim out(c, c)
    Dim i As Long, j As Long
    With WorksheetFunction
        For i = 1 To c
            For j = i To c
                out(i, j) = .Covar(.index(args(1), 0, i), .index(args(1), 0, j))
                out(j, i) = out(i, j)
            Next j
        Next i
    End With
    fn_cov = out
End Function

' X = all(A)
' X = all(A,dim)
'
' all(...) tests if all elements in A evaluates to true.
' In practice, all() is a natural extension of the logical AND
' operator.
Private Function fn_all(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    If IsEmpty(args(1)) Then fn_all = Utils_ToMatrix(True): Exit Function
    Dim x As Long, i As Long, r As Variant
    Utils_SetupReducedDimOperation args, r, x
    With WorksheetFunction
        For i = 1 To UBound(r, 2 - x)
            r(x * i + (1 - x), (1 - x) * i + x) = .And(.index(args(1), x * i, (1 - x) * i))
        Next i
    End With
    fn_all = r
End Function

' X = any(A)
' X = any(A,dim)
'
' any(...) tests if any of the elements in A evaluates to true.
' In practice, any() is a natural extension of the logical OR
' operator.
Private Function fn_any(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    If IsEmpty(args(1)) Then fn_any = Utils_ToMatrix(False): Exit Function
    Dim x As Long, i As Long, r As Variant
    Utils_SetupReducedDimOperation args, r, x
    With WorksheetFunction
        For i = 1 To UBound(r, 2 - x)
            r(x * i + (1 - x), (1 - x) * i + x) = .Or(.index(args(1), x * i, (1 - x) * i))
        Next i
    End With
    fn_any = r
End Function

' b = binom(n,k)
Private Function fn_binom(args As Variant) As Variant
    Utils_AssertArgsCount args, 2, 2
    Dim n As Long, k As Long
    n = args(1)(1, 1)
    k = args(2)(1, 1)
    fn_binom = Utils_ToMatrix(WorksheetFunction.Fact(n) / (WorksheetFunction.Fact(k) * WorksheetFunction.Fact(n - k)))
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
    Utils_AssertArgsCount args, 1, 2
    If IsEmpty(args(1)) Then fn_sum = Utils_ToMatrix(0): Exit Function
    Dim x As Long, i As Long, r As Variant
    Utils_SetupReducedDimOperation args, r, x
    With WorksheetFunction
        For i = 1 To UBound(r, 2 - x)
            r(x * i + (1 - x), (1 - x) * i + x) = .Sum(.index(args(1), x * i, (1 - x) * i))
        Next i
    End With
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
    Utils_AssertArgsCount args, 1, 2
    If IsEmpty(args(1)) Then fn_prod = Utils_ToMatrix(1): Exit Function
    Dim x As Long, i As Long, r As Variant
    Utils_SetupReducedDimOperation args, r, x
    With WorksheetFunction
        For i = 1 To UBound(r, 2 - x)
            r(x * i + (1 - x), (1 - x) * i + x) = .Product(.index(args(1), x * i, (1 - x) * i))
        Next i
    End With
    fn_prod = r
End Function

' X = mean(A)
' X = mean(A,dim)
Private Function fn_mean(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    Dim x As Long, i As Long, r As Variant
    Utils_SetupReducedDimOperation args, r, x
    With WorksheetFunction
        For i = 1 To UBound(r, 2 - x)
            r(x * i + (1 - x), (1 - x) * i + x) = .Average(.index(args(1), x * i, (1 - x) * i))
        Next i
    End With
    fn_mean = r
End Function

' X = median(A)
' X = median(A,dim)
Private Function fn_median(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    Dim x As Long, i As Long, r As Variant
    Utils_SetupReducedDimOperation args, r, x
    With WorksheetFunction
        For i = 1 To UBound(r, 2 - x)
            r(x * i + (1 - x), (1 - x) * i + x) = .Median(.index(args(1), x * i, (1 - x) * i))
        Next i
    End With
    fn_median = r
End Function

' X = prctile(A,p)
' X = prctile(A,p,dim)
'
' p must be a number between 0 and 1
Private Function fn_prctile(args As Variant) As Variant
    Utils_AssertArgsCount args, 2, 3
    Dim x As Long, i As Long, r As Variant
    Utils_SetupReducedDimOperation args, r, x, 3
    With WorksheetFunction
        For i = 1 To UBound(r, 2 - x)
            r(x * i + (1 - x), (1 - x) * i + x) = .Percentile(.index(args(1), x * i, (1 - x) * i), args(2)(1, 1))
        Next i
    End With
    fn_prctile = r
End Function

' X = percentrank(A,x)
' X = percentrank(A,x,dim)
'
' The inverse of the prctile
Private Function fn_percentrank(args As Variant) As Variant
    Utils_AssertArgsCount args, 2, 3
    Dim x As Long, i As Long, r As Variant
    Utils_SetupReducedDimOperation args, r, x, 3
    With WorksheetFunction
        For i = 1 To UBound(r, 2 - x)
            r(x * i + (1 - x), (1 - x) * i + x) = .PercentRank(.index(args(1), x * i, (1 - x) * i), args(2)(1, 1))
        Next i
    End With
    fn_percentrank = r
End Function

' X = count(A)
' X = count(A,dim)
'
' X = count(A) counts the number of elements in A which do not evaluate to false
Private Function fn_count(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    If IsEmpty(args(1)) Then fn_count = 0: Exit Function
    Dim x As Long, i As Long, j As Long, r As Variant
    Utils_SetupReducedDimOperation args, r, x
    For i = 1 To UBound(args(1), 1)
        For j = 1 To UBound(args(1), 2)
            r(i * x + (1 - x), j * (1 - x) + x) _
                = r(i * x + (1 - x), j * (1 - x) + x) - CBool(args(1)(i, j))
        Next j
    Next i
    fn_count = r
End Function

' X = counta(A)
' X = counta(A,dim)
'
' X = counta(A) counts the number of elements in A which are not empty
Private Function fn_counta(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    If IsEmpty(args(1)) Then fn_counta = 0: Exit Function
    Dim x As Long, i As Long, j As Long, r As Variant
    Utils_SetupReducedDimOperation args, r, x
    For i = 1 To UBound(args(1), 1)
        For j = 1 To UBound(args(1), 2)
            r(i * x + (1 - x), j * (1 - x) + x) _
                = r(i * x + (1 - x), j * (1 - x) + x) - CBool(Not IsEmpty(args(1)(i, j)))
        Next j
    Next i
    fn_counta = r
End Function

' M = max(A)
' M = max(A,[],dim)
' M = max(A,B)
Private Function fn_max(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 3
    Dim r As Variant, x As Long, y As Long, i As Long
    If UBound(args) = 1 Or UBound(args) = 3 Then
        If UBound(args) = 3 Then Utils_Assert IsEmpty(args(2)), "2nd argument must be empty matrix, []."
        If IsEmpty(args(1)) Then fn_max = Utils_ToMatrix(DOUBLE_MIN): Exit Function
        Utils_SetupReducedDimOperation args, r, x, 3
        With WorksheetFunction
            For i = 1 To UBound(r, 2 - x)
                r(x * i + (1 - x), (1 - x) * i + x) _
                    = .MAX(.index(args(1), x * i, (1 - x) * i))
            Next i
        End With
    Else
        Dim r1 As Long, c1 As Long
        Dim r2 As Long, c2 As Long
        Utils_Size args(1), r1, c1
        Utils_Size args(2), r2, c2
        Utils_Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2), "max(): Wrong dimensions."
        ReDim r(MAX(r1, r2), MAX(c1, c2))
        For x = 1 To UBound(r, 1)
            For y = 1 To UBound(r, 2)
                r(x, y) = MAX(args(1)(MIN(x, r1), MIN(y, c1)), args(2)(MIN(x, r2), MIN(y, c2)))
            Next y
        Next x
    End If
    fn_max = r
End Function

' M = min(A)
' M = min(A,[],dim)
' M = min(A,B)
Private Function fn_min(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 3
    Dim r As Variant, x As Long, y As Long, i As Long
    If UBound(args) = 1 Or UBound(args) = 3 Then
        If UBound(args) = 3 Then Utils_Assert IsEmpty(args(2)), "2nd argument must be empty matrix, []."
        If IsEmpty(args(1)) Then fn_min = Utils_ToMatrix(DOUBLE_MAX): Exit Function
        Utils_SetupReducedDimOperation args, r, x, 3
        With WorksheetFunction
            For i = 1 To UBound(r, 2 - x)
                r(x * i + (1 - x), (1 - x) * i + x) _
                    = .MIN(.index(args(1), x * i, (1 - x) * i))
            Next i
        End With
    Else
        Dim r1 As Long, c1 As Long
        Dim r2 As Long, c2 As Long
        Utils_Size args(1), r1, c1
        Utils_Size args(2), r2, c2
        Utils_Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2), "min: bad dimensions."
        ReDim r(MAX(r1, r2), MAX(c1, c2))
        For x = 1 To UBound(r, 1)
            For y = 1 To UBound(r, 2)
                r(x, y) = MIN(args(1)(MIN(x, r1), MIN(y, c1)), args(2)(MIN(x, r2), MIN(y, c2)))
            Next y
        Next x
    End If
    fn_min = r
End Function

' X = size(A)
'
' X = size(A) returns a 1-by-2 vector with the number of rows and columns in A.
Private Function fn_size(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim out: ReDim out(1, 2)
    Utils_Size args(1), out(1, 1), out(1, 2)
    fn_size = out
End Function

' X = diag(A)
'
' X = diag(A) returns a matrix with A in the diagonal if A is a vector,
' or a column vector with the diagonal of A if A is a matrix.
Private Function fn_diag(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim rows As Long, cols As Long, r As Variant, i As Long
    Utils_Size args(1), rows, cols
    If Utils_IsVectorShape(rows, cols) Then
        r = fn_repmat(Array(0, rows * cols, rows * cols))
        For i = 1 To UBound(r, 1)
            r(i, i) = args(1)(MIN(i, rows), MIN(i, cols))
        Next i
    Else
        ReDim r(MIN(rows, cols), 1)
        For i = 1 To UBound(r, 1)
            r(i, 1) = args(1)(i, i)
        Next i
    End If
    fn_diag = r
End Function

' X = rand
' X = rand(n)
' X = rand(n,m)
' X = rand([n m])
'
' rand(...) returns pseudorandom values drawn from the standard
' uniform distribution on the open interval (0,1)
Private Function fn_rand(args As Variant) As Variant
    Utils_AssertArgsCount args, 0, 2
    Dim n As Long, m As Long
    Utils_GetSizeFromArgs args, n, m, 1
    If n < 1 Or m < 1 Then Exit Function
    Dim out: ReDim out(n, m)
    For n = 1 To UBound(out, 1)
        For m = 1 To UBound(out, 2)
            out(n, m) = Rnd
        Next m
    Next n
    fn_rand = out
End Function

' X = randi(imax)
' X = randi(imax,n)
' X = randi(imax,n,m)
' X = randi(imax,[n m])
' X = randi([imin imax], ...)
'
' randi(...) returns pseudorandom integer values drawn from the
' discrete uniform distribution on the interval [1, imax] or
' [imin, imax].
Private Function fn_randi(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 3
    Dim imin As Long, imax As Long, n As Long, m As Long
    Utils_GetSizeFromArgs args, n, m
    If n < 1 Or m < 1 Then Exit Function
    If Utils_Numel(args(1)) = 1 Then
        imin = 1
        imax = args(1)
    Else
        imin = args(1)(1, 1)
        imax = args(1)(MIN(2, UBound(args(1), 1)), MIN(2, UBound(args(1), 2)))
    End If
    Dim out: ReDim out(n, m)
    For n = 1 To UBound(out, 1)
        For m = 1 To UBound(out, 2)
            out(n, m) = CLng(Rnd * (imax - imin)) + imin
        Next m
    Next n
    fn_randi = out
End Function

' X = randn
' X = randn(n)
' X = randn(n,m)
' X = randn([n m])
'
' randn(...) returns pseudorandom values drawn from the standard
' normal distribution with mean 0 and variance 1.
Private Function fn_randn(args As Variant) As Variant
    Utils_AssertArgsCount args, 0, 2
    Dim n As Long, m As Long
    Utils_GetSizeFromArgs args, n, m, 1
    If n < 1 Or m < 1 Then Exit Function
    Dim r: ReDim r(n, m)
    Dim p(2) As Double, tmp As Double, cached As Boolean
    For n = 1 To UBound(r, 1)
        For m = 1 To UBound(r, 2)
            If Not cached Then
                Do
                    p(1) = 2 * Rnd - 1
                    p(2) = 2 * Rnd - 1
                    tmp = p(1) * p(1) + p(2) * p(2)
                Loop Until tmp <= 1
                tmp = Sqr(-2 * Log(tmp) / tmp)
                p(1) = p(1) * tmp
                p(2) = p(2) * tmp
            End If
            cached = Not cached
            r(n, m) = p(2 + cached)
        Next m
    Next n
    fn_randn = r
End Function

' X = normcdf(A)
Private Function fn_normcdf(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim r As Long, c As Long, i As Long, j As Long
    Utils_Size args(1), r, c
    With WorksheetFunction
        For i = 1 To r
            For j = 1 To c
                args(1)(i, j) = .NormSDist(args(1)(i, j))
            Next j
        Next i
    End With
    fn_normcdf = args(1)
End Function

' X = repmat(A,n)
' X = repmat(A,n,m)
' X = repmat(A,[n m])
'
' X = repmat(A,n) creates a large matrix X consisting of an n-by-n tiling of A.
' X = repmat(A,n,m) creates a large matrix X consisting of an n-by-m tiling of A.
' X = repmat(A,[n m]) creates a large matrix X consisting of an n-by-m tiling of A.
Private Function fn_repmat(args As Variant) As Variant
    Dim out, r As Long, c As Long, n As Long, m As Long, i As Long, j As Long
    Utils_AssertArgsCount args, 2, 3
    Utils_GetSizeFromArgs args, n, m
    Utils_Size args(1), r, c
    If r * n < 1 Or c * m < 1 Then Exit Function
    ReDim out(r * n, c * m)
    For n = 0 To n - 1
        For m = 0 To m - 1
            For i = 1 To r
                For j = 1 To c
                    out(n * r + i, m * c + j) = args(1)(i, j)
                Next j
            Next i
        Next m
    Next n
    fn_repmat = out
End Function

' B = reshape(A,n,m)
' B = reshape(A,[],m)
' B = reshape(A,n,[])
'
' B = reshape(A,n,m) returns the n-by-m matrix B whose elements are taken
' column-wise from A. An error results if A does not have n*m elements.
' Either n or m can be the empty matrix [] in which case the length of the
' dimension is calculated automatically.
Private Function fn_reshape(args As Variant) As Variant
    Utils_AssertArgsCount args, 3, 3
    Dim rows As Long, cols As Long, idx As Long
    Dim r_i As Long, r_j As Long, arg_i As Long, arg_j As Long
    Utils_Size args(1), rows, cols
    If IsEmpty(args(2)) Then
        args(3) = args(3)(1, 1)
        If args(3) = 0 Then Exit Function
        Utils_Assert rows * cols Mod args(3) = 0, "reshape: number of elements not evenly divisible by m"
        args(2) = rows * cols / args(3)
    ElseIf IsEmpty(args(3)) Then
        args(2) = args(2)(1, 1)
        If args(2) = 0 Then Exit Function
        Utils_Assert rows * cols Mod args(2) = 0, "reshape: number of elements not evenly divisible by n"
        args(3) = rows * cols / args(2)
    Else
        args(2) = args(2)(1, 1)
        args(3) = args(3)(1, 1)
        Utils_Assert rows * cols = args(2) * args(3), "reshape: number of elements is not equal to n*m"
    End If
    Utils_Assert args(2) >= 0 And args(3) >= 0, "new size must be non-negative"
    If args(2) = 0 And args(3) = 0 Then Exit Function
    Dim r: ReDim r(args(2), args(3))
    For idx = 1 To rows * cols
        Utils_Ind2Sub rows, idx, arg_i, arg_j
        Utils_Ind2Sub CLng(args(2)), idx, r_i, r_j
        r(r_i, r_j) = args(1)(arg_i, arg_j)
    Next idx
    fn_reshape = r
End Function

' M = hcat(A, B, C, ...)
'
' hcat concatenates scalars and matrices horizontally
' Parameters must either be scalars or matrices with the same number of rows
' Scalars are expanded to a column vector with the correct number of identical entries
' hcat(A,B,C) is equivalent to [ A B C ]
Private Function fn_hcat(args As Variant) As Variant
    Dim arg_count As Long
    arg_count = Utils_Stack_Size(args)
    If arg_count = 0 Then Exit Function
    If arg_count = 1 Then
        fn_hcat = args(1)
        Exit Function
    End If
    Dim rows As Long, cols As Long, i As Long, r As Long, c As Long
    For i = 1 To arg_count
        Utils_Size args(i), r, c
        Utils_Assert rows = 0 Or rows = 1 Or r = 0 Or r = 1 Or rows = r, "hcat: different row counts"
        rows = MAX(rows, r)
        cols = cols + c
    Next i
    If rows = 0 Or cols = 0 Then Exit Function
    Dim out, n As Long, m As Long
    ReDim out(1 To rows, 1 To cols)
    cols = 0
    For i = 1 To arg_count
        Utils_Size args(i), r, c
        For n = 1 To rows
            For m = 1 To c
                out(n, cols + m) = args(i)(MIN(n, r), m)
            Next m
        Next n
        cols = cols + c
    Next i
    fn_hcat = out
End Function

' M = vcat(A, B, C, ...)
'
' hcat concatenates scalars and matrices vertically
' Parameters must either be scalars or matrices with the same number of columns
' Scalars are expanded to a row vector with the correct number of identical entries
' vcat(A,B,C) is equivalent to [ A; B; C ]
Private Function fn_vcat(args As Variant) As Variant
    Dim arg_count As Long
    arg_count = Utils_Stack_Size(args)
    If arg_count = 0 Then Exit Function
    If arg_count = 1 Then
        fn_vcat = args(1)
        Exit Function
    End If
    Dim rows As Long, cols As Long, i As Long, r As Long, c As Long
    For i = 1 To arg_count
        Utils_Size args(i), r, c
        Utils_Assert cols = 0 Or cols = 1 Or c = 0 Or c = 1 Or cols = c, "vcat: different column counts"
        cols = MAX(cols, c)
        rows = rows + r
    Next i
    If rows = 0 Or cols = 0 Then Exit Function
    Dim out, n As Long, m As Long
    ReDim out(1 To rows, 1 To cols)
    rows = 0
    For i = 1 To arg_count
        Utils_Size args(i), r, c
        For n = 1 To r
            For m = 1 To cols
                out(rows + n, m) = args(i)(n, MIN(m, c))
            Next m
        Next n
        rows = rows + r
    Next i
    fn_vcat = out
End Function

Private Function Utils_Match(lookupvalue, lookuparray, notfoundvalue) As Variant
    On Error GoTo notfound:
    Utils_Match = WorksheetFunction.Match(lookupvalue, lookuparray, 0)
    Exit Function
notfound:
    Utils_Match = notfoundvalue
End Function

' X = match(lookupvalues, lookuparray)
' X = match(lookupvalues, lookuparray, notfoundvalue)
'
' Returns the linear indices for the lookup values found in the lookuparray.
' Will return #N/A for values not found unless a notfoundvalue is specified
Private Function fn_match(args As Variant) As Variant
    Utils_AssertArgsCount args, 2, 3
    Dim notfoundvalue As Variant
    notfoundvalue = Utils_GetOptionalScalarArg(args, 3, [NA()])
    Dim r As Long, c As Long, i As Long, j As Long
    Utils_Size args(1), r, c
    For i = 1 To r
        For j = 1 To c
            args(1)(i, j) = Utils_Match(args(1)(i, j), args(2), notfoundvalue)
        Next j
    Next i
    fn_match = args(1)
End Function

' X = tostring(A)
'
' tostring(A) converts all entries of A into strings
Private Function fn_tostring(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim r As Long, c As Long, i As Long, j As Long
    Utils_Size args(1), r, c
    For i = 1 To r
        For j = 1 To c
            args(1)(i, j) = args(1)(i, j) & ""
        Next j
    Next i
    fn_tostring = args(1)
End Function

' Z = if(B,X,Y)
'
' Z = if(B,X,Y) returns X if B evaluates to true; otherwise Y.
' Note: X is only evaluated if B is true and Y is only evaluated if B is false
Private Function fn_if(args As Variant) As Variant
    Utils_AssertArgsCount args, 3, 3
    args(1) = calc_tree(args(1))
    If IsEmpty(args(1)) Then
        fn_if = calc_tree(args(3))
    Else
        fn_if = calc_tree(args(3 + CLng(CBool(args(1)(1, 1)))))
    End If
End Function

' X = iferror(A,B)
'
' X = iferror(A,B) returns A if the evaluation of A does not result in a error;
' otherwise, B is returned.
Private Function fn_iferror(args As Variant) As Variant
    Utils_AssertArgsCount args, 2, 2
    On Error GoTo ErrorHandler:
    fn_iferror = calc_tree(args(1))
    Exit Function
ErrorHandler:
    fn_iferror = calc_tree(args(2))
End Function

' Y = diff(X)
' Y = diff(X,n)
' Y = diff(X,n,dim)
Private Function fn_diff(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 3
    If IsEmpty(args(1)) Then Exit Function
    Dim x As Long: x = Utils_CalcDimDirection(args, 3)
    If UBound(args(1), 1 + x) < 2 Then Exit Function
    Dim i As Long, j As Long
    Dim r: ReDim r(UBound(args(1), 1) - (1 - x), UBound(args(1), 2) - x)
    For i = 2 - x To UBound(args(1), 1)
        For j = 1 + x To UBound(args(1), 2)
            r(i - (1 - x), j - x) = args(1)(i, j) - args(1)(i - (1 - x), j - x)
        Next j
    Next i
    fn_diff = r
    Dim n As Long: n = Utils_GetOptionalScalarArg(args, 2, 1)
    If n > 1 Then fn_diff = fn_diff(Array(r, n - 1, 1 + x))
End Function

' B = unique(A)
'
' B = unique(A) returns a column vector with all the unique elements of A.
' The values of B will be in sorted order.
Private Function fn_unique(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim numel As Long, i As Long, counter As Long, save
    numel = Utils_Numel(args(1))
    If numel < 1 Then Exit Function
    args(1) = fn_reshape(Array(args(1), Empty, Utils_ToMatrix(1)))
    args(1) = fn_sort(Array(args(1)))
    ReDim save(1 To numel - 1)
    counter = 1
    For i = 1 To UBound(save)
        save(i) = (0 <> Utils_Compare(args(1)(i, 1), args(1)(i + 1, 1)))
        counter = counter - CLng(save(i))
    Next i
    Dim r: ReDim r(1 To counter, 1)
    r(1, 1) = args(1)(1, 1)
    counter = 2
    For i = 1 To UBound(save)
        If save(i) Then
            r(counter, 1) = args(1)(i + 1, 1)
            counter = counter + 1
        End If
    Next i
    fn_unique = r
End Function

Private Function fn_droperror(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim numel As Long, i As Long, counter As Long, save
    numel = Utils_Numel(args(1))
    If numel < 1 Then Exit Function
    args(1) = fn_reshape(Array(args(1), Empty, Utils_ToMatrix(1)))
    ReDim save(1 To numel)
    counter = 0
    For i = 1 To UBound(save)
        save(i) = Not IsError(args(1)(i, 1))
        counter = counter - CLng(save(i))
    Next i
    If counter = 0 Then Exit Function
    Dim r: ReDim r(1 To counter, 1)
    counter = 1
    For i = 1 To UBound(save)
        If save(i) Then
            r(counter, 1) = args(1)(i, 1)
            counter = counter + 1
        End If
    Next i
    fn_droperror = r
End Function

' B = sorttable(A, column)
' B = sorttable(A, column, "descend")
Private Function fn_sorttable(args As Variant) As Variant
    Utils_AssertArgsCount args, 2, 3
    
    Dim rows As Long, cols As Long, i As Long, j As Long
    Utils_Size args(1), rows, cols
    
    ' Get all input parameters
    Dim col As Long, ascend As Boolean
    col = args(2)(1, 1)
    ascend = Not Utils_IsFlagSet(args, "descend")
    
    ' Create array containing row indices
    ' Initially, the indices are just 1,2,3,...,n and
    ' then the sorting will be done on these indices.
    Dim indices: ReDim indices(1 To rows, 1 To 1)
    For i = 1 To rows
        indices(i, 1) = i
    Next i
    
    ' Do the actual sorting of the column
    Utils_QuickSortCol args(1), indices, 1, rows, col, 1, ascend
    
    ' Build the table with the sorted values
    Dim out: ReDim out(1 To rows, 1 To cols)
    For i = 1 To rows
        For j = 1 To cols
            out(i, j) = args(1)(indices(i, 1), j)
        Next j
    Next i
    fn_sorttable = out
    Utils_Conform out
End Function

' B = sort(A)
' B = sort(A,dim)
' B = sort(...,"descend")
' B = sort(...,"indices")
'
' sort() sorts the entries in each row or column.
'
' Flags:
' "descend":  Sort descending
' "indices":  Return sorted indices instead of values
Private Function fn_sort(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 4
    
    ' Get all input parameters
    Dim sortRows As Boolean, ascend As Boolean
    sortRows = (1 = Utils_CalcDimDirection(args))
    ascend = Not Utils_IsFlagSet(args, "descend")
    
    ' Transpose input matrix if rows must be sorted
    If sortRows Then
        args(1) = WorksheetFunction.Transpose(args(1))
        Utils_Conform args(1)
    End If
    
    ' Create an equal-sized array containing indices
    ' Initially, the indices are just 1,2,3,...,n and
    ' then the sorting will be done on these indices.
    Dim rows As Long, cols As Long, i As Long, j As Long
    Utils_Size args(1), rows, cols
    Dim indices: ReDim indices(1 To rows, 1 To cols)
    For i = 1 To rows
        For j = 1 To cols
            indices(i, j) = i
        Next j
    Next i
    
    ' Do the actual sorting of each column
    For j = 1 To cols
        Utils_QuickSortCol args(1), indices, 1, rows, j, j, ascend
    Next j
    
    ' Return the sorted indices if that was specified;
    ' otherwise build and return a matrix with the sorted values
    If Utils_IsFlagSet(args, "indices") Then
        fn_sort = indices
    Else
        Dim r: ReDim r(1 To rows, 1 To cols)
        For i = 1 To rows
            For j = 1 To cols
                r(i, j) = args(1)(indices(i, j), j)
            Next j
        Next i
        fn_sort = r
    End If
    
    ' Remember to "transpose back" if we sorted the rows
    If sortRows Then fn_sort = WorksheetFunction.Transpose(fn_sort)
    Utils_Conform fn_sort
End Function

' Implementation of the quick-sort algorithm - is a helper for fn_sort(), fn_sorttable()
' Sorts on a column by swapping indices.
' No actual swapping of values in the original matrix is done.
'
' It sorts the column "col" in the range from "first" to "last"
Private Function Utils_QuickSortCol(arr As Variant, ByRef indices As Variant, first As Long, last As Long, arr_col As Long, indices_col As Long, ascend As Boolean)
    If first >= last Then Exit Function
    Dim tmp As Variant
    Dim pivot As Variant: pivot = arr(indices(first, indices_col), arr_col)
    Dim left As Long: left = first
    Dim right As Long: right = last
    Dim ascendprefix As Long: ascendprefix = -1 - 2 * Sgn(ascend)
    While left <= right
        While ascendprefix * Utils_Compare(arr(indices(left, indices_col), arr_col), pivot) < 0
            left = left + 1
        Wend
        While ascendprefix * Utils_Compare(pivot, arr(indices(right, indices_col), arr_col)) < 0
            right = right - 1
        Wend
        If left <= right Then
            tmp = indices(left, indices_col)
            indices(left, indices_col) = indices(right, indices_col)
            indices(right, indices_col) = tmp
            left = left + 1
            right = right - 1
        End If
    Wend
    Utils_QuickSortCol arr, indices, first, right, arr_col, indices_col, ascend
    Utils_QuickSortCol arr, indices, left, last, arr_col, indices_col, ascend
End Function

' Called from Utils_QuickSortCol. Compares numerics and strings.
Private Function Utils_Compare(arg1 As Variant, arg2 As Variant) As Variant
    If IsNumeric(arg1) Then
        If IsNumeric(arg2) Then
            Utils_Compare = arg1 - arg2
        Else
            Utils_Compare = -1
        End If
    Else
        If IsNumeric(arg2) Then
            Utils_Compare = 1
        Else
            Utils_Compare = StrComp(CStr(arg1), CStr(arg2))
        End If
    End If
End Function

' X = arrayfun(func,A1,...,An)
'
' arrayfun(...) calls the in-cell Excel function with name <func> and passes elements from
' A1 to An, where n is the number of inputs to func.
' A1 to An must all be scalars or matrices with the same size
Private Function fn_arrayfun(args As Variant) As Variant
    Utils_AssertArgsCount args, 2, 100
    Utils_Assert TypeName(args(1)) = "String", "apply: 1st argument must be an Excel function name."
    Dim i As Long, r1 As Long, c1 As Long, r2 As Long, c2 As Long
    r1 = -1: c1 = -1
    For i = 2 To Utils_Stack_Size(args)
        Utils_Size args(i), r2, c2
        Utils_Assert (r1 < 0 And c1 < 0) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2) Or ((r1 = 1 Or r1 = r2) And c2 = 1) Or (r2 = 1 And (c1 = 1 Or c1 = c2)), "apply(): Wrong input sizes."
        r1 = MAX(r1, r2): c1 = MAX(c1, c2)
    Next i
    Dim v, r: ReDim r(r1, c1)
    For r1 = 1 To UBound(r, 1)
        For c1 = 1 To UBound(r, 2)
            v = Empty
            For i = 2 To Utils_Stack_Size(args)
                Utils_Size args(i), r2, c2
                Utils_Stack_Push args(i)(MIN(r1, r2), MIN(c1, c2)), v
            Next i
            r(r1, c1) = Evaluate(args(1) & "(" & Join(v, ",") & ")")
        Next c1
    Next r1
    fn_arrayfun = r
End Function

' B = join(A)
' B = join(A,joiner)
' B = join(A,joiner,dim)
'
' B = concat(...) concatenates the elements of A along the first
' dimension whose size does not equal 1.
Private Function fn_join(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 3
    Dim i As Long, j As Long, x As Long, r_i As Long, r_j As Long, joiner As String
    x = Utils_CalcDimDirection(args, 3)
    joiner = Utils_GetOptionalScalarArg(args, 2, "")
    Dim r: ReDim r(x * UBound(args(1), 1) + (1 - x), (1 - x) * UBound(args(1), 2) + x)
    For i = 1 To UBound(args(1), 1)
        For j = 1 To UBound(args(1), 2)
            r_i = x * i + (1 - x)
            r_j = (1 - x) * j + x
            If (1 - x) * i + x * j = 1 Then
                r(r_i, r_j) = args(1)(i, j)
            Else
                r(r_i, r_j) = r(r_i, r_j) & joiner & args(1)(i, j)
            End If
        Next j
    Next i
    fn_join = r
End Function

' B = fn_expand(A)
' B = fn_expand(A,n)
' B = fn_expand(A,,m)
' B = fn_expand(A,n,m)
'
' expand(A) returns the matrix beginning in cell A and expanding down and to the right
' as far as there are contiguous non-empty cells.
' Set n or m to specifically fix the number of rows or columns.
' If n <= 0 or m <= 0, expand() will return the empty matrix.
Private Function fn_expand(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 3
    Utils_Assert _
        args(1)(1) = "eval_variable" And TypeName(arguments(args(1)(2))) = "Range", _
        "expand(): 1st argument must be a cell"
    Dim cell As Range: Set cell = arguments(args(1)(2))
    Dim rows As Variant: If UBound(args) > 1 Then rows = calc_tree(args(2))
    Dim cols As Variant: If UBound(args) > 2 Then cols = calc_tree(args(3))
    If IsEmpty(rows) Then
        If IsEmpty(cell.Offset(1, 0)) Then
            rows = 1
        Else
            rows = cell.End(xlDown).Row - cell.Row + 1
        End If
    Else
        rows = rows(1, 1)
        If rows <= 0 Then Exit Function
    End If
    If IsEmpty(cols) Then
        If IsEmpty(cell.Offset(0, 1)) Then
            cols = 1
        Else
            cols = cell.End(xlToRight).Column - cell.Column + 1
        End If
    Else
        cols = cols(1, 1)
        If cols <= 0 Then Exit Function
    End If
    fn_expand = cell.Resize(rows, cols)
    Utils_Conform fn_expand
End Function

Private Function fn_e(args As Variant) As Variant
    Utils_AssertArgsCount args, 0, 0
    fn_e = Utils_ToMatrix(Exp(1))
End Function

Private Function fn_pi(args As Variant) As Variant
    Utils_AssertArgsCount args, 0, 0
    fn_pi = Utils_ToMatrix(WorksheetFunction.Pi())
End Function

' v = version
'
' Returns a string with the current version of the Q library.
Private Function fn_version(args As Variant) As Variant
    Utils_AssertArgsCount args, 0, 0
    fn_version = Utils_ToMatrix(version)
End Function


