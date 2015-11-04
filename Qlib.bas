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
'   - =Q("a+b+c",3,4,5)        -> 12
'   - =Q("eye(3)")             -> the 3x3 identity matrix
'   - =Q("mean(a)",A1:D5)      -> vector with the mean of each column in cells A1:D5
'   - =Q("a.*b",A1:D5,F1:I5)   -> element wise multiplication of cells A1:D5 and F1:I5
'   - =Q("a([1 3],end)",A1:D5) -> get the last entries in row 1 and 3 of cells A1:D5
'   - =Q("sort(a)",A1:D5)      -> sort each column of cells A1:D5
'   - =Q("3+4;ans^2")          -> 49
'                                 Multiple expressions are separated by ";" or linebreak.
'                                 "ans" returns the last result and the very last
'                                 result is then returned by Q().
'
' Features:
'   - All standard MATLAB operators: :,::,+,-,*,/,.*,./,^,.^,||,&&,|,&,<,<=,>,>=,==,~=,~,'
'   - Most used MATLAB functions: eye,zeros,ones,sum,cumsum,cumprod,prod,mean,median,corr,
'     cov,prctile,std,isequal,fix,rand,randn,repmat,find,sqrt,exp,sort and many more...
'   - Indexing via a(2,:) or a(5,3:end)
'   - Concatenate matrices with '[]', i.e. [ a b; c d]
'   - Multiple expressions separated by ";" or a linebreak.
'     The variable ans contains the result of the previous expression
'   - Excel functions: if,iferror
'   - Prefix function calls with ! to call external VBA functions not found in Q.
'
' For the newest version, go to:
' http://github.com/nielsls/Q
'
' 2015, Niels Lykke Sørensen

Option Explicit
Option Base 1

Private Const VERSION = "1.54"
    
Private Const NUMERICS = "0123456789"
Private Const ALPHAS = "abcdefghijklmnopqrstuvwxyz"
Private Const SINGLE_OPS = "()[],;:+-#'"
Private Const COMBO_OPS = ".|&<>~=*/^!"

Private expression As String
Private expressionIndex As Long
Private currentToken As String
Private previousTokenIsSpace As Boolean
Private arguments As Variant
Private endValues As Variant ' A stack of numbers providing the right value of the "end" constant
Private errorMsg As String
Private ans As Variant       ' Result of last answer when multiple expressions are used as input

' Entry point - the only public function in the library
Public Function Q(expr As Variant, ParamArray args() As Variant) As Variant
    On Error GoTo ErrorHandler

    expression = expr
    arguments = args
    endValues = Empty
    errorMsg = ""
    ans = [NA()]
    expressionIndex = 1
    Tokens_Advance ' Find first token in input string
    
    Dim root As Variant
    Do
        Select Case currentToken
            Case ""
                Exit Do
            Case ";", vbLf
                Tokens_Advance
            Case Else
                root = Parse_Binary()
                'Utils_DumpTree root    'Uncomment for debugging
                ans = eval_tree(root)
                Utils_Assert _
                    currentToken = "" Or currentToken = ";" Or currentToken = vbLf, _
                    "'" & currentToken & "' not allowed here"
        End Select
    Loop While True
    
    Q = ans
    If IsEmpty(Q) Then Q = [NA()]  'Makes sure the empty matrix is not converted to a 0
    Exit Function
    
ErrorHandler:
    ' Cannot rely on usual error message transfer only as
    ' this does not work when using Application.Run()
    If errorMsg = "" Then errorMsg = Err.Description
    Q = "ERROR - " & errorMsg
End Function

'*******************************************************************
' Evaluator invariants:
'   - All variables must be either strings, scalars or 2d arrays.
'     A variable cannot be a 1x1 array; then it must be a scalar.
'     Use Utils_Conform() to correctly shape all variables.
'   - Variable names: [a-z]
'   - Function names: [a-z][a-z0-9_]*
'   - Numbers support e/E exponent
'   - The empty matrix/scalar [] has per definition 0 rows, 0 cols,
'     dimension 0 and is internally represented by the default
'     Variant value Empty
'*******************************************************************

'*********************
'*** TOKEN CONTROL ***
'*********************

Private Sub Tokens_Advance()
    previousTokenIsSpace = Tokens_AdvanceWhile(" ")
    If expressionIndex > Len(expression) Then currentToken = "": Exit Sub
    
    Dim startIndex As Long: startIndex = expressionIndex
    Select Case Asc(Mid(expression, expressionIndex, 1))
    
        Case Asc("""")
            expressionIndex = expressionIndex + 1
            Tokens_AdvanceWhile """", True
            Utils_Assert expressionIndex <= Len(expression), "Unfinished string literal"
            expressionIndex = expressionIndex + 1
            
        Case Asc("a") To Asc("z")
            Tokens_AdvanceWhile NUMERICS & ALPHAS & "_"
            
        Case Asc("0") To Asc("9")
            Tokens_AdvanceWhile NUMERICS
            If Tokens_AdvanceWhile(".", False, True) Then
                Tokens_AdvanceWhile NUMERICS
            End If
            If Tokens_AdvanceWhile("eE", False, True) Then
                Tokens_AdvanceWhile NUMERICS & "-", False, True
                Tokens_AdvanceWhile NUMERICS
            End If
            
        Case Asc(vbLf) 'New line
            expressionIndex = expressionIndex + 1
                
        Case Else
            If Not Tokens_AdvanceWhile(SINGLE_OPS, False, True) Then
                Tokens_AdvanceWhile COMBO_OPS
            End If
         
    End Select
    
    currentToken = Mid(expression, startIndex, expressionIndex - startIndex)
    Utils_Assert expressionIndex > startIndex Or expressionIndex > Len(expression), _
        "Illegal char: " & Mid(expression, expressionIndex, 1)
End Sub

Private Sub Tokens_AssertAndAdvance(token As String)
    Utils_Assert token = currentToken, "Missing token: " & token
    Tokens_Advance
End Sub

Private Function Tokens_AdvanceWhile(str As String, _
    Optional stopAtStr As Boolean = False, _
    Optional singleCharOnly As Boolean = False) As Boolean
    While expressionIndex <= Len(expression) _
        And stopAtStr <> (InStr(str, Mid(expression, expressionIndex, 1)) > 0)
        expressionIndex = expressionIndex + 1
        Tokens_AdvanceWhile = True
        If singleCharOnly Then Exit Function
    Wend
End Function

'**********************************
'*** BUILD ABSTRACT SYNTAX TREE ***
'**********************************

' Each leaf in the parsed tree is an array with two entries.
' First entry is a string with the name of the function that should be executed.
' Second entry is a 1-dim array with its arguments.

' Parse chain:
'
' Binary
'   Prefix
'     Postfix
'       List
'       Atomic
'         Binary
'         Matrix
'           List

'Returns true if token is a suitable operator
Private Function Parse_FindOp(token As String, opType As String, ByRef op As Variant) As Boolean
    op = Null
    
    Select Case opType
        Case "binary"
            Select Case token
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

Private Function Parse_Matrix() As Variant
    While currentToken <> "]"
        Utils_Stack_Push Parse_List(True), Parse_Matrix
        If currentToken = ";" Then Tokens_Advance
        Utils_Assert currentToken <> "", "Missing ']'"
    Wend
End Function

Private Function Parse_List(Optional isSpaceSeparator As Boolean = False) As Variant
    Do While InStr(";)]", currentToken) = 0
        Utils_Stack_Push Parse_Binary(), Parse_List
        If currentToken = "," Then
            Tokens_Advance
        ElseIf Not (previousTokenIsSpace And isSpaceSeparator) Then
            Exit Do
        End If
    Loop
End Function

Private Function Parse_Binary(Optional lastPrec As Long = -999) As Variant
    Parse_Binary = Parse_Prefix()
    Dim op: Do While Parse_FindOp(currentToken, "binary", op)
        If op(2) + CLng(op(3)) < lastPrec Then Exit Do
        Tokens_Advance
        Parse_Binary = Array("op_" & op(1), Array(Parse_Binary, Parse_Binary(CLng(op(2)))))
    Loop
End Function

Private Function Parse_Prefix() As Variant
    Dim op
    If Parse_FindOp(currentToken, "unaryprefix", op) Then
        Tokens_Advance
        Parse_Prefix = Array("op_" & op, Array(Parse_Prefix()))
    Else
        Parse_Prefix = Parse_Postfix()
    End If
End Function

Private Function Parse_Postfix() As Variant
    Parse_Postfix = Parse_Atomic
    Dim op: Do
        If Parse_FindOp(currentToken, "unarypostfix", op) Then
            Parse_Postfix = Array("op_" & op, Array(Parse_Postfix))
            Tokens_Advance
        ElseIf currentToken = "(" Then
            Tokens_Advance
            Parse_Postfix = Array("eval_index", Array(Parse_Postfix, Parse_List()))
            Tokens_AssertAndAdvance ")"
        Else
            Exit Do
        End If
    Loop While True
End Function

Private Function Parse_Atomic() As Variant
    Utils_Assert currentToken <> "", "Missing argument"
    Select Case Asc(currentToken) ' Filter on first char of token
            
        Case Asc(":")
            Parse_Atomic = Array("eval_colon", Empty)
            Tokens_Advance
            
        Case Asc("(")
            Tokens_Advance
            Parse_Atomic = Parse_Binary()
            Tokens_AssertAndAdvance ")"
            
        Case Asc("[")  ' Found a matrix concatenation
            Tokens_Advance
            Parse_Atomic = Array("eval_concat", Parse_Matrix())
            Tokens_AssertAndAdvance "]"
    
        Case Asc("""") ' Found a constant string
            Parse_Atomic = Array("eval_constant", Array(Mid(currentToken, 2, Len(currentToken) - 2)))
            Tokens_Advance
            
        Case Asc("0") To Asc("9") ' Found a numeric constant
            Parse_Atomic = Array("eval_constant", Array(Val(currentToken)))
            Tokens_Advance
                    
        Case Asc("a") To Asc("z")
            If currentToken = "end" Then
                Parse_Atomic = Array("eval_end", Empty)
                Tokens_Advance
            ElseIf currentToken = "ans" Then
                Parse_Atomic = Array("eval_ans", Empty)
                Tokens_Advance
            ElseIf Len(currentToken) = 1 Then ' Found an input variable
                Parse_Atomic = Array("eval_arg", Array(Asc(currentToken) - Asc("a")))
                Tokens_Advance
            Else                   ' Found a function call
                Parse_Atomic = "fn_" & currentToken
                Tokens_Advance
                If currentToken = "(" Then
                    Tokens_AssertAndAdvance "("
                    Parse_Atomic = Array(Parse_Atomic, Parse_List())
                    Tokens_AssertAndAdvance ")"
                Else
                    Parse_Atomic = Array(Parse_Atomic, Empty)
                End If
            End If
            
        Case Else
            Utils_Assert False, "Unexpected token: " & currentToken
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

Private Function IFF(a As Boolean, b As Variant, c As Variant) As Variant
    If a Then IFF = b Else IFF = c
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
    Select Case Utils_Dimensions(v)
        Case 0: Utils_Numel = IFF(IsEmpty(v), 0, 1)
        Case 1: Utils_Numel = UBound(v)
        Case 2: Utils_Numel = UBound(v, 1) * UBound(v, 2)
        Case Else: Utils_Assert False, "Dimension > 2"
    End Select
End Function

' Makes sure that a 1x1 matrix is transformed to a scalar
' and a 1-dim vector is transformed to a 2-dim vector of size 1xN
Private Sub Utils_Conform(ByRef v As Variant)
    Select Case Utils_Dimensions(v)
        Case 1:
            If UBound(v) = 1 Then
                v = v(1)
            Else
                Dim r: ReDim r(1, UBound(v))
                Dim i As Long
                For i = 1 To UBound(r, 2)
                    r(1, i) = v(i)
                Next i
                v = r
            End If
            
        Case 2:
            If UBound(v, 1) = 1 And UBound(v, 2) = 1 Then v = v(1, 1)
            
        Case Is > 2:
            Utils_Assert False, "Dimension > 2"
    End Select
End Sub

Private Sub Utils_ForceMatrix(v As Variant)
    If Utils_Dimensions(v) = 0 Then
        Dim r: ReDim r(1, 1)
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
        Case 0: If Not IsEmpty(v) Then r = 1: c = 1
        Case 1: r = UBound(v): c = 1
        Case 2: r = UBound(v, 1): c = UBound(v, 2)
        Case Else: Utils_Assert False, "Dimension > 2"
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

' Transforms all entries in the vector from trees to values
Private Sub Utils_CalcArgs(args As Variant)
    Dim i As Long
    For i = 1 To Utils_Stack_Size(args)
        args(i) = eval_tree(args(i))
    Next i
End Sub

' Test if a flag was supplied in the args, i.e. such as "descend" in the sort function
Private Function Utils_IsFlagSet(args As Variant, flag As String) As Boolean
    Dim i As Long
    For i = UBound(args) To 1 Step -1
        If StrComp(TypeName(args(i)), "String") = 0 Then
            If StrComp(args(i), flag, vbTextCompare) = 0 Then
                Utils_IsFlagSet = True
                Exit Function
            End If
        End If
    Next i
End Function

'do cols -> return 0, do rows -> return 1
Private Function Utils_CalcDimDirection(args As Variant, Optional dimIndex As Long = 2) As Long
    If UBound(args) >= dimIndex Then
        If IsNumeric(args(dimIndex)) Then
            Utils_CalcDimDirection = args(dimIndex) - 1
            Exit Function
        End If
    End If
    Utils_CalcDimDirection = IFF(Utils_Rows(args(1)) = 1, 1, 0)
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
                    n = args(index)
                    m = n
                Case 2
                    n = args(index)(1, 1)
                    m = args(index)(MIN(2, UBound(args(index), 1)), MIN(2, UBound(args(index), 2)))
                Case Else
                    Utils_Assert False, "Bad input format"
            End Select
        Case Is = index + 1
            n = args(index)
            m = args(index + 1)
        Case Else
            Utils_Assert False, "Bad size input"
    End Select
End Function

' Provides a Q function with an easy way of obtaining the value of an optional argument
Private Function Utils_GetOptionalArg(args As Variant, index As Long, defaultValue As Variant)
    If Utils_Stack_Size(args) >= index Then
        Utils_GetOptionalArg = args(index)
    Else
        Utils_GetOptionalArg = defaultValue
    End If
End Function

' Do the initial calculations that are the same for every binary operation
Private Function Utils_SetupBinaryOperation(args As Variant, r As Variant, _
    ByRef r1 As Long, ByRef c1 As Long, ByRef r2 As Long, ByRef c2 As Long, _
    Optional preCalcArgs As Boolean = True) As Variant
    If preCalcArgs Then Utils_CalcArgs args
    Utils_ForceMatrix args(1): Utils_Size args(1), r1, c1
    Utils_ForceMatrix args(2): Utils_Size args(2), r2, c2
    Utils_Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2), _
        "Dimension mismatch"
    ReDim r(MAX(r1, r2), MAX(c1, c2))
End Function

' Is called from each Q function to ensure the given function has
' has been called with the right number of arguments
Private Sub Utils_AssertArgsCount(args As Variant, lb As Long, ub As Long)
    Dim size As Long: size = Utils_Stack_Size(args)
    Utils_Assert size >= lb, "Too few arguments"
    Utils_Assert size <= ub, "Too many arguments"
End Sub

' Allows each Q function to fail gracefully with a nice error message.
Private Sub Utils_Assert(expr As Boolean, Optional msg As String = "Unknown error")
    If expr Then Exit Sub
    errorMsg = msg
    Err.Raise vbObjectError + 999
End Sub

'**********************
'*** EVAL FUNCTIONS ***
'**********************

' The main eval function
' Supply an abstract syntax tree and get a value
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
            eval_tree = Run(root(1), root(2))
    End Select
End Function

Private Function eval_constant(args As Variant) As Variant
    eval_constant = args(1)
End Function

Private Function eval_arg(args As Variant) As Variant
    If args(1) > UBound(arguments) Then
        Utils_Assert False, "Argument '" & Chr(Asc("a") + args(1)) & "' not found."
    End If
    eval_arg = CVar(arguments(args(1)))
    Utils_Conform eval_arg
End Function

Private Function eval_end(args As Variant) As Variant
    Utils_Assert Utils_Stack_Size(endValues) > 0, """end"" not allowed here."
    eval_end = Utils_Stack_Peek(endValues)
End Function

Private Function eval_ans(args As Variant) As Variant
    eval_ans = ans
End Function

Private Function eval_colon(args As Variant) As Variant
    Utils_Assert False, "colon not allowed here"
End Function

Private Function Utils_IsVectorShape(r As Long, c As Long) As Boolean
    Utils_IsVectorShape = (r = 1 And c > 1) Or (r > 1 And c = 1)
End Function

Private Function eval_indexarg(root As Variant, endValue As Long) As Variant
    Dim r As Variant
    If root(1) = "eval_colon" Then
        ReDim r(endValue, 1)
        Dim idx As Long
        For idx = 1 To endValue
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
    Dim r, i1, i2
    Dim matrows As Long, matcols As Long, r1 As Long, c1 As Long, r2 As Long, c2 As Long
    Dim idx As Long, r_i As Long, r_j As Long, arg_i As Long, arg_j As Long, i1_i As Long, i1_j As Long, i2_i As Long, i2_j As Long
    
    args(1) = eval_tree(args(1))
    Utils_ForceMatrix args(1)
    Utils_Size args(1), matrows, matcols
    
    Select Case Utils_Stack_Size(args(2))
        
        Case 0:
            r = args(1)
        
        Case 1:
            On Error GoTo ErrorHandler
            i1 = eval_indexarg(args(2)(1), matrows * matcols)
            Utils_Size i1, r1, c1
            If r1 = 0 Or c1 = 0 Then Exit Function
            Utils_ForceMatrix i1
            If Utils_IsVectorShape(r1, c1) And Utils_IsVectorShape(matrows, matcols) Then
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
            On Error GoTo 0
            
        Case 2:
            On Error GoTo ErrorHandler
            i1 = eval_indexarg(args(2)(1), matrows)
            i2 = eval_indexarg(args(2)(2), matcols)
            Utils_Size i1, r1, c1
            Utils_Size i2, r2, c2
            If r1 = 0 Or c1 = 0 Or r2 = 0 Or c2 = 0 Then Exit Function
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
            On Error GoTo 0
            
        Case Else:
            Utils_Assert False, "Too many index arguments"
    
    End Select

    Utils_Conform r
    eval_index = r
    Exit Function
    
ErrorHandler:
    Utils_Assert False, Err.Description
End Function

' Evaluates matrix concatenation []
Private Function eval_concat(args As Variant) As Variant

    ' Get matrices and check their sizes are compatible for concatenation
    Dim totalRows As Long, totalCols As Long
    Dim requiredRows As Long, requiredCols As Long
    Dim rows As Long, cols As Long, i As Long, j As Long
    requiredCols = 0
    For i = 1 To Utils_Stack_Size(args) ' loop over each row
        totalCols = 0
        requiredRows = 0
        For j = 1 To Utils_Stack_Size(args(i)) ' loop over each column
            args(i)(j) = eval_tree(args(i)(j))
            Utils_Size args(i)(j), rows, cols
            If requiredRows = 0 Then
                requiredRows = rows
            Else
                Utils_Assert requiredRows = rows Or rows = 0, "Concatenation: Different row counts"
            End If
            totalCols = totalCols + cols
        Next j
        If requiredCols = 0 Then
            requiredCols = totalCols
        Else
            Utils_Assert requiredCols = totalCols Or totalCols = 0, "Concatenation: Different column counts"
        End If
        totalRows = totalRows + requiredRows
    Next i
    totalCols = requiredCols 'Needed in case last row was []
    
    ' Perform the actual concatenation by copying input matrices
    If totalRows = 0 Or totalCols = 0 Then Exit Function
    Dim r: ReDim r(totalRows, totalCols)
    Dim x As Long, y As Long
    totalRows = 0
    For i = 1 To Utils_Stack_Size(args)
        totalCols = 0
        For j = 1 To Utils_Stack_Size(args(i))
            If IsEmpty(args(i)(j)) Then
                rows = 0
                cols = 0
            Else
                Utils_ForceMatrix args(i)(j)
                Utils_Size args(i)(j), rows, cols
                For x = 1 To rows
                    For y = 1 To cols
                        r(totalRows + x, totalCols + y) = args(i)(j)(x, y)
                    Next y
                Next x
            End If
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
        Case Else: Utils_Assert False, "Cannot evaluate " & args(1)(1) & ": Too many arguments"
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
    Utils_Assert False, "Operator ||: Could not convert argument to boolean value"
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
    Utils_Assert False, "Operator &&: Could not convert argument to boolean value"
End Function

' Matches operator &
Private Function op_and(args As Variant) As Variant
    Dim r, r1 As Long, c1 As Long, r2 As Long, c2 As Long, x As Long, y As Long
    Utils_SetupBinaryOperation args, r, r1, c1, r2, c2
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
    Dim r, r1 As Long, c1 As Long, r2 As Long, c2 As Long, x As Long, y As Long
    Utils_SetupBinaryOperation args, r, r1, c1, r2, c2
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
    Dim r, r1 As Long, c1 As Long, r2 As Long, c2 As Long, x As Long, y As Long
    Utils_SetupBinaryOperation args, r, r1, c1, r2, c2
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
    Dim r, r1 As Long, c1 As Long, r2 As Long, c2 As Long, x As Long, y As Long
    Utils_SetupBinaryOperation args, r, r1, c1, r2, c2
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
    Dim r, r1 As Long, c1 As Long, r2 As Long, c2 As Long, x As Long, y As Long
    Utils_SetupBinaryOperation args, r, r1, c1, r2, c2
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
    Dim r, r1 As Long, c1 As Long, r2 As Long, c2 As Long, x As Long, y As Long
    Utils_SetupBinaryOperation args, r, r1, c1, r2, c2
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
    Dim r, r1 As Long, c1 As Long, r2 As Long, c2 As Long, x As Long, y As Long
    Utils_SetupBinaryOperation args, r, r1, c1, r2, c2
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
    Dim r, r1 As Long, c1 As Long, r2 As Long, c2 As Long, x As Long, y As Long
    Utils_SetupBinaryOperation args, r, r1, c1, r2, c2
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
    If m < 0 Then Exit Function
    Dim r: ReDim r(1, 1 + m)
    For i = 0 To m
        r(1, 1 + i) = start + i * step
    Next i
    Utils_Conform r
    op_colon = r
End Function

' Matches operator +
Private Function op_plus(args As Variant) As Variant
    Dim r, r1 As Long, c1 As Long, r2 As Long, c2 As Long, x As Long, y As Long
    Utils_SetupBinaryOperation args, r, r1, c1, r2, c2
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
    Dim r, r1 As Long, c1 As Long, r2 As Long, c2 As Long, x As Long, y As Long
    Utils_SetupBinaryOperation args, r, r1, c1, r2, c2
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
        Dim r: ReDim r(rows, cols)
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
        Utils_Assert UBound(args(1), 2) = UBound(args(2), 1), "mtimes(): Matrix sizes not compatible"
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
    Dim r, r1 As Long, c1 As Long, r2 As Long, c2 As Long, x As Long, y As Long
    Utils_SetupBinaryOperation args, r, r1, c1, r2, c2
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
    Utils_Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1)
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
    Dim r, r1 As Long, c1 As Long, r2 As Long, c2 As Long, x As Long, y As Long
    Utils_SetupBinaryOperation args, r, r1, c1, r2, c2
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
        Utils_Assert False, "mpower: Input must be scalars"
    End If
End Function

' Matches operator .^
Private Function op_power(args As Variant) As Variant
    Dim r, r1 As Long, c1 As Long, r2 As Long, c2 As Long, x As Long, y As Long
    Utils_SetupBinaryOperation args, r, r1, c1, r2, c2
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
    op_transpose = WorksheetFunction.Transpose(eval_tree(args(1)))
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
    fn_islogical = False
    Utils_ForceMatrix args(1)
    Dim i As Long, j As Long
    For i = 1 To UBound(args(1), 1)
        For j = 1 To UBound(args(1), 2)
            If Not WorksheetFunction.IsLogical(args(1)(i, j)) Then Exit Function
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
        If CDbl(args(1)) <> 0 Then fn_find = 1
    Else
        Dim counter As Long, i As Long, j As Long
        For i = 1 To rows
            For j = 1 To cols
                If CDbl(args(1)(i, j)) <> 0 Then counter = counter + 1
            Next j
        Next i
        If counter = 0 Then Exit Function
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
    Utils_AssertArgsCount args, 1, 1
    Utils_ForceMatrix args(1)
    Dim r: ReDim r(UBound(args(1), 1), UBound(args(1), 2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = WorksheetFunction.RoundDown(args(1)(x, y), 0)
        Next y
    Next x
    Utils_Conform r
    fn_fix = r
End Function

' X = round(A)
' X = round(A,k)
'
' X = round(A) rounds the elements of A.
' X = round(A,k) rounds the elements of A with k decimal places. Default is 0
Private Function fn_round(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    Dim k As Long: If UBound(args) > 1 Then k = args(2)
    Utils_ForceMatrix args(1)
    Dim r: ReDim r(UBound(args(1), 1), UBound(args(1), 2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = WorksheetFunction.Round(args(1)(x, y), k)
        Next y
    Next x
    Utils_Conform r
    fn_round = r
End Function

' X = ceil(A)
'
' X = ceil(A) rounds each element of A to the nearest integer greater than
' or equal to that element.
Private Function fn_ceil(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Utils_ForceMatrix args(1)
    Dim r: ReDim r(UBound(args(1), 1), UBound(args(1), 2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = WorksheetFunction.Ceiling(args(1)(x, y), 1)
        Next y
    Next x
    Utils_Conform r
    fn_ceil = r
End Function

' X = floor(A)
'
' X = floor(A) rounds each element of A to the nearest integer smaller than
' or equal to that element.
Private Function fn_floor(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Utils_ForceMatrix args(1)
    Dim r: ReDim r(UBound(args(1), 1), UBound(args(1), 2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = WorksheetFunction.Floor(args(1)(x, y), 1)
        Next y
    Next x
    Utils_Conform r
    fn_floor = r
End Function

Private Function fn_inv(args As Variant) As Variant
    If Utils_Dimensions(args(1)) = 0 Then
        fn_inv = 1# / args(1)
    Else
        Utils_Assert UBound(args(1), 1) = UBound(args(1), 2), "matrix not quadratic"
        fn_inv = WorksheetFunction.MInverse(args(1))
    End If
End Function

' X = exp(A)
'
' X = exp(A) returns the exponential for each element of A.
Private Function fn_exp(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Utils_ForceMatrix args(1)
    Dim i As Long, j As Long
    Dim r: ReDim r(UBound(args(1), 1), UBound(args(1), 2))
    For i = 1 To UBound(r, 1)
        For j = 1 To UBound(r, 2)
            r(i, j) = Exp(args(1)(i, j))
        Next j
    Next i
    Utils_Conform r
    fn_exp = r
End Function

' X = log(A)
'
' X = log(A) returns the natural logarithm of the elements of A.
Private Function fn_log(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Utils_ForceMatrix args(1)
    Dim i As Long, j As Long
    Dim r: ReDim r(UBound(args(1), 1), UBound(args(1), 2))
    For i = 1 To UBound(r, 1)
        For j = 1 To UBound(r, 2)
            r(i, j) = Log(args(1)(i, j))
        Next j
    Next i
    Utils_Conform r
    fn_log = r
End Function

' X = sqrt(A)
'
' X = sqrt(A) returns the square root of the elements of A.
Private Function fn_sqrt(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Utils_ForceMatrix args(1)
    Dim i As Long, j As Long
    Dim r: ReDim r(UBound(args(1), 1), UBound(args(1), 2))
    For i = 1 To UBound(r, 1)
        For j = 1 To UBound(r, 2)
            r(i, j) = Sqr(args(1)(i, j))
        Next j
    Next i
    Utils_Conform r
    fn_sqrt = r
End Function

' r = rows(A)
'
' r = rows(A) returns the number of rows in A.
Private Function fn_rows(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    fn_rows = Utils_Rows(args(1))
End Function

' c = cols(A)
'
' c = cols(A) returns the number of columns in A.
Private Function fn_cols(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    fn_cols = Utils_Cols(args(1))
End Function

' n = numel(A)
'
' n = numel(A) returns the number of elements in A.
Private Function fn_numel(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    fn_numel = Utils_Numel(args(1))
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
    fn_zeros = fn_repmat(Array(0, n, m))
    Utils_Conform fn_zeros
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
    fn_ones = fn_repmat(Array(1, n, m))
    Utils_Conform fn_ones
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
    Dim r: ReDim r(n, m)
    For n = 1 To UBound(r, 1)
        For m = 1 To UBound(r, 2)
            r(n, m) = -CLng(n = m)
        Next m
    Next n
    Utils_Conform r
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
    fn_true = fn_repmat(Array(True, n, m))
    Utils_Conform fn_true
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
    fn_false = fn_repmat(Array(False, n, m))
    Utils_Conform fn_false
End Function

' X = xor(A,B)
Private Function fn_xor(args As Variant) As Variant
    Dim r, r1 As Long, c1 As Long, r2 As Long, c2 As Long, x As Long, y As Long
    Utils_SetupBinaryOperation args, r, r1, c1, r2, c2, False
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = CBool(args(1)(MIN(x, r1), MIN(y, c1))) Xor CBool(args(2)(MIN(x, r2), MIN(y, c2)))
        Next y
    Next x
    Utils_Conform r
    fn_xor = r
End Function

' X = tick2ret(A)
' X = tick2ret(A,method)
' X = tick2ret(A,method,dim)
Private Function fn_tick2ret(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 3
    Utils_ForceMatrix args(1)
    Dim x As Long: x = Utils_CalcDimDirection(args, 3)
    Dim simple As Boolean: simple = Not Utils_IsFlagSet(args, "continuous")
    Dim r: ReDim r(UBound(args(1), 1) - (1 - x), UBound(args(1), 2) - x)
    Dim i As Long, j As Long
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

' X = isnum(A)
'
' X = isnum(A) returns a matrix with the same of A
' indicating whether its elements are numeric.
Private Function fn_isnum(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    If IsEmpty(args(1)) Then fn_isnum = False: Exit Function
    Dim i As Long, j As Long
    Utils_ForceMatrix args(1)
    For i = 1 To UBound(args(1), 1)
        For j = 1 To UBound(args(1), 2)
            args(1)(i, j) = IsNumeric(args(1)(i, j))
        Next j
    Next i
    Utils_Conform args(1)
    fn_isnum = args(1)
End Function

' X = iserror(A)
'
' X = iserror(A) returns a matrix same size as A indicating
' whether each entry is an error or not.
Private Function fn_iserror(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Utils_ForceMatrix args(1)
    Dim r: ReDim r(UBound(args(1), 1), UBound(args(1), 2))
    Dim x As Long, y As Long
    For x = 1 To UBound(r, 1)
        For y = 1 To UBound(r, 2)
            r(x, y) = IsError(args(1)(x, y))
        Next y
    Next x
    Utils_Conform r
    fn_iserror = r
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
    If IsEmpty(args(1)) Then fn_cumsum = 0: Exit Function
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
' along the first dimension (the columns); cumprod(A,2) works along the
' second dimension (the rows).
Private Function fn_cumprod(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    If IsEmpty(args(1)) Then fn_cumprod = 1: Exit Function
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

' B = cummax(A)
' B = cummax(A,dim)
Private Function fn_cummax(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    Dim i As Long, j As Long, x As Long
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args)
    For i = 2 - x To UBound(args(1), 1)
        For j = 1 + x To UBound(args(1), 2)
            args(1)(i, j) = MAX(args(1)(i, j), args(1)(i - (1 - x), j - x))
        Next j
    Next i
    Utils_Conform args(1)
    fn_cummax = args(1)
End Function

' B = cummin(A)
' B = cummin(A,dim)
Private Function fn_cummin(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    Dim i As Long, j As Long, x As Long
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args)
    For i = 2 - x To UBound(args(1), 1)
        For j = 1 + x To UBound(args(1), 2)
            args(1)(i, j) = MIN(args(1)(i, j), args(1)(i - (1 - x), j - x))
        Next j
    Next i
    Utils_Conform args(1)
    fn_cummin = args(1)
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
    Utils_AssertArgsCount args, 1, 2
    Dim i As Long, x As Long
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args)
    Dim r: ReDim r(x * UBound(args(1), 1) + (1 - x), (1 - x) * UBound(args(1), 2) + x)
    For i = 1 To UBound(r, 2 - x)
        r(x * i + (1 - x), (1 - x) * i + x) _
            = WorksheetFunction.StDev(WorksheetFunction.index(args(1), x * i, (1 - x) * i))
    Next i
    Utils_Conform r
    fn_std = r
End Function

' X = corr(A)
'
' X = corr(A) returns a correlation matrix for the columns of A.
Private Function fn_corr(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim c As Long: c = UBound(args(1), 2)
    Dim r: ReDim r(c, c)
    Dim i As Long, j As Long
    For i = 1 To c
        r(i, i) = 1
        For j = i + 1 To c
            r(i, j) = WorksheetFunction.Correl( _
                WorksheetFunction.index(args(1), 0, i), _
                WorksheetFunction.index(args(1), 0, j))
            r(j, i) = r(i, j)
        Next j
    Next i
    fn_corr = r
End Function

' X = cov(A)
'
' X = cov(A) returns a covariance matrix for the columns of A.
Private Function fn_cov(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim c As Long: c = UBound(args(1), 2)
    Dim r: ReDim r(c, c)
    Dim i As Long, j As Long
    For i = 1 To c
        For j = i To c
            r(i, j) = WorksheetFunction.Covar( _
                WorksheetFunction.index(args(1), 0, i), _
                WorksheetFunction.index(args(1), 0, j))
            r(j, i) = r(i, j)
        Next j
    Next i
    fn_cov = r
End Function

' X = all(A)
' X = all(A,dim)
'
' all(...) tests if all elements in A evaluates to true.
' In practice, all() is a natural extension of the logical AND
' operator.
Private Function fn_all(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    If IsEmpty(args(1)) Then fn_all = True: Exit Function
    Dim x As Long, i As Long
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args)
    Dim r: ReDim r(x * UBound(args(1), 1) + (1 - x), (1 - x) * UBound(args(1), 2) + x)
    For i = 1 To UBound(r, 2 - x)
        r(x * i + (1 - x), (1 - x) * i + x) _
            = WorksheetFunction.And(WorksheetFunction.index(args(1), x * i, (1 - x) * i))
    Next i
    Utils_Conform r
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
    If IsEmpty(args(1)) Then fn_any = False: Exit Function
    Dim x As Long, i As Long
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args)
    Dim r: ReDim r(x * UBound(args(1), 1) + (1 - x), (1 - x) * UBound(args(1), 2) + x)
    For i = 1 To UBound(r, 2 - x)
        r(x * i + (1 - x), (1 - x) * i + x) _
            = WorksheetFunction.Or(WorksheetFunction.index(args(1), x * i, (1 - x) * i))
    Next i
    Utils_Conform r
    fn_any = r
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
    If IsEmpty(args(1)) Then fn_sum = 0: Exit Function
    Dim x As Long, i As Long, r As Variant
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args)
    ReDim r(x * UBound(args(1), 1) + (1 - x), (1 - x) * UBound(args(1), 2) + x)
    For i = 1 To UBound(r, 2 - x)
        r(x * i + (1 - x), (1 - x) * i + x) _
            = WorksheetFunction.Sum(WorksheetFunction.index(args(1), x * i, (1 - x) * i))
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
    Utils_AssertArgsCount args, 1, 2
    If IsEmpty(args(1)) Then fn_prod = 1: Exit Function
    Dim i As Long, x As Long, r As Variant
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args)
    ReDim r(x * UBound(args(1), 1) + (1 - x), (1 - x) * UBound(args(1), 2) + x)
    For i = 1 To UBound(r, 2 - x)
        r(x * i + (1 - x), (1 - x) * i + x) _
            = WorksheetFunction.Product(WorksheetFunction.index(args(1), x * i, (1 - x) * i))
    Next i
    Utils_Conform r
    fn_prod = r
End Function

' X = mean(A)
' X = mean(A,dim)
Private Function fn_mean(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    Dim i As Long, x As Long
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args)
    Dim r: ReDim r(x * UBound(args(1), 1) + (1 - x), (1 - x) * UBound(args(1), 2) + x)
    For i = 1 To UBound(r, 2 - x)
        r(x * i + (1 - x), (1 - x) * i + x) _
            = WorksheetFunction.Average(WorksheetFunction.index(args(1), x * i, (1 - x) * i))
    Next i
    Utils_Conform r
    fn_mean = r
End Function

' X = median(A)
' X = median(A,dim)
Private Function fn_median(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 2
    Dim i As Long, x As Long
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args)
    Dim r: ReDim r(x * UBound(args(1), 1) + (1 - x), (1 - x) * UBound(args(1), 2) + x)
    For i = 1 To UBound(r, 2 - x)
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
    Utils_AssertArgsCount args, 1, 3
    Dim i As Long, x As Long
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args, 3)
    Dim r: ReDim r(x * UBound(args(1), 1) + (1 - x), (1 - x) * UBound(args(1), 2) + x)
    For i = 1 To UBound(r, 2 - x)
        r(x * i + (1 - x), (1 - x) * i + x) _
            = WorksheetFunction.percentile(WorksheetFunction.index(args(1), x * i, (1 - x) * i), args(2))
    Next i
    Utils_Conform r
    fn_prctile = r
End Function

' M = max(A)
' M = max(A,[],dim)
' M = max(A,B)
Private Function fn_max(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 3
    Dim r As Variant, x As Long, y As Long, i As Long
    
    Utils_ForceMatrix args(1)
    Dim r1 As Long, c1 As Long
    Utils_Size args(1), r1, c1
    
    If UBound(args) = 1 Or UBound(args) = 3 Then
        If UBound(args) = 3 Then Utils_Assert IsEmpty(args(2)), "2nd argument must be empty matrix, []."
        x = Utils_CalcDimDirection(args, 3)
        ReDim r(x * r1 + (1 - x), (1 - x) * c1 + x)
        For i = 1 To UBound(r, 2 - x)
            r(x * i + (1 - x), (1 - x) * i + x) _
                = WorksheetFunction.MAX(WorksheetFunction.index(args(1), x * i, (1 - x) * i))
        Next i
    Else
        Utils_ForceMatrix args(2)
        Dim r2 As Long, c2 As Long
        Utils_Size args(2), r2, c2
        Utils_Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2), "max(): Wrong dimensions."
        ReDim r(MAX(r1, r2), MAX(c1, c2))
        For x = 1 To UBound(r, 1)
            For y = 1 To UBound(r, 2)
                r(x, y) = MAX(args(1)(MIN(x, r1), MIN(y, c1)), args(2)(MIN(x, r2), MIN(y, c2)))
            Next y
        Next x
    End If
    
    Utils_Conform r
    fn_max = r
End Function

' M = min(A)
' M = min(A,[],dim)
' M = min(A,B)
Private Function fn_min(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 3
    Dim r As Variant, x As Long, y As Long, i As Long
    
    Utils_ForceMatrix args(1)
    Dim r1 As Long, c1 As Long
    Utils_Size args(1), r1, c1
    
    If UBound(args) = 1 Or UBound(args) = 3 Then
        If UBound(args) = 3 Then Utils_Assert IsEmpty(args(2)), "2nd argument must be empty matrix, []."
        x = Utils_CalcDimDirection(args, 3)
        ReDim r(x * r1 + (1 - x), (1 - x) * c1 + x)
        For i = 1 To UBound(r, 2 - x)
            r(x * i + (1 - x), (1 - x) * i + x) _
                = WorksheetFunction.MIN(WorksheetFunction.index(args(1), x * i, (1 - x) * i))
        Next i
    Else
        Utils_ForceMatrix args(2)
        Dim r2 As Long, c2 As Long
        Utils_Size args(2), r2, c2
        Utils_Assert (r1 = 1 And c1 = 1) Or (r2 = 1 And c2 = 1) Or (r1 = r2 And c1 = c2), "min(): Wrong dimensions."
        ReDim r(MAX(r1, r2), MAX(c1, c2))
        For x = 1 To UBound(r, 1)
            For y = 1 To UBound(r, 2)
                r(x, y) = MIN(args(1)(MIN(x, r1), MIN(y, c1)), args(2)(MIN(x, r2), MIN(y, c2)))
            Next y
        Next x
    End If
    
    Utils_Conform r
    fn_min = r
End Function

' b = isequal(A,B)
Private Function fn_isequal(args As Variant) As Variant
    Utils_AssertArgsCount args, 2, 2
    fn_isequal = False
    Dim dim1 As Long: dim1 = Utils_Dimensions(args(1))
    Dim dim2 As Long: dim2 = Utils_Dimensions(args(2))
    If dim1 = 0 And dim2 = 0 Then
        fn_isequal = (args(1) = args(2))
    ElseIf dim1 = 2 And dim2 = 2 Then
        Dim r1 As Long, c1 As Long: Utils_Size args(1), r1, c1
        Dim r2 As Long, c2 As Long: Utils_Size args(2), r2, c2
        If r1 <> r2 Or c1 <> c2 Then Exit Function
        Dim i As Long, j As Long
        For i = 1 To r1
            For j = 1 To c1
                If args(1)(i, j) <> args(2)(i, j) Then Exit Function
            Next j
        Next i
        fn_isequal = True
    End If
End Function

' X = isempty(A)
'
' Returns true if A equals the empty matrix [].
Private Function fn_isempty(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    fn_isempty = IsEmpty(args(1))
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

' X = diag(A)
'
' X = diag(A) returns a matrix with A in the diagonal if A is a vector,
' or a vector with the diagonal of A if A is a matrix.
Private Function fn_diag(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim rows As Long, cols As Long, r As Variant, i As Long
    Utils_ForceMatrix args(1)
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
    Utils_Conform r
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
    Dim r: ReDim r(n, m)
    For n = 1 To UBound(r, 1)
        For m = 1 To UBound(r, 2)
            r(n, m) = Rnd
        Next m
    Next n
    Utils_Conform r
    fn_rand = r
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
    If Utils_Numel(args(1)) = 1 Then
        imin = 1
        imax = args(1)
    Else
        imin = args(1)(1, 1)
        imax = args(1)(MIN(2, UBound(args(1), 1)), MIN(2, UBound(args(1), 2)))
    End If
    Dim r: ReDim r(n, m)
    For n = 1 To UBound(r, 1)
        For m = 1 To UBound(r, 2)
            r(n, m) = CLng(Rnd * (imax - imin)) + imin
        Next m
    Next n
    Utils_Conform r
    fn_randi = r
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
    Dim r: ReDim r(n, m)
    Dim c As Long: c = 3
    Dim p(2) As Double, tmp As Double
    For n = 1 To UBound(r, 1)
        For m = 1 To UBound(r, 2)
            If c > 2 Then
                Do
                    p(1) = 2 * Rnd - 1
                    p(2) = 2 * Rnd - 1
                    tmp = p(1) * p(1) + p(2) * p(2)
                Loop Until tmp <= 1
                tmp = Sqr(-2 * Log(tmp) / tmp)
                p(1) = p(1) * tmp
                p(2) = p(2) * tmp
                c = 1
            End If
            r(n, m) = p(c)
            c = c + 1
        Next m
    Next n
    Utils_Conform r
    fn_randn = r
End Function

' X = normcdf(A)
Private Function fn_normcdf(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    If Utils_Dimensions(args(1)) = 0 Then
        fn_normcdf = WorksheetFunction.NormSDist(args(1))
    Else
        Dim i As Long, j As Long
        Dim r: ReDim r(UBound(args(1), 1), UBound(args(1), 2))
        For i = 1 To UBound(r, 1)
            For j = 1 To UBound(r, 2)
                r(i, j) = WorksheetFunction.NormSDist(args(1)(i, j))
            Next j
        Next i
        fn_normcdf = r
    End If
End Function

' X = repmat(A,n)
' X = repmat(A,n,m)
' X = repmat(A,[n m])
'
' X = repmat(A,n) creates a large matrix X consisting of an n-by-n tiling of A.
' X = repmat(A,n,m) creates a large matrix X consisting of an n-by-m tiling of A.
' X = repmat(A,[n m]) creates a large matrix X consisting of an n-by-m tiling of A.
Private Function fn_repmat(args As Variant) As Variant
    Dim r1 As Long, c1 As Long, n As Long, m As Long, i As Long, j As Long
    Utils_AssertArgsCount args, 2, 3
    Utils_GetSizeFromArgs args, n, m
    Utils_ForceMatrix args(1)
    Utils_Size args(1), r1, c1
    Dim r: ReDim r(r1 * n, c1 * m)
    For n = 0 To n - 1
        For m = 0 To m - 1
            For i = 1 To r1
                For j = 1 To c1
                    r(n * r1 + i, m * c1 + j) = args(1)(i, j)
                Next j
            Next i
        Next m
    Next n
    Utils_Conform r
    fn_repmat = r
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
    Utils_ForceMatrix args(1)
    Utils_Size args(1), rows, cols
    If IsEmpty(args(2)) Then args(2) = rows * cols / args(3)
    If IsEmpty(args(3)) Then args(3) = rows * cols / args(2)
    Dim r: ReDim r(args(2), args(3))
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
' tostring(A) converts all entries of A into strings
Private Function fn_tostring(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Utils_ForceMatrix args(1)
    Dim n As Long, m As Long
    Utils_Size args(1), n, m
    Dim r: ReDim r(n, m)
    For n = 1 To UBound(r, 1)
        For m = 1 To UBound(r, 2)
            r(n, m) = args(1)(n, m) & ""
        Next m
    Next n
    Utils_Conform r
    fn_tostring = r
End Function

' X = if(a,B,C)
'
' X = if(a,B,C) returns B if a evaluates to true; otherwise C.
' Note: B is only evaluated if a is true and C is only evaluated if a is false
Private Function fn_if(args As Variant) As Variant
    Utils_AssertArgsCount args, 3, 3
    fn_if = eval_tree(args(3 + CLng(CBool(eval_tree(args(1))))))
End Function

' X = iferror(A,B)
'
' X = iferror(A,B) returns A if the evaluation of A does not result in a error; then B is returned instead.
Private Function fn_iferror(args As Variant) As Variant
    Utils_AssertArgsCount args, 2, 2
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
    Utils_AssertArgsCount args, 1, 2
    If IsEmpty(args(1)) Then fn_count = 0: Exit Function
    Dim x As Long, i As Long, j As Long
    Dim rows As Long, cols As Long
    Utils_ForceMatrix args(1)
    Utils_Size args(1), rows, cols
    x = Utils_CalcDimDirection(args)
    Dim r: ReDim r(x * rows + (1 - x), (1 - x) * cols + x)
    For i = 1 To rows
        For j = 1 To cols
            r(i * x + (1 - x), j * (1 - x) + x) _
                = r(i * x + (1 - x), j * (1 - x) + x) - CBool(args(1)(i, j))
        Next j
    Next i
    Utils_Conform r
    fn_count = r
End Function

' Y = diff(X)
' Y = diff(X,n)
' Y = diff(X,n,dim)
Private Function fn_diff(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 3
    If IsEmpty(args(1)) Then Exit Function
    Utils_ForceMatrix args(1)
    Dim x As Long: x = Utils_CalcDimDirection(args, 3)
    If UBound(args(1), 1 + x) < 2 Then Exit Function
    Dim i As Long, j As Long
    Dim r: ReDim r(UBound(args(1), 1) - (1 - x), UBound(args(1), 2) - x)
    For i = 2 - x To UBound(args(1), 1)
        For j = 1 + x To UBound(args(1), 2)
            r(i - (1 - x), j - x) = args(1)(i, j) - args(1)(i - (1 - x), j - x)
        Next j
    Next i
    Utils_Conform r
    fn_diff = r
    Dim n As Long: n = Utils_GetOptionalArg(args, 2, 1)
    If n > 1 Then fn_diff = fn_diff(Array(r, n - 1, 1 + x))
End Function

' B = unique(A)
'
' B = unique(A) returns a column vector with all the unique elements of A.
' The values of B are in sorted order.
Private Function fn_unique(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 1
    Dim rows As Long, cols As Long, i As Long, Count As Long, save
    args(1) = fn_reshape(Array(args(1), Empty, 1))
    args(1) = fn_sort(Array(args(1)))
    Utils_ForceMatrix args(1)
    Utils_Size args(1), rows, cols
    ReDim save(1 To rows - 1)
    Count = 1
    For i = 1 To UBound(save)
        save(i) = (0 <> Utils_Compare(args(1)(i, 1), args(1)(i + 1, 1)))
        Count = Count - CLng(save(i))
    Next i
    Dim r: ReDim r(1 To Count, 1)
    r(1, 1) = args(1)(1, 1)
    Count = 2
    For i = 1 To UBound(save)
        If save(i) Then
            r(Count, 1) = args(1)(i + 1, 1)
            Count = Count + 1
        End If
    Next i
    Utils_Conform r
    fn_unique = r
End Function

' B = sort(A)
' B = sort(A,dim)
' B = sort(...,"descend")
' B = sort(...,"indices")
'
' sort() sorts the entries in each row or column.
'
' "descend"  Sort descending
' "indices"  Return sorted indices instead of values
Private Function fn_sort(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 4
    
    ' Get all input parameters
    Dim sortRows As Boolean, ascend As Boolean, returnIndices As Boolean
    sortRows = (1 = Utils_CalcDimDirection(args))
    ascend = Not Utils_IsFlagSet(args, "descend")
    returnIndices = Utils_IsFlagSet(args, "indices")
    
    ' Transpose input matrix if rows must be sorted
    If sortRows Then
        args(1) = WorksheetFunction.Transpose(args(1))
        Utils_Conform args(1)
    End If
    
    ' Make sure the input is a matrix so we can access it like (i,j)
    Utils_ForceMatrix args(1)

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
        Utils_QuickSortCol args(1), indices, 1, rows, j, ascend
    Next j
    
    ' Return the sorted indices if that was specified;
    ' otherwise build and return a matrix with the sorted values
    If returnIndices Then
        fn_sort = indices
    Else
        Dim r:  ReDim r(1 To rows, 1 To cols)
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

' Implementation of the quick-sort algorithm - is a helper for fn_sort()
' Sorts on columns by swapping indices.
' No actual swapping of values in the original matrix is done.
'
' It sorts the column "col" in the range from "first" to "last"
Private Function Utils_QuickSortCol(arr As Variant, indices As Variant, first As Long, last As Long, col As Long, ascend As Boolean)
    If first >= last Then Exit Function
    Dim tmp As Variant
    Dim pivot As Variant: pivot = arr(indices(first, col), col)
    Dim left As Long: left = first
    Dim right As Long: right = last
    Dim ascendprefix As Long: ascendprefix = -1 - 2 * Sgn(ascend)
    While left <= right
        While ascendprefix * Utils_Compare(arr(indices(left, col), col), pivot) < 0
            left = left + 1
        Wend
        While ascendprefix * Utils_Compare(pivot, arr(indices(right, col), col)) < 0
            right = right - 1
        Wend
        If left <= right Then
            tmp = indices(left, col)
            indices(left, col) = indices(right, col)
            indices(right, col) = tmp
            left = left + 1
            right = right - 1
        End If
    Wend
    Utils_QuickSortCol arr, indices, first, right, col, ascend
    Utils_QuickSortCol arr, indices, left, last, col, ascend
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
    Utils_Assert TypeName(args(1)) = "String", "apply(): 1st argument must be an Excel function name."
    Dim i As Long, r1 As Long, c1 As Long, r2 As Long, c2 As Long
    r1 = -1: c1 = -1
    For i = 2 To Utils_Stack_Size(args)
        Utils_ForceMatrix args(i)
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
    Utils_Conform r
    fn_arrayfun = r
End Function

' B = concat(A)
' B = concat(A,joiner)
' B = concat(A,joiner,dim)
'
' B = concat(...) concatenates the elements of A along the first
' dimension whose size does not equal 1.
Private Function fn_concat(args As Variant) As Variant
    Utils_AssertArgsCount args, 1, 3
    Dim i As Long, j As Long, x As Long, joiner As String
    Utils_ForceMatrix args(1)
    x = Utils_CalcDimDirection(args, 3)
    If UBound(args) > 1 Then joiner = args(2)
    Dim r: ReDim r(x * UBound(args(1), 1) + (1 - x), (1 - x) * UBound(args(1), 2) + x)
    For i = 1 To UBound(args(1), 1)
        For j = 1 To UBound(args(1), 2)
            If (1 - x) * i + x * j = 1 Then
                r(x * i + (1 - x), (1 - x) * j + x) = args(1)(i, j)
            Else
                r(x * i + (1 - x), (1 - x) * j + x) = r(x * i + (1 - x), (1 - x) * j + x) & joiner & args(1)(i, j)
            End If
        Next j
    Next i
    Utils_Conform r
    fn_concat = r
End Function

' v = version
'
' Returns a string with the current version of the Q library.
Private Function fn_version(args As Variant) As Variant
    Utils_AssertArgsCount args, 0, 0
    fn_version = VERSION
End Function
