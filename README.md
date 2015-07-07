Q 
====

###Use MATLAB syntax in Excel

Q is a small, easy-to-use VBA library featuring a single public function, Q(), containing an expression parser.
Q() is able to parse and evaluate a subset of the MATLAB programming language.
It features almost all MATLAB operators, selected standard functions and has complete support for submatrices, '( )', and concatenation, '[ ]'.

###Example usage
 Formula | Result 
---------|--------
`=Q("2+2")` | 4 
`=Q("a+b+c",3,4,5)` | 12
`=Q("a(:,4)",A1:D5)` | The 4th column in cells A1:D5
`=Q("eye(3)")`  |  The 3x3 identity matrix
`=Q("mean(a)",A1:D5)`      |Row vector with the mean of each column in cells A1:D5
`=Q("a.*b",A1:D5,F1:I5)`   | Element wise multiplication of cells A1:D5 and F1:I5
`=Q("a([1 3],end)",A1:D5)` | 2x1 matrix with the last entries in row 1 and 3 of cells A1:D5
`=Q("sort(a)",A1:D5)` | sort each column of cells A1:D5
`=Q("3+4;ans^2")` | 49<br />Multiple expressions are separated by ";" or line break. <br />Variable "ans" always contains the previous result.

###Features
  - All standard operators: :,::,+,-,\*,/,.*,./,^,.^,||,&&,|,&,<,<=,>,>=,==,~=,~,'
  - Most used functions: <i>eye,zeros,ones,sum,cumsum,cumprod,prod,
    mean,median,prctile,std,isequal,fix,rand,randn,repmat,reshape,find,sort,sqrt,exp,inv</i>...
  - Indexing via fx. `a(2,:)` or `a(5,3:end)`
  - Concatenate matrices with '[ ]', i.e. `[ a b; c d]`
  - Excel functions: <i>if,iferror</i>
  - Prefix function calls with ! to call external VBA functions not found in Q.

###How to use
1. Open up Excel
2. Press Alt+F11 to open the VBA editor
3. Choose Insert -> Module
4. Copy-paste the contents of [Qlib.bas](https://raw.githubusercontent.com/nielsls/Q/master/Qlib.bas) to your new module
5. Well done! Go test in Excel by typing `=Q("2+2")` in a cell

2015, Niels Lykke Sørensen
