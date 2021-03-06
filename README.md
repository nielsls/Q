# Q 

### Use MATLAB syntax in Excel
Q is a small, easy-to-use VBA library featuring a single public function, Q().<br/>
Q() is able to parse and evaluate a subset of the MATLAB programming language.
It features almost all MATLAB operators, selected standard functions and has complete support for matrix indexing, '( )', and concatenation, '[ ]'.

MATLAB is NOT required and there are no external dependencies. 

### Example usage
 Formula | Result 
---------|--------
`=Q("2+2")` | 4 
`=Q("A+B+C",3,4,5)` | 12
`=Q("A(:,4)",A1:D5)` | The 4th column in cells A1:D5
`=Q("eye(3)")`  |  The 3x3 identity matrix
`=Q("mean(A)",A1:D5)`      |Row vector with the mean of each column in cells A1:D5
`=Q("A.*B",A1:D5,F1:I5)`   | Element wise multiplication of cells A1:D5 and F1:I5
`=Q("A([1 3],end)",A1:D5)` | 2x1 matrix with the last entries in row 1 and 3 of cells A1:D5
`=Q("sort(A)",A1:D5)` | Sort each column of cells A1:D5
`=Q("3+4;ans^2")` | 49<br />Multiple expressions are separated by ";" or a line break. <br />Variable "ans" always contains the previous result.

### Features
* All standard operators: `: + - * / ^ .* ./ .^ || && | & < <= > >= == ~=` 
* Most used functions: `eye,zeros,ones,sum,cumsum,cumprod,prod,mean,median,prctile, std,isequal,fix,rand,randn,repmat,reshape,find,sort,sqrt,exp,inv...`
* Indexing via e.g. `A(2,:)` or `A(5,3:end)`
* Concatenate matrices with '[ ]', i.e. `[ A B; C D ]`
* Multiple expressions separated by ";" or a line break.
* Excel functions: `if,iferror`
* Prefix function calls with ! to call external VBA functions not found in Q.

### How to use
1. Open up Excel
2. Press Alt+F11 to open the VBA editor
3. Choose Insert -> Module
4. Copy-paste the contents of [Qlib.bas](https://raw.githubusercontent.com/nielsls/Q/master/Qlib.bas) to your new module
5. Well done! Go test in Excel by typing `=Q("2+2")` in a cell

2017, Niels Lykke Sørensen
