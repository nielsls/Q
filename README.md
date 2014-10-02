Q
=

A MATLAB-like matrix parser for Microsoft Excel

Q features a single public function, Q(), containing an expression parser.
The parser is able to parse and evaluate a subset of the MATLAB programming language.
It features almost all MATLAB operators, selected standard functions
and has complete support for submatrices, '()', and concatenation, '[]'.

Example usage:
  - =Q("2+2")                -> 4
  - =Q("a+b+c",3,4,5)        -> 12
  - =Q("eye(3)")             -> the 3x3 identity matrix
  - =Q("mean(a)",A1:D5)      -> row vector with the mean of each column in cells A1:D5
  - =Q("a.*b",A1:D5,F1:I5)   -> element wise multiplication of cells A1:D5 and F1:I5
  - =Q("a([1 3],end)",A1:D5) -> 2x1 matrix with the last entries in row 1 and 3 of cells A1:D5

Supported MATLAB-like features:
  - All standard operators: :,::,+,-,*,/,.*,./,^,.^,||,&&,|,&,<,<=,>,>=,==,~=,~,'
  - Most used functions: eye,zeros,ones,sum,cumsum,cumprod,prod,
    mean,median,prctile,std,isequal,fix,rand,randn,find,sqrt,exp,inv...
  - Indexing via fx. a(2,:) or a(5,3:end)
  - Concatenate matrices with '[]', i.e. [ a b; c d]

2014, Niels Lykke Sørensen, niesre@danskebank.dk
