# excel-problem-solving

This hobby project uses Microsoft Excel (without VBA) to solve numeric puzzles as those proposed by Project Euler.

When possible, two solutions are presented:  
- one using Excel spreadsheet capabilities (use of rows and columns),  
- another using one-liners, taking advantage of array formulas.

Problems solved :

**Project Euler 1: Multiples of 3 and 5** (i) in the form of a spreadsheet, (ii) as an one-liner with array formulas, (iii) with recursion

**Project Euler 2: Even Fibonacci numbers** (i) in the form of a spreadsheet, and (ii) as an one-liner with array formulas

Fibonacci :
```
LAMBDA(N; IF(N=1; 0; INDEX( REDUCE({0;1}; SEQUENCE(N-1); LAMBDA(x;y; IF({1;0}; INDEX(x;2); SUM(x)))); 2 )))
```

(...)

**Project Euler 5: Smallest Multiple** as an one-liner with array formulas

**Project Euler 6: Sum Square Difference** (i) in the form of a spreadsheet, and (ii) as an one-liner with array formulas

(...)

**Project Euler 8: Largest Product in a Series** (i) in the form of a spreadsheet, and (ii) as an one-liner with array formulas

String to digits array:
```
LAMBDA(S;VALUE(MID(S;SEQUENCE(1;LEN(S));1)));
```

(...)

**Project Euler 11: Largest Product in a Grid** (i) in the form of a spreadsheet, and (ii) as an one-liner with array formulas

Note : functions working on arrays: SEQUENCE, BITOR, SUM, MOD...

Working with lambda: MAP, REDUCE, LAMBDA

Useful: LET

Useful functions: OFFSET

(end of README)
