# excel-problem-solving

This hobby project uses Microsoft Excel (without VBA) to solve numeric puzzles as those proposed by Project Euler.

When possible, two solutions are presented:  
- one using Excel spreadsheet capabilities (use of rows and columns),  
- another using one-liners, taking advantage of array formulas.

Problems solved :

**Project Euler 1: Multiples of 3 and 5** (i) in the form of a spreadsheet, (ii) as an one-liner with array formulas, (iii) with recursion

**Project Euler 2: Even Fibonacci numbers** (i) in the form of a spreadsheet, and (ii) as an one-liner with array formulas

N-th Fibonacci number :
```
LAMBDA(n; IF(n=1; 0; INDEX( REDUCE({0;1}; SEQUENCE(n-1); LAMBDA(x;y; IF({1;0}; INDEX(x;2); SUM(x)))); 2 )))
```

Make Fibonacci sequence:
```
LAMBDA(m
   REDUCE(0;
      SEQUENCE(m);
      LAMBDA(current_array; n;
         IF(n=0; 0;
         IF(n=1; VSTACK(0;1);
            LET(new_fibo; HSTACK(INDEX(current_array;n;1) + INDEX(current_array;n-1;1));
                VSTACK(current_array;new_fibo)))))))
```

(...)

**Project Euler 5: Smallest Multiple** (i) as an one-liner with array formulas, and (ii) with recursion

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
