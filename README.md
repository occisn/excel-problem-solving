# excel-problem-solving

This hobby project uses Microsoft Excel (without VBA) to solve numeric puzzles as those proposed by Project Euler.

When possible, several solutions are presented:  
- use of Excel spreadsheet capabilities (use of rows and columns),  
- use of recursion,  
- use of one-liners, taking advantage of array formulas.

## Table of contents

**Tools:** [recursion](#tools-recursion), [array formulas](#tools-array-formulas), [useful functions](#tools-useful-functions)

**Project Euler problems:**
[1](#project-euler-1-multiples-of-3-or-5), [2](#project-euler-2-even-fibonacci-numbers), ...,  [5](#project-euler-5-smallest-multiple), [6](#project-euler-6-sum-square-difference), ...,  [8](#project-euler-8-largest-product-in-a-series), ...,  [11](#project-euler-11-largest-product-in-a-grid)

## Tools: recursion

Recursion can be achieved by the use of `LET` and `LAMBDA`.

For instance, this formula calculates 10!:
```
= LET(
SUB; LAMBDA(ME;N; IF(N=0; 1; N*ME(ME; N-1)));
SUB(SUB; 10))
```
Value: 3628800

The same with tail-call recursion:
```
=LET(
SUB; LAMBDA(ME;ACC;N; IF(N=0; ACC; ME(ME;ACC*N; N-1)));
SUB(SUB; 1; 10))
```
Value: 3628800


## Tools: array formulas

Create a vertical array:  
- `SEQUENCE(10)` creates a vertical array 1...10  
- trick: `IF({1;0}; 4; 5)` creates a vertical array with two elements: 4 and 5

Return specific element of an array:  
- `INDEX(array; row)`  
- `INDEX(array;; column)`

Operations on arrays:  
- `MAP`  
- `SUM`  
- `REDUCE`  
- `MAX`

Some usual functions accept array and return array.  
For instance:  
- `MID` (substring) with a sequence of position in string; returns a horizontal array

## Tools: useful functions

`OFFSET(cell; 0; 0; 1; 4)` refers to the range spanning from 'cell' with height 1 and width 4  
`OFFSET(cell; 1; 2)` refers to the cell located 1 row below and 2 columns on the right of 'cell' 

substring: `MID`

convert string to number: `VALUE`

**Digits of a number:**  
- number of digits: `LEN`  
- M-th digit of N (counting from right): `MOD(QUOTIENT(N; 10^(M-1)); 10)`  
- M-th digit of N (couting from left): `MOD(QUOTIENT(N; 10^(LEN(N)-M)); 10)`  
- sum of digits of N (with separate identification of each digits):  
`SUM(MAP(SEQUENCE(LEN(N)); LAMBDA(M; MOD(QUOTIENT(N; 10^(M-1)); 10))))`   
- sum of digits of N (via 2-cell accumulator):  
`INDEX(REDUCE(IF({1;0}; 0; N); SEQUENCE(LEN(N)); LAMBDA(ACC;N; IF({1;0}; INDEX(ACC; 1) + MOD(INDEX(ACC; 2); 10); QUOTIENT(INDEX(ACC; 2); 10)))); 1)`
- product of digits of N (via 2-cell accumulator):  
`INDEX(REDUCE(IF({1;0}; 1; N); SEQUENCE(LEN(N)); LAMBDA(ACC;N; IF({1;0}; INDEX(ACC; 1) * MOD(INDEX(ACC; 2); 10); QUOTIENT(INDEX(ACC; 2); 10)))); 1)`

**Digits of a number expressed as a string:**  
- convert string S to horizontal array of digits: `VALUE(MID(S; SEQUENCE(1; LEN(S)); 1))`  
- multiply digits of a number expressed as a string S: `PRODUCT(VALUE(MID(S; SEQUENCE(1; LEN(S)); 1)))`

## Project Euler 1: Multiples of 3 or 5

_If we list all the natural numbers below 10 that are multiples of 3 or 5, we get 3, 5, 6 and 9. The sum of these multiples is 23. Find the sum of all the multiples of 3 or 5 below 1000._
[(source)](https://projecteuler.net/problem=1)

Three solutions are proposed:  
(i) in the form of a spreadsheet,  
(ii) as an one-liner with array formulas,  
(iii) with recursion

Recursion is implemented in the same way as factorial above.

The one-liner with array formulas is essentially the following. It creates a sequence 1...999, filters it, and sums it.
```
=LET(
MULTIPLE_OF_3_OR_5; LAMBDA(N; OR(MOD(N;3)=0; MOD(N;5)=0));
FILTER_MULTIPLE_OF_3_OR_5; LAMBDA(N; IF(MULTIPLE_OF_3_OR_5(N); N; 0));
SUM( MAP( SEQUENCE(999); FILTER_MULTIPLE_OF_3_OR_5)))
```

## Project Euler 2: Even Fibonacci Numbers

_Each new term in the Fibonacci sequence is generated by adding the previous two terms. 
By starting with 1 and 2, the first 10 terms will be: 1, 2, 3, 5, 8, 13, 21, 34, 55, 89, ...
By considering the terms in the Fibonacci sequence whose values do not exceed four million, find the sum of the even-valued terms._ [(source)](https://projecteuler.net/problem=2) 

Two solutions are proposed:  
(i) in the form of a spreadsheet, and  
(ii) as an one-liner with array formulas

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

## Project Euler 5: Smallest Multiple

_2520 is the smallest number that can be divided by each of the numbers from 1 to 10 without any remainder. What is the smallest positive number that is evenly divisible by all of the numbers from 1 to 20?_ [(source)](https://projecteuler.net/problem=5)

Two solutions are proposed:  
(i) as an one-liner with array formulas, and  
(ii) with recursion

The recursion solution is similar to factorial above.

The one-liner with array formulas is essentially the following. It creates a 1...20 sequence, then reduces it with the 'leas common multiple' (LCM) function:
```
=REDUCE(1; SEQUENCE(20); LAMBDA(a;v; LCM(a;v)))
```

## Project Euler 6: Sum Square Difference

_The sum of the squares of the first ten natural numbers is  
1^2 + 2^2 + ... + 10^2 = 385  
The square of the sum of the first ten natural numbers is  
(1 + 2 + ... + 10)^2 = 552 = 3025  
Hence the difference between the sum of the squares of the first ten natural numbers and the square of the sum is 3025 − 385 = 2640.  
Find the difference between the sum of the squares of the first one hundred natural numbers and the square of the sum._  
[(source)](https://projecteuler.net/problem=6)

Two solutions are proposed:  
(i) in the form of a spreadsheet, and  
(ii) as an one-liner with array formulas

The one-liner with array formulas is essentially the following. It creates a 1...100 sequence, then uses `MAP` to square it number by number, and concludes.
```
=SUM( SEQUENCE(100) )^2
      - SUM( MAP(SEQUENCE(100); LAMBDA(a;a^2)) )
```

## Project Euler 8: Largest Product in a Series

_The four adjacent digits in the 1000-digit number that have the greatest product are 9x9x8x9 = 5832.  
73167176531330624919225119674426574742355349194934  
96983520312774506326239578318016984801869478851843  
85861560789112949495459501737958331952853208805511  
12540698747158523863050715693290963295227443043557  
66896648950445244523161731856403098711121722383113  
62229893423380308135336276614282806444486645238749  
30358907296290491560440772390713810515859307960866  
70172427121883998797908792274921901699720888093776  
65727333001053367881220235421809751254540594752243  
52584907711670556013604839586446706324415722155397  
53697817977846174064955149290862569321978468622482  
83972241375657056057490261407972968652414535100474  
82166370484403199890008895243450658541227588666881  
16427171479924442928230863465674813919123162824586  
17866458359124566529476545682848912883142607690042  
24219022671055626321111109370544217506941658960408  
07198403850962455444362981230987879927244284909188  
84580156166097919133875499200524063689912560717606  
05886116467109405077541002256983155200055935729725  
71636269561882670428252483600823257530420752963450  
Find the thirteen adjacent digits in the 1000-digit number that have the greatest product.   What is the value of this product?_  
[(source)](https://projecteuler.net/problem=8)

Two solutions are proposed :  
(i) in the form of a spreadsheet, and  
(ii) as an one-liner with array formulas (with two variants)

## Project Euler 11: Largest Product in a Grid

_In the 20 x 20 grid below, four numbers along a diagonal line have been marked in red.  
08 02 22 97 38 15 00 40 00 75 04 05 07 78 52 12 50 77 91 08  
[...]  
01 70 54 71 83 51 54 69 16 92 33 48 61 43 52 01 89 19 67 48  
The product of these numbers is 26 x 63 x 78 x 14 = 1788696.  
What is the greatest product of four adjacent numbers in the same direction (up, down, left, right, or diagonally) in the 20 x 20 grid?_  
[(source)](https://projecteuler.net/problem=11)

Two solutions are proposed:  
(i) in the form of a spreadsheet;  
(ii) as an one-liner with array formulas

(end of README)
