# Excel_avg
auto calculate Excel‘s average

        const int minRow = 1;
        const int maxRow = 999;

        const int minColumn = 1;
        const int maxColumn = 99;

【1】 ------------------------------  "a b.xlsx"

search for an excel file "XXX.xlsx"
 1   1   2   3
[a]  4   5   6           <---
 3   7   8   9
[b]  1   2   3           <---
 5   4   5   6
 6   7   8   9

Output average from row [a] to row [b] to specified excel "a b.xlsx"
2020/11/29 210101
XXX.xlsx    4   5   6

【2】 ------------------------------  "a .xlsx"
search for an excel file "YYY.xlsx"
 1   1   2   3
[a]  4   5   6           <---
 3   7   8   9
 4   1   2   3           <---

Output average from row [a] to the last line to specified excel "a .xlsx"
2020/11/29 210101
XXX.xlsx    4   5   6
