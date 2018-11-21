'' Text File: test/03-range-index.vb
'' Excel File: 03-range-index.xlsx
'' Elapsed Time (Parsing+Topological Sort): 0 s. or 0 min.
RANGE!A1:B6 = [[1.1 1.2][2.1 2.2][3.1 3.2][4.1 4.2][5.1 5.2][6.1 6.2]]
RANGE!A10 = INDEX(RANGE!A1:B6,2,2)
