'' Text File: test/02-vertical-range-index.vb
'' Excel File: 02-vertical-range-index.xlsx
'' Excel Formulas Number: 1
'' Elapsed Time (parsing + topological sort): 1 s. or 0 min.
VERTICAL_RANGE!I1:I7 = [ 0.0 1.0 2.0 3.0 4.0 5.0 6.0 ]
VERTICAL_RANGE!B9 = INDEX(VERTICAL_RANGE!I1:I7,4)
