'' Text File: test/01-horizontal-range-index.vb
'' Excel File: 01-horizontal-range-index.xlsx
'' Elapsed Time (parsing + topological sort): 0 s. or 0 min.
HORIZONTAL_RANGE!B1:H1 = [ 0.0 1.0 2.0 3.0 4.0 5.0 6.0 ]
HORIZONTAL_RANGE!B5 = INDEX(HORIZONTAL_RANGE!B1:H1,3)
