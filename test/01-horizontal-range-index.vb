' 
' Text File: test/01-horizontal-range-index.vb
' Excel File: test/01-horizontal-range-index.xlsx
' Excel Formulas Number: 1
' Elapsed Time (parsing + topological sort): 0 s. or 0 min.
' As Raw Text - Start
' B1 = 0.0
' C1 = 1.0
' D1 = 2.0
' E1 = 3.0
' F1 = 4.0
' G1 = 5.0
' H1 = 6.0
' HORIZONTAL_RANGE!B5 = INDEX(B1:H1,3)
' A7 = 
' As Raw Text - End
B1:H1 = [ 0.0 1.0 2.0 3.0 4.0 5.0 6.0 ]
B5 = INDEX(B1:H1,3)
