' 
' Text File: test/HORIZONTAL_RANGE.vb
' Excel File: test/HORIZONTAL_RANGE.xlsx
' Excel Formulas Number: 1
' Elapsed Time (parsing + topological sort): 0 s. or 0 min.
' As Raw Text - Start
' A1 = 1.0
' B1 = 3.0
' C1 = 6.0
' D1 = 8.0
' Foglio1!A3 = MATCH(3,A1:D1,0)
' As Raw Text - End
A1:D1 = [ 1.0 3.0 6.0 8.0 ]
A3 = MATCH(3,A1:D1,0)
