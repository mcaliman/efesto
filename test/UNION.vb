' 
' Text File: test/UNION.vb
' Excel File: test/UNION.xlsx
' Excel Formulas Number: 1
' Elapsed Time (parsing + topological sort): 0 s. or 0 min.
' As Raw Text - Start
' A1 = 1.0
' B1 = 2.0
' A2 = 3.0
' B2 = 5.0
' C2 = 6.0
' B3 = 7.0
' C3 = 8.0
' Foglio1!A5 = SUM(A1:B2, C2:C3)
' As Raw Text - End
A1:B2 = [[1.0 2.0][3.0 5.0]]
C2:C3 = [ 6.0 8.0 ]
A5 = SUM(A1:B2,C2:C3)
