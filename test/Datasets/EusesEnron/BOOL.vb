' 
' Text File: test/Datasets/EusesEnron/BOOL.vb
' Excel File: test/Datasets/EusesEnron/BOOL.xlsx
' Excel Formulas Number: 2
' Elapsed Time (parsing + topological sort): 0 s. or 0 min.
' As Raw Text - Start
' Boolean!A1 = IF(AND(A3=1,A4=TRUE),A5,A6)
' A3 = 1.0
' Boolean!A4 = TRUE
' A5 = IFTRUE
' A6 = IFFALSE
' As Raw Text - End
A3 = 1.0
A4 = TRUE
A5 = "IFTRUE"
A6 = "IFFALSE"
A1 = IF(AND(A3=1,A4=TRUE),A5,A6)
