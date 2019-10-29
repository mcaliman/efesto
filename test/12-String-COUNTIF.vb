' 
' Text File: test/12-String-COUNTIF.vb
' Excel File: test/12-String-COUNTIF.xlsx
' Excel Formulas Number: 1
' Elapsed Time (parsing + topological sort): 0 s. or 0 min.
' As Raw Text - Start
' C1 = 1.0
' D1 = 2.0
' E1 = 3.0
' COUNTIF!B4 = COUNTIF(C1:E1,">=2")
' As Raw Text - End
C1:E1 = [ 1.0 2.0 3.0 ]
B4 = COUNTIF(C1:E1,">=2")
