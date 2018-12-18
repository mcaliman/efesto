'' Text File: test/12-String-COUNTIF.vb
'' Excel File: 12-String-COUNTIF.xlsx
'' Excel Formulas Number: 1
'' Elapsed Time (parsing + topological sort): 2 s. or 0 min.
COUNTIF!C1:E1 = [ 1.0 2.0 3.0 ]
COUNTIF!B4 = COUNTIF(COUNTIF!C1:E1,">=2")
