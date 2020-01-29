' 
' Text File: test/13-NamedRange.vb
' Excel File: test/13-NamedRange.xlsx
' Excel Formulas Number: 1
' Elapsed Time (parsing + topological sort): 0 s. or 0 min.
' As Raw Text - Start
' A1 = 1.0
' A2 = 2.0
' A3 = 3.0
' A4 = 4.0
' A5 = 5.0
' A6 = 6.0
' NamedRange!A8 = SUM(slist)
' As Raw Text - End
NamedRange!slist = [ 1.0 2.0 3.0 4.0 5.0 6.0 ]
A8 = SUM(NamedRange!slist)
