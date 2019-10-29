' 
' Text File: test/area-2.vb
' Excel File: test/area-2.xlsx
' Excel Formulas Number: 2
' Elapsed Time (parsing + topological sort): 0 s. or 0 min.
' As Raw Text - Start
' A1 = 11.0
' B1 = 21.0
' A2 = 12.0
' B2 = 22.0
' A3 = 13.0
' B3 = 23.0
' A1 = 11.0
' B1 = 21.0
' A2 = 12.0
' B2 = 22.0
' A3 = 13.0
' B3 = 23.0
' A4 = 14.0
' B4 = 24.0
' UseArea1AndArea2!A1 = INDEX(Area2Name,1,2)
' UseArea1AndArea2!A2 = INDEX(Area1!A1:B3,2,2)
' As Raw Text - End
Area1!A1:B3 = [[11.0 21.0][12.0 22.0][13.0 23.0]]
Area2!Area2Name = [[11.0 21.0][12.0 22.0][13.0 23.0][14.0 24.0]]
UseArea1AndArea2!A2 = INDEX(Area1!A1:B3,2,2)
UseArea1AndArea2!A1 = INDEX(Area2!Area2Name,1,2)
