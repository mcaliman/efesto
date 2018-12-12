'' Text File: test/14-Boolean.vb
'' Excel File: 14-Boolean.xlsx
'' Elapsed Time (parsing + topological sort): 2 s. or 0 min.
Boolean!A3 = 1.0
Boolean!A4 = TRUE
Boolean!A5 = "IFTRUE"
Boolean!A6 = "IFFALSE"
Boolean!A1 = IF(AND(Boolean!A3=1,Boolean!A4=TRUE),Boolean!A5,Boolean!A6)
