# Algorithm

## Title
TODO
## Abstract
TODO
### A Analyzer
TODO
### P Parser
TODO
### S Topological Sorter
TODO
### C Compiler (or Translator)
TODO

Basic rules

```
⟨Start⟩ ::= = ⟨Formula⟩ | ⟨ArrayFormula⟩ | ⟨Metadata⟩ 
⟨ArrayFormula⟩ ::= {= ⟨Formula⟩ }
⟨Formula⟩ ::= ⟨Constant⟩ | ⟨Reference⟩ | ⟨FunctionCall⟩ | ⟨ParenthesisFormula⟩ | ⟨ConstantArray⟩ | RESERVED_NAME
⟨ParenthesisFormula⟩ ::= ( ⟨Formula⟩ )
⟨Constant⟩ ::= ⟨Number⟩ | TEXT | BOOL | DATETIME | ERROR  
⟨Number⟩::= INT | FLOAT
⟨FunctionCall⟩ ::=  ⟨EXCEL_FUNCTION⟩ | ⟨Unary⟩ | ⟨PercentFormula⟩ | ⟨Binary⟩
```

Constant(s)
```
⟨Constant⟩ ::= ⟨Number⟩ | TEXT | BOOL | DATETIME | ERROR
⟨Number⟩::= INT | FLOAT
```
