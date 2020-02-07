# Algorithm

```
⟨Start⟩ ::= = ⟨Formula⟩ | ⟨ArrayFormula⟩ | ⟨Metadata⟩ 
⟨ArrayFormula⟩ ::= {= ⟨Formula⟩ }
⟨Formula⟩ ::= ⟨Constant⟩ | ⟨Reference⟩ | ⟨FunctionCall⟩ | ⟨ParenthesisFormula⟩ | ⟨ConstantArray⟩ | RESERVED_NAME
⟨ParenthesisFormula⟩ ::= ( ⟨Formula⟩ )
⟨Constant⟩ ::= ⟨Number⟩ | TEXT | BOOL | DATETIME | ERROR  
⟨Number⟩::= INT | FLOAT
⟨FunctionCall⟩ ::=  ⟨EXCEL_FUNCTION⟩ | ⟨Unary⟩ | ⟨PercentFormula⟩ | ⟨Binary⟩
```