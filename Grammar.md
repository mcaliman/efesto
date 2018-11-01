# Pseudo BNF Grammar

⟨Start⟩ ::= '=' ⟨Formula⟩ | ⟨ArrayFormula⟩ | ⟨Metadata⟩ 
⟨ArrayFormula⟩ ::= '{=' ⟨Formula⟩ '}'

⟨Formula⟩ ::= ⟨Constant⟩ | ⟨Reference⟩ | ⟨FunctionCall⟩ | ⟨ParenthesisFormula⟩ | ⟨ConstantArray⟩
| RESERVED_NAME

⟨ParenthesisFormula⟩ ::= '(' ⟨Formula⟩ ')'
⟨Constant⟩ ::= INT | FLOAT | TEXT | BOOL | DATETIME | ERROR  

⟨FunctionCall⟩ ::=  ⟨ExcelFunction⟩ | ⟨Unary⟩ | ⟨PercentFormula⟩ | ⟨Binary⟩

-- todo begin
⟨Arguments⟩ ::= ϵ | ⟨Argument⟩ { ‘,’ ⟨Argument⟩ }
⟨Argument⟩ ::= ⟨Formula⟩ | ϵ

Impl.note: abstract class Argument extends Formula

-- todo end



⟨Unary⟩ = ⟨Plus⟩  | ⟨Minus⟩ 
⟨Plus⟩  ::= '+' ⟨Formula⟩ 
⟨Minus⟩ ::= '-' ⟨Formula⟩ 

⟨Binary⟩    ::= ⟨Add⟩ | ⟨Sub⟩ | ⟨Mult⟩ | ⟨Divide⟩ | ⟨Lt⟩ | ⟨Gt⟩ | ⟨Eq⟩ | ⟨Leq⟩ | ⟨GtEq⟩ | ⟨Neq⟩ 
| ⟨Concat⟩
| ⟨Power⟩

⟨Add⟩      ::= ⟨Formula⟩ '+' ⟨Formula⟩
⟨Sub⟩      ::= ⟨Formula⟩ '-' ⟨Formula⟩
⟨Mult⟩     ::= ⟨Formula⟩ '*' ⟨Formula⟩
⟨Divide⟩   ::= ⟨Formula⟩ '/' ⟨Formula⟩
⟨Lt⟩       ::= ⟨Formula⟩ '<' ⟨Formula⟩
⟨Gt⟩       ::= ⟨Formula⟩ '>' ⟨Formula⟩
⟨Eq⟩       ::= ⟨Formula⟩ '=' ⟨Formula⟩
⟨Leq⟩      ::= ⟨Formula⟩ '<=' ⟨Formula⟩
⟨GtEq⟩     ::= ⟨Formula⟩ '>=' ⟨Formula⟩
⟨Neq⟩      ::= ⟨Formula⟩ '<>' ⟨Formula⟩

⟨Concat⟩ ::= ⟨Formula⟩ ‘&’ ⟨Formula⟩
⟨Power⟩ ::= ⟨Formula⟩ ‘^’ ⟨Formula⟩

⟨PercentFormula⟩ ::= ⟨Formula⟩ ‘%’

⟨Reference⟩ ::= 
⟨ReferenceItem⟩
| ⟨RangeReference⟩
| ⟨Intersection⟩ 
| ‘(’ ⟨Union⟩ ‘)’
| ‘(’ ⟨Reference⟩ ‘)’
| ⟨PrefixReferenceItem⟩
| ⟨Prefix⟩ UDF* ⟨Arguments⟩ ‘)’(notImp.)  
| ⟨DynamicDataExchange⟩(notImp.)

⟨RangeReference⟩ ::= ⟨Reference⟩ ‘:’ ⟨Reference⟩ 

⟨Intersection⟩# ::= ⟨Reference⟩ ‘ ’ ⟨Reference⟩       //Implemented as "Binary"

⟨Union⟩ ::= ⟨Reference⟩# | ⟨Reference⟩ ‘,’ ⟨Union⟩     //Implemented as "Binary"

⟨PrefixReferenceItem⟩# ::= ⟨Prefix⟩ ⟨ReferenceItem⟩  

⟨ReferenceItem⟩ ::= CELL
| ⟨NamedRange⟩
//| ⟨StructuredReference⟩
| 
//| VERTICAL_RANGE
//| HORIZONTAL_RANGE
//| UDF ⟨Arguments⟩ ‘)’
| ERROR_REF
| ⟨ReferenceFunction⟩
| ⟨ConditionalReferenceFunction⟩  
 
 ⟨ReferenceFunction⟩ ::= REFERENCE_FUNCTION //Not.Correctly implemented, inherits from Funzion (not ReferenceItem) 
 
 REFERENCE_FUNCTION ::= Excel built-in reference function INDEX | OFFSET | INDIRECT

 ⟨ConditionalReferenceFunction⟩  ::= REF_FUNCTION_COND //Not.Correctly implemented, inherits from Function (not ReferenceItem) 
 REF_FUNCTION_COND ::= IF | CHOOSE
 
 
⟨Prefix⟩* ::= SHEET
| ‘’’ SHEET_QUOTED
| ⟨File⟩ SHEET
| ‘’’ ⟨File⟩ SHEET_QUOTED
| FILE* ‘!’
| MULTIPLE_SHEETS
| ⟨File⟩ MULTIPLE_SHEETS

⟨ExcelFunction⟩ ::= ⟨Function⟩ ⟨Arguments⟩ ‘)’ | ⟨ExcelBuiltInFunction⟩

⟨Function⟩ ::= FUNCTION | UDF

⟨ExcelBuiltInFunction⟩ ::= EXCEL-FUNCTION '(' ⟨Arguments⟩ ')'

EXCEL-FUNCTION ::= 'ABS' | ... Excel built-in function (Any entry from the file Excel-built-in-function-list.md) 




---



   
       

    

    













Lexical tokens used in grammar

RESERVED-NAME An Excel reserved name    _xlnm\. [A-Z_]+
CELL          Cell reference             $? [A-Z]+ $? [0-9]+        2


REFERENCE-FUNCTION

Metadata è d supporto


