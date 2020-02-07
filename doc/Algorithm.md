# Algorithm

## Title
TODO
## Abstract
TODO
## A Analyzer
TODO
## P Parser
TODO
## S Topological Sorter
TODO
## C Compiler (or Translator) 
TODO

### Constant
TODO
⟨Constant⟩ ::= ⟨Number⟩ | TEXT | BOOL | DATETIME | ERROR  
⟨Number⟩::= INT | FLOAT

INT is not used, all Numbers are FLOAT (e.g. 10 --> 10.0 where --> is a relation)

assume 
    C is a cell address (e.g. A15), float is a float value 
so
    C = float --> (def C float) 
in example
    A1 = 10.0 --> (def A1 10.0)
    
for TEXT terminal /  lexical token
C = "Text" --> (def C "Text")

for BOOL values
C = boolean --> (def C Boolean/boolean)
e.g.
C = TRUE --> (def C Boolean/TRUE) 

for DATETIME values
C = datetime (in Excel date time value is implemented as numbers)  so we can detect this format property and 
use a clojure macro to convert as Clojure/Java Date/Time/LocalDateTime value
C = datetime --> (def C excel-date(datetime))

for ERROR type
C = #err 
where 
    C is cell address like A15 and #err is ERROR like #REF      
C = #err --> (def C #err)

---
### Binary Operation
TODO


⟨Binary⟩   ::= ⟨Add⟩ | ⟨Sub⟩ | ⟨Mult⟩ | ⟨Divide⟩ | ⟨Lt⟩ | ⟨Gt⟩ | ⟨Eq⟩ | ⟨Leq⟩ | ⟨GtEq⟩ | ⟨Neq⟩  | ⟨Concat⟩ | ⟨Power⟩
⟨Add⟩      ::= ⟨Formula⟩+⟨Formula⟩
⟨Sub⟩      ::= ⟨Formula⟩-⟨Formula⟩
⟨Mult⟩     ::= ⟨Formula⟩*⟨Formula⟩
⟨Divide⟩   ::= ⟨Formula⟩/⟨Formula⟩
⟨Lt⟩       ::= ⟨Formula⟩<⟨Formula⟩
⟨Gt⟩       ::= ⟨Formula⟩>⟨Formula⟩
⟨Eq⟩       ::= ⟨Formula⟩=⟨Formula⟩
⟨Leq⟩      ::= ⟨Formula⟩<=⟨Formula⟩
⟨GtEq⟩     ::= ⟨Formula⟩>=⟨Formula⟩
⟨Neq⟩      ::= ⟨Formula⟩<>⟨Formula⟩
⟨Concat⟩   ::= ⟨Formula⟩&⟨Formula⟩
⟨Power⟩    ::= ⟨Formula⟩^⟨Formula⟩

for the subset of binary operation
operator = + | - | * | / | < | > | <= | >= | = 
assume C = A op B --> (def C (op A B))

case operator = <> (NotEq)
C = A <> B --> (def C (not= A B))

for the operator Power ^ 
C = A '^' B --> (def C (Math/pow A B))

case operator Concat &
C = A '&' B --> (def C (str A B))
 
---
### Range and Cell Reference
TODO
---
### Conditional Reference Functions (IF and CHOOSE functions)
TODO
---
### Reference Functions (IF and CHOOSE functions)
TODO


