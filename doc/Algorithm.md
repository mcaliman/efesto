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

1. Constant
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

2. Binary Operation
TODO

3. Range and Cell Reference
TODO

4. Conditional Reference Functions (IF and CHOOSE functions)
TODO

5. Reference Functions (IF and CHOOSE functions)
TODO


