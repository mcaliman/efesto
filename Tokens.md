# Tokens

## Priorities

| Token                  | Priority | 
| :---                   | :---     | 
| SHEET                  | 5        |
| SHEET_QUOTED           | 5        |
| FILE                   | 5        |
| EXCEL_FUNCTION         | 5        |
| REF_FUNCTION           | 5        |
| REF_FUNCTION_COND      | 5        |
| UDF                    | 4        |
| NAME_PREFIXED          | 3        |
| CELL                   | 2        |
| MULTIPLE_SHEETS        | 1        |
| BOOL                   | 0        |
| NUMBER                 | 0        | 
| STRING                 | 0        |
| DDECALL                | 0        |
| ERROR                  | 0        |
| ERROR-REF              | 0        |
| FILEPATH               | 0        |
| HORIZONTAL_RANGE       | 0        |
| VERTICAL_RANGE         | 0        |
| FILENAME               | -1       |            
| RESERVED_NAME          | -1       |
| NAME                   | -2       |
| SR_COLUMN              | -3       |


                  

* BOOL 
    * Description: Boolean literal 
    * Contents: TRUE | FALSE 
* CELL 
    * Cell reference 
    * $? [A-Z]+ $? [0-9]+ 
* DDECALL 
    * Dynamic Data Exchange link 
    * ’ ([^ ’] | ”)+ ’ 
* ERROR 
    * Error literal 
    * '#NULL!' | '#DIV/0!' | '#VALUE!' | '#NAME?' | '#NUM!' | '#N/A' 
* ERROR_REF 
    * Reference error literal 
    * '#REF!'
* EXCEL_FUNCTION 
    * Excel built-in function 
    * (Any entry from the function list3) \( 
* FILE 
    * External file reference using number 
    * \[ [0-9]+ \] 
* FILENAME 
    * External file reference using name 
    * \[ 4+ \] 
* FILEPATH 
    * Windows file path 
    * [A-Z] : \\ (4+ \\)* 
* HORIZONTAL_RANGE 
    * Range of rows 
    * $? [0-9]+ : $? [0-9]+ 
* MULTIPLE_SHEETS 
    * Multiple sheet references 
    * ((2+ : 2+)|( ’ (3 | ”)+ : (3 | ”)+ ’ )) ! 
* NAME 
    * User Defined Name 
    * [A-Z_\\][A-Z0-9\\_.1]* 
* NAME_PREFIXED 
    * User defined name which starts with a string that could be another token 
    * (TRUE | FALSE | [A-Z]+[0-9]+) [A-Z0-9_.1]+ 
* NUMBER 
    * An integer, floating point or scientific notation number literal 
    * [0-9]+ ,? [0-9]* (e [0-9]+)? 
* REF_FUNCTION 
    * Excel built-in reference function 
    * (INDEX | OFFSET | INDIRECT)\( 
* REF_FUNCTION_COND 
    * Excel built-in conditional reference function 
    * (IF | CHOOSE)\( 
* RESERVED_NAME 
    * An Excel reserved name 
    * _xlnm\. [A-Z_]+ 
* SHEET 
    * The name of a worksheet 
    * 2+ ! 
* SHEET_QUOTED 
    * Quoted worksheet name 
    * 3+ ’ ! 
* STRING 
    * String literal 
    * " ([^ "] | "")* " 
* SR_COLUMN 
    * Structured reference column 
    * \[ [A-Z0-9\\_.1]+ \] 
* UDF 
    * User Defined Function 
    * (_xll\.)? [A-Z_\][A-Z0-9_\\.1]* ( 
* VERTICAL_RANGE 
    * Range of columns 
    * $? [A-Z]+ : $? [A-Z]+ 
    
# Placeholder character 
Placeholder for Specification
* 1 Extended characters Non-control Unicode characters x80 and up
* 2 Sheet characters Any character except ’ * [ ] \ : / ? ( ) ; { } # " = < > & + - * / ^ % , ␣
* 3 Enclosed sheet characters Any character except ’ * [ ] \ : / ?
* 4 Filename characters Any character except " * [ ] \ : / ? < > j    