# PML (Programmable Macro Language) - Complete Coding Reference

**For E3D/PDMS Design Software**

---

## Table of Contents

1. [Introduction to PML](#1-introduction-to-pml)
2. [Language Basics](#2-language-basics)
3. [Data Types and Variables](#3-data-types-and-variables)
4. [Operators and Expressions](#4-operators-and-expressions)
5. [Control Flow Structures](#5-control-flow-structures)
6. [Arrays and Collections](#6-arrays-and-collections)
7. [Functions and Methods](#7-functions-and-methods)
8. [Object-Oriented Programming](#8-object-oriented-programming)
9. [Macros](#9-macros)
10. [File Operations](#10-file-operations)
11. [Error Handling](#11-error-handling)
12. [Forms and GUI](#12-forms-and-gui)
13. [Gadgets Reference](#13-gadgets-reference)
14. [Menus](#14-menus)
15. [Form Layout](#15-form-layout)
16. [Database Integration](#16-database-integration)
17. [Best Practices](#17-best-practices)
18. [Complete Function Reference](#18-complete-function-reference)

---

## 1. Introduction to PML

### 1.1 What is PML?

PML (Programmable Macro Language) is a file-based interpreted language for AVEVA products (E3D, PDMS). It exists in three versions:

- **PML1**: Basic macro/command language with form capabilities
- **PML2**: Advanced object-oriented version (current standard)
- **PML.NET**: Integration with Microsoft .NET controls

### 1.2 Key Features

- **Variable Types**: STRING, REAL, BOOLEAN, ARRAY
- **Object-Oriented**: Methods, functions, user-defined types
- **Dynamic Loading**: Forms, functions, objects loaded on demand
- **GUI Support**: Complete forms and menus system
- **Database Integration**: Direct access to PDMS/E3D database elements

### 1.3 Character Format

- **Unicode Support**: UTF-8 internal format (from PDMS 12.1+)
- **File Format**: UTF-8 with BOM or strict ASCII (decimal 32-127)
- **Unicode in Code**: Can use Unicode in variable names, form names, method names

---

## 2. Language Basics

### 2.1 Case Independence

PML is **case-insensitive** for keywords and variable names:

```pml
!MyVar = 100
!myvar = 200  -- Same variable
!MYVAR = 300  -- Still the same variable
```

### 2.2 Comments

```pml
-- This is a single-line comment

$* This is also a comment

/* This is a
   multi-line
   comment */
```

### 2.3 Special Characters

#### The Dollar Sign ($)

- **Variable Expansion**: `$!varname` or `$!!varname`
- **Macro Parameters**: `$1`, `$2`, ... `$9`
- **Text Delimiters**: `$<` and `$>`
- **Command Continuation**: `$/`

```pml
!name = 'Fred'
$P Name is: $!name  -- Output: Name is: Fred

$M/macro.mac $1 $2 $3  -- Call macro with parameters
```

### 2.4 Text Delimiters

Multiple options for string delimiters:

```pml
!text1 = 'Hello World'
!text2 = |Hello World|
!text3 = $<Hello World$>  -- Useful for macro parameters
```

### 2.5 Filename Extensions

- `.mac` - PML1 macros
- `.pmlobj` - PML2 object type definitions
- `.pmlfnc` - PML2 function definitions
- `.pmlfrm` - PML2 form definitions
- `.pmlcmd` - PML command sequences

### 2.6 Abbreviations

PML supports command abbreviations (minimum unique prefix):

```pml
NEW EQUIP  -- Can be abbreviated to:
N EQ
```

---

## 3. Data Types and Variables

### 3.1 Built-in Types

#### STRING
```pml
!name = 'John Doe'
!path = 'C:\temp\file.txt'
!empty = ''
```

#### REAL
```pml
!length = 100.5
!diameter = 50
!pi = 3.14159
!temp = -40.0
```

#### BOOLEAN
```pml
!flag = TRUE
!isValid = FALSE
```

#### ARRAY
```pml
!myArray = ARRAY()
!numbers = ARRAY(1, 2, 3, 4, 5)
```

### 3.2 Variable Scope

#### Local Variables (!)
```pml
!localVar = 'Local'  -- Only accessible in current scope
```

#### Global Variables (!!)
```pml
!!globalVar = 'Global'  -- Accessible everywhere
```

### 3.3 Creating Variables

#### PML1 Style
```pml
VAR !name NAME  -- Get current element name
VAR !pos POS IN WORLD  -- Get position
VAR !value (23 * 1.8 + 32)  -- Calculate value
VAR !list COLLECT ALL ELBOW FOR CE  -- Collect database elements
```

#### PML2 Style
```pml
!name = NAME  -- Direct assignment
!pos = POS IN WORLD
!value = 23 * 1.8 + 32
```

### 3.4 Variable Naming Rules

- Maximum 16 characters
- Can contain letters and numbers
- **Never start with a number**
- **Never use period (.) in names**
- Case-insensitive
- Use meaningful names

**Good Examples:**
```pml
!elementName
!pipe_diameter
!count01
!isValid
```

**Bad Examples:**
```pml
!1stElement    -- Starts with number
!pipe.dia      -- Contains period
!x             -- Not descriptive
```

### 3.5 Variable Types and Conversion

```pml
-- Type checking
!type = !myVar.ObjectType()  -- Returns 'STRING', 'REAL', 'BOOLEAN', 'ARRAY'

-- String to Real
!numText = '123.45'
!number = !numText.Real()

-- Real to String
!value = 100.5
!text = !value.String()

-- Boolean to String
!flag = TRUE
!text = !flag.String()  -- Returns 'TRUE'
```

### 3.6 UNSET and UNDEFINED

```pml
-- UNSET: Variable declared but no value assigned
!var = UNSET
IF !var.IS(UNSET) THEN
  $P Variable is unset
ENDIF

-- UNDEFINED: Variable not declared
IF !undeclaredVar.IS(UNDEFINED) THEN
  $P Variable is undefined
ENDIF
```

---

## 4. Operators and Expressions

### 4.1 Arithmetic Operators

```pml
!a = 10
!b = 3

!sum = !a + !b        -- Addition: 13
!diff = !a - !b       -- Subtraction: 7
!prod = !a * !b       -- Multiplication: 30
!quot = !a / !b       -- Division: 3.333...
!neg = -!a            -- Negation: -10
```

### 4.2 Comparison Operators

```pml
!a EQ !b    -- Equal
!a NE !b    -- Not equal
!a GT !b    -- Greater than
!a LT !b    -- Less than
!a GE !b    -- Greater than or equal
!a LE !b    -- Less than or equal
```

### 4.3 Logical Operators

```pml
!result = !a AND !b   -- Logical AND
!result = !a OR !b    -- Logical OR
!result = NOT !a      -- Logical NOT
```

### 4.4 String Operations

#### Concatenation
```pml
!first = 'Hello'
!second = 'World'
!full = !first & ' ' & !second  -- 'Hello World'
!full = !first + ' ' + !second  -- Also works

-- PML1 style
VAR !full ('$!first' + ' ' + '$!second')
```

#### String Methods
```pml
!text = 'Hello World'

!len = !text.Length()           -- 11
!upper = !text.Upcase()         -- 'HELLO WORLD'
!lower = !text.Downcase()       -- 'hello world'
!sub = !text.Substring(1, 5)    -- 'Hello'
!pos = !text.Match('World')     -- 7
```

### 4.5 Mathematical Functions

```pml
!angle = 45
!rad = !angle * (3.14159 / 180)

!sine = SIN(!rad)
!cosine = COS(!rad)
!tangent = TAN(!rad)
!arcsine = ASIN(0.5)
!arccosine = ACOS(0.5)
!arctangent = ATAN(1.0)

!square = POW(5, 2)       -- 25
!squareRoot = SQR(25)     -- 5
!absolute = ABS(-10)      -- 10
!integer = INT(3.7)       -- 3
!nearest = NINT(3.7)      -- 4
!logarithm = LOG(100)     -- Natural log
!log10 = ALOG(100)        -- Log base 10
```

### 4.6 Operator Precedence

From highest to lowest:

1. `()` - Parentheses
2. `*` `/` - Multiplication, Division
3. `+` `-` - Addition, Subtraction
4. `EQ` `NE` `GT` `LT` `GE` `LE` - Comparison
5. `NOT` - Logical NOT
6. `AND` - Logical AND
7. `OR` - Logical OR

```pml
!result = (60 * 2 / 3 + 5)  -- = 45
-- Evaluated as: ((60 * 2) / 3) + 5
```

### 4.7 Units in Expressions

```pml
!length = 1000MM        -- Millimeters
!length2 = 1M           -- Meters
!diameter = 6IN         -- Inches
!temp = 100DEG          -- Degrees
```

---

## 5. Control Flow Structures

### 5.1 IF Statements

#### Basic IF
```pml
IF !value GT 100 THEN
  $P Value is greater than 100
ENDIF
```

#### IF-ELSE
```pml
IF !value GT 100 THEN
  $P Value is greater than 100
ELSE
  $P Value is 100 or less
ENDIF
```

#### IF-ELSEIF-ELSE
```pml
IF !value GT 100 THEN
  $P Value is greater than 100
ELSEIF !value EQ 100 THEN
  $P Value equals 100
ELSE
  $P Value is less than 100
ENDIF
```

#### Nested IF
```pml
IF !value1 GT 0 THEN
  IF !value2 GT 0 THEN
    $P Both values are positive
  ELSE
    $P Only value1 is positive
  ENDIF
ENDIF
```

#### IF TRUE Expression (Inline)
```pml
!result = IF !condition THEN !value1 ELSE !value2
```

### 5.2 DO Loops

#### Basic DO Loop
```pml
DO !i FROM 1 TO 10
  $P Iteration: $!i
ENDDO
```

#### DO Loop with Step
```pml
DO !i FROM 0 TO 100 STEP 10
  $P Value: $!i
ENDDO

DO !i FROM 10 TO 1 STEP -1
  $P Countdown: $!i
ENDDO
```

#### Open DO Loop
```pml
!count = 0
DO
  !count = !count + 1
  $P Count: $!count
  BREAK IF !count GE 10
ENDDO
```

#### DO VALUES (Array Iteration)
```pml
!names = ARRAY('John', 'Jane', 'Bob')
DO !name VALUES !names
  $P Name: $!name
ENDDO
```

#### DO INDEX (Array Index Iteration)
```pml
!values = ARRAY(10, 20, 30, 40, 50)
DO !idx INDEX !values
  $P Index $!idx = $!values[$!idx]
ENDDO
```

### 5.3 Loop Control

#### BREAK - Exit loop immediately
```pml
DO !i FROM 1 TO 100
  IF !i EQ 50 THEN
    BREAK
  ENDIF
  $P !i
ENDDO
```

#### BREAKIF - Conditional break
```pml
DO !i FROM 1 TO 100
  BREAKIF !i GT 50
  $P !i
ENDDO
```

#### SKIP - Skip to next iteration
```pml
DO !i FROM 1 TO 10
  IF !i EQ 5 THEN
    SKIP  -- Skip iteration when i = 5
  ENDIF
  $P !i
ENDDO
```

#### SKIPIF - Conditional skip
```pml
DO !i FROM 1 TO 10
  SKIPIF !i EQ 5  -- Skip when i = 5
  $P !i
ENDDO
```

### 5.4 Labels and GOTO

```pml
!count = 0
LABEL /START
  !count = !count + 1
  $P Count: $!count
  GOLABEL /START IF !count LT 10

LABEL /END
$P Finished
```

**Important**: GOTO should be used sparingly. Prefer loops and structured control flow.

#### Illegal Jumping
```pml
-- CANNOT jump into or out of loops/IF blocks
DO !i FROM 1 TO 10
  -- GOLABEL /OUTSIDE  -- ILLEGAL!
ENDDO

LABEL /OUTSIDE
```

---

## 6. Arrays and Collections

### 6.1 Creating Arrays

```pml
-- Empty array
!myArray = ARRAY()

-- Array with values
!numbers = ARRAY(1, 2, 3, 4, 5)
!names = ARRAY('John', 'Jane', 'Bob')
!mixed = ARRAY('Text', 123, TRUE)

-- Adding elements
!myArray[1] = 'First'
!myArray[2] = 'Second'
!myArray[3] = 'Third'
```

### 6.2 Multi-dimensional Arrays

```pml
!matrix = ARRAY()
!matrix[1] = ARRAY(1, 2, 3)
!matrix[2] = ARRAY(4, 5, 6)
!matrix[3] = ARRAY(7, 8, 9)

$P Value: $!matrix[2][3]  -- Output: 6
```

### 6.3 Array Methods

```pml
!myArray = ARRAY('A', 'B', 'C')

-- Get array size
!size = !myArray.Size()

-- Append element
!myArray.Append('D')

-- Delete element
!myArray.Delete(2)  -- Removes 'B'

-- Set element
!myArray.Set(1, 'Z')  -- Changes 'A' to 'Z'

-- Get element
!value = !myArray[2]

-- Check if empty
!isEmpty = !myArray.Empty()

-- Clear all elements
!myArray.EmptyArray()
```

### 6.4 Array Sorting

```pml
!numbers = ARRAY(5, 2, 8, 1, 9, 3)

-- Sort ascending
!numbers.Sort()

-- Sort descending
!numbers.SortDescending()

-- Sort with custom comparison
USING !numbers
  SORT VALUES ASCENDING
ENDUSING
```

#### Sorting Strings
```pml
!names = ARRAY('Charlie', 'Alice', 'Bob')
!names.Sort()  -- Alphabetical: Alice, Bob, Charlie
```

### 6.5 String Split to Array

```pml
!text = 'apple,banana,orange,grape'
!fruits = !text.Split(',')
-- Result: ARRAY('apple', 'banana', 'orange', 'grape')

!words = 'Hello World PML'.Split(' ')
-- Result: ARRAY('Hello', 'World', 'PML')
```

### 6.6 Array Subtotaling (PML1)

```pml
VAR !data COLLECT ALL PIPE WITH DIAMETER GT 100
VAR !data SORT VALUES ASCENDING
VAR !data SUBTOTAL ON DIAMETER
```

### 6.7 Collections from Database

```pml
-- Collect database elements
VAR !pipes COLLECT ALL PIPE FOR CE
VAR !elbows COLLECT ALL ELBOW WITH RATING EQ '150'
VAR !equipment COLLECT ALL EQUI FOR ZONE

-- Iterate through collection
DO !pipe VALUES !pipes
  $P Pipe name: $(!pipe.Name)
ENDDO
```

### 6.8 Block Evaluation with Arrays

```pml
!pipes = COLLECT ALL PIPE FOR ZONE

-- Evaluate expression for all elements
!names = EVALUATE (NAME) FOR ALL FROM !pipes
!diameters = EVALUATE (DIAM) FOR ALL FROM !pipes

-- With filtering
!bigPipes = EVALUATE (NAME) FOR ALL FROM !pipes WITH DIAM GT 100
```

### 6.9 COMPOSE Command for Formatted Output

The `COMPOSE` command builds formatted text from arrays - powerful for reports:

```pml
-- Collect data
var !selection collect all pipe for zone
var !names raw evaluate (NAME) for all from !selection
var !diams raw evaluate (DIAM) for all from !selection
var !lengths raw evaluate (LENGTH) for all from !selection

-- Create sort order
var !sortorder evaluate (vval(!diams)) indices !selection

-- Build formatted lines with COMPOSE
var !separator |;|
var !delimiter | |

var !lines compose sort !sortorder quote (vtext(!delimiter)) nosuml $
  col val !names q $
  sep (vtext(!separator)) $
  col val !diams q $
  sep (vtext(!separator)) $
  col val !lengths q

-- Output lines
do !line values !lines
   write |$!line|
enddo
```

**COMPOSE Options:**
- `sort !sortorder` - Apply sorting indices
- `quote (delim)` - Quote values containing separator
- `nosuml` - Skip subtotal/summary lines
- `col val !array q` - Column from array, quoted
- `sep (sep)` - Separator between columns

**Dynamic COMPOSE String:**
```pml
-- Build compose command dynamically
if(|$!write| eq |write|)then
   var !composeCom |var !lines compose sort !sortorder quote (vtext(!del)) nosuml|
else
   var !composeCom |$!write compose sort !sortorder quote (vtext(!del)) nosuml|
endif

-- Execute dynamic command
$!composeCom $
col val !col1 q $
sep (vtext(!sep)) $
col val !col2 q $
sep (vtext(!sep)) $
col val !col3 q
```

### 6.10 RAW Keyword for EVALUATE

Use `raw` to prevent type conversion in evaluations:

```pml
-- With RAW: keeps original format
var !col1 raw evaluate (DIAM) for all from !selection
-- Result: Strings exactly as displayed in database

-- Without RAW: converts to native types  
var !col2 evaluate (DIAM) for all from !selection
-- Result: REAL numbers

-- Essential for formatted output
var !names raw evaluate (NAME of OWNER of OWNER) for all from !selection
```

---

## 7. Functions and Methods

### 7.1 PML Functions (PML2)

#### Defining a Function

```pml
define function !!CalculateArea(!width is REAL, !height is REAL)
  !area = !width * !height
  return !area
endfunction
```

#### Calling a Function

```pml
!result = !!CalculateArea(10.5, 20.0)
$P Area: $!result
```

#### Function with Multiple Return Types

```pml
define function !!GetElementInfo(!elemName is STRING)
  !info = OBJECT ARRAY()
  !info[1] = !elemName
  !info[2] = DIAMETER OF !elemName
  !info[3] = POSITION OF !elemName
  return !info
endfunction
```

#### Function with ANY Type

```pml
define function !!ProcessValue(!value is ANY)
  !type = !value.ObjectType()
  IF !type EQ 'STRING' THEN
    return !value.Upcase()
  ELSEIF !type EQ 'REAL' THEN
    return !value * 2
  ELSE
    return !value
  ENDIF
endfunction
```

### 7.2 Procedures (Functions without Return)

```pml
define method .DisplayMessage(!msg is STRING)
  $P ==============================
  $P $!msg
  $P ==============================
endmethod
```

### 7.3 Function Storage

Functions should be stored in files with `.pmlfnc` extension:

**File: CalculateArea.pmlfnc**
```pml
define function !!CalculateArea(!width is REAL, !height is REAL)
  return !width * !height
endfunction
```

**PML Functions vs Macros:**
- **PML Functions (`.pmlfnc`)** are the recommended approach for storing command sequences
- Functions provide argument type checking and return types
- Functions replace older PML Macros (`.mac`) for most use cases
- Use macros only for simple command sequences without parameters

**Automatic Loading from PMLLIB:**
- Place `.pmlfnc` files in any directory defined in the `PMLLIB` environment variable
- PML automatically scans these directories on startup
- Functions are loaded on-demand when first called
- No manual loading required

**Example PMLLIB Setup:**
```batch
REM Windows
set PMLLIB=C:\MyPMLFunctions;C:\AVEVA\PDMS\pmllib

REM Linux
export PMLLIB=/home/user/pml:/opt/aveva/pdms/pmllib
```

**User Customization Best Practice:**
- **NEVER modify original AVEVA files**
- Store all user customizations in a separate directory
- Add your custom directory to the **beginning** of PMLLIB path
- Your functions will override AVEVA defaults when needed

```batch
REM User area first, then AVEVA area
set PMLLIB=C:\Users\MyName\CustomPML;C:\AVEVA\PDMS\pmllib
```

### 7.4 Built-in Functions

```pml
-- String functions
!upper = 'hello'.Upcase()
!lower = 'HELLO'.Downcase()
!len = 'hello'.Length()
!sub = 'hello world'.Substring(1, 5)
!match = 'hello world'.Match('world')

-- Math functions
!sqrt = SQR(25)
!power = POW(2, 8)
!abs = ABS(-10)

-- Type conversion
!str = !number.String()
!num = !string.Real()

-- Array functions
!size = !array.Size()
!array.Append(!item)
!array.Delete(!index)
```

---

## 8. Object-Oriented Programming

### 8.1 Defining Object Types

**File: Person.pmlobj**
```pml
define object PERSON
  member .FirstName is STRING
  member .LastName is STRING
  member .Age is REAL
  member .Email is STRING
enddefine
```

### 8.2 Creating Object Instances

```pml
!employee = object PERSON()
!employee.FirstName = 'John'
!employee.LastName = 'Doe'
!employee.Age = 30
!employee.Email = 'john.doe@company.com'
```

### 8.3 Object Methods

#### Basic Object Methods

```pml
define object RECTANGLE
  member .Width is REAL
  member .Height is REAL
  
  define method .Area()
    return .Width * .Height
  endmethod
  
  define method .Perimeter()
    return 2 * (.Width + .Height)
  endmethod
  
  define method .Scale(!factor is REAL)
    .Width = .Width * !factor
    .Height = .Height * !factor
  endmethod
enddefine
```

Using the object:
```pml
!rect = object RECTANGLE()
!rect.Width = 10
!rect.Height = 20

!area = !rect.Area()
!perimeter = !rect.Perimeter()

!rect.Scale(2.0)  -- Double the size
```

#### Database Reference (DBREF) Introspection Methods

These methods are critical for dynamic database querying, especially when working with elements having unknown or variable attributes:

**Attribute() - Get/Set Object Member by Name**
```pml
!elem = !!CE  -- Current element (DBREF)
!attrName = 'DIAM'

-- Get attribute value dynamically
!value = !elem.Attribute(!attrName)  -- Returns ANY type
$P Diameter: $!value

-- Set attribute value dynamically
!elem.Attribute('XPOS') = 1000
```

**Attributes() - List All Object Members**
```pml
!elem = object DBREF('/PIPE-1')

-- Get array of all attribute names
!attrList = !elem.Attributes()  -- Returns ARRAY OF STRINGS

-- Iterate through all attributes dynamically
DO !attrName VALUES !attrList
  !value = !elem.Attribute(!attrName)
  $P $!attrName = $!value
ENDDO
```

This is essential for:
- Generating comprehensive reports with all available attributes
- Working with pipe specifications containing varying component types
- Creating generic utilities that work across different element types

**Set() - Check If Attribute Has Value**
```pml
!elem = !!CE

-- Check if attribute is set (not UNSET)
!hasValue = !elem.Attribute('DIAM').Set()

IF !hasValue THEN
  !diam = !elem.Attribute('DIAM')
  $P Diameter is set: $!diam
ELSE
  $P Diameter is UNSET
ENDIF

-- Common pattern: check if UNSET
IF !elem.Attribute('DESCRIPTION').Set().not() THEN
  $P Description is UNSET
ENDIF
```

**String() - Convert Object Value to String**
```pml
!elem = !!CE
!value = !elem.Attribute('POSITION')

-- Convert to string for file output or display
!textValue = !value.String()
$P Position as text: $!textValue

-- Write to file
!file = object FILE()
!file.Filename = 'C:\temp\output.txt'
!file.OpenWrite()
!file.WriteLine(!value.String())
!file.Close()
```

**Example: Dynamic Attribute Report Generator**
```pml
define function !!GenerateElementReport(!elem is DBREF)
  !report = ARRAY()
  
  -- Get all attributes
  !attrs = !elem.Attributes()
  
  -- Build report lines
  DO !attrName VALUES !attrs
    !value = !elem.Attribute(!attrName)
    
    -- Check if attribute has a value
    IF !value.Set() THEN
      !line = !attrName & ' = ' & !value.String()
      !report.Append(!line)
    ENDIF
  ENDDO
  
  return !report
endfunction

-- Usage
!elem = object DBREF('/ZONE-1/PIPE-1')
!reportLines = !!GenerateElementReport(!elem)

DO !line VALUES !reportLines
  $P $!line
ENDDO
```

### 8.4 Constructor Methods

```pml
define object CIRCLE
  member .Radius is REAL
  member .CenterX is REAL
  member .CenterY is REAL
  
  -- Constructor with no arguments
  define method .CIRCLE()
    .Radius = 1.0
    .CenterX = 0.0
    .CenterY = 0.0
  endmethod
  
  -- Constructor with arguments
  define method .CIRCLE(!r is REAL, !x is REAL, !y is REAL)
    .Radius = !r
    .CenterX = !x
    .CenterY = !y
  endmethod
  
  define method .Area()
    return 3.14159 * .Radius * .Radius
  endmethod
enddefine
```

Usage:
```pml
!circle1 = object CIRCLE()           -- Uses default constructor
!circle2 = object CIRCLE(5, 10, 20)  -- Uses parameterized constructor

!area = !circle2.Area()
```

### 8.5 Method Overloading

```pml
define object CALCULATOR
  -- Add two numbers
  define method .Add(!a is REAL, !b is REAL)
    return !a + !b
  endmethod
  
  -- Add three numbers
  define method .Add(!a is REAL, !b is REAL, !c is REAL)
    return !a + !b + !c
  endmethod
  
  -- Add elements of an array
  define method .Add(!numbers is ARRAY)
    !sum = 0
    DO !num VALUES !numbers
      !sum = !sum + !num
    ENDDO
    return !sum
  endmethod
enddefine
```

### 8.6 Invoking Methods from Other Methods

```pml
define object CIRCLE
  member .Radius is REAL
  
  define method .Area()
    return 3.14159 * .Radius * .Radius
  endmethod
  
  define method .DisplayInfo()
    !area = .Area()  -- Call another method
    $P Circle Radius: $.Radius
    $P Circle Area: $!area
  endmethod
enddefine
```

---

## 9. Macros

### 9.1 Basic Macros (PML1)

**File: CreateBox.mac**
```pml
NEW EQUIP /MYBOX
NEW BOX
XLEN 1000
YLEN 2000
ZLEN 3000
```

Running:
```pml
$M/CreateBox
```

### 9.2 Macros with Parameters

**File: CreateBox.mac**
```pml
-- Parameters: $1=name, $2=xlen, $3=ylen, $4=zlen
NEW EQUIP /$1
NEW BOX
XLEN $2
YLEN $3
ZLEN $4
```

Running:
```pml
$M/CreateBox BOX001 1000 2000 3000
```

### 9.3 Advanced Macro with Text Parameters

```pml
$M/CreatePipe $<Pipe Description$> 150 10000
```

**File: CreatePipe.mac**
```pml
-- $1 = description (can contain spaces)
-- $2 = diameter
-- $3 = length
NEW PIPE /$1
DIAM $2
LENGTH $3
```

### 9.4 Synonyms

```pml
-- Define synonym
$S XXX = NEW ELBOW SELECT WITH STYP LR ORI P1 IS N

-- Use synonym
XXX

-- Parameterized synonym
$S YYY = NEW BOX XLEN $S1 YLEN $S2 ZLEN $S3
YYY 100 200 300

-- Delete synonym
$S XXX =

-- Delete all synonyms (DANGEROUS!)
$SK

-- Turn synonyms off/on
$S-
$S+
```

### 9.5 Numbered Variables in Macros

```pml
VAR 1 NAME
VAR 2 'Hello World'
VAR 3 (99)
VAR 4 (99 * 3 / 6 + 0.5)
VAR 117 POS IN SITE
VAR 118 (NAME OF OWNER OF OWNER)
VAR 119 'hello' + 'world' + 'how are you'
```

### 9.6 Conditional Macro Execution Pattern

Production macros often use parameter-based conditional execution for different modes:

```pml
-- Multi-mode macro: Definition mode, Prompt mode, Run mode
-- Usage: $M/MyReport DEFINE    -- Set up parameters
--        $M/MyReport           -- Interactive prompt
--        $M/MyReport RUN ...   -- Direct execution

if (|$1| eq |DEFINE|) then
   -- DEFINE mode: Initialize report parameters
   var _REPFORMAT |CSV|
   var _REPSEPARATOR |;|
   var _REPFILENAME |C:\output\report.csv|
   var !!RepColumns[1] raw |NAME|
   var !!RepColumns[2] raw |TYPE|
   var !!RepColumns[3] raw |DIAMETER|
   return
endif

if (|$1| eq |PROMPT| or |$1| eq ||) then
   -- PROMPT mode: Show interactive dialog
   $M "%PMLUI%/CLIB/REP/RSELECT" $<RUN$> $<$0$> $
     $<$_REPFILENAME$> $<pipe$> $<$> $<ce$>
   return
endif

-- RUN mode: Execute with parameters $2, $3, etc.
var !fileName (trim(|$2|))
var !selType (upcase(trim(|$3|)))
-- Process report...
```

**Pattern Benefits:**
- Single macro handles configuration, UI, and execution
- DEFINE mode initializes global variables
- PROMPT mode provides interactive interface
- RUN mode enables batch/scripted execution
- Cleaner than multiple separate macros

### 9.7 Macros vs Functions - Best Practice

**Recommendation:** Use PML Functions (`.pmlfnc`) instead of Macros (`.mac`) for most scenarios.

**PML Functions Advantages:**
- Type-checked parameters
- Return values
- Better error handling
- Automatic loading from PMLLIB
- More maintainable code

**When to Use Macros:**
- Simple command sequences
- No parameters needed
- Legacy code compatibility
- Interactive command shortcuts

**Migration Example:**

**Old Macro (CreateBox.mac):**
```pml
-- Parameters: $1=name, $2=xlen, $3=ylen, $4=zlen
NEW EQUIP /$1
NEW BOX
XLEN $2
YLEN $3
ZLEN $4
```

**New Function (CreateBox.pmlfnc):**
```pml
define function !!CreateBox(!name is STRING, !xlen is REAL, !ylen is REAL, !zlen is REAL)
  -- Type checking automatic
  -- Better error messages
  
  NEW EQUIP /(!name)
  NEW BOX
  XLEN !xlen
  YLEN !ylen
  ZLEN !zlen
  
  return !!CE  -- Return created element
endfunction
```

---

## 10. File Operations

### 10.1 Creating File Objects

```pml
!file = object FILE()
!file.Filename = 'C:\temp\output.txt'
```

### 10.2 Writing to Files

```pml
!file = object FILE()
!file.Filename = 'C:\temp\output.txt'
!file.OpenWrite()

!file.WriteLi ne('First line')
!file.WriteLine('Second line')
!file.WriteLine('Third line')

!file.Close()
```

### 10.3 Reading from Files

```pml
!file = object FILE()
!file.Filename = 'C:\temp\input.txt'
!file.OpenRead()

DO
  !line = !file.ReadLine()
  BREAK IF !file.IsEndOfFile()
  $P Line: $!line
ENDDO

!file.Close()
```

### 10.4 Appending to Files

```pml
!file = object FILE()
!file.Filename = 'C:\temp\output.txt'
!file.OpenAppend()

!file.WriteLine('Appended line')

!file.Close()
```

### 10.5 File Methods

```pml
!file = object FILE()

-- Check if file exists
!exists = !file.Exists('C:\temp\file.txt')

-- Delete file
!file.Delete()

-- Copy file
!file.Copy('C:\temp\source.txt', 'C:\temp\dest.txt')

-- Get file size
!size = !file.Size()

-- Check if end of file
!eof = !file.IsEndOfFile()
```

### 10.6 Reading/Writing Arrays

#### WriteArray() Method
```pml
!myArray = ARRAY('Line 1', 'Line 2', 'Line 3')

-- Write array to file (line by line)
!file = object FILE()
!file.Filename = 'C:\temp\data.txt'
!file.OpenWrite()
!file.WriteArray(!myArray)
!file.Close()

-- Read file into array
!file.OpenRead()
!loadedArray = !file.ReadArray()
!file.Close()
```

#### WriteFile() Method - Efficient Array Output

The **WriteFile()** method writes array contents directly to file in a single operation:

```pml
!outputArray = ARRAY()
!outputArray.Append('Header Line')
!outputArray.Append('Data Line 1')
!outputArray.Append('Data Line 2')
!outputArray.Append('Data Line 3')

-- Write entire array to file at once
!file = object FILE()
!success = !file.WriteFile('C:\temp\output.txt', !outputArray, 'OVERWRITE')

IF !success THEN
  $P File written successfully
ELSE
  $P Error writing file
ENDIF
```

**WriteFile() Modes:**
- `'OVERWRITE'` - Replace existing file
- `'APPEND'` - Add to existing file

**Complete Database Export Example:**
```pml
define function !!ExportElementAttributes(!elem is DBREF, !filepath is STRING)
  !outputLines = ARRAY()
  
  -- Add header
  !outputLines.Append('Element: ' & !elem.Fullname())
  !outputLines.Append('Type: ' & !elem.Type)
  !outputLines.Append('---')
  
  -- Get all attributes
  !attrs = !elem.Attributes()
  
  -- Add each attribute
  DO !attrName VALUES !attrs
    !value = !elem.Attribute(!attrName)
    
    IF !value.Set() THEN
      !line = !attrName & ',' & !value.String()
      !outputLines.Append(!line)
    ENDIF
  ENDDO
  
  -- Write to file efficiently
  !file = object FILE()
  !success = !file.WriteFile(!filepath, !outputLines, 'OVERWRITE')
  
  return !success
endfunction

-- Usage
!elem = object DBREF('/ZONE-1/PIPE-1')
!success = !!ExportElementAttributes(!elem, 'C:\temp\pipe_data.csv')
```

**Best Practice:** Use `WriteFile()` for batch operations instead of multiple `WriteLine()` calls for better performance.

### 10.7 Error Handling with Files

```pml
!file = object FILE()
!file.Filename = 'C:\temp\test.txt'

handle ANY
  !file.OpenRead()
  DO
    !line = !file.ReadLine()
    BREAK IF !file.IsEndOfFile()
    $P $!line
  ENDDO
  !file.Close()
elsehandle
  $P Error reading file: $!!Error.Message
endhandle
```

### 10.8 PML1 File Operations with OPENFILE/CLOSEFILE

Legacy macros use PML1 file commands with file handles:

```pml
-- Open file with overwrite confirmation
var !fileName 'C:\output\report.csv'
openfile "$!fileName" overwrite !output
handle (41,324)  -- File exists error
   confirm |OK to overwrite $!fileName?|
   handle (61,115)  -- User cancelled
      return error 1 'File $!fileName exists. Use OVERWRITE mode to force write.'
   endhandle
   if(not $_CALERT_USERDATA) then
      return  -- User clicked No
   endif
   openfile "$!fileName" overwrite !output
   handle any
      return error 1 |Unable to open file $!fileName|
   endhandle
elsehandle any
   return error 1 |Unable to open file $!fileName|
endhandle

-- Write to file using handle
var !write |writefile $!output|
$!write 'Header Line'
$!write 'Data Line 1'
$!write 'Data Line 2'

-- Always close file
closefile $!output
```

**Key Patterns:**
- `openfile` creates file handle variable (`!output`)
- `writefile $!output` writes line to file
- `closefile $!output` closes file
- Nested error handling for user confirmation
- Special error codes: `(41,324)` = file exists, `(61,115)` = user cancel

### 10.9 VTEXT() Function for Safe Variable Expansion

Use `vtext()` to safely expand variables in string comparisons:

```pml
-- Safe string comparison
if(vtext(!fileName) ne ||) then
   -- File name provided
else
   -- No file name, use display output
endif

-- Without vtext() can cause issues with special characters
if(vtext(!selType) eq || or vtext(!selType) eq |ALL|) then
   var !criteria |ALL|
else
   var !criteria (|ALL (| + vtext(!selType) + |)|)
endif
```

**Why VTEXT():**
- Prevents errors with UNSET or special characters
- Returns empty string for UNSET instead of error
- Essential for building dynamic command strings

---

## 11. Error Handling

### 11.1 Basic Error Handling

```pml
HANDLE ANY
  -- Code that might cause errors
  NEW EQUIP /MYEQUIP
  XPOS 1000
ELSEHANDLE
  $P An error occurred
ENDHANDLE
```

### 11.2 Specific Error Handling

```pml
HANDLE OBJNOTFND
  !elem = object DBREF('/MYEQUIP')
ELSEHANDLE
  $P Element not found
ENDHANDLE
```

### 11.3 Error Object

```pml
HANDLE ANY
  -- Code
ELSEHANDLE
  $P Error Number: $!!Error.Number
  $P Error Message: $!!Error.Message
  $P Error Type: $!!Error.Type
ENDHANDLE
```

### 11.4 ONERROR Settings

```pml
-- Continue on error (default)
ONERROR CONTINUE

-- Stop on error
ONERROR STOP

-- Call function on error
ONERROR CALL !!MyErrorHandler

define function !!MyErrorHandler()
  $P Error: $!!Error.Message
  -- Handle error
endfunction
```

### 11.5 Nested Error Handling

```pml
HANDLE ANY
  HANDLE OBJNOTFND
    !elem = object DBREF('/MISSING')
  ELSEHANDLE
    $P Inner: Object not found
  ENDHANDLE
ELSEHANDLE
  $P Outer: General error
ENDHANDLE
```

### 11.6 Error Recovery

```pml
!success = FALSE
HANDLE ANY
  NEW EQUIP /TESTEQUIP
  !success = TRUE
ELSEHANDLE
  $P Failed to create equipment
  -- Attempt recovery
  NEW EQUIP /TESTEQUIP_ALT
  !success = TRUE
ENDHANDLE

IF !success THEN
  $P Operation succeeded
ENDIF
```

### 11.7 ONERROR with GOLABEL for Resource Cleanup

Advanced pattern: Use `onerror golabel` to ensure cleanup happens even on errors:

```pml
-- Save original state
var !units units
mm bore mm distance

-- Set up error recovery BEFORE operations
onerror golabel /closefile

-- Open file
openfile "$!fileName" overwrite !output

-- Perform operations that might fail
var !selection collect $!criteria
handle any
   if(vtext(!fileName) ne ||) then
      closefile $!output
   endif
   $!units  -- Restore units
   return error 1 |Selection is invalid: $!!ERROR.TEXT|
endhandle

-- Process data...
-- Write to file...

-- Normal completion
if (vtext(!fileName) ne ||) then
   closefile $!output
endif
onerror golabel /restoreUnits
$!units  -- Restore units
return

-- Error handlers with cleanup
label /closefile
handle any
   if (vtext(!fileName) ne ||) then
      closefile $!output
   endif
   $!units
   return error
endhandle

label /restoreUnits
handle any
   $!units
   return error
endhandle
```

**Pattern Benefits:**
- Guarantees file closure even on errors
- Restores system state (units, settings)
- Multiple error exit points with consistent cleanup
- Labels group related error handling

### 11.8 DEFINED() Function for Conditional Logic

Check if variable exists before using:

```pml
var !graphicsMode 1

if (defined(!graphicsMode)) then
   prompt 'Processing $!count elements...'
endif

-- Use for optional features
handle (61,115)  -- User cancel
elsehandle none
   var !graphicsMode 1  -- Only set if not cancelled
endhandle

if (defined(!graphicsMode)) then
   -- Show progress only in graphics mode
   prompt 'Evaluating expressions...'
endif
```

**Use Cases:**
- Optional progress indicators
- Feature flags
- Mode detection
- Preventing undefined variable errors

---

## 12. Forms and GUI

### 12.1 Simple Form Definition

**File: MyForm.pmlfrm**
```pml
setup form !!MyForm

-- Form properties
title 'My First Form'
okbutton yes
cancelbutton yes

-- Add gadgets
paragraph .instructions |Enter your details:|

text .username tag 'Username:' width 20

password .password tag 'Password:' width 20

button .submit |Submit| callback .submitCallback

define method .submitCallback()
  !user = !!MyForm.username.Value
  !pass = !!MyForm.password.Value
  $P User: $!user
endmethod

exit
```

### 12.2 Loading and Showing Forms

```pml
-- Load form
!!MyForm = object FORM()

-- Show form as dialog
!!MyForm.Show()

-- Show form as document
!!MyForm.ShowDocument()

-- Hide form
!!MyForm.Hide()

-- Kill form
!!MyForm.Kill()
```

### 12.3 Form Types

```pml
setup form !!MyForm

-- Dialog form (default)
formtype dialog

-- Document form
formtype document

-- Main form (application window)
formtype main

exit
```

### 12.4 Form Callbacks

```pml
setup form !!MyForm

-- Initialization callback
initialise .initCallback

define method .initCallback()
  $P Form is initializing
endmethod

-- OK button callback
okmethod .okCallback

define method .okCallback()
  $P OK button clicked
endmethod

-- Cancel button callback
cancelmethod .cancelCallback

define method .cancelCallback()
  $P Cancel button clicked
endmethod

-- Close callback
quitmethod .quitCallback

define method .quitCallback()
  $P Form is closing
endmethod

-- First shown callback
firstshown .firstShownCallback

define method .firstShownCallback()
  $P Form shown for first time
endmethod

exit
```

### 12.5 Form Variables

```pml
setup form !!MyForm

-- Define form variables
member .FilePath is STRING
member .Count is REAL
member .IsValid is BOOLEAN
member .DataList is ARRAY

define method .initCallback()
  .FilePath = 'C:\temp\default.txt'
  .Count = 0
  .IsValid = TRUE
  .DataList = ARRAY()
endmethod

exit
```

### 12.6 Form Properties

```pml
setup form !!MyForm

title 'Application Title'
width 40        -- Grid units
height 30       -- Grid units
resize yes      -- Allow resizing
okbutton yes
cancelbutton yes
applybutton yes
helpbutton yes

exit
```

---

## 13. Gadgets Reference

### 13.1 Text Input

```pml
-- Simple text input
text .name tag 'Name:' width 20

-- With default value
text .address tag 'Address:' width 30 value 'Default Address'

-- Password field
password .pass tag 'Password:' width 15

-- Multiline text
textpane .description tag 'Description:' width 40 height 10
```

### 13.2 Buttons

```pml
-- Regular button
button .ok |OK| callback .okCallback

-- Button with pixmap
button .save pixmap '/icons/save.png' width 32 height 32

-- Toggle button
button .toggle |Toggle| toggle callback .toggleCallback

-- Link label button
button .link |Click Here| linklabel callback .linkCallback
```

### 13.3 Toggles and Radio Buttons

```pml
-- Toggle (checkbox)
toggle .option1 |Enable feature|

-- Multiple toggles
toggle .opt1 |Option 1|
toggle .opt2 |Option 2|
toggle .opt3 |Option 3|

-- Radio buttons (in a frame)
frame .radioGroup |Select One|
  rtoggle .radio1 |Choice 1|
  rtoggle .radio2 |Choice 2|
  rtoggle .radio3 |Choice 3|
endframe
```

### 13.4 Lists

```pml
-- Single selection list
list .items tag 'Select Item:' width 20 height 8 single

-- Multiple selection list
list .items tag 'Select Items:' width 20 height 8 multiple

-- Populate list
!!MyForm.items.InsertItem('Item 1')
!!MyForm.items.InsertItem('Item 2')
!!MyForm.items.InsertItem('Item 3')

-- Get selected
!selected = !!MyForm.items.Value  -- Index of selection
!text = !!MyForm.items.TextValue  -- Text of selection
```

### 13.5 Option (Dropdown)

```pml
-- Define option gadget
option .country tag 'Country:' width 15
  item 'USA'
  item 'UK'
  item 'Germany'
  item 'France'
endoption

-- Combo box (editable dropdown)
combobox .custom tag 'Custom:' width 15
  item 'Option 1'
  item 'Option 2'
  item 'Option 3'
endcombobox
```

### 13.6 Numeric Input

```pml
numericinput .count tag 'Count:' from 0 to 100 step 1

-- Get/set value
!value = !!MyForm.count.Value
!!MyForm.count.Value = 50
```

### 13.7 Slider

```pml
-- Horizontal slider
slider .volume horizontal from 0 to 100 width 20

-- Vertical slider
slider .level vertical from 0 to 100 height 15

-- With callback
slider .value horizontal from 0 to 100 callback .sliderCallback

define method .sliderCallback()
  !val = .value.Value
  $P Slider value: $!val
endmethod
```

### 13.8 Frames

```pml
-- Normal frame
frame .group1 |Group 1|
  toggle .opt1 |Option 1|
  toggle .opt2 |Option 2|
endframe

-- Tabset frame
frame .tabs tabset
  frame .tab1 |Tab 1|
    paragraph .info1 |Content for tab 1|
  endframe
  
  frame .tab2 |Tab 2|
    paragraph .info2 |Content for tab 2|
  endframe
endframe

-- Panel frame
frame .panel panel
  text .field1 tag 'Field 1:'
  text .field2 tag 'Field 2:'
endframe

-- Foldable panel
frame .foldable folduppanel |Click to expand|
  text .hidden1 tag 'Hidden Field 1:'
  text .hidden2 tag 'Hidden Field 2:'
endframe
```

### 13.9 Paragraph (Display Text)

```pml
paragraph .title |This is display text| width 30

-- Multiline paragraph
paragraph .info width 40 lines 3 |
  Line 1 of information
  Line 2 of information
  Line 3 of information
|
```

### 13.10 View Gadgets

```pml
-- Alpha view (text output)
alphaview .output width 60 height 20

-- 2D graphics view
g2dview .drawing width 40 height 30

-- 3D graphics view
g3dview .model width 50 height 40
```

### 13.11 Line (Separator)

```pml
-- Horizontal line
line .sep1 horizontal width 40

-- Vertical line
line .sep2 vertical height 20
```

### 13.12 Database Selector

```pml
selector .elements tag 'Select Elements:' width 30 height 10

-- Populate with database elements
!!MyForm.elements.SetDBList(!elementList)
```

---

## 14. Menus

### 14.1 Bar Menu

```pml
setup form !!MyForm

-- Define bar menu
barmenu .mainMenu
  menu .file |File|
    menufield .new |New| callback .newCallback
    menufield .open |Open| callback .openCallback
    menufield .save |Save| callback .saveCallback
    menufield .separator separator
    menufield .exit |Exit| callback .exitCallback
  endmenu
  
  menu .edit |Edit|
    menufield .cut |Cut| callback .cutCallback
    menufield .copy |Copy| callback .copyCallback
    menufield .paste |Paste| callback .pasteCallback
  endmenu
  
  menu .help |Help|
    menufield .about |About| callback .aboutCallback
  endmenu
endbarmenu

define method .newCallback()
  $P New clicked
endmethod

-- Other callbacks...

exit
```

### 14.2 Popup Menu

```pml
-- Define popup menu
menu .popup popup
  menufield .option1 |Option 1| callback .opt1Callback
  menufield .option2 |Option 2| callback .opt2Callback
  menufield .option3 |Option 3| callback .opt3Callback
endmenu

-- Show popup menu
define method .showPopup()
  .popup.Popup()
endmethod
```

### 14.3 Toggle Menu Items

```pml
menu .view |View|
  menufield .showToolbar |Show Toolbar| toggle callback .toolbarCallback
  menufield .showStatus |Show Status Bar| toggle callback .statusCallback
endmenu

define method .toolbarCallback()
  !checked = .showToolbar.Checked
  IF !checked THEN
    $P Toolbar shown
  ELSE
    $P Toolbar hidden
  ENDIF
endmethod
```

### 14.4 Dynamic Menu Modification

```pml
-- Add menu item
.fileMenu.InsertMenuItem('newItem', 'New Item', .newItemCallback)

-- Remove menu item
.fileMenu.DeleteMenuItem('newItem')

-- Enable/disable menu item
.saveMenuItem.Active = FALSE  -- Disable
.saveMenuItem.Active = TRUE   -- Enable

-- Check/uncheck toggle menu item
.showToolbar.Checked = TRUE
```

---

## 15. Form Layout

### 15.1 Layout Modes

```pml
setup form !!MyForm

-- Variable character layout (recommended)
layout varchars

-- Fixed character layout
layout fixchars

exit
```

### 15.2 Auto-placement

```pml
-- Set path direction
path down     -- Place next gadget below
path right    -- Place next gadget to the right
path up       -- Place next gadget above
path left     -- Place next gadget to the left

-- Set spacing
hdist 2.0     -- Horizontal distance in grid units
vdist 1.5     -- Vertical distance in grid units

-- Example
path down
vdist 0.5
text .field1 tag 'Field 1:' width 20
text .field2 tag 'Field 2:' width 20
text .field3 tag 'Field 3:' width 20
```

### 15.3 Alignment

```pml
-- Horizontal alignment
halign left
halign center
halign right

-- Vertical alignment
valign top
valign center
valign bottom

-- Example
path right
halign left
valign center
button .btn1 |Button 1|
button .btn2 |Button 2|
button .btn3 |Button 3|
```

### 15.4 Absolute Positioning

```pml
-- Position at grid coordinates
button .ok at x 10 y 20

-- Position relative to other gadgets
button .cancel at xmin.ok ymax.ok + 1

-- Position relative to form
button .apply at xmax form - size ymax form
```

### 15.5 Relative Positioning

```pml
-- Position relative to previous gadget
text .field1 at 5 5 width 20
text .field2 at xmin ymax + 0.5 width 20
text .field3 at xmin.field2 ymax + 0.5 width 20

-- Position relative to form edges
button .ok at xmax form - size ymax form - size
button .cancel at xmin.ok - size - 1 ymax.ok
```

### 15.6 Anchoring (Resizable Forms)

```pml
-- Anchor gadget edges
frame .main anchor all  -- Anchor to all edges (resize with form)

text .field1 anchor top + left  -- Stay at top-left
button .ok anchor bottom + right  -- Stay at bottom-right

list .items anchor left + right + top + bottom  -- Resize with form
```

### 15.7 Docking

```pml
-- Dock to edges
frame .toolbar dock top
frame .statusbar dock bottom
frame .sidebar dock left
frame .main dock all  -- Fill remaining space
```

### 15.8 Sizing

```pml
-- Absolute size
button .ok width 10 height 2

-- Relative size
button .cancel width.ok height.ok

-- Size to fit container
text .field1 width to max form - padding

-- Size to another gadget
frame .group1 width to xmax.lastGadget height 15
```

---

## 16. Database Integration

### 16.1 Accessing Database Elements

```pml
-- Get current element
!ce = !!CE

-- Get element by path
!elem = object DBREF('/ZONE-1/EQUIP-1')

-- Check if element exists
!exists = !elem.Exists()
```

### 16.2 Reading Attributes

```pml
!elem = object DBREF('/MYEQUIP')

-- Read attributes
!name = !elem.Name
!type = !elem.Type
!owner = !elem.Owner

-- Specific attributes
!xpos = !elem.Xpos
!ypos = !elem.Ypos
!zpos = !elem.Zpos

-- Position
!pos = !elem.Position
```

### 16.3 Setting Attributes

```pml
!elem = object DBREF('/MYEQUIP')

-- Set attributes
!elem.Xpos = 1000
!elem.Ypos = 2000
!elem.Zpos = 3000

-- Set description
!elem.Description = 'My Equipment'
```

### 16.4 Creating Elements

```pml
-- Create new equipment
NEW EQUIP /MYEQUIP
XPOS 1000
YPOS 2000
ZPOS 3000

-- Get reference to created element
!newEquip = !!CE
```

### 16.5 Navigating Hierarchy

```pml
!elem = object DBREF('/ZONE-1/EQUIP-1/NOZZLE-1')

-- Navigate up
!owner = !elem.Owner          -- Parent element
!grandparent = !owner.Owner   -- Grandparent

-- Navigate down
!firstMember = !elem.First    -- First child
!lastMember = !elem.Last      -- Last child

-- Navigate siblings
!next = !elem.Next            -- Next sibling
!prev = !elem.Previous        -- Previous sibling
```

### 16.6 Collections

```pml
-- Collect specific type
!pipes = COLLECT ALL PIPE FOR CE
!elbows = COLLECT ALL ELBOW FOR /ZONE-1

-- Collect with filter
!bigPipes = COLLECT ALL PIPE WITH DIAM GT 200 FOR ZONE

-- Iterate collection
DO !pipe VALUES !pipes
  !name = !pipe.Name
  !diam = !pipe.Diameter
  $P Pipe: $!name, Diameter: $!diam
ENDDO
```

### 16.7 COLLECTION Object for Efficient Queries

The COLLECTION object provides efficient, filtered querying of the database hierarchy:

**Basic COLLECTION Usage**
```pml
-- Create collection object
!collect = object COLLECTION()

-- Set scope (hierarchy element to search under)
!pipe = object DBREF('/ZONE-1/PIPE-1')
!collect.scope(!pipe)

-- Set type of elements to find
!collect.type('ELBOW')

-- Execute search and get results
!elbows = !collect.results()  -- Returns ARRAY of DBREF objects

-- Process results
DO !elbow VALUES !elbows
  $P Found elbow: $(!elbow.Name)
ENDDO
```

**Complete COLLECTION Pattern**
```pml
define function !!GetPipeComponents(!pipePath is STRING, !compType is STRING)
  -- Create collection
  !collect = object COLLECTION()
  
  -- Set scope to pipe
  !pipeRef = object DBREF(!pipePath)
  !collect.scope(!pipeRef)
  
  -- Set component type (e.g., 'ELBOW', 'VALVE', 'TEE')
  !collect.type(!compType)
  
  -- Get and return results
  !components = !collect.results()
  return !components
endfunction

-- Usage
!valves = !!GetPipeComponents('/ZONE-1/PIPE-1', 'VALVE')
$P Found $(!valves.Size()) valves
```

**Using Current Element (!!CE)**
```pml
-- !!CE is a global DBREF pointing to current database element
!currentElem = !!CE

-- Collect all members under current element
!collect = object COLLECTION()
!collect.scope(!!CE)
!collect.type('MEMBER')
!members = !collect.results()

DO !member VALUES !members
  $P Member: $(!member.Name), Type: $(!member.Type)
ENDDO
```

**COLLECTION vs COLLECT Comparison**
```pml
-- Traditional COLLECT (PML1 style)
!pipes = COLLECT ALL PIPE FOR ZONE

-- COLLECTION object (PML2 style - more flexible)
!collect = object COLLECTION()
!collect.scope(object DBREF('/ZONE-1'))
!collect.type('PIPE')
!pipes = !collect.results()

-- COLLECTION allows dynamic type specification
!elementType = 'ELBOW'  -- Could come from user input
!collect.type(!elementType)
!results = !collect.results()
```

**Pipe Specification Example**
```pml
define function !!GetAllComponentsInSpec(!specName is STRING)
  !allComponents = ARRAY()
  
  -- Get spec element
  !spec = object DBREF('/CATALOGUE/' & !specName)
  
  -- Collect all section types
  !collect = object COLLECTION()
  !collect.scope(!spec)
  !collect.type('SCTN')
  !sections = !collect.results()
  
  -- For each section, collect components
  DO !section VALUES !sections
    !collect.scope(!section)
    !collect.type('MEMBER')
    !components = !collect.results()
    
    DO !comp VALUES !components
      !allComponents.Append(!comp)
    ENDDO
  ENDDO
  
  return !allComponents
endfunction
```

### 16.8 Traditional Queries

```pml
-- Query elements
!name = NAME OF /MYEQUIP
!type = TYPE OF OWNER
!pos = POSITION OF /NOZZLE-1

-- Query with expressions
!diameter = EVALUATE (DIAM) FOR !pipe
!names = EVALUATE (NAME) FOR ALL FROM !collection
```

### 16.9 Modification Tracking

```pml
-- Undo support
UNDO ON
  -- Make changes
  NEW PIPE /PIPE-1
  DIAM 150
  LENGTH 5000
UNDO OFF

-- Redo
REDO
```

---

## 17. Best Practices

### 17.1 Code Organization

#### Meaningful Variable Names
```pml
-- Use meaningful names
!pipelineDiameter = 150  -- Good
!pd = 150                -- Bad

-- Group related code
-- All database operations
!pipes = COLLECT ALL PIPE FOR ZONE
DO !pipe VALUES !pipes
  -- Process pipe
ENDDO

-- All file operations
!file = object FILE()
!file.Filename = 'output.txt'
!file.OpenWrite()
-- Write data
!file.Close()
```

#### File Organization and PMLLIB Structure

**Recommended Directory Structure:**
```
C:\MyPML\                    (Add to PMLLIB)
├── functions\               (General utilities)
│   ├── StringUtils.pmlfnc
│   ├── MathUtils.pmlfnc
│   └── DatabaseUtils.pmlfnc
├── reports\                 (Report generators)
│   ├── PipeReport.pmlfnc
│   └── EquipmentList.pmlfnc
├── forms\                   (GUI forms)
│   ├── MainDialog.pmlfrm
│   └── Settings.pmlfrm
├── objects\                 (Custom objects)
│   └── Point3D.pmlobj
└── pml.index               (Auto-generated - DO NOT EDIT)
```

**PMLLIB Best Practices:**

1. **Separate User Code from AVEVA Code**
   ```batch
   REM CORRECT - User area first
   set PMLLIB=C:\MyPML;C:\ProjectPML;C:\AVEVA\PDMS\pmllib
   
   REM WRONG - Don't put custom code in AVEVA directories
   REM set PMLLIB=C:\AVEVA\PDMS\pmllib
   ```

2. **Never Modify AVEVA Files**
   - AVEVA updates will overwrite your changes
   - Create new files in your own directories
   - Override functions by placing your version earlier in PMLLIB path

3. **Automatic Loading**
   - PML scans all PMLLIB directories at startup
   - Creates `pml.index` files automatically
   - Functions load on first call (lazy loading)
   - No manual loading required

4. **Rebuilding Indexes**
   ```pml
   -- After adding new .pmlfnc files, rebuild index
   pml rehash
   
   -- Or rebuild all PMLLIB directories (slower)
   pml rehash all
   
   -- Reload specific function after editing
   pml reload function !!MyFunction
   ```

5. **Function Naming Conventions**
   ```pml
   -- Use descriptive names
   !!CalculatePipeLength      -- Good
   !!CPL                       -- Bad
   
   -- Prefix by category
   !!Report_GeneratePipeList
   !!Report_ExportToExcel
   !!Util_StringToArray
   !!Util_FileExists
   !!DB_GetAllPipes
   !!DB_FindElementByName
   ```

6. **Version Control**
   - Store your PMLLIB directories in version control (Git, SVN)
   - Exclude auto-generated `pml.index` files
   - Document dependencies in README files

### 17.2 Comments

```pml
-- Use comments to explain WHY, not WHAT
!factor = 1.732  -- Square root of 3 for triangular calculation

-- Document complex algorithms
-- This function calculates the intersection point of two lines
-- using parametric equations and Cramer's rule
define function !!LineIntersection(!line1 is OBJECT, !line2 is OBJECT)
  -- Implementation
endfunction
```

### 17.3 Error Handling

```pml
-- Always handle potential errors
HANDLE ANY
  !elem = object DBREF(!path)
  !data = !elem.SomeAttribute
ELSEHANDLE
  $P Error accessing element: $!path
  $P Error: $!!Error.Message
  !data = UNSET  -- Provide default
ENDHANDLE
```

### 17.4 Function Design

```pml
-- Functions should do one thing well
define function !!CalculatePipeLength(!diameter is REAL, !flowRate is REAL)
  -- Single purpose: calculate required length
  !length = !flowRate / (!diameter * 0.01)
  return !length
endfunction

-- Use meaningful parameter names
define function !!ValidateInput(!username is STRING, !password is STRING)
  -- Not: (!a is STRING, !b is STRING)
  -- Implementation
endfunction
```

### 17.5 Resource Cleanup

```pml
-- Always close files
!file = object FILE()
HANDLE ANY
  !file.OpenRead()
  -- Process file
ELSEHANDLE
  $P File error
ENDHANDLE
-- Always close, even on error
IF !file.IsOpen() THEN
  !file.Close()
ENDIF
```

### 17.6 Performance

```pml
-- Cache frequently accessed values
!ce = !!CE
!ceName = !ce.Name  -- Cache instead of calling !!CE.Name repeatedly

-- Use appropriate collection methods
!names = EVALUATE (NAME) FOR ALL FROM !elements  -- Fast
-- vs
DO !elem VALUES !elements
  !names.Append(!elem.Name)  -- Slower
ENDDO

-- Minimize database access
!elements = COLLECT ALL PIPE FOR ZONE  -- Single query
-- Better than individual element access
```

### 17.7 Code Reusability

```pml
-- Create utility functions
define function !!FormatNumber(!value is REAL, !decimals is REAL)
  !format = '0.' & '0'.Repeat(!decimals.Int())
  return !value.String(!format)
endfunction

-- Use objects for related data
define object POINT3D
  member .X is REAL
  member .Y is REAL
  member .Z is REAL
  
  define method .DistanceTo(!other is POINT3D)
    !dx = .X - !other.X
    !dy = .Y - !other.Y
    !dz = .Z - !other.Z
    return SQR(!dx*!dx + !dy*!dy + !dz*!dz)
  endmethod
enddefine
```

### 17.8 User Feedback and Progress Indicators

Provide clear feedback during long operations:

```pml
-- Show progress messages
prompt 'Selecting elements from database...'
var !selection collect all pipe for zone

prompt 'Evaluating expressions for $!selSize elements...'
var !names evaluate (NAME) for all from !selection

prompt 'Formatting output...'
-- Process data...

prompt 'Writing report...'
-- Write files...

prompt dismiss  -- Clear prompt when done
```

**Conditional Progress (Graphics Mode Only):**
```pml
-- Detect graphics mode
handle (61,115)  -- User cancel error = graphics mode
elsehandle none
   var !graphicsMode 1
endhandle

-- Show prompts only in graphics mode
if (defined(!graphicsMode)) then
   prompt 'Processing...'
endif

-- Always dismiss at end
if (defined(!graphicsMode)) then
   prompt dismiss
endif
```

**Pattern Benefits:**
- Users know script is running
- Prevents "hanging" perception
- No prompts in batch/background mode
- Professional user experience

### 17.9 Production Macro Template

Complete template incorporating best practices:

```pml
$* Report Generator v1.0
$* Usage: $M/MyReport [DEFINE|PROMPT|RUN filename type filter]

-- === MODE SELECTION ===
if (|$1| eq |DEFINE|) then
   -- Initialize global parameters
   var !!REP_FORMAT |CSV|
   var !!REP_SEPARATOR |;|
   var !!REP_COLUMNS ARRAY('NAME', 'TYPE', 'DIAMETER')
   return
endif

if (|$1| eq |PROMPT| or |$1| eq ||) then
   -- Show interactive dialog
   $M "%PMLUI%/MyDialog" RUN $0
   return
endif

-- === PARAMETER PROCESSING ===
var !fileName (trim(|$2|))
var !elemType (upcase(trim(|$3|)))
var !filter (trim(|$4|))

-- === STATE MANAGEMENT ===
var !units units
mm bore mm distance
onerror golabel /cleanup

-- === FILE OPERATIONS ===
if (vtext(!fileName) ne ||) then
   openfile "$!fileName" overwrite !output
   handle (41,324)
      confirm |OK to overwrite $!fileName?|
      handle (61,115)
         $!units
         return error 1 'Operation cancelled by user'
      endhandle
      if(not $_CALERT_USERDATA) then
         $!units
         return
      endif
      openfile "$!fileName" overwrite !output
   elsehandle any
      $!units
      return error 1 |Unable to open file: $!!ERROR.TEXT|
   endhandle
   var !write |writefile $!output|
else
   var !write |write|
endif

-- === DETECT GRAPHICS MODE ===
handle (61,115)
elsehandle none
   var !graphicsMode 1
endhandle

-- === MAIN PROCESSING ===
if (defined(!graphicsMode)) then
   prompt 'Collecting elements...'
endif

var !selection collect all $!elemType for ce
handle any
   if(vtext(!fileName) ne ||) then
      closefile $!output
   endif
   $!units
   return error 1 |Collection failed: $!!ERROR.TEXT|
endhandle

var !count (arraysize(!selection))
if (defined(!graphicsMode)) then
   prompt 'Processing $!count elements...'
endif

-- Evaluate attributes
var !names raw evaluate (NAME) for all from !selection
var !types raw evaluate (TYPE) for all from !selection

-- Build output
var !header |Name;Type|
$!write (vtext(!header))

var !lines compose quote (| |) $
  col val !names q $
  sep (|;|) $
  col val !types q

do !line values !lines
   $!write (vtext(!line))
enddo

-- === CLEANUP ===
if (vtext(!fileName) ne ||) then
   closefile $!output
endif
$!units
if (defined(!graphicsMode)) then
   prompt dismiss
endif
return

-- === ERROR HANDLERS ===
label /cleanup
handle any
   if (vtext(!fileName) ne ||) then
      closefile $!output
   endif
   $!units
   if (defined(!graphicsMode)) then
      prompt dismiss
   endif
   return error
endhandle
```

---

## 18. Complete Function Reference

### 18.1 String Functions

```pml
-- Length
!len = 'Hello'.Length()  -- Returns 5

-- Case conversion
!upper = 'hello'.Upcase()      -- 'HELLO'
!lower = 'HELLO'.Downcase()    -- 'hello'

-- Substring
!sub = 'Hello World'.Substring(1, 5)  -- 'Hello'
!sub = 'Hello World'.Substring(7)     -- 'World'

-- Search
!pos = 'Hello World'.Match('World')   -- Returns 7
!pos = 'Hello World'.Match('Bye')     -- Returns 0 (not found)

-- Replace
!new = 'Hello World'.Replace('World', 'PML')  -- 'Hello PML'

-- Trim
!trimmed = '  Hello  '.Trim()         -- 'Hello'
!trimmed = '  Hello  '.TrimLeft()     -- 'Hello  '
!trimmed = '  Hello  '.TrimRight()    -- '  Hello'

-- Split
!parts = 'a,b,c,d'.Split(',')  -- ARRAY('a', 'b', 'c', 'd')

-- Join
!joined = !array.Join(',')  -- 'a,b,c,d'

-- Repeat
!repeated = 'ab'.Repeat(3)  -- 'ababab'

-- Type conversion
!number = '123.45'.Real()
!text = (123.45).String()
```

### 18.2 Math Functions

```pml
-- Trigonometric
!sine = SIN(0.5)
!cosine = COS(0.5)
!tangent = TAN(0.5)
!arcsine = ASIN(0.5)
!arccosine = ACOS(0.5)
!arctangent = ATAN(0.5)

-- Powers and roots
!power = POW(2, 8)        -- 2^8 = 256
!squareRoot = SQR(25)     -- 5
!cube = POW(!value, 3)

-- Logarithms
!natural = LOG(100)       -- Natural logarithm
!base10 = ALOG(100)       -- Log base 10

-- Rounding
!absolute = ABS(-10.5)    -- 10.5
!integer = INT(3.7)       -- 3 (truncate)
!nearest = NINT(3.7)      -- 4 (round)
!floor = INT(3.7)         -- 3
!ceiling = INT(3.7 + 0.999999)  -- 4

-- Negate
!negative = NEGATE(10)    -- -10
```

### 18.3 Array Functions

```pml
!arr = ARRAY('a', 'b', 'c')

-- Size
!size = !arr.Size()

-- Add elements
!arr.Append('d')          -- Add to end
!arr.Set(2, 'x')          -- Set element 2

-- Remove elements
!arr.Delete(2)            -- Delete element 2
!arr.EmptyArray()         -- Clear all

-- Check empty
!isEmpty = !arr.Empty()

-- Sort
!arr.Sort()               -- Ascending
!arr.SortDescending()     -- Descending

-- Search
!index = !arr.Find('b')   -- Returns index or 0

-- Copy
!newArr = !arr.Copy()
```

### 18.4 Type Checking

```pml
-- Get type
!type = !var.ObjectType()  -- 'STRING', 'REAL', 'BOOLEAN', 'ARRAY'

-- Check type
!isString = !var.IS(STRING)
!isReal = !var.IS(REAL)
!isBoolean = !var.IS(BOOLEAN)
!isArray = !var.IS(ARRAY)

-- Check state
!isUnset = !var.IS(UNSET)
!isUndefined = !var.IS(UNDEFINED)
```

### 18.5 File System Functions

```pml
!file = object FILE()

-- File operations
!exists = !file.Exists('C:\temp\file.txt')
!file.Delete()
!file.Copy(!source, !dest)
!file.Move(!source, !dest)
!size = !file.Size()

-- Directory operations
!dir = object DIRECTORY()
!files = !dir.GetFiles('C:\temp')
!folders = !dir.GetDirectories('C:\temp')
!exists = !dir.Exists('C:\temp')
```

### 18.6 Date/Time Functions

```pml
-- Get current date/time
!now = NOW()
!date = DATE()
!time = TIME()

-- Format date/time
!formatted = !now.String('DD/MM/YYYY HH:MM:SS')
```

### 18.7 System Functions

```pml
-- Environment variables
!value = GETENV('PMLLIB')
!pdmsexe = GETENV('PDMSEXE')
!pdmsui = GETENV('PDMSUI')

-- Execute system command
SYSTEM('dir C:\temp')

-- Get username
!user = USERNAME()

-- Get hostname  
!host = HOSTNAME()

-- Module information
!module = MODULENAME()

-- Database information
!dbname = DBNAME()

-- Session ID
!sessionId = SESSIONID()
```

### 18.8 Project and Time Functions

```pml
-- Current units setting
var !units units
-- Returns current units (e.g., 'MM BORE MM DISTANCE')

-- Restore units
$!units  -- Expands to saved units command

-- Date and time
var !date clock date  -- Current date
var !time clock time  -- Current time

-- Project information
var !projCode proj code  -- Project code
var !projNumb proj numb  -- Project number
var !projName proj name  -- Project name
var !projDesc proj desc  -- Project description
var !projMess proj mess  -- Project message
```

**State Management Pattern:**
```pml
-- Save current state
var !units units
var !savedCE !!CE

-- Change state for processing
mm bore mm distance
ce /ZONE-1

-- Do work...

-- Restore state
$!units
ce $(!savedCE.Fullname())
```

### 18.8 Database Query Functions

```pml
-- Element queries
!name = NAME OF /EQUIP-1
!type = TYPE OF OWNER
!owner = OWNER OF /PIPE-1

-- Attribute queries
!diam = DIAM OF /PIPE-1
!length = LENGTH OF /PIPE-1
!position = POSITION OF /NOZZLE-1

-- Wildcard matching
!match = MATCHWILD('PIPE-100', 'PIPE-*')  -- TRUE

-- Bad reference check
!isBad = BADREF(!element)

-- Element state
!created = CREATE(!element)
!deleted = DELETED(!element)
!modified = MODIFIED(!element)

-- Empty check
!isEmpty = EMPTY(!element)
```

### 18.9 Geometric Functions

```pml
-- Distance between points
!dist = DISTANCE(!pos1, !pos2)

-- Angle between directions
!angle = ANGLE(!dir1, !dir2)

-- Cross product
!cross = CROSS(!dir1, !dir2)

-- Dot product
!dot = DOT(!dir1, !dir2)
```

### 18.10 Conversion Functions

```pml
-- String to number
!number = !string.Real()
!int = !string.Int()

-- Number to string
!text = !number.String()
!formatted = !number.String('F2')  -- 2 decimal places

-- Boolean conversion
!bool = !string.Boolean()
!text = !bool.String()

-- Unit conversion
!meters = 1000MM.ConvertTo('M')  -- 1.0 M
```

---

## 19. Advanced Topics

### 19.1 PML Directives

PML directives control PML system behavior:

```pml
-- Reload an object definition after changes
pml reload object OBJECTNAME

-- Rebuild PML file index in first PMLLIB directory
pml rehash

-- Rebuild all PML file indexes (SLOW - use sparingly)
pml rehash all

-- Re-read all pml.index files
pml index

-- Query location of PML file
q var !!PML.GetPathname('filename.pmlfnc')
```

### 19.2 PMLLIB Environment Variable and File Organization

The `PMLLIB` environment variable defines search paths for PML files:

```batch
REM Windows example - User directories FIRST
set PMLLIB=C:\Users\MyName\CustomPML;C:\ProjectPML;C:\AVEVA\PDMS\pmllib

REM Linux example
export PMLLIB=/home/user/custom_pml:/project/pml:/opt/aveva/pdms/pmllib
```

**Critical Best Practices:**

1. **User Customizations Must Be Separate**
   - **NEVER modify original AVEVA files** - updates will overwrite changes
   - Store all customizations in separate user directories
   - Place user directories **at the beginning** of PMLLIB path
   - Your functions will override AVEVA defaults when names match

2. **Search Path Order Matters**
   ```batch
   REM Correct order: User -> Project -> AVEVA
   set PMLLIB=C:\MyPML;C:\SharedPML;C:\AVEVA\PDMS\pmllib
   
   REM PML searches directories left-to-right
   REM First match wins
   ```

3. **Automatic File Indexing**
   - PML scans all directories in PMLLIB path at startup
   - Creates `pml.index` file in each directory automatically
   - Index contains references to all `.pmlfnc`, `.pmlobj`, `.pmlfrm` files
   - **DO NOT manually edit pml.index files**

4. **File Extensions and Loading**
   - `.pmlfnc` - Functions (loaded on first call)
   - `.pmlobj` - Object definitions (loaded on first use)
   - `.pmlfrm` - Forms (loaded when shown)
   - `.mac` - Macros (not indexed, must be explicitly called)

5. **Rebuilding Indexes**
   ```pml
   -- After adding new files to first PMLLIB directory
   pml rehash
   
   -- Rebuild ALL PMLLIB directory indexes (SLOW - use sparingly)
   pml rehash all
   
   -- Re-read all pml.index files without rebuilding
   pml index
   
   -- Reload specific function after editing
   pml reload function !!MyFunction
   pml reload object MYOBJECT
   ```

6. **Directory Structure Example**
   ```
   C:\Users\John\PML\          (First in PMLLIB)
   ├── pml.index               (Auto-generated)
   ├── MyUtilities.pmlfnc
   ├── CustomReport.pmlfnc
   ├── MyDialog.pmlfrm
   └── CustomObject.pmlobj
   
   C:\ProjectX\PML\            (Second in PMLLIB)
   ├── pml.index
   ├── ProjectUtils.pmlfnc
   └── ProjectForms.pmlfrm
   
   C:\AVEVA\PDMS\pmllib\      (Last in PMLLIB - AVEVA defaults)
   ├── pml.index
   └── [AVEVA files - DO NOT MODIFY]
   ```

7. **Function Override Example**
   ```pml
   -- File: C:\AVEVA\PDMS\pmllib\StandardReport.pmlfnc
   define function !!GenerateReport()
     $P Standard AVEVA Report
   endfunction
   
   -- File: C:\MyPML\StandardReport.pmlfnc (earlier in PMLLIB)
   define function !!GenerateReport()
     $P Custom Project Report
     -- Your custom implementation
   endfunction
   
   -- When called, YOUR version executes (first in path)
   !!GenerateReport()  -- Outputs: "Custom Project Report"
   ```

8. **Checking File Locations**
   ```pml
   -- Query where PML finds a function
   q var !!PML.GetPathname('MyUtilities.pmlfnc')
   
   -- Returns full path, showing which directory was used
   ```

### 19.3 Database Reference (DBREF) Objects

```pml
-- Create database reference
!elem = object DBREF('/ZONE-1/EQUIP-1')

-- Check if element exists
!exists = !elem.Exists()

-- Get element properties
!name = !elem.Name
!type = !elem.Type
!fullname = !elem.Fullname

-- Navigate hierarchy
!owner = !elem.Owner
!first = !elem.First
!last = !elem.Last
!next = !elem.Next
!prev = !elem.Previous

-- Check element state
!isCreated = !elem.IsCreated()
!isDeleted = !elem.IsDeleted()
!isModified = !elem.IsModified()

-- Bad reference check
!isBad = BADREF(!elem)
```

### 19.4 Geometric Objects

#### POSITION
```pml
-- Create position
!pos = object POSITION(1000, 2000, 3000)

-- Access components
!x = !pos.East
!y = !pos.North
!z = !pos.Up

-- Set components
!pos.East = 1500
!pos.North = 2500
!pos.Up = 3500

-- Position in different coordinate systems
!worldPos = POS IN WORLD
!sitePos = POS IN SITE
!zonePos = POS IN ZONE
```

#### DIRECTION
```pml
-- Create direction
!dir = object DIRECTION(1, 0, 0)  -- X-axis

-- Standard directions
!north = object DIRECTION(0, 1, 0)
!up = object DIRECTION(0, 0, 1)

-- Access components
!x = !dir.East
!y = !dir.North
!z = !dir.Up

-- Normalize direction
!normalized = !dir.Normalize()
```

#### ORIENTATION
```pml
-- Create orientation
!ori = object ORIENTATION()

-- Get orientation from element
!elemOri = ORIENTATION OF /PIPE-1

-- Set orientation
!ori.East = object DIRECTION(1, 0, 0)
!ori.North = object DIRECTION(0, 1, 0)
!ori.Up = object DIRECTION(0, 0, 1)
```

### 19.5 Precision and Tolerance

```pml
-- Real number comparisons use tolerance
!a = 1.0
!b = 1.0000001

-- These may be considered equal due to tolerance
IF !a EQ !b THEN
  $P Equal within tolerance
ENDIF

-- For exact comparison, use subtraction
IF ABS(!a - !b) LT 0.0001 THEN
  $P Equal within 0.0001
ENDIF
```

### 19.6 Session and System Information

```pml
-- Current user
!user = USERNAME()

-- Current module
!module = MODULENAME()

-- Current database
!db = DBNAME()

-- Working units
!distUnit = UNIT DIST
!angleUnit = UNIT ANGL

-- Session ID
!sessionId = SESSIONID()

-- Get environment variable
!pmllib = GETENV('PMLLIB')
!pdmsexe = GETENV('PDMSEXE')
```

### 19.7 Block Evaluation

```pml
-- Evaluate expression for collection
!pipes = COLLECT ALL PIPE FOR ZONE

-- Get all names
!names = EVALUATE (NAME) FOR ALL FROM !pipes

-- Get with filter
!bigPipes = EVALUATE (NAME) FOR ALL FROM !pipes WITH DIAM GT 200

-- Calculate values
!lengths = EVALUATE (LENGTH) FOR ALL FROM !pipes

-- Complex expressions
!areas = EVALUATE (DIAM * DIAM * 3.14159 / 4) FOR ALL FROM !pipes
```

### 19.8 Using USING Blocks

```pml
-- Sort array using USING block
USING !myArray
  SORT VALUES ASCENDING
ENDUSING

-- Block evaluate with USING
USING !collection
  !result = EVALUATE (NAME) FOR ALL WITH DIAM GT 100
ENDUSING
```

### 19.9 Rules (Late Evaluation)

```pml
-- Define variable for late evaluation
!rule = RULE (DIAM * 2)

-- Rule evaluated when accessed
!pipes = COLLECT ALL PIPE FOR ZONE
DO !pipe VALUES !pipes
  !doubleDiam = !rule  -- Evaluated for current element
  $P Pipe: $(!pipe.Name), Double Diam: $!doubleDiam
ENDDO
```

### 19.10 Undo/Redo Support

```pml
-- Enable undo
UNDO ON
  NEW PIPE /PIPE-1
  DIAM 150
  LENGTH 5000
UNDO OFF

-- Undo last operation
UNDO

-- Redo
REDO

-- Check undo status
!canUndo = CANUNDO()
!canRedo = CANREDO()
```

### 19.11 Alpha Log

```pml
-- Write to alpha log
ALPHALOG 'This message goes to alpha.log'

-- Query alpha log status
!isOn = ALPHALOGGING()

-- Turn alpha logging on/off
ALPHALOG ON
ALPHALOG OFF
```

### 19.12 PML Tracing

```pml
-- Enable PML tracing
TRACE ON

-- Trace specific commands
TRACE COMMAND ON
TRACE FUNCTION ON
TRACE METHOD ON

-- Disable tracing
TRACE OFF

-- Trace to file
TRACE FILE 'C:\temp\trace.log'
```

### 19.13 Suspending PML Execution

```pml
-- Pause execution (for debugging)
PAUSE

-- Wait for user input
WAIT

-- Delay execution (milliseconds)
DELAY 1000  -- Wait 1 second
```

### 19.14 Querying PML State

```pml
-- Query current PML file stack
q pmlstack

-- Query variable values
q var !myVariable

-- Query all variables
q var *

-- Check what can be typed next
q next
```

### 19.15 MATCHWILD Function

```pml
-- Wildcard matching
!text = 'PIPE-100-6IN-CS'

-- Match pattern (* = any characters, ? = single char)
!matches = MATCHWILD(!text, 'PIPE-*')      -- TRUE
!matches = MATCHWILD(!text, '*-6IN-*')     -- TRUE
!matches = MATCHWILD(!text, 'PIPE-???-*')  -- TRUE
!matches = MATCHWILD(!text, 'VALVE-*')     -- FALSE

-- Case-insensitive matching
!matches = MATCHWILD(!text.Upcase(), 'PIPE-*')
```

### 19.16 Working with Current Element (CE)

```pml
-- Get current element
!currentElem = !!CE

-- Set current element
CE /ZONE-1/EQUIP-1

-- Save and restore CE
!savedCE = !!CE
-- Do work with different CE
CE /OTHER/ELEMENT
-- Restore
CE $(!savedCE.Fullname)
```

---

## Appendix A: Keywords Reference

### Reserved Keywords

```
AND, ANY, ARRAY, AT, BOOLEAN, BREAK, BREAKIF, BY,
CALLBACK, DEFINE, DELETE, DO, DOWNCASE, ELSEHANDLE,
ELSE, ELSEIF, ENDDO, ENDFUNCTION, ENDHANDLE, ENDIF,
ENDMETHOD, ENDUSING, EQ, EXIT, FALSE, FOR, FROM,
FUNCTION, GE, GT, HANDLE, IF, INDEX, IS, LE, LT,
MEMBER, METHOD, NE, NOT, OBJECT, OF, OR, REAL,
RETURN, SETUP, SKIP, SKIPIF, STEP, STRING, THEN,
TO, TRUE, UNDEFINED, UNSET, UPCASE, USING, VALUES,
ALPHALOG, BADREF, CE, COLLECT, DELAY, DEFINED,
EVALUATE, GETENV, GOLABEL, LABEL, MODULENAME, PAUSE,
RULE, SESSIONID, TRACE, UNDEFINED, UNDO, REDO, WAIT
```

---

## Appendix B: Common Error Messages

| Error | Meaning | Solution |
|-------|---------|----------|
| Object not found (OBJNOTFND) | Element doesn't exist in database | Check path, verify element exists with HANDLE |
| Invalid syntax | PML syntax error | Check command format, brackets, keywords |
| Type mismatch | Wrong data type used | Verify variable types match expected |
| File not found | File path incorrect | Check file exists, verify path |
| Division by zero | Attempted to divide by zero | Add check before division |
| Array index out of bounds | Invalid array index | Check array size before accessing |
| Name already in use | Duplicate element name | Use unique names or check existing |
| Attribute undefined | Element doesn't have attribute | Check element type, use HANDLE |
| Bad reference (BADREF) | Reference to deleted element | Check with BADREF() function |
| Undefined variable | Variable not declared | Declare variable or check spelling |
| Method not found | Method doesn't exist for object | Check object type, verify method name |
| PML file not found | .pmlfnc/.pmlobj file missing | Check PMLLIB path, run pml rehash |

---

## Appendix C: File Organization

### Recommended Directory Structure

```
%PMLLIB%/
├── functions/
│   ├── utilities.pmlfnc
│   ├── calculations.pmlfnc
│   └── validators.pmlfnc
├── objects/
│   ├── Point3D.pmlobj
│   ├── Rectangle.pmlobj
│   └── Circle.pmlobj
├── forms/
│   ├── MainForm.pmlfrm
│   ├── SettingsForm.pmlfrm
│   └── ReportForm.pmlfrm
├── macros/
│   ├── initialize.mac
│   ├── batch_process.mac
│   └── export_data.mac
└── resources/
    └── pixmaps/
        ├── icons/
        └── images/
```

---

## Appendix D: Quick Reference Card

### Variable Declaration
```pml
!local = value              -- Local variable
!!global = value            -- Global variable
```

### Control Structures
```pml
IF condition THEN ... ENDIF
DO !i FROM 1 TO 10 ... ENDDO
DO !v VALUES !array ... ENDDO
HANDLE ANY ... ELSEHANDLE ... ENDHANDLE
```

### Functions
```pml
define function !!Name(!param is TYPE)
  return !value
endfunction
```

### Objects
```pml
define object NAME
  member .Field is TYPE
  define method .Method()
    -- code
  endmethod
enddefine
```

### Forms
```pml
setup form !!FormName
  title 'Title'
  text .field tag 'Label:' width 20
  button .ok |OK| callback .okMethod
exit
```

---

**End of PML Complete Reference**

*This reference document consolidates all information from the E3D/PDMS PML training materials and manuals. It is designed to be a complete standalone reference for PML programming.*

---

**Document Version**: 1.0  
**Last Updated**: December 2025  
**Based on**: E3D PML2, PDMS 12.1+
