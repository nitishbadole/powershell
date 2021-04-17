###################################################################################
# S1.0 Execution policy
# S2.0 Help
# S3.0 Remote invocation
# S4.0 Variables
# S5.0 Looping
# S6.0 Switch
# S7.0 Major cmdlets (foreach-object, where-object, select-object)
# S8.0 Quoting, string expansion
# S9.0 Formatting, outputting, files and redirection
# S10.0 Functions
# S11.0 Advanced functions
# S12.0 Scripts
# S13.0 Files and providers
# S14.0 XML

<#
This is a multi-line comment.
It only works in PowerShell v2.
#>


###################################################################################
# S0.0 TODO
# How to make a new object in a pipeline, e.g. PS SQLSERVER:\sql\accaldev1\default\databases\laadmin\tables> ls | foreach { "{0}`t{1},{2}" -f $_.Name, $_.IndexSpaceUsed, $_.DataSpaceUsed }
# How to "grep" a list of members.
# How to control what properties are displayed when outputting data.


###################################################################################
# S1.0 Execution policy
Get-Help –Online about_execution_policies    # TODO
Get-ExecutionPolicy
Set-ExecutionPolicy remotesigned             # Allow local scripts (created by me) to be run. But does not allow remote scripts to run.
                                             # See Payette, Section 8.1.1, p277.

###################################################################################
# S2.0 Help
get-help about*                      # help about concepts
get-help dir
get-help -online dir                 # This is sometimes more complete and up to date.
get-command -verb stop 
get-command -noun process
get-command -CommandType alias       # List all aliases
get-command cat                      # Same as the 'get-alias' below.
get-alias cat                        # Find the command for an alias
get-alias -definition get-childitem  # Find an alias for a command
get-member [-static]                 # List members of an object. Easy to use to discover static methods.
get-member -inputobject $a           

###################################################################################
# S3.0 Remote invocation
get-help -online invoke-command                # alias is icm
(gcm mmc.exe).FileVersionInfo                  # Example: get info about mmc on the local machine
icm $mach1 { (gcm mmc.exe).FileVersionInfo }   # Remote example: run the same command on the machine identified by variable $mach1

Enter-PSSession server1                        # Start an interactive session on a remote machine.
Exit-PSSession                                 # And quit it.


###################################################################################
# S4.0 Variables
$n = 3                # Scalar
$a = 1,2 + 3,4        # Array of 4 items
cd $HOME              # Environment variables use Unix syntax, not %HOME% syntax
[int] $a              # Declare a variable of a particular type.
$a,$b,$c = 1,2,3,4    # Multiple assignment.

$( ... )              # Subexpression operator - allows any code where a value expression would be required
@( ... )              # Array subexpression operator - as above, but guarantees the result will be an array.

# Suffixes. kb, mb, gb, tb, pb. Upper or lower case. They are powers of 2, not 10.
$size = 20kb
$size

# Type conversion rule: left hand operand "wins" for type conversion.
"002" -eq 2             # 2 becomes a string "2", not equal to "002"
2 -eq "002"             # "002" becomes an int 2, is equal
[int] "002" -eq 2       # Cast to int.
$a = [char[]]"Hello"    # Convert a string to a character array.
$b = -join $a; $b       # Convert a character array to a string.

# Hash tables.
$emptyHash = @{}
$user = @{ FirstName = "John"; LastName = "Smith"; PhoneNumber = "555-1212" }    
$user.PhoneNumber               # Property access
$user["LastName"]               # Smith
$user["FirstName", "LastName"]  # John Smith
$user["Unk"]                    # null
$user.Keys
$user.Values
$user.Count

# Note that PowerShell generally treats hash tables as scalars, not enumerables.
# This means that using it in a "foreach" only loops once. See Bruce Payette, p87.
# You have to call GetEnumerator yourself.
foreach ($pair in $user.GetEnumerator())
{
    $pair.key + " is " + $pair.value
}

# Arrays. Construct using the comma operator.
$a = 1,2,3
$a += 4,5,6,7
$a
$a = 1,"hello",3.14    
$a[0]
$a[1]
$a[2]
# polymorphic, it is a System.Object[].
$a.GetType().FullName

"Hello $a"    # Can be interpolated into strings.
$a[99]        # no-op!
$a[99] = 3    # crashes!

$oneElement = ,1                       # Create array with one element
$emptyArray = @()                      # Create array with zero elements
$emptyArray.GetType().FullName
$oneElement = @(1)                     # Alternative way of creating an arrray with one element.

# New List<int>().
$list = new-object system.collections.generic.list[int]
$list.Add(23)
$list.Add(34)
$list.Count
$list.Add("Hello")                     # Error.

# Calling static methods.
[string]::Join(' + ', ("one","two","three"))    # Call String.Join on an array of strings.
[Math]::Pi

# Operators. The standard operators are case-insensitive.
# To get case-sensitive versions, prefix with "c", e.g. "ceq".
# There is also an "i" prefix for case-insensitivity, but this is redundant.

# Comparison      : eq  ne  gt  ge  lt  le
# Containment     : contains  notcontaints
# Patterns        : like  notlike  match  notmatch  replace  split  join
# Logical         : and  or  not  xor  band  bor  bnot  bxor   (for making compound expressions)

5 -eq 7
5 -ge 5
5 -gt 4
"a" -eq "A"      # true
"a" -ieq "A"     # true
"a" -ceq "A"     # false

# Comparisons on collections.
1,2,3,5,3,3,6,78,8,3 -gt 4       # Acts as a filter
$false,$true -eq $false          # Find all elements equal to $false and return them!
$false,$true -contains $false    # See difference w.r.t. the above.
1..10 -contains 2
"hello","world" -contains "HELLO"    # True
"hello","world" -ccontains "HELLO"   # False

# WILDCARD Pattern matching (for filenames, like and notlike)
"one" -like "o*"                     # True
"one","oval","lords" -like "o*"      # "one","oval"   
"one","oval","lords" -notlike "o*"   # "lords"
"hello" -like "*el*"                 # True
"hello" -like "*el?"                 # False
"hello" -like "*ell?"                # True
"alpha" -like "[a-z]lpha"            # True
"alpha" -like "[b-z]*"               # False
"alpha" -like "[abc]*"               # True

# REGEX Pattern matching (for strings, match, notmatch, replace)
# They set the $matches variable (a hash).
"Hello" -match "[jkl]"               # True
"Hello" -notmatch "[jkl]"            # False
"Hello" -replace "He","Ye"           # "Yello"
"Hello" -replace "He"                # "llo"
"Hello" -match "(He)l(lo)"; $matches                              # ( ) identify submatches.
"abcdef" -match "(?<o1>a)(?<o2>((?<e3>b)(?<e4>c))de)f"; $matches  # ?<name> is for named submatches.

"a,b,c,d" -replace "\s*,\s*","+"                                # "a+b+c+d".  Replace all commas & their surrounding whitespace with "+" signs.
"rose is red, sky is blue" -replace 'is (red|blue)','was $1'    # $1 is the first match BECAUSE WE ARE IN SINGLE QUOTES.
# ${name} - sub named match, $$ - a literal $, $& - the entire match
# There are others, see Bruce Payette, p 138.

# -join works on collections of STRINGS only. The parentheses are necessary unless
# you use a variable. This is because unary operators such as join have a higher
# precedence than binary operators such as comma.
$out = -join ("1","2","3")      # This is UNARY join.
$out                            # The string "1,2,3"

$out = -join "1","2","3"
$out                            # An array of 3 elements.

$in = "1","2","3"
$out = -join $in
$out                            # The string "1,2,3"


# Binary join: <collection> -join "str"
1..10 -join "*"

# Unary split. Always splits on whitespace.
-split "hello cruel world"      # An array of 3 elements.
# Binary split. "a, b ,c" split "\w*,\w*",n,MatchType,Options
'a:b:c:d:e' -split ':'
'a:b:c:d:e' -split ':',3
$c = "Mercury,Venus,Earth,Mars,Jupiter,Saturn,Uranus,Neptune"
$c -split {$_ -eq "e" -or $_ -eq "p"}
'a*b*c' -split "*",0,"simplematch"           # Don't interpret the split pattern as a regular expression.


# Arrays.
@("bbb", "aaa", "ddd" | sort)[0]       # @ always converts subexpression into an array.
"hello"[-1]                            # -1 is the last element.
"hello"[2..4]                          # Slicing using the range operator, which is just shorthand for...
"hello"[2,3,4]                         # slicing using indexes...
"hello"[4,0]                           # which is more flexible.
"hello"[-4..-1]                        # Last 4 elements.

# Dot operator.
"hello".("len" + "gth")                # Properties can be invoked dynamically.
$prop = "length";                      # Another example of dynamic invocation.
"hello".$prop

$method = "sin"                        # We can also invoke methods dynamically.
[math]::$method                        # Returns info about the method (a PSMethodInfo object)
[math]::$method.invoke(3.14)

# Namespaces.
$global:myvar = 3                      # Create myvar in the global namespace.
$env:SystemRoot                        # Environment variables.
"kjlkjlj" > c:\temp\output.txt
${c:\temp\output.txt}                  # Use the C drive as a namespace, read a file from it.
${c:new.txt} = ${c:old.txt} -replace 'is (red|blue)','was $1'   # Write to a new file (can also do with redirection).
${c:old.txt} = ${c:old.txt} -replace 'is (red|blue)','was $1'   # Overwrite the existing file (can also do with redirection).

# Splatting.
function s ($x, $y, $z) { "x=$x, y=$y, z=$z, args=$args" }
$list = 1,2,3
s $list
s @list                                # Use @ to break lists apart.
$list += 5,6,7                         # Extra elements get appended to the $args variable.
s @list
# Hash tables will be splatted by matching the names of keys to the names of arguments.


###################################################################################
# S5.0 Looping
for ($i=0; $i -lt 5; $i++) { $i }                     # Basic loop
for ($i=0; $($y = $i*2; $i -lt 5); $i++) { $y }       # Each component is actually a pipeline / subpexpression
foreach ($i in 1..10) { if ($i % 2) {"$i is odd"}}
1..10 | foreach { if ($_ % 2) {"$_ is odd"}}

$i = 0
while ($i++ -lt 10)
{
    if ($i % 2) { "$i is odd" }
}

$svchosts = get-process svchost
for ($($total=0;$i=0); $i -lt $svchosts.count; $i++)
{
    $total += $svchosts[$i].handles
}
$total

###################################################################################
# S6.0 Switch
switch (2) { 1 { "One" } 2 { "two" } 2 {"another 2"} default { "default" } }        # All matches are executed (use "break").
switch ('abc') {'abc' {"one"} 'ABC' {"two"}}                                        # Defaults to case-insensitive.
switch -case ('abc') {'abc' {"one"} 'ABC' {"two"}}                                  # -casesensitive is the option. (Can also use this with regexes).
switch -wildcard ('abc') {a* {"astar"} *c {"starc"}}                                # Pattern matching. $_ is the element that matched.
switch -regex ('abc') {^a {"a*: $_"} 'c$' {"*c: $_"}}                               # Regex matching.
switch -regex ('abc') {'(^a)(.*$)' {$matches}}                                      # Can use $matches[0] etc. 0 = overall key, 1 = first submatch, 2 = the remainder.

switch (5)                                                                          # Switching using expressions.
{
    {$_ -gt 3} {"greater than three"}
    {$_ -gt 7} {"greater than 7"}
}

switch (1,2,3,4,5,6)                                                                # Can use switch as a loop.
{
    {$_ % 2} {"Odd $_"; continue}                                                   # continue = start next loop iteration; break = exit the loop
    4 {"FOUR"}
    default {"Even $_"}
}

$dll = $txt = $log = 0                                                              # Example: Counting files in a directory.
switch -wildcard (dir c:\windows) {*.dll {$dll++} *.txt {$txt++} *.log {$log++}}
"dlls: $dll text files: $txt log files: $log"

$au = $du = $su = 0
switch -regex -file c:\windows\windowsupdate.log                                    # Process a file line by line.
{
    'START.*Finding updates.*AutomaticUpdates' {$au++}
    'START.*Finding updates.*Defender' {$du++}
    'START.*Finding updates.*SMS' {$su++}
}
"Automatic:$au Defender:$du SMS:$su"


###################################################################################
# S7.0 Major cmdlets (foreach-object, where-object, select-object)
dir *.txt | foreach-object { $c += 1; $l += $_.length }       # $_ is the current object.
dir *.txt | foreach { $c += 1; $l += $_.length }              # 
dir *.txt | % { $c += 1; $l += $_.length }                    # These 3 are all equivalent. The last 2 are aliases.
gps svchost | % { $t = 0 }{ $t += $_.handles }{ $t }          # Can take 3 script blocks, before, during and after.
Get-Process | %{$_.modules} | sort -u modulename              # Note that foreach unravels collection.

1..10 | where-object {!($_ -band 1)}
1..10 | where {!($_ -band 1)}
1..10 | ? {!($_ -band 1)}                                     # These 3 are all equivalent.

$result = $(for ($i=1; $i -le 10; $i++) {$i})                 # Using a statement as a value.
$result = for ($i=1; $i -le 10; $i++) {$i}                    # Again.
$var = if (! $var) { 12 } else {$var}                         # Same with "IF" statements.

get-help -online select-object
  Select-Object [[-Property] <Object[]> ] [-ExcludeProperty <String[]> ] [-ExpandProperty <String> ]   # Pick properties from an object
  Select-Object [-First <Int32> ] [-Last <Int32> ] [-Skip <Int32> ] [-Unique] [-Wait]                  # Pick object instances from a collection

dir c: | sort -property length -descending | select-object -first 3
dir c: | sort -property length -descending | select-object -property name


###################################################################################
# S8.0 Quoting, string expansion
cat 'c:\temp\hello world.txt'          # Use single quotes for paths with spaces
cat c:\temp\hello` world.txt           # Use backtick to quote a single character
$v = "files"
cd "c:\program $v"                     # Double quotes do variable expansion.
cd "c:\program `$v"                    # And you can use backtick to suppress expansion.
"`$v is:\n$v"                          # Not special to Powershell...
"`$v is:`n$v"                          # This gives a newline. Escape sequences are only processed in double quotes.

"2 + 2 is $(2 + 2)"                    # Subexpressions can be expanded.
"Numbers 1 thru 10: $(for ($i=1; $i -le 10; $i++) { $i })."

# Here-strings. Note that the delimiters must be on their own lines as shown, or it won't work.
# The double-quoted version will do expansion.
$here = @"
long
multi-line
string
"@

function ql { $args }                   # Quote-list (returns args as a list)
function qs { "$args" }                 # Quote-string (returns args as a single string)


###################################################################################
# S9.0 Formatting, outputting, files and redirection. Powershell defines default formatters and outputters
get-command format-*
ls c:\ | ft                         # format-table or format-list are the most used
ls c:\ | Format-Custom -Depth 1     # format-custom can be handy for debugging

gcm out-*
ls c:\ | out-file -encoding utf8 'c:\temp\listings.txt' 
ls c:\ | out-gridview

2 + 2 > c:\temp\foo.txt                 # Redirect Stdout - Overwrite
2 + 3 >> c:\temp\foo.txt                # Redirect Stdout - Append.
dir nosuchfile.txt 2> err.txt           # Redirect Stderr - Overwrite.
dir nosuchfile.txt 2>> err.txt          # Redirect Stderr - Append.
dir nosuchfile.txt 2>&1>                # Redirect Stderr to Stdout (and print to console)
myScript 2>&1 > output.txt              # Redirect Stderr to Stdout (and send to output.txt)
myScript > output.txt 2>&1              # Alternative syntax for the above.
$a = myScript 2>&1                      # Capture output and error objects in $a.
$error = $($output = myScript) 2>&1     # Capture output objects and error objects in separate variables.

type c:\temp\foo.txt
mkdir foo > $null               # Same as piping to Out-Null:   mkdir foo | out-null

# How to format strings.
"Hello {0}, World {1}" -f "Phil",[DateTime]::now                       # Use -f with an array.
[string]::Format("Hello {0}, World at {1}", "Phil", [DateTime]::Now)   # Call the static method like in C#.

@'
line1
line2
line3
'@ > out.txt

$text = get-content out.txt; $text; $text.length    # Get file as an array of strings.
$one = $text -join "`r`n"; $one; $one.length        # Convert to a single string.

###################################################################################
# S10.0 Functions.
# Note that functions in Powershell are commands, not like functions in other languages.
function PassArgsSeparately
{
    "Hello $args, how are you?"
}
PassArgsSeparately Phil and Bob


function CountArgs
{
    "`$args = $args"
    "`$args.count = " + $args.count
}
CountArgs a and b and c
CountArgs a,b,c             # Passes 1 argument, an array of strings.


function MultipleAssignment
{
    $x, $y = $args
    $x + $y
}
MultipleAssignment 2 5


function FormalParameters_Subtract($from, $count)
{
    $from - $count
}
FormalParameters_Subtract 10 2
FormalParameters_Subtract -count 3 -from 30
#FormalParameters_Subtract hello world          # Runtime failure


function TypedParameters_Multiply([int] $a, [int] $b)
{
    $a * $b
}

TypedParameters_Multiply 4 6
TypedParameters_Multiply 4 6.3      # Rounds down to 6
TypedParameters_Multiply 4 6.56     # Rounds up to 7
TypedParameters_Multiply 4 "3"      # Automatic coercion
#TypedParameters_Multiply 4 hello    # Failure


function NoOverloading([int] $a, [int] $b)
{
    $a * $b
}
function NoOverloading([int] $a, [int] $b, $c)    # This replaces the first version
{
    $a - $b
}

NoOverloading 5 3


function VariableArgs([int] $a, [int] $b)
{
    "VariableArgs: `$a = $a, `$b = $b, `$args = $args"
}
VariableArgs 1
VariableArgs 1 2
VariableArgs 1 2 3
VariableArgs 1 2 3 4 hello


function DefaultArgumentValues([int] $a = 1, [int] $b = 5)
{
    "DefaultArgumentValues: `$a = $a, `$b = $b, `$args = $args"
}
DefaultArgumentValues
DefaultArgumentValues 50
DefaultArgumentValues 51 52
DefaultArgumentValues 51 52 53
DefaultArgumentValues 51 52 53 54 hello


function DefaultArgumentValues_CanBeExpressions([datetime] $d = $(get-date))
{
    $d.dayofweek
}
DefaultArgumentValues_CanBeExpressions


function Switches_AreOptional([switch] $please)
{
    if ($please)
    {
        "You're welcome"
    } 
    else
    {
        "Don't be rude"
    }
}
Switches -please
Switches


function Booleans_AreMandatoryIfSpecified([bool] $x)
{
    $x
}

Booleans_AreMandatory
#Booleans_AreMandatory -x            # Error
Booleans_AreMandatory -x $true


function Functions_ReturnValuesAsAnArray()
{
    # The result of a function is the result of each expression.
    # They are captured and returned as an array.
    1
    2
    3.14
    999
}
Functions_ReturnValuesAsAnArray
$result = Functions_ReturnValuesAsAnArray
$result.length                      # Prints 4
$result[2]


function Functions_Scalar()
{
    # Except for scalar functions, which return scalar values.
    72
}
$result = Functions_Scalar
$result.length                      # Prints 4
$result


function Functions_BewareOfPhantomReturns()
{
    $al = new-object system.collections.arraylist
    $args | foreach { $al.add($_) }
    $al
}
# Returns 0 1 2 3 because that is the value returned by add(), which is then
# collected into the return array.
Functions_BewareOfPhantomReturns a b c d    


function Functions_BewareOfPhantomReturnsFixed()
{
    $al = new-object system.collections.arraylist
    $args | foreach { [void] $al.add($_) }
    $al
}
# Returns 0 1 2 3 because that is the value returned by add(), which is then
# collected into the return array.
Functions_BewareOfPhantomReturnsFixed a b c d    


function Pipeline_SpecialInputVariable()
{
    # $input is an enumerator over all inputs from the pipeline
    $total = 0
    $input | foreach { $total += $_ }
    $total
}
1..5 | Pipeline_SpecialInputVariable
Pipeline_SpecialInputVariable


# A filter has the same declaration syntax as a function, but uses "filter" instead of "function".
# Function: when used in a pipeline halts streaming - the previous element runs to completion before the function runs ONCE. Uses $input.
# Filter: is run once FOR EACH element in the pipeline. Uses $_ to contain the current pipeline element.
#    >>>> In practice, for "foreach" cmdlet and the process block remove the need for "filter" and it is not used.


function MyCmdlet($x)
{
    begin {
        $c = 0
        "In Begin, c is $c, x is $x"
    }
    process {
        $c++
        "In Process, c is $c, x is $x, `$_ is $_"
    }
    end {
        "In End, c is $c, x is $x"
    }
}
1..5 | MyCmdlet 22
MyCmdlet 22              # Runs the Process block once.

# Pipeline processing algorithm (it is single-threaded)
# In a pipeline of N commands, all the begin{} blocks are run, then the
# process{} block of the first command is run. If it produces an object
# then it is passed to the next command, until the end of the pipeline
# is reached. Then all the end{} blocks are run.


###################################################################################
# S11.0 Advanced functions
function AllParameterMetaData()
{
    # Help is only displayed when the shell is prompting for missing parameter
    # values, so it isn't very useful.
    param(
        # All possible values.
        [Parameter(Mandatory=$true,
        Position=0,
        ParameterSetName="set1",
        ValueFromPipeline=$false,
        ValueFromPipelineByPropertyName=$true,
        ValueFromRemainingArguments=$false,
        HelpMessage="some help this is")]
        [int]
        $p1 = 0
    )    
    begin {
    }
    process {
    }
    end {
    }
}


function TypicalParameterMetaData()
{
    param(
        [Parameter(Position = 0, Mandatory = $true)][int] $p1,
        [Parameter(Position = 1, Mandatory = $true)][int] $p2,
        [Parameter(Position = 2, Mandatory = $true)][int] $p3
    )
}


function ValueFromPipelineOrParam()
{
    param(
        [Parameter(ValueFromPipeline = $true)] $x
    )
    process {
        $x
    }
}
ValueFromPipelineOrParam 123
123, 12 | ValueFromPipelineOrParam


function ValueFromPipelineByPropertyName()
{
    param(
        # Instead of binding the entire object in the pipeline, just bind 1 property (called DayOfWeek)
        [Parameter(ValueFromPipelineByPropertyName=$true)] $DayOfWeek
    )
    process {
        $DayOfWeek
    }
}
Get-Date | ValueFromPipelineByPropertyName
ValueFromPipelineByPropertyName (Get-Date)    # But this still binds the entire date object
ValueFromPipelineByPropertyName ((Get-Date).DayOfWeek)    # So you have to extract it yourself


###################################################################################
# S12.0 Scripts
# To run a script from the command line
# (good) because it inteprets "C:\my scripts\script.ps1" as the command
#   powershell.exe File "C:\my scripts\script.ps1" data.csv

# or (bad, because it interprets "C:\my" as the command and everyting else as arguments.
#   powershell.exe Command "C:\my scripts\script.ps1" data.csv           # gives 2 args, splut at "C:\my"

# Use the param statement to specify script arguments. It must be the first
# executable statement in the script.
@'
param($name="world")
"Hello $name!"
'@ > hello.ps1


return            # Exit a script, but only if called at the top level (not in a function).
exit [code]       # Exit a script from any level. The code is available from cmd.exe as ERRORLEVEL, e.g. "echo %ERRORLEVEL%".
. ./script.ps1    # Dot source a script. Evaluates variables in the current scope. e.g. you can change variables defined in the outer script.

powershell -File "my script.ps1" data.csv   # Run a script from a (CMD.EXE/DOS) batch file.
powershell -Command "$a = 1"                # Run a one-liner.


###################################################################################
# S13.0 Files and providers
get-psprovider                # List all providers.
dir env:*
dir alias:*gc*                # List aliases.
cd HKLM:
cd C:
New-PSDrive -Name docs -PSProvider filesystem -Root (Resolve-Path ~/*Documents)    # Create a new PS drive "cd docs:"
notepad (Resolve-Path docs:/junk.txt).ProviderPath                                 # Convert a PSPath into a Provider Path.
                              
1..3 | foreach { "This is file $_" > "file$_.txt"}          # Generate some files (there may be other text files in this folder).
Get-Content *.txt                                           # Read them in. Concatenates all files into an array of strings.
(gc *.txt).gettype()

# Set-Content writes objects pass-thru to the file. It calls ToString() on anything that is not a string.
# Out-File on the other hand tries to format the things that you give it (like to the console).
1..3 | %{ "This is file $_" | Set-Content -Encoding ascii file$_.txt }       

# How to grep a file.
Select-String "wildcard description" $pshome/en-US/about*.txt


###################################################################################
# S14.0 XML
$d = [xml] "<top><a>one</a><b>two</b><c>3</c></top>"        # XML can be accessed "pathwise".
$d
$d.top
$d.top.a
$el = $d.CreateElement("d");                                # Creates <d></d>
$el.set_InnerText("Hello")
$d.top.AppendChild($el)

$attr = $d.CreateAttribute("BuiltBy")
$attr.Value = "Windows PowerShell"
$d.DocumentElement.SetAttributeNode($attr)

$d["BuiltBy"]
