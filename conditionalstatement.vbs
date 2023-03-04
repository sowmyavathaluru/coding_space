Dim str1,str2             'if else statements 
str1 = "java"
str2 =  "python" 
str3 = len(str1) 
str4 = len(str2) 
if str3 > str4 Then 
'MsgBox "Both the strings are equal" 
else 
'MsgBox" both the strings are not equal" 
end if

firstname = "vathaluru"        'if elseif statements 
lastname = "sowmya"
str = (Left(firstname,3))
str1 = (Left(lastname,3))
'MsgBox str 
if str = str1  Then 
'MsgBox "both strings having left elements are same"  
elseif str > str1 Then 
'MsgBox "string1 is greater than string2" 
else 
'MsgBox "both the strings are equal" 
end if   

Dim value                  
value = InputBox("enter a value") 
MsgBox IsNumeric (value) 
'MsgBox TypeName(value)
if IsNumeric(value) = true Then  

if value >= 10 And value <= 100 Then 
MsgBox "value is a small number " 

elseif value > 100 AND value <= 1000 THEN 
MsgBox "value is a medium Number" 

Elseif value > 1000 AND value <=10000 THEN 
MsgBox "value is a big number" 

ElseIf value > 10000 Then 
MsgBox "value is a high number" 

Else 
MsgBox "value is either zero or negative number" 
End if 
else 
MsgBox "invalid input "
End if 
    
Dim a,b ,operation 
a=10 
b=20 
'operation = InputBox("enter a operation")
select case operation 

case "add"
'msgbox "addition of a,b is :" & a+b  
case "sub"
'msgbox "subraction of a,b is :" & a-b 
case "mul"
'msgbox "multification of a,b is :" & a*b 
case "div"
'msgbox "division of a,b is :" & a/b 
case "mod"
'msgbox "modulus of a,b is :" & a MOD b
case Else 
'MsgBox "invalid operation" 
End Select