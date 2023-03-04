dim array1
array1 = Array("red","blue","green","pink")
For Each element in array1 
'MsgBox element 
next 

'for i = 0 to UBound(Array1) 
'MsgBox Array1(i) 
'Next

'MsgBox UBound(Array1)   'no of elements
'MsgBox lBound(array1)   ' least index 
array1(2) = "black" 
'MsgBox array1(2) 

Redim Preserve array1(5)       ' Redim preserve used for to extend an array length 
array1(4) = "brown" 
'MsgBox array1(0) 
array1(5) = "grey" 
'MsgBox array1(5)
'array1(6) = "orange" 
'MsgBox array1(6)  

a1 = Split("Red $ Blue $ Yellow","$")  ' splitting the string and gives result red , blue ,yellow 
for each a in a1 
'MsgBox a 
next  

a = array("Red","Blue","Yellow")
b = join(a,"$")                       ' join the every string with $ symbol 
'MsgBox b 
 
a = array("Red","Blue","Yellow")     '
b = Filter(a,"B")
c = Filter(a,"e")
d = Filter(a,"Y") 
For each x in b
'MsgBox x                   ' filter is checking tht given array and then particular alphabet is there gives the entire word like  'd fo red 
Next 
For each x in c
'MsgBox x 
Next 
For each x in d
'MsgBox x 
Next
 
a = array("Red","Blue","Yellow")
b = "12345"
c = 234
MsgBox  IsArray(a)     ' checking give input is an array or not and gives true or false 
MsgBox IsArray(b) 
MsgBox IsArray(c) 
