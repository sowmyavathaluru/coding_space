Set myArrayList = CreateObject( "System.Collections.ArrayList" ) 
set x =  CreateObject( "System.Collections.ArrayList" ) 
x.Add "blue"          'add is a method
x.Add "red" 
x.Add "black" 
MsgBox x.Count         'count is a property 
for each value in x 
'MsgBox value 
Next  
'x.Clear 
x(1) = "green"
'MsgBox x(1)   

class myclass 
Function add()
'MsgBox 5 + 2 
end Function
Function newsub() 
'MsgBox 10-2
End Function
end Class
set y = new myclass 
y.Add 
y.newsub  

array1 = Array("red","blue","green","pink","black",true,1,20,10.5) 
set list1 = CreateObject("system.collections.ArrayList") 

for each element in array1 
list1.Add element
next 
'MsgBox TypeName(list1) 
'MsgBox list1.count 
 
array2 = Array(1,3,56,45,67,8,95,34,23,12,4)
set list_even = CreateObject("system.collections.ArrayList") 
set list_odd = CreateObject("system.collections.ArrayList") 
for each num in array2 
if num mod 2 = 0 Then 
list_even.Add num 
else 
list_odd.add num 
end if 
next 
MsgBox list_even
MsgBox list_odd
  
