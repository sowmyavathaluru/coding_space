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
'MsgBox list_even.count 
'MsgBox list_odd.Count   


list_even.sort() 
list_even.reverse() 
For each ele in list_even 
MsgBox ele 
next 

list_odd.sort() 
list_odd.reverse() 
For each ele in list_odd 
MsgBox ele 
next 
@@@@@@@@@@@@@@@@@@@@@
list_odd.sort() 
list_odd.reverse() 
For each ele in list_odd 
MsgBox ele 
next 