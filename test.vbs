'MsgBox ("1" & 2) 
'MsgBox("1" + "2")

str1 = "I" 
str2 = "i" 
'MsgBox (InStr(str1,str2)) 
'MsgBox(InStrRev(str1,str2))  

'MsgBox(StrComp(str2,str1))

str1 = "she is a good girl" 
'MsgBox(Mid(str1,10,6)) 

str1 = "c:\Users\SVATHALU\Desktop\vbs\test.vbs"
str2 = "c:\Users\SVATHALU\vbs\new.vbs" 

a =2 
b = 3 
'MsgBox a = b  
'MsgBox (TypeName(3))

str1 = "Drive c:\abcdef\SVATHALU\Desktop\vbscript.txt" 
str2 = "c:\Users\SVATHALU\vbs\new.vbs" 
'MsgBox InStrRev(str1,"\") + 1   '31
'str = mid(str2, InStrRev(str2,".")+1) 
'MsgBox str 
'MsgBox InStr(str1,".") - (InStrRev(str1,"\") + 1)  '35

'str = mid(str1, InStrRev(str1,"\")+1,InStr(str1,".") - (InStrRev(str1,"\") + 1)) 
'MsgBox str 
MsgBox (Instr(10,str1,"\"))-(Instr(str1,"\")+1)

MsgBox (Instr(str1,"\")+1)  '10
location_of_first_slash =  (Instr(str1,"\")+1)

'MsgBox (Instr(4,str1,"\"))
str4 = mid(str1,(Instr(str1,"\")+1),(Instr(10,str1,"\"))-(Instr(str1,"\")+1)) 
MsgBox str4 




