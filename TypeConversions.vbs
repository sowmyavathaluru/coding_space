a = 3 
MsgBox Typename(a) 'integer
a = 3.2 
msgBox TypeName(a)  'Double 
x = "6.1"
MsgBox TypeName(x)  'string
y = "5"
MsgBox x+y    '6.15 
MsgBox CDbl(x)+CInt(y)  '11.1  
MsgBox CInt(x)+CInt(y)  '11
z = 1
c = "24/04/2021"
MsgBox CDate(c) 
MsgBox IsDate(c)  ' True 