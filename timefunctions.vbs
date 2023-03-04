a = now() 
MsgBox(a) 

Dim var1 
var1 = "4:30:55 "

var2 = Hour(var1) 
MsgBox var2  

var3 = Minute(var1) 
MsgBox var3 

var4 = Second(var1)
MsgBox var4 

var5 = Time() 
MsgBox var5 

var6 = Timer() 
MsgBox var6  

var7 = TimeSerial(9,45,56) 'format (hour,minute,second)
MsgBox var7

var8 = TimeValue("20:30") 
MsgBox var8