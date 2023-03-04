dim a 
a = date () 
'MsgBox a 

b = CDate("jan 10 2020") 
'MsgBox b 

date2 = #20/2/2021# 
MsgBox(DateAdd("yyyy",1,date2)) 'date add function 
'MsgBox(DateAdd("w",1,date2))
'MsgBox(DateAdd("ww",1,date2))
MsgBox(DateAdd("m",1,date2))
'MsgBox(DateAdd("q",1,date2))
'MsgBox(DateAdd("n",1,date2))
'MsgBox(DateAdd("m",1,date2))
'MsgBox(DateAdd("s",1,date2))
'MsgBox(DateAdd("y",2,date2))
msgBox(DateAdd("d",2,date2))
                                 
strdate = "25-01-2001 02:30:40"     'date difference 
enddate = "12-06-2002 06:20:50" 
'MsgBox (DateDiff("yyyy",strdate,enddate))                          
'MsgBox (DateDiff("y",strdate,enddate))
'MsgBox (DateDiff("d",strdate,enddate))
'MsgBox (DateDiff("m",strdate,enddate))
'MsgBox (DateDiff("q",strdate,enddate))
'MsgBox (DateDiff("w",strdate,enddate))
'MsgBox (DateDiff("ww",strdate,enddate))
'MsgBox (DateDiff("h",strdate,enddate))
'MsgBox (DateDiff("n",strdate,enddate))
'MsgBox (DateDiff("s",strdate,enddate))