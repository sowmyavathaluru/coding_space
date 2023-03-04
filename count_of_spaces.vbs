string_1 = "pushpa is a good girl " 
MsgBox UBound(Split(string_1," "))
for each ele in Split(string_1," ") 
MsgBox ele.Count
Next 

