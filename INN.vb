Dim totQty

totQty = me.Symbols.Count

if totQty = 0  then
	Exit Sub
end if

for i = totQty - 1 to 0 Step -1
        if  instr( 1, "0123456789/\|", me.Symbols(i).Symbol ) = 0 then 
            me.Symbols.Delete(i)  
        End if
      next 
	  
	  
'======================================================================================


Dim vFieldName, vFieldVal, i
Dim IssKPP, IssINN

vFieldName = "IssINNKPP"
vFieldVal = me.Field(vFieldName).Text

if Len(vFieldVal) < 10 then
    me.Field(vFieldName).Text = ""
    Exit Sub
end if

i = 0

for i = Len(vFieldVal) - 1 to 2 Step -1
        if  instr( 1, "/\|", Mid( vFieldVal, i , 1 ) ) <> 0 then 
            Exit for    
        End if
      next
    
if i > 1 then
    IssKPP = Right (vFieldVal, Len(vFieldVal) - i)
    IssINN = Left (vFieldVal, i - 1)
else	
    IssINN = vFieldVal
end if

if Len(IssINN) > 12 then
    IssINN = Left(IssINN, 12)
end if

if Len(IssINN) < 10 then
    IssINN = ""
end if

if Len(IssKPP) > 9 then
    IssKPP = Left(IssKPP, 9)
end if

if Len(IssKPP) < 9 then
    IssKPP = ""
end if


me.Field("IssINN").Text = IssINN
me.Field("IssKPP").Text = IssKPP




