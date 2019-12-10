Dim vFieldVal, i

vFieldVal = me.Field("DocNum").Text

if Len(vFieldVal) = 0 then
    exit sub
end if


for i = Len(vFieldVal) - 1 to 2 Step -1
    if  instr( 1, " ", Mid( vFieldVal, i , 1 ) ) <> 0 then 
        me.Symbols(i - 1).Symbol = "_"   
    End if
next  