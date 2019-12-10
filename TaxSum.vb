dim vField
const Sep = ","

vField = "TaxSum"

if Len(me.Field("TaxRate").Text) < 8 and Len(me.Field("TaxRate").Text) > 0 then
	if InStr(LCase(me.Field("TaxRate").Text), "ндс") or InStr(LCase(me.Field("TaxRate").Text), "без") or instr( 1, "0",Mid(me.Field("TaxRate").Text, 1, 1)) <> 0 then
		me.Field(vField).Text = "0" & Sep & "00"
		eXIT sUB
	end if
end if

Rules.NormalizeMoneyField Me, vField

'============================================================================================================================


dim vFieldName, vField, vFieldSumwTax

vFieldName = "TaxSum"

vField = me.Field(vFieldName).Value

if me.Field("TaxIncluded").Text = "true" then
    if Len(me.Field("SumWithTax").Text) > 0 and Len(me.Field("TaxRate").Text) > 0 then
        vField = me.Field("SumWithTax").Value * (1 - 1/((100 + me.Field("TaxRate").Value) * 0.01))
    else    
        vField = 0
    end if    
else
    if Len(me.Field("Price").Text) > 0 and Len(me.Field("TaxRate").Text) > 0 then    
        vField = me.Field("Price").Value * me.Field("TaxRate").Value / 100
    else
        vField = 0
    end if
end if

me.Field(vFieldName).Text = Round(vField, 2)