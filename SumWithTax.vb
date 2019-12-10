const Sep = ","
const DeltaFloat = 0.1

dim vField, totQty, i, vFieldCalc, condCanCompare, vFieldCost, vFieldTaxSum, vFieldCalcIsStrong

vField = "SumWithTax"
condCanCompare = false
vFieldCalcIsStrong = false

if Len(me.Field("Cost").Text) > 0 and Len(me.Field("TaxSum").Text) > 0 then

	if me.Field("Cost").Value > 0  then
		vFieldCost = me.Field("Cost").Value
	else
		vFieldCost = me.Field("Cost").Text
		vFieldCost = Replace(vFieldCost, "-", ",")
		vFieldCost = Replace(vFieldCost, ".", ",")
		vFieldCost = Replace(vFieldCost, "=", ",")
		If IsNumeric(vFieldCost) then
			vFieldCost = CDbl(vFieldCost)
		else
			vFieldCost = CDbl("0,00")
		end if
	end if

	if me.Field("TaxSum").Value > 0 then
		vFieldTaxSum = me.Field("TaxSum").Value
	else 
		vFieldTaxSum = me.Field("TaxSum").Text
		vFieldTaxSum = Replace(vFieldTaxSum, "-", ",")
		vFieldTaxSum = Replace(vFieldTaxSum, ".", ",")
		vFieldTaxSum = Replace(vFieldTaxSum, "=", ",")
		If IsNumeric(me.Field("TaxSum").Text) then
			vFieldTaxSum = CDbl(vFieldTaxSum) 
		else
			vFieldTaxSum = CDbl("0,00")		
		end if
	end if

    vFieldCalc = vFieldCost + vFieldTaxSum
	
	if vFieldCalc > 0 then
		if not (me.Field("TaxSum").IsSuspicious = true or me.Field("Cost").IsSuspicious = true) then
			vFieldCalcIsStrong = true
		end if
		
		condCanCompare = true
	end if

end if


if condCanCompare = true then	
	vFieldVal = me.FIELD(vField).text
	if Len(vFieldVal) = 0 then 		
		me.FIELD(vField).Value = vFieldCalc
		me.ErrorMessage = "Значение последнего столбца пустое, было вычислено из соседних столбцов -> " & vFieldCalc
		me.NeedVerification = true
		' добавление дробной части, если значение целое
		vFieldVal = me.FIELD(vField).text
		if Len(vFieldVal) > 3 then
			if (instr( 1, ".,-=",Mid(vFieldVal, Len(vFieldVal) - 2, 1)) = 0) then  
			   me.FIELD(vField).text = vFieldVal & Sep & "00"
			end if 
		else
			me.FIELD(vField).text = vFieldVal & Sep & "00"
		end if
		
		Exit Sub
	end if
	
	if abs(vFieldVal - vFieldCalc) > DeltaFloat then
		me.CheckSucceeded = false
		me.ErrorMessage = "Значение последнего столбца не совпадает с суммой из соседних столбцов -> " & vFieldCalc				
	end if  
end if


'========================================================================================


if me.NeedVerification then
    me.NeedVerification = false
    me.CheckSucceeded = false
end if