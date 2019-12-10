Dim totQty, deleteIndex, i, maxInterval

totQty = me.Symbols.Count

if totQty = 0  then
	Exit Sub
end if
    
deleteIndex = -1

maxInterval = 100

for i = totQty - 1 to 0 Step -1
    if  instr( 1, "0123456789", me.Symbols(i).Symbol ) <> 0 then 
        maxInterval = abs(me.Symbols(i).Rect.Top - me.Symbols(i).Rect.Bottom) * 2
        exit for
    End if
next 


if (totQty > 4) then
   
    for i=0 to totQty - 5
		if (me.Symbols(i+1).Rect.Right - me.Symbols(i).Rect.Right > maxInterval) then
			deleteIndex = i        
		end if
	next

	if (deleteIndex <> -1) then
		me.Text = Right(Me.Text, totQty - deleteIndex)
	end if

end if

if me.Symbols.Count > 1 then

    if abs(me.Symbols(1).Rect.Right - me.Symbols(0).Rect.Right) > maxInterval then
    			me.Symbols.Delete(0)    
    end if

end if

'=======================================================================


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

totQty = me.Symbols.Count

if totQty > 0 then	
	if totQty > 2 then
		if  instr( 1, "1", me.Symbols(me.Symbols.Count - 1).Symbol ) <> 0 then 
			if abs(me.Field(vField).Regions(0).SurroundingRect.Right - me.Symbols(totQty - 1).Rect.Right) < 2 then
				me.Symbols.Delete(totQty - 1)    
			end if
		end if
	end if
	totQty = me.Symbols.Count
	'чистка мусора с конца
	for i = totQty - 1 to 0 Step -1
			if instr( 1, "0123456789", me.Symbols(i).Symbol ) = 0 then 
				me.Symbols.Delete(i)			
			else
					exit for
			end if
		next
end if


totQty = me.Symbols.Count

if totQty > 3 then

	if (me.Symbols(totQty - 3).Symbol = " ") then
		me.Symbols(totQty - 3).Symbol = Sep
	end if
    
    for i = totQty - 2 to 1 Step -1
        if  instr( 1, ".,-=0123456789", me.Symbols(i).Symbol ) = 0 then 
            me.Symbols.Delete(i)  
        End if
    next 
    
    if  instr( 1, "0123456789", me.Symbols(0).Symbol ) = 0 then 
        me.Symbols.Delete(0)
        totQty = totQty - 1  
    End if  
else if totQty > 0 then
		for i = totQty - 1 to 0 Step -1
			if  instr( 1, "0123456789", me.Symbols(i).Symbol ) = 0 then 
				me.Symbols.Delete(i)  
			End if
		next 
    end if     
end if	

vFieldVal = Me.Field(vField).Text    

if Len(vFieldVal) > 3 then
	
    if me.Field(vField).IsSuspicious = true then    
        me.CheckSucceeded = false
        me.ErrorMessage = "Есть неуверенно распознанные символы"
        me.FocusedField = me.Field(vField)
    end if
    
    ' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    ' нормализация суммы (удаляем разделители тысяч, записываем наш десятичный разделитель - Sep)
    
    Dim intResult, fract
        
      ' ищем десятичный разделитель
      for i = Len(vFieldVal) - 1 to 2 Step -1
        if  instr( 1, ".,-=", Mid( vFieldVal, i , 1 ) ) <> 0 then 
            Exit for    
        End if
      next
    
      if i > 1 then

        fract = Right (vFieldVal, Len(vFieldVal) - i)
        
        if (Len(fract) > 2) then
            fract = Mid(fract, 1, 2)
        else if (Len(fract) = 1) then
            fract = fract & "0"
            end if
        end if
        
        intResult = Left (vFieldVal, i)
        
        ' очищаем целую часть суммы от разделителей тысяч   
        intResult = Replace (intResult, ".", "")
        intResult = Replace (intResult, ",", "")
        intResult = Replace (intResult, "-", "")
        intResult = Replace (intResult, "=", "")	
		
		if Len(intResult) > 1 and me.Symbols(0).Symbol = "0" then
			intResult = Replace(intResult,"0","9",1,1)
			me.Symbols(0).IsSuspicious = true
		end if                 
		
        me.FIELD(vField).text = intResult & Sep & fract
	else
		if me.Symbols(0).Symbol = "0" and Len(me.Text) = 3 then
			me.FIELD(vField).text = me.Symbols(0).Symbol & Sep & Right(me.Text, me.Symbols.Count - 1)
			me.Symbols(1).IsSuspicious = true
			me.CheckSucceeded = false
			me.ErrorMessage = "Возможна ошибка в дробной части"
			Exit Sub
		end if 
		
		if (instr( 1, ".,-=",Mid(vFieldVal, Len(vFieldVal) - 2, 1)) = 0) then  
			me.Field(vField).Text = Mid(vFieldVal, 1, Len(vFieldVal) - 2) & Sep & Mid(vFieldVal, Len(vFieldVal) - 1, 2)
			me.Symbols(me.Symbols.Count - 3).IsSuspicious = true       
			me.CheckSucceeded = false
			me.ErrorMessage = "Возможна ошибка в дробной части"
		end if 

    end if		
'else 
	'me.CheckSucceeded = false
    'me.ErrorMessage = "Низкое качество распознавания"
end if 


if condCanCompare = true then	
	vFieldVal = me.FIELD(vField).text
	if Len(vFieldVal) = 0 then 		
		me.FIELD(vField).Value = vFieldCalc
		me.Symbols(0).IsSuspicious = true
		me.ErrorMessage = "Значение последнего столбца пустое, было вычислено из соседних столбцов -> " & vFieldCalc
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
	else
		me.CheckSucceeded = true
	end if  
end if