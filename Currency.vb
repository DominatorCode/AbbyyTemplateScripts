Dim totQty, deleteIndex, i, maxInterval

totQty = Me.Symbols.Count

if totQty = 0  then
	Exit Sub
end if
    
deleteIndex = -1

maxInterval = 100

for i = totQty - 1 to 0 Step -1
    if  instr( 1, "0123456789", Me.Symbols(i).Symbol ) <> 0 then 
        maxInterval = abs(Me.Symbols(i).Rect.Top - Me.Symbols(i).Rect.Bottom) * 2
        exit for
    End if
next 




if (totQty > 4) then
   
    for i=0 to totQty - 5
		if (Me.Symbols(i+1).Rect.Right - Me.Symbols(i).Rect.Right > maxInterval) then
			deleteIndex = i        
		end if
	next

	if (deleteIndex <> -1) then
		Me.Text = Right(Me.Text, totQty - deleteIndex)
	end if

end if

if Me.Symbols.Count > 1 then

    if abs(Me.Symbols(1).Rect.Right - Me.Symbols(0).Rect.Right) > maxInterval then
    			Me.Symbols.Delete(0)    
    end if

end if

'======================================================================

const Sep = ","

Public Sub NormalizeMoneyField(this, vField)

dim totQty, i, vFieldVal


totQty = this.Symbols.Count

if totQty > 0 then
	'чистка мусора с конца	
	for i = totQty - 1 to 0 Step -1
		if instr( 1, "0123456789", this.Symbols(i).Symbol ) = 0 then 
			this.Symbols.Delete(i)			
		else
			exit for
		end if
	next
end if

totQty = this.Symbols.Count

if totQty > 3  then
	
	if (this.Symbols(totQty - 3).Symbol = " ") then
		this.Symbols(totQty - 3).Symbol = Sep
	end if
    
    for i = totQty - 2 to 1 Step -1
        if  instr( 1, ".,-=0123456789", this.Symbols(i).Symbol ) = 0 then 
            this.Symbols.Delete(i)  
        End if
    next 
    
    if  instr( 1, "0123456789", this.Symbols(0).Symbol ) = 0 then 
        this.Symbols.Delete(0)
    End if
end if
	
totQty = this.Symbols.Count
	
vFieldVal = this.Field(vField).Text

if Len(vFieldVal) > 3 then

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
	
	if Len(intResult) > 1 and this.Symbols(0).Symbol = "0" then
		intResult = Replace(intResult,"0","9",1,1)
		this.CheckSucceeded = false
	end if	
	
	this.FIELD(vField).text = intResult & Sep & fract
		   
	Exit Sub
  else
	
	if this.Symbols(0).Symbol = "0" and Len(this.Text) = 3 then
		this.FIELD(vField).text = this.Symbols(0).Symbol & Sep & Right(this.Text, this.Symbols.Count - 1)
		this.CheckSucceeded = false
		this.ErrorMessage = "Возможна ошибка в дробной части"
		Exit Sub
	end if	
	
	if (instr( 1, ".,-=",Mid(vFieldVal, Len(vFieldVal) - 2, 1)) = 0) then  
		this.Field(vField).Text = Mid(vFieldVal, 1, Len(vFieldVal) - 2) & Sep & Mid(vFieldVal, Len(vFieldVal) - 1, 2)
        this.Symbols(this.Symbols.Count - 3).IsSuspicious = true       
		this.CheckSucceeded = false
		this.ErrorMessage = "Возможна ошибка в дробной части"
    end if 
	
  end if
'else
	'this.CheckSucceeded = false
    'this.ErrorMessage = "Отсутствует дробная часть"
end if


End Sub


'========================================================================================

dim vField

vField = "Price"

Rules.ClearTableCurrencyEnd Me, vField
Rules.NormalizeMoneyField Me, vField
Rules.ErrorsHighlight Me, vField



'=========================================================================================

Public Sub ClearTableCurrencyEnd(this, vField)

dim totQty

totQty = this.Symbols.Count
if totQty > 2 then
	if  instr( 1, "1", this.Symbols(this.Symbols.Count - 1).Symbol ) <> 0 then 
		if abs(this.Field(vField).Regions(0).SurroundingRect.Right - this.Symbols(totQty - 1).Rect.Right) < 2 then
			this.Symbols.Delete(totQty - 1)    
		end if
	end if
end if

End Sub




'=============================================================================================

if InStr(Me.Field("TaxRate").Text, "ндс") or InStr(Me.Field("TaxRate").Text, "без") or Me.Field("TaxRate").Text = "0" then
    Me.Field(vField).Text = "0"
    eXIT sUB
end if


