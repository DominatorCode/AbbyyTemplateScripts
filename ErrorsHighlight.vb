Public Sub ErrorsHighlight(this, vField)

dim totQty, i, vFieldVal


totQty = this.Symbols.Count

if totQty = 0 then
	this.CheckSucceeded = false
	this.ErrorMessage = "Пустое значение"
	Exit Sub
else		
	for i = totQty - 1 to 0 Step -1
		if  instr( 1, "0123456789", this.Symbols(i).Symbol ) <> 0 then 
			if this.Symbols(i).IsSuspicious then	
				this.CheckSucceeded = false
				this.ErrorMessage = "Есть неуверенно распознанные символы"
				Exit for
			end if
		End if
	next 
end if

End Sub




Public Sub ErrorsHighlightText(this, vField)
Dim totQty, i

totQty = this.Symbols.Count

if totQty > 0 then

 Set re = New RegExp
 re.Pattern = "[A-ZА-Я0-9]"
 re.IgnoreCase = True
 re.Global = True
    for i = totQty - 1 to 0 Step -1
     if re.Test(this.Symbols(i).Symbol) and this.Symbols(i).IsSuspicious then 
       this.CheckSucceeded = false
       Exit for 
     End if 
    next

End if

End Sub
