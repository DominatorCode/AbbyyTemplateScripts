Dim totQty, maxInterval

totQty = me.Symbols.Count

if (totQty < 3) then
    
    for i = totQty - 1 to 0 Step -1
        if  instr( 1, "16802", me.Symbols(i).Symbol ) = 0 then 
            me.Symbols.Delete(i)
        End if
    next
end if

if (totQty < 3) then    
    if me.Text = "6" or me.Text = "16" or me.Text = "8" then
        me.Text = "18"
    end if
    
    if me.Text = "2" then
        me.Text = "12"
    end if
    
    if me.Symbols.Count = 1 then
        me.Text = ""
    end if
    if totQty > 1 then 
		me.Text = me.Text & "%"
	end if       
    exit sub
end if
totQty = me.Symbols.Count
if totQty < 5 then    
	maxInterval = 100

	for i = totQty - 1 to 0 Step -1
		if  instr( 1, "1802", me.Symbols(i).Symbol ) <> 0 then 
			maxInterval = abs(me.Symbols(i).Rect.Top - me.Symbols(i).Rect.Bottom) * 2
			exit for
		End if
	next 


	if (abs(me.Symbols(totQty - 1).Rect.Left - me.Symbols(totQty - 2).Rect.Left) > maxInterval) then
		me.Symbols.Delete(totQty - 1)
	end if

	if (abs(me.Symbols(1).Rect.Right - me.Symbols(0).Rect.Right) > maxInterval) then
		me.Symbols.Delete(0)
	end if
end if

totQty = me.Symbols.Count

if totQty < 5 then

    if me.Text = "0" or me.Text = "0%" then
        Exit sub
    end if

    for i = totQty - 1 to 0 Step -1
        if  instr( 1, "18602", me.Symbols(i).Symbol ) = 0 then 
            me.Symbols.Delete(i)
        End if
    next
end if
totQty = me.Symbols.Count
if totQty < 5 then    
    if me.Text = "6" or me.Text = "16" or me.Text = "8" then
        me.Text = "18"
    end if
    
    if me.Text = "2" then
        me.Text = "12"
    end if
    
    if me.Symbols.Count = 1 then
        me.Text = ""
    end if
	
	if totQty > 1 then 
		me.Text = me.Text & "%"
	end if
	Exit Sub
end if

if totQty < 8 then
	for i = totQty - 1 to totQty - 4 Step -1
		if  instr( 1, "1802безндсБНДСЕЗ ", me.Symbols(i).Symbol ) = 0 then 
			me.Symbols.Delete(i)
		End if
	next
end if
totQty = me.Symbols.Count
if totQty < 8 then

    for i = totQty - 5 to 1 Step -1
        if  instr( 1, "1802безндс/БНДСЕЗ ", me.Symbols(i).Symbol ) = 0 then 
            me.Symbols.Delete(i)
        End if
    next

	if  instr( 1, "120Бб", me.Symbols(0).Symbol ) = 0 then 
			me.Symbols.Delete(0)
	End if 

	if me.Text = "6" or me.Text = "16" or me.Text = "8" then
			me.Text = "18"
	end if

	if me.Text = "2" then
			me.Text = "12"
	end if

	if me.Symbols.Count = 1 then
		me.Text = ""
	end if

	if me.Text = "ез НДС" then
		me.Text = "Без НДС"
	end if
    
end if    
totQty = me.Symbols.Count
if totQty < 7 and totQty > 1 then
	me.Text = me.Text & "%"
end if

'=================================================================

if Len(me.Text) > 2 then

    if InStr(me.Text,"10") <> 0 then
        me.Text = "10%"
        Exit Sub
    end if
    
    if InStr(me.Text,"20") <> 0 then
        me.Text = "20%"
        Exit Sub
    end if

end if

