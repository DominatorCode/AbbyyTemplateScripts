Dim vDate, vLen, vFieldName, DateTmp
vFieldName = "DocDate"

if Len(me.Field("DocDate2").Text) > 0 then
    if IsDate(me.Field("DocDate2").Text) then
        DateTmp = FormatDateTime(me.Field("DocDate2").Text, 2)
		if Len(DateTmp) > 0 then
            DateTmp = me.Field("DocDate2").Text
        end if
    else
        DateTmp = ""
    end if
end if


if Len(DateTmp) > 0 then
   vDate = DateTmp
    if Mid(vDate,5,1) = "-"  and isNumeric(Mid (vDate,4,1)) then
        Exit Sub
    end if

    vDate = Replace (vDate, "/", "")
    vDate = Replace (vDate, ".", "")
    vDate = Replace (vDate, "-", "")

    vDate = CStr (vDate)
    vLen = Len (vDate)

    if IsNumeric (vDate) then
       if vLen = 6 then
            ' vDate = Mid (vDate,1,2) & "." & Mid (vDate,3,2) & ".20" & Mid (vDate,5,2)
            vDate = ".20" & Mid (vDate,5,2) & "-" & Mid (vDate,3,2) & "-" & Mid (vDate,1,2)
       elseif (vLen = 8) then
            ' vDate = Mid (vDate,1,2) & "." & Mid (vDate,3,2) & "." & Mid (vDate,5,4)
            vDate = Mid (vDate,5,4) & "-" & Mid (vDate,3,2) & "-" & Mid (vDate,1,2)
       else
          me.Field(vFieldName).Text = " "   
          Exit Sub
       end if
    else
        me.Field(vFieldName).Text = " "
        Exit Sub
    end if
        
    me.Field(vFieldName).Text = vDate    
    
else if Len(me.Field("DocDate1").Text) > 0 then
    vField = me.Field("DocDate1").Text 

    Dim vY
      vY = Year(vField)
     ' содержимое опозналось как дата, но на самом деле она некорректная
      if CStr(vY) = "1601" Then  
		 me.Field("DocDate1").Text = ""			
         Exit Sub
      end if
    
    Dim vDay, vMonth, vYear
    
    vDay = Day(vField)
    if Len(vDay) = 1 then
        vDay = "0" & vDay
    end if
    
    vMonth = Month(vField)
    if Len(vMonth) = 1 then
        vMonth = "0" & vMonth
    end if
    
    vYear = Year(vField)
    
    if vYear > Year(Now) then
        vYear = Year(Now)
    end if
	
	if vYear < Year(Now) - 50 then
        vYear = Year(Now)
    end if
    
    ' me.Field(vFieldName).Text = CStr(vDay)& "." & CStr(vMonth)& "." & CStr(vYear)
    me.Field(vFieldName).Text = CStr(vYear) & "-" & CStr(vMonth)& "-" & CStr(vDay)
    end if
end if