Dim vDate, vLen, vFieldName
vFieldName = "DocDateWord"

Dim vField

if Len(me.Field(vFieldName).Text) = 0 then
    Exit Sub
end if

totQty = me.Symbols.Count

Set re = New RegExp
re.Pattern = "[A-Z0-9А-Я ]"
re.IgnoreCase = True
re.Global = True

for i = totQty - 1 to 0 Step -1
        if not re.Test(me.Symbols(i).Symbol) then
            me.Symbols.Delete(i)  
        End if
Next


totQty = me.Symbols.Count

if totQty > 4 then

    ' удаляем мусор в конце   
    Set re = New RegExp
    re.Pattern = "[0-9]"
    re.IgnoreCase = True
    re.Global = True
    
    for i = totQty - 1 to 1 Step -1
        if re.Test(me.Symbols(i).Symbol) then 
            Exit for    
        End if
    next 
    if i > 2 then
        me.Text = Left(me.Text, i + 1)
    end if

end if

vField = me.Text
SetLocale("ru")

if IsDate(vField) then
    me.Field(vFieldName).Text = FormatDateTime(vField, 2)
end if