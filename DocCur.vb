dim vField

vField = "DocCur"

if me.FIELD(vField).text = "" or Len(me.FIELD(vField).text) <> 3 then
    me.Field (vField).text = "руб."
else 
    if me.FIELD(vField).VALUE = "643" then
        me.Field (vField).text = "руб."
    end if   
end if

if me.FIELD(vField).VALUE = "810" then
    me.Field (vField).text = "RUR"
else
    if me.FIELD(vField).VALUE = "978" then
        me.Field (vField).text = "EUR"
    end if
end if

if me.FIELD(vField).VALUE = "840" then
    me.Field (vField).text = "USD"
end if