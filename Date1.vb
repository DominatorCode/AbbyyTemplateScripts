Dim vDate, vFieldName

vFieldName = "DocDate1"

Dim vField

if Len(me.Field(vFieldName).Text) = 0 then
    Exit Sub
end if


vField = me.Field(vFieldName).Text
SetLocale("cs")

if not IsDate(vField) then
    me.Field(vFieldName).Text = ""
end if