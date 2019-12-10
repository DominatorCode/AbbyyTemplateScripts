Dim R
R = Replace (me.FIELD("Country").text, vbCR, " ")
R = Replace (R, vbLF, " ")

me.FIELD("Country").text = R

if instr(lcase(R), "китай") or instr(lcase(R), "народн") or instr(lcase(R), "китай") or instr(lcase(R), "йская")then
    me.Text = "КНР"
end if