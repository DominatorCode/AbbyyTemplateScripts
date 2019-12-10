' =============================== ======================================

dim vFieldCalc, condIsEqual, vFieldCost, vFieldTaxSum
'dim vField
'vField = "TaxSum"

condIsEqual = false
vFieldCalcIsStrong = false

if Len(me.Field("CostSum").Text) > 0 and Len(me.Field("Sum").Text) > 0 then

 if me.Field("CostSum").Value > 0  then
  vFieldCost = me.Field("CostSum").Value
 else If IsNumeric(me.Field("CostSum").Text) then
   vFieldCost = CDbl(me.Field("CostSum").Text)
  else
   vFieldCost = CDbl("0,00")
  end if
 end if

 if me.Field("Sum").Value > 0 then
  vFieldTaxSum = me.Field("Sum").Value
 else If IsNumeric(me.Field("Sum").Text) then
   vFieldTaxSum = CDbl(me.Field("Sum").Text) 
  else
   vFieldCost = CDbl("0,00")  
  end if
 end if

    vFieldCalc = vFieldCost + vFieldTaxSum
 
 if vFieldCalc > 0 then 
  if vFieldTaxSum - vFieldCost = 0 then    
    condIsEqual = true
  end if
 end if

end if


if me.Value = 0 and condIsEqual = true then
    me.CheckSucceeded = true
    Exit Sub
end if

if Len(me.Text) = 0 and condIsEqual = true then
    me.Text = "none"
end if



'============================================= =================================================

if me.Text = "none" or me.Field("TaxSum_Calc").Value = 0 then
    me.Value = 0
    me.Text = "0,00"
    me.CheckSucceeded = true
end if


