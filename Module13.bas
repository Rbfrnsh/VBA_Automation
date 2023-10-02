Attribute VB_Name = "Module13"
Sub concat()

  With Range("Y2:Y" & Range("X" & Rows.Count).End(3).Row)
    .Formula = "=X2 & MID(""thstndrdth"",MIN(9,2*RIGHT(X2)*(MOD(X2-11,100)>2)+1),2) & "" Tenant"""
    .Value = .Value
  End With

End Sub



