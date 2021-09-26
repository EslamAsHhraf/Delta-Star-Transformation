Option Explicit
Dim x,r1,r2,r3,r12,r31,r23
x=InputBox("Please"&Vblf&"Enter 1 if you want (star to delta)"&Vblf& "Enter 2 if you want (delta to star)"&Vblf&"Then click on ok","Choose The Way","1 or 2")
If x=1 then
r1=InputBox("Enter R1"&Vblf&"Then click on ok","Star To Delta","The number should be bigger than zaro")
If r1<=0 Or Not IsNumeric(r1) Then
r1=InputBox("Error !!"&Vblf&"You input"&Vblf& "R1= "&r1&Vblf&"You didn't enter number bigger than zero "&Vblf&"Enter R1 again"&Vblf&"Then click on ok","Star To Delta","The number should be bigger than zaro")
End If
r2=InputBox("Enter R2"&Vblf&"Then click on ok","Star To Delta","The number should be bigger than zaro")
If r2<=0 Or Not IsNumeric(r2) Then
r2=InputBox("Error !!"&Vblf&"You input"&Vblf& "R2= "&r2&Vblf&"You didn't enter number bigger than zero "&Vblf&"Enter R2 again"&Vblf&"Then click on ok","Star To Delta","The number should be bigger than zaro")
End If
r3=InputBox("Enter R3"&Vblf&"Then click on ok","Star To Delta","The number should be bigger than zaro")
If r3<=0 Or Not IsNumeric(r3) Then
r3=InputBox("Error !!"&Vblf&"You input"&Vblf& "R3= "&r3&Vblf&"You didn't enter number bigger than zero "&Vblf&"Enter R3 again"&Vblf&"Then click on ok","Star To Delta","The number should be bigger than zaro")
End If
If r1<=0  Or Not IsNumeric(r1) Then
MsgBox"Error"&Vblf&"You input R1= "&r1&Vblf&"R1 isn't bigger than 0"&Vblf&"Close the program and try again",16,"Error"
ElseIf r2<=0 Or Not IsNumeric(r2) Then
MsgBox"Error"&Vblf&"You input R2= "&r2&Vblf&"R2 isn't bigger than 0"&Vblf&"Close the program and try again",16,"Error"
ElseIf r3<=0 Or Not IsNumeric(r3) Then
MsgBox"Error"&Vblf&"You input R3= "&r3&Vblf&"R3 isn't bigger than 0"&Vblf&"Close the program and try again",16,"Error"
Else
r12 =  r1+(r1 * r2) / r3+ r2 
r31 =r3+ (r1 * r3) / r2 +r1 
r23 =  r2+ (r3 * r2) / r1 +r3
MsgBox "Results: "&Vblf&"Your inputs:"&Vblf&"R1= "&r1 &Vblf& "R2= "&r2 &Vblf& "R3= "&r3 &Vblf& "The output:" &Vblf& "R12= "&r12 &Vblf& "R23= "&r23 &Vblf& "R31= "&r31,0,"Star To Delta"
End If
ElseIf x=2 then
r12=InputBox("Enter R12"&Vblf&"Then click on ok","Delta To Star","The number should be bigger than zaro")
If r12<=0 Or Not IsNumeric(r12) Then
r12=InputBox("Error !!"&Vblf&"You input"&Vblf&"R12= "&r12&Vblf&"You didn't enter number bigger than zero "&Vblf&"Enter R12 again"&Vblf&"Then click on ok","Delta To Star","The number should be bigger than zaro")
End If
r23=InputBox("Enter R23"&Vblf&"Then click on ok","Delta To Star","The number should be bigger than zaro")
If r23<=0 Or Not IsNumeric(r23) Then
r23=InputBox("Error !!"&Vblf&"You input "&Vblf&"R23= "&r23&Vblf&"You didn't enter number bigger than zero "&Vblf&"Enter R23 again"&Vblf&"Then click on ok","Delta To Star","The number should be bigger than zaro")
End If
r31=InputBox("Enter R31"&Vblf&"Then click on ok","Delta To Star","The number should be bigger than zaro")
If r31<=0 Or Not IsNumeric(r31) Then
r31=InputBox("Error !!"&Vblf&"You input"&Vblf&"R31= "&r31&Vblf&"You didn't enter number bigger than zero "&Vblf&"Enter R31 again"&Vblf&"Then click on ok","Delta To Star","The number should be bigger than zaro")
End If
If r12<=0  Or Not IsNumeric(r12) Then
MsgBox"Error"&Vblf&"You input R12= "&r12&Vblf&"R12 isn't bigger than 0"&Vblf&"Close the program and try again",16,"Error"
ElseIf r23<=0 Or Not IsNumeric(r23) Then
MsgBox"Error"&Vblf&"You input R23= "&r23&Vblf&"R23 isn't bigger than 0"&Vblf&"Close the program and try again",16,"Error"
ElseIf r31<=0 Or Not IsNumeric(r31) Then
MsgBox"Error"&Vblf&"You input R31= "&r31&Vblf&"R31 isn't bigger than 0"&Vblf&"Close the program and try again",16,"Error"
Else
r1 = (r12*r31)/(r12+0+r23+0+r31)
r2 = (r23*r12)/(r12+0+r23+0+r31)
r3 =  (r23*r31)/(r12+0+r23+0+r31)
MsgBox"Results:"&Vblf& "Your inputs:"&Vblf&"R12= "&r12 &Vblf& "R23= "&r23 &Vblf& "R31= "&r31 &Vblf& "The output:" &Vblf& "R1= "&r1 &Vblf& "R2= "&r2 &Vblf& "R3= "&r3,0,"Delta To Star"
End If
Else
MsgBox"Error"&Vblf&"Please enter 1 or 2 only"&Vblf&"Close the program and try again",16,"Error"
End If