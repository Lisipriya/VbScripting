Dim x
x = InputBox("Enter your value")

Msgbox ("Do Until")

Do Until x=15
Msgbox (x)
x=x+1
Loop

Msgbox ("Do While")
x = InputBox("Enter your value")
Do While x<15
If x= 10 Then Exit Do
Msgbox (x)
x=x+1
Loop

Dim arr(3)

For i=0 to 3
arr(i) = Inputbox("Enter array item:")
Next

For Each x in arr
Msgbox(x)
Next