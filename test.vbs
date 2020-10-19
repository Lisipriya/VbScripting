
Dim temp
 Function Celsius(fDegrees)
     Celsius = (fDegrees - 32) * 5 / 9
 End Function
temp = InputBox("Please enter the temperature in degrees F.", 1)
MsgBox "The temperature is " & round(Celsius(temp)) & " degrees C."
