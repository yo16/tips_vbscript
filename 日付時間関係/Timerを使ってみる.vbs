Option Explicit

Dim startTime,endTime

startTime = Timer

Dim idx
For idx = 1 to 500000
Next

endTime = Timer



msgBox endTime - startTime
