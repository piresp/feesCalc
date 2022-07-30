Sub Go()

restValue = Cells(6, 3).Value
clientPayment = Cells(8, 3).Value

nParcel = 0
lastParcel = 0
Row = 3

While restValue > 0
    lastParcel = restValue
    restValue = restValue - clientPayment
    
    Cells(Row, 6) = lastParcel
    Cells(Row, 5) = nParcel
    Cells(10, 3) = lastParcel
    
    restValue = restValue * 1.045
    Row = Row + 1
    Cells(9, 3) = nParcel
    nParcel = nParcel + 1
Wend
End Sub
