Sub fillWorkout()

Dim colCount As Integer
Dim Match As Integer

colCount = 1
Z = randomExercise(colCount)
Column = colCount
Worksheets("Workout").Cells(5 + colCount + 1, 2).Value = Worksheets("Uebungen").Cells(Z, Column).Value
Worksheets("Workout").Cells(5 + colCount + 1, 3).Value = Worksheets("Uebungen").Cells(Z, Column + 1).Value
Worksheets("Workout").Cells(5 + colCount + 1, 4).Value = Worksheets("Uebungen").Cells(Z, Column + 2).Value
Worksheets("Workout").Cells(5 + colCount + 1, 5).Value = Worksheets("Uebungen").Cells(Z, Column + 3).Value

y = Z
Do While y = Z
    y = randomExercise(colCount)
    Column = colCount
    Worksheets("Workout").Cells(5 + colCount, 2).Value = Worksheets("Uebungen").Cells(y, Column).Value
    Worksheets("Workout").Cells(5 + colCount, 3).Value = Worksheets("Uebungen").Cells(y, Column + 1).Value
    Worksheets("Workout").Cells(5 + colCount, 4).Value = Worksheets("Uebungen").Cells(y, Column + 2).Value
    Worksheets("Workout").Cells(5 + colCount, 5).Value = Worksheets("Uebungen").Cells(y, Column + 3).Value
Loop

colCount = 2
For i = 2 To 8
    Z = randomExercise(colCount)
    Column = colCount + 3 * (colCount - 1)
    Worksheets("Workout").Cells(5 + colCount + 1, 2).Value = Worksheets("Uebungen").Cells(Z, Column).Value
    Worksheets("Workout").Cells(5 + colCount + 1, 3).Value = Worksheets("Uebungen").Cells(Z, Column + 1).Value
    Worksheets("Workout").Cells(5 + colCount + 1, 4).Value = Worksheets("Uebungen").Cells(Z, Column + 2).Value
    Worksheets("Workout").Cells(5 + colCount + 1, 5).Value = Worksheets("Uebungen").Cells(Z, Column + 3).Value
    colCount = colCount + 1
Next i

Match = 1
Do While Match > 0
    Match = 0
    Zufall
    For i = 6 To 14
        If Worksheets("Workout").Cells(i, 2).Value = Worksheets("Workout").Cells(15, 2).Value Then
            Match = Match + 1
        End If
    Next i
Loop

End Sub

Sub Zufall()

Dim colCount As Integer
randomGroup = Int((8 * Rnd) + 1)
colCount = randomGroup

Z = randomExercise(colCount)
Column = colCount + 3 * (colCount - 1)
Worksheets("Workout").Cells(15, 2).Value = Worksheets("Uebungen").Cells(Z, Column).Value
Worksheets("Workout").Cells(15, 3).Value = Worksheets("Uebungen").Cells(Z, Column + 1).Value
Worksheets("Workout").Cells(15, 4).Value = Worksheets("Uebungen").Cells(Z, Column + 2).Value
Worksheets("Workout").Cells(15, 5).Value = Worksheets("Uebungen").Cells(Z, Column + 3).Value
colCount = colCount + 1

End Sub



Function randomExercise(colCount As Integer) As Integer

countExercise = 1

Column = colCount + 3 * (colCount - 1)

Do While Not IsEmpty(Worksheets("Uebungen").Cells(3 + countExercise, Column))
    countExercise = countExercise + 1
Loop

randomExercise = Int((countExercise * Rnd) + 3)


End Function