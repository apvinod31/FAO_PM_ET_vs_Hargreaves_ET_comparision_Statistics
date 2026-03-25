Attribute VB_Name = "Module2"
Sub ET_Comparison_FAO_Hargreaves()

Dim lastRow As Long, i As Long
Dim obs As Double, sim As Double
Dim meanObs As Double
Dim sumObs As Double, sumErr As Double
Dim sumSqErr As Double, sumSqObsDev As Double

Dim sumX As Double, sumY As Double
Dim sumXY As Double, sumX2 As Double, sumY2 As Double
Dim n As Long

Dim NSE As Double, PBIAS As Double, R2 As Double

' Find last row
lastRow = Cells(Rows.Count, "A").End(xlUp).Row
n = lastRow - 1

' ===== Mean Observed (FAO PM) =====
For i = 2 To lastRow
    sumObs = sumObs + Cells(i, "C").Value
Next i

meanObs = sumObs / n

' ===== Loop for statistics =====
For i = 2 To lastRow

    sim = Cells(i, "B").Value
    obs = Cells(i, "C").Value

    sumSqErr = sumSqErr + (obs - sim) ^ 2
    sumSqObsDev = sumSqObsDev + (obs - meanObs) ^ 2

    sumErr = sumErr + (sim - obs)

    sumX = sumX + obs
    sumY = sumY + sim
    sumXY = sumXY + obs * sim
    sumX2 = sumX2 + obs ^ 2
    sumY2 = sumY2 + sim ^ 2

Next i

' ===== NSE =====
NSE = 1 - (sumSqErr / sumSqObsDev)

' ===== PBIAS =====
PBIAS = (sumErr / sumObs) * 100

' ===== R2 =====
R2 = ((n * sumXY - sumX * sumY) ^ 2) / _
     ((n * sumX2 - sumX ^ 2) * (n * sumY2 - sumY ^ 2))

' ===== Output Results =====
Range("E2") = "NSE"
Range("F2") = NSE

Range("E3") = "PBIAS (%)"
Range("F3") = PBIAS

Range("E4") = "R²"
Range("F4") = R2

' ===== Scatter Plot (Observed vs Simulated) =====
Dim cht1 As ChartObject
Set cht1 = ActiveSheet.ChartObjects.Add(Left:=350, Width:=420, Top:=40, Height:=300)

cht1.Chart.ChartType = xlXYScatter
cht1.Chart.SetSourceData Source:=Range("C2:C" & lastRow & ",B2:B" & lastRow)

cht1.Chart.HasTitle = True
cht1.Chart.ChartTitle.Text = "Observed (FAO PM) vs Hargreaves ET"

cht1.Chart.Axes(xlCategory).HasTitle = True
cht1.Chart.Axes(xlCategory).AxisTitle.Text = "Observed ET (mm/day)"

cht1.Chart.Axes(xlValue).HasTitle = True
cht1.Chart.Axes(xlValue).AxisTitle.Text = "Hargreaves ET (mm/day)"

' ===== Time Series Plot =====
Dim cht2 As ChartObject
Set cht2 = ActiveSheet.ChartObjects.Add(Left:=350, Width:=420, Top:=360, Height:=300)

cht2.Chart.ChartType = xlLine

With cht2.Chart
    .SeriesCollection.NewSeries
    .SeriesCollection(1).Name = "Observed ET (FAO PM)"
    .SeriesCollection(1).XValues = Range("A2:A" & lastRow)
    .SeriesCollection(1).Values = Range("C2:C" & lastRow)

    .SeriesCollection.NewSeries
    .SeriesCollection(2).Name = "Hargreaves ET"
    .SeriesCollection(2).XValues = Range("A2:A" & lastRow)
    .SeriesCollection(2).Values = Range("B2:B" & lastRow)

    .HasTitle = True
    .ChartTitle.Text = "Daily ET Comparison Time Series"
End With

MsgBox "Statistical comparison and graphs generated successfully!"

End Sub

