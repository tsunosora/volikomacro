Sub F280_320()
Dim s As Shape, s1 As Shape, SR As ShapeRange
ActiveDocument.Unit = cdrMeter
Set SR = ActiveSelectionRange
For Each s In SR
Set s1 = ActiveLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK F280_320 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
Sub F280_220()
Dim s As Shape, s1 As Shape, SR As ShapeRange
ActiveDocument.Unit = cdrMeter
Set SR = ActiveSelectionRange
For Each s In SR
Set s1 = ActiveLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK F280_220 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
Sub F280_160()
Dim s As Shape, s1 As Shape, SR As ShapeRange
ActiveDocument.Unit = cdrMeter
Set SR = ActiveSelectionRange
For Each s In SR
Set s1 = ActiveLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK F280_160 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
Sub F380_320()
Dim s As Shape, s1 As Shape, SR As ShapeRange
ActiveDocument.Unit = cdrMeter
Set SR = ActiveSelectionRange
For Each s In SR
Set s1 = ActiveLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK F380_320 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
Sub STIKER_CAMEL_150()
Dim s As Shape, s1 As Shape, SR As ShapeRange
ActiveDocument.Unit = cdrMeter
Set SR = ActiveSelectionRange
For Each s In SR
Set s1 = ActiveLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK CAMEL_150 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
Sub STIKER_CAMEL_126()
Dim s As Shape, s1 As Shape, SR As ShapeRange
ActiveDocument.Unit = cdrMeter
Set SR = ActiveSelectionRange
For Each s In SR
Set s1 = ActiveLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK CAMEL_126 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
Sub STIKER_INFLEX_150()
Dim s As Shape, s1 As Shape, SR As ShapeRange
ActiveDocument.Unit = cdrMeter
Set SR = ActiveSelectionRange
For Each s In SR
Set s1 = ActiveLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK INFLEX_150 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
