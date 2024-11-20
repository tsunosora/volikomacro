Sub F280_320()
Dim s As shape, s1 As shape, sr As ShapeRange
ActiveDocument.Unit = cdrMeter
Set sr = ActiveSelectionRange
For Each s In sr
Set s1 = activeLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK F280_320 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
Sub F280_220()
Dim s As shape, s1 As shape, sr As ShapeRange
ActiveDocument.Unit = cdrMeter
Set sr = ActiveSelectionRange
For Each s In sr
Set s1 = activeLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK F280_220 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
Sub F280_160()
Dim s As shape, s1 As shape, sr As ShapeRange
ActiveDocument.Unit = cdrMeter
Set sr = ActiveSelectionRange
For Each s In sr
Set s1 = activeLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK F280_160 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
Sub F340_320()
Dim s As shape, s1 As shape, sr As ShapeRange
ActiveDocument.Unit = cdrMeter
Set sr = ActiveSelectionRange
For Each s In sr
Set s1 = activeLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK F340_320 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
Sub F380_320()
Dim s As shape, s1 As shape, sr As ShapeRange
ActiveDocument.Unit = cdrMeter
Set sr = ActiveSelectionRange
For Each s In sr
Set s1 = activeLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK F380_320 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
Sub F440_320()
Dim s As shape, s1 As shape, sr As ShapeRange
ActiveDocument.Unit = cdrMeter
Set sr = ActiveSelectionRange
For Each s In sr
Set s1 = activeLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK F440_320 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
Sub STIKER_CAMEL_150()
Dim s As shape, s1 As shape, sr As ShapeRange
ActiveDocument.Unit = cdrMeter
Set sr = ActiveSelectionRange
For Each s In sr
Set s1 = activeLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK CAMEL_150 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
Sub STIKER_CAMEL_126()
Dim s As shape, s1 As shape, sr As ShapeRange
ActiveDocument.Unit = cdrMeter
Set sr = ActiveSelectionRange
For Each s In sr
Set s1 = activeLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK CAMEL_126 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
Sub STIKER_INFLEX_150()
Dim s As shape, s1 As shape, sr As ShapeRange
ActiveDocument.Unit = cdrMeter
Set sr = ActiveSelectionRange
For Each s In sr
Set s1 = activeLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK INFLEX_150 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
Sub STIKER_MAXDEC_155()
Dim s As shape, s1 As shape, sr As ShapeRange
ActiveDocument.Unit = cdrMeter
Set sr = ActiveSelectionRange
For Each s In sr
Set s1 = activeLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK_ST_MAXDEC_155 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
Sub STIKER_ORAJET_155()
Dim s As shape, s1 As shape, sr As ShapeRange
ActiveDocument.Unit = cdrMeter
Set sr = ActiveSelectionRange
For Each s In sr
Set s1 = activeLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK_ST_ORAJET_155 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
Sub STIKER_ORAJET_105()
Dim s As shape, s1 As shape, sr As ShapeRange
ActiveDocument.Unit = cdrMeter
Set sr = ActiveSelectionRange
For Each s In sr
Set s1 = activeLayer.CreateArtisticText(s.CenterX, s.CenterX - 1, Round(s.SizeWidth, 2) & " x " & Round(s.SizeHeight, 2) & " (m) " & "| CETAK_ST_ORAJET_105 | CUSTOMER | FINISHING |PCS")
s1.TopY = s.TopY
Next s
End Sub
Sub Lubang_Keling_Sudut()
Dim sr As ShapeRange
Dim sTemp As shape
Const dblOffset As Double = -0.5
Const dblDiameter As Double = 0.59

    Set sr = ActiveSelectionRange
    If sr.Count > 0 Then
        'top left
        Set sTemp = activeLayer.CreateEllipse2(sr.LeftX - dblOffset, sr.TopY + dblOffset, dblDiameter / 2)
        'bottom left
        Set sTemp = activeLayer.CreateEllipse2(sr.LeftX - dblOffset, sr.BottomY - dblOffset, dblDiameter / 2)
        'top right
        Set sTemp = activeLayer.CreateEllipse2(sr.RightX + dblOffset, sr.TopY + dblOffset, dblDiameter / 2)
        'bottom right
        Set sTemp = activeLayer.CreateEllipse2(sr.RightX + dblOffset, sr.BottomY - dblOffset, dblDiameter / 2)
    Else
        MsgBox "Nothing is selected.", vbExclamation, "Circles Off Corners"
    End If
End Sub
