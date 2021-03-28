VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Calendar1_Click()
Range("C2").Value = Calendar1.Value
Calendar1.Visible = True

End Sub
Private Sub CommandButton1_Click()
 Dim c   As Range
 Dim s As Range
 Dim x As Integer
 Dim y As Integer
 Dim b As Integer
 
 On Error Resume Next
    Set c = Sheets("Sayfa3").Range("a:a").Find(CDate(Range("c2")), LookIn:=xlValues, LookAt:=xlWhole)
    c.Offset(0, 1).Value = Range("h7").Value
    c.Offset(0, 2).Value = Range("h12").Value
    c.Offset(0, 3).Value = Range("h17").Value
    c.Offset(0, 4).Value = Range("h22").Value
    c.Offset(0, 5).Value = Range("h27").Value
    c.Offset(0, 6).Value = Range("h32").Value
    c.Offset(0, 7).Value = Range("h37").Value
    c.Offset(0, 8).Value = Range("h42").Value
    c.Offset(0, 9).Value = Range("h47").Value
    c.Offset(0, 10).Value = Range("h52").Value
    c.Offset(0, 11).Value = Range("h57").Value
    c.Offset(0, 12).Value = Range("h62").Value
    c.Offset(0, 13).Value = Range("h67").Value
    c.Offset(0, 14).Value = Range("h72").Value
    c.Offset(0, 15).Value = Range("h77").Value
    Set s = Sheets("Sayfa4").Range("a:ak").Find(CDate(Range("c2")), LookIn:=xlValues, LookAt:=xlWhole)
    x = 1
    y = -1
    
    
    While x < 76
    While y < 6
    s.Offset(x, y).Value = Cells(x + 2, y + 3).Value
    s.Offset(x, y).Interior.ColorIndex = Cells(x + 2, y + 3).Interior.ColorIndex
    y = y + 1
    Wend
    y = -1
    x = x + 1
    Wend
    
Range("e4:e6").Value = 0
Range("e9:e11").Value = 0
Range("e14:e16").Value = 0
Range("e19:e21").Value = 0
Range("e24:e26").Value = 0
Range("e29:e31").Value = 0
Range("e34:e36").Value = 0
Range("e39:e41").Value = 0
Range("e44:e46").Value = 0
Range("e49:e51").Value = 0
Range("e54:e56").Value = 0
Range("e59:e61").Value = 0
Range("e64:e66").Value = 0
Range("e69:e71").Value = 0
Range("e74:e76").Value = 0
b = 0
If c.Offset(0, 21).Value <> c.Offset(1, 21).Value Then
While c.Offset(0, 21).Value = c.Offset(b, 21).Value

b = b - 1
Wend

c.Average(0, 18) = WorksheetFunction.Average(Range("c.offset(0,16):c.Offset(b,16)"))


End If





End Sub

