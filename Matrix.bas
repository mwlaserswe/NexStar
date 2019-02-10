Attribute VB_Name = "Matrix"

Public Sub MatrixAddition(matrix1() As Double, ByVal row1 As Integer, ByVal col1 As Integer, matrix2() As Double, ByVal row2 As Integer, ByVal col2 As Integer, matrix3() As Double)
    Dim i As Integer
    Dim j As Integer
    
    
    For i = 0 To row1 - 1
      For j = 0 To col2 - 1
        matrix3(i, j) = matrix1(i, j) + matrix2(i, j)
      Next
    Next
End Sub

'Public Sub PolarKarthesisch(AZ As Double, ELEV As Double, V As Vector)
'  'https://de.wikipedia.org/wiki/Kugelkoordinaten
'  '   Absatz "Andere Konventionen"
'
'  V.x = Cos(ELEV) * Cos(AZ)
'  V.y = Cos(ELEV) * Sin(AZ)
'  V.z = Sin(ELEV)
'End Sub

Public Function PolarKarthesisch(HourAngle As Double, Declination As Double) As Vector
  'https://de.wikipedia.org/wiki/Kugelkoordinaten
  '   Absatz "Andere Konventionen"

  PolarKarthesisch.x = Cos(-HourAngle) * Cos(Declination)
  PolarKarthesisch.Y = Sin(-HourAngle) * Cos(Declination)
  PolarKarthesisch.z = Sin(Declination)
End Function


Public Function CrossProduct(V1 As Vector, V2 As Vector) As Vector
  'http://james-ramsden.com/calculate-the-cross-product-c-code/
  
  CrossProduct.x = V1.Y * V2.z - V2.Y * V1.z
  CrossProduct.Y = (V1.x * V2.z - V2.x * V1.z) * -1
  CrossProduct.z = V1.x * V2.Y - V2.x * V1.Y
End Function


Public Sub MatrixProduct(matrix1() As Double, ByVal row1 As Integer, ByVal col1 As Integer, matrix2() As Double, ByVal row2 As Integer, ByVal col2 As Integer, matrix3() As Double)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    
    For i = 0 To row1 - 1
        For j = 0 To col2 - 1
            For k = 0 To row2 - 1
                Temp = Temp + (matrix1(i, k) * matrix2(k, j))
            Next
            matrix3(i, j) = Temp
            Temp = 0
        Next
    Next
End Sub



