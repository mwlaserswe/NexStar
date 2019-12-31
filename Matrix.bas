Attribute VB_Name = "Matrix"
Option Explicit


'Public Function MatrixAddition(matrix1() As Double, matrix2() As Double, ByVal row As Integer, ByVal col As Integer) As Double
'    Dim i As Integer
'    Dim j As Integer
'
'
'    For i = 0 To row - 1
'      For j = 0 To col - 1
'        MatrixAddition(i, j) = matrix1(i, j) + matrix2(i, j)
'      Next
'    Next
'End Function

Public Function AddAzAlt(v1 As AzAlt, v2 As AzAlt) As AzAlt
    AddAzAlt.Az = v1.Az + v2.Az
    AddAzAlt.Alt = v1.Alt + v2.Alt
End Function

Public Function SubAzAlt(v1 As AzAlt, v2 As AzAlt) As AzAlt
    SubAzAlt.Az = v1.Az - v2.Az
    SubAzAlt.Alt = v1.Alt - v2.Alt
End Function


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


Public Function CrossProduct(v1 As Vector, v2 As Vector) As Vector
  'http://james-ramsden.com/calculate-the-cross-product-c-code/
  
  CrossProduct.x = v1.Y * v2.z - v2.Y * v1.z
  CrossProduct.Y = (v1.x * v2.z - v2.x * v1.z) * -1
  CrossProduct.z = v1.x * v2.Y - v2.x * v1.Y
End Function


Public Function ScalarProduct(scalar As Double, V As Vector) As Vector
    ScalarProduct.x = scalar * V.x
    ScalarProduct.Y = scalar * V.Y
    ScalarProduct.z = scalar * V.z
End Function


Public Function LenghtVector(V As Vector) As Double
    LenghtVector = Sqr(V.x * V.x + V.Y * V.Y + V.z * V.z)
End Function


Public Function AngleBetweenVectors(v1 As Vector, v2 As Vector) As Double
    ' per cross product:  http://ne.lo-net2.de/selbstlernmaterial/m/ag/skp/skp_ww_gw.pdf
    Dim len1 As Double
    Dim len2 As Double
    Dim len3 As Double
    len1 = LenghtVector(CrossProduct(v1, v2))
    len2 = LenghtVector(v1)
    len3 = LenghtVector(v2)

    AngleBetweenVectors = arcsin(len1 / (len2 * len3))

    ' per scalar product:  https://matheguru.com/lineare-algebra/winkel-zwischen-zwei-vektoren.html
    ' Dim len1 As Double
    ' Dim len2 As Double
    ' len1 = LenghtVector(v1)
    ' len2 = LenghtVector(v2)
    ' Dim scalar As Double
    ' scalar = v1.x * v2.x + v1.Y * v2.Y + v1.z * v2.z
    '
    ' AngleBetweenVectors = (Pi / 2) - arcsin(scalar / (len1 * len2))
    
End Function


Public Sub MatrixProduct(matrix1() As Double, ByVal row1 As Integer, ByVal col1 As Integer, matrix2() As Double, ByVal row2 As Integer, ByVal col2 As Integer, matrix3() As Double)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Dim Temp As Double
    
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


Public Sub Calculate_Inverse(System_DIM As Long, Matrix_A() As Double, Inverse_Matrix() As Double)

    Dim k As Long
    Dim N As Long
    Dim M As Long
    Dim line_1 As Long
    Dim temporary_1 As Double
    Dim elem1 As Double
    Dim multiplier_1 As Double
    Dim MAX_DIM As Long
    Dim message As String
    Dim response
    
    Dim Operations_Matrix(10, 20) As Double

    Dim Solution_Problem As Boolean

    'Uses Gauss elimination method in order to calculate the inverse matrix [A]-1
    'Method: Puts matrix [A] at the left and the singular matrix [I] at the right:
    '[ a11 a12 a13 | 1 0 0 ]
    '[ a21 a22 a23 | 0 1 0 ]
    '[ a31 a32 a33 | 0 0 1 ]
    'Then using line operations, we try to build the singular matrix [I] at the left.
    'After we have finished, the inverse matrix [A]-1 (bij) will be at the right:
    '[ 1 0 0 | b11 b12 b13 ]
    '[ 0 1 0 | b21 b22 b23 ]
    '[ 0 0 1 | b31 b32 b33 ]
    
    On Error GoTo errhandler 'In case the inverse cannot be found (Determinant = 0)
    
    Solution_Problem = False
    MAX_DIM = 10
    
    'Assign values from matrix [A] at the left
    For N = 0 To System_DIM - 1
        For M = 0 To System_DIM - 1
            Operations_Matrix(M, N) = Matrix_A(M, N)
        Next
    Next
    
    'Assign values from singular matrix [I] at the right
    For N = 0 To System_DIM - 1
        For M = 0 To System_DIM - 1
            If N = M Then
                Operations_Matrix(M, N + System_DIM) = 1
            Else
                Operations_Matrix(M, N + System_DIM) = 0
            End If
        Next
    Next
    
    'Build the Singular matrix [I] at the left
    For k = 0 To System_DIM - 1
       'Bring a non-zero element first by changes lines if necessary
       If Operations_Matrix(k, k) = 0 Then
          For N = k To System_DIM - 1
            If Operations_Matrix(N, k) <> 0 Then line_1 = N: Exit For 'Finds line_1 with non-zero element
          Next N
          'Change line k with line_1
          For M = k To System_DIM * 2 - 1
             temporary_1 = Operations_Matrix(k, M)
             Operations_Matrix(k, M) = Operations_Matrix(line_1, M)
             Operations_Matrix(line_1, M) = temporary_1
          Next M
       End If
       
        elem1 = Operations_Matrix(k, k)
       For N = k To 2 * System_DIM - 1
        Operations_Matrix(k, N) = Operations_Matrix(k, N) / elem1
       Next N
       
       'For other lines, make a zero element by using:
       'Ai1=Aij-A11*(Aij/A11)
       'and change all the line using the same formula for other elements
       For N = 0 To System_DIM - 1
            If N = k And N = MAX_DIM Then Exit For 'Finished
            If N = k And N < MAX_DIM Then N = N + 1 'Do not change that element (already equals to 1), go for next one
          If Operations_Matrix(N, k) <> 0 Then 'if it is zero, stays as it is
             multiplier_1 = Operations_Matrix(N, k) / Operations_Matrix(k, k)
             For M = k To 2 * System_DIM - 1
                Operations_Matrix(N, M) = Operations_Matrix(N, M) - Operations_Matrix(k, M) * multiplier_1
             Next M
          End If
       Next N
    Next k
    
    'Assign the right part to the Inverse_Matrix
    For N = 0 To System_DIM - 1
        For k = 0 To System_DIM - 1
            Inverse_Matrix(N, k) = Operations_Matrix(N, System_DIM + k)
        Next k
    Next N
    
    Exit Sub
    
errhandler:
    message = "An error occured during the calculation process. Determinant of Matrix [A] is probably equal to zero."
    response = MsgBox(message, vbCritical)
    Solution_Problem = True

End Sub

