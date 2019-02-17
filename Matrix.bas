Attribute VB_Name = "Matrix"
Option Explicit

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


Public Function ScalarProduct(Scalar As Double, v As Vector) As Vector
    ScalarProduct.x = Scalar * v.x
    ScalarProduct.Y = Scalar * v.Y
    ScalarProduct.z = Scalar * v.z
End Function


Public Function LenghtVector(v As Vector) As Double
    LenghtVector = Sqr(v.x * v.x + v.Y * v.Y + v.z * v.z)
End Function




Public Sub MatrixProduct(matrix1() As Double, ByVal row1 As Integer, ByVal col1 As Integer, matrix2() As Double, ByVal row2 As Integer, ByVal col2 As Integer, matrix3() As Double)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Dim temp As Double
    
    For i = 0 To row1 - 1
        For j = 0 To col2 - 1
            For k = 0 To row2 - 1
                temp = temp + (matrix1(i, k) * matrix2(k, j))
            Next
            matrix3(i, j) = temp
            temp = 0
        Next
    Next
End Sub


'Public Sub Calculate_Inverse(System_DIM As Long, Matrix_A() As Double, Inverse_Matrix() As Double)
'
'    Dim k As Long
'    Dim N As Long
'    Dim m As Long
'    Dim line_1 As Long
'    Dim temporary_1 As Double
'    Dim elem1 As Double
'    Dim multiplier_1 As Double
'    Dim MAX_DIM As Long
'    Dim message As String
'    Dim response
'
'    Dim Operations_Matrix(10, 20) As Double
'
'    Dim Solution_Problem As Boolean
'
''Uses Gauss elimination method in order to calculate the inverse matrix [A]-1
''Method: Puts matrix [A] at the left and the singular matrix [I] at the right:
''[ a11 a12 a13 | 1 0 0 ]
''[ a21 a22 a23 | 0 1 0 ]
''[ a31 a32 a33 | 0 0 1 ]
''Then using line operations, we try to build the singular matrix [I] at the left.
''After we have finished, the inverse matrix [A]-1 (bij) will be at the right:
''[ 1 0 0 | b11 b12 b13 ]
''[ 0 1 0 | b21 b22 b23 ]
''[ 0 0 1 | b31 b32 b33 ]
'
'On Error GoTo errhandler 'In case the inverse cannot be found (Determinant = 0)
'
'Solution_Problem = False
'MAX_DIM = 10
'
''Assign values from matrix [A] at the left
'For N = 1 To System_DIM
'    For m = 1 To System_DIM
'        Operations_Matrix(m, N) = Matrix_A(m, N)
'    Next
'Next
'
''Assign values from singular matrix [I] at the right
'For N = 1 To System_DIM
'    For m = 1 To System_DIM
'        If N = m Then
'            Operations_Matrix(m, N + System_DIM) = 1
'        Else
'            Operations_Matrix(m, N + System_DIM) = 0
'        End If
'    Next
'Next
'
'
''Build the Singular matrix [I] at the left
'For k = 1 To System_DIM
'   'Bring a non-zero element first by changes lines if necessary
'   If Operations_Matrix(k, k) = 0 Then
'      For N = k To System_DIM
'        If Operations_Matrix(N, k) <> 0 Then line_1 = N: Exit For 'Finds line_1 with non-zero element
'      Next N
'      'Change line k with line_1
'      For m = k To System_DIM * 2
'         temporary_1 = Operations_Matrix(k, m)
'         Operations_Matrix(k, m) = Operations_Matrix(line_1, m)
'         Operations_Matrix(line_1, m) = temporary_1
'      Next m
'   End If
'
'    elem1 = Operations_Matrix(k, k)
'   For N = k To 2 * System_DIM
'    Operations_Matrix(k, N) = Operations_Matrix(k, N) / elem1
'   Next N
'
'   'For other lines, make a zero element by using:
'   'Ai1=Aij-A11*(Aij/A11)
'   'and change all the line using the same formula for other elements
'   For N = 1 To System_DIM
'        If N = k And N = MAX_DIM Then Exit For 'Finished
'        If N = k And N < MAX_DIM Then N = N + 1 'Do not change that element (already equals to 1), go for next one
'      If Operations_Matrix(N, k) <> 0 Then 'if it is zero, stays as it is
'         multiplier_1 = Operations_Matrix(N, k) / Operations_Matrix(k, k)
'         For m = k To 2 * System_DIM
'            Operations_Matrix(N, m) = Operations_Matrix(N, m) - Operations_Matrix(k, m) * multiplier_1
'         Next m
'      End If
'   Next N
'Next k
'
''Assign the right part to the Inverse_Matrix
'For N = 1 To System_DIM
'    For k = 1 To System_DIM
'        Inverse_Matrix(N, k) = Operations_Matrix(N, System_DIM + k)
'    Next k
'Next N
'
'Exit Sub
'
'errhandler:
'message = "An error occured during the calculation process. Determinant of Matrix [A] is probably equal to zero."
'response = MsgBox(message, vbCritical)
'Solution_Problem = True
'
'End Sub



Public Sub Calculate_Inverse(System_DIM As Long, Matrix_A() As Double, Inverse_Matrix() As Double)

    Dim k As Long
    Dim N As Long
    Dim m As Long
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
    For m = 0 To System_DIM - 1
        Operations_Matrix(m, N) = Matrix_A(m, N)
    Next
Next

'Assign values from singular matrix [I] at the right
For N = 0 To System_DIM - 1
    For m = 0 To System_DIM - 1
        If N = m Then
            Operations_Matrix(m, N + System_DIM) = 1
        Else
            Operations_Matrix(m, N + System_DIM) = 0
        End If
    Next
Next

'    m11 = Operations_Matrix(0, 0)
'    m12 = Operations_Matrix(0, 1)
'    m13 = Operations_Matrix(0, 2)
'    m21 = Operations_Matrix(1, 0)
'    m22 = Operations_Matrix(1, 1)
'    m23 = Operations_Matrix(1, 2)
'    m31 = Operations_Matrix(2, 0)
'    m32 = Operations_Matrix(2, 1)
'    m33 = Operations_Matrix(2, 2)
'
'    m14 = Operations_Matrix(0, 3)
'    m15 = Operations_Matrix(0, 4)
'    m16 = Operations_Matrix(0, 5)
'    m24 = Operations_Matrix(1, 3)
'    m25 = Operations_Matrix(1, 4)
'    m26 = Operations_Matrix(1, 5)
'    m34 = Operations_Matrix(2, 3)
'    m35 = Operations_Matrix(2, 4)
'    m36 = Operations_Matrix(2, 5)

'Build the Singular matrix [I] at the left
For k = 0 To System_DIM - 1
   'Bring a non-zero element first by changes lines if necessary
   If Operations_Matrix(k, k) = 0 Then
      For N = k To System_DIM - 1
        If Operations_Matrix(N, k) <> 0 Then line_1 = N: Exit For 'Finds line_1 with non-zero element
      Next N
      'Change line k with line_1
      For m = k To System_DIM * 2 - 1
         temporary_1 = Operations_Matrix(k, m)
         Operations_Matrix(k, m) = Operations_Matrix(line_1, m)
         Operations_Matrix(line_1, m) = temporary_1
      Next m
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
         For m = k To 2 * System_DIM - 1
            Operations_Matrix(N, m) = Operations_Matrix(N, m) - Operations_Matrix(k, m) * multiplier_1
Mainform.List1.AddItem N & " " & m & " " & Operations_Matrix(N, m)
         Next m
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

