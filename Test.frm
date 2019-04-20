VERSION 5.00
Begin VB.Form Test 
   Caption         =   "Test"
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   2760
      TabIndex        =   20
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2880
      TabIndex        =   19
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton C_TestKalibrierung_3 
      Caption         =   "3. Test Kal. using sub()"
      Height          =   435
      Left            =   240
      TabIndex        =   18
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton C_TestKalibrierung_2 
      Caption         =   "2. Test der Kalibrierung"
      Height          =   435
      Left            =   240
      TabIndex        =   17
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command2"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   4080
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   4920
      TabIndex        =   14
      Top             =   1080
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Demo Stern"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton C_TestKalibrierung_1 
      Caption         =   "1. Test der Kalibrierung"
      Height          =   435
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton C_TestSiderialTime 
      Caption         =   "Siderial Time"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label L_AltStar 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label L_AzStar 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Azimuth"
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Altitude"
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label L_HourAngle 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Hour Angle"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label L_SiderialTime 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Siderial Time"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label L_SiderialTimeHMS 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "--"
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub C_TestKalibrierung_1_Click()
   ' matrix_method_rev_d.pdf Seite 37
    Dim tmp As Vector

    Dim tst As MyTime
    Dim InitTimerad As Double
    Dim ObservTime1Rad As Double
    Dim ObservTime2Rad As Double
    Dim RA1Rad As Double
    Dim RA2Rad As Double
    Dim RA1Deg As Double
    Dim RA2Deg As Double
    Dim DEC1Rad As Double
    Dim DEC2Rad As Double
    Dim TelHorizAngle1 As Double
    Dim TelHorizAngle2 As Double
    Dim TelElevAngle1 As Double
    Dim TelElevAngle2 As Double


    tst.H = 21
    tst.M = 0
    tst.s = 0
    InitTimerad = TimeToRad(tst)

    tst.H = 21
    tst.M = 27
    tst.s = 56
    ObservTime1Rad = TimeToRad(tst)

    tst.H = 0
    tst.M = 7
    tst.s = 54
    RA1Rad = TimeToRad(tst)
    RA1Deg = RadToDeg(RA1Rad)
    DEC1Rad = DegToRad(29.038)
    TelHorizAngle1 = DegToRad(99.25)
    TelElevAngle1 = DegToRad(83.87)

    tst.H = 21
    tst.M = 37
    tst.s = 2
    ObservTime2Rad = TimeToRad(tst)

    tst.H = 2
    tst.M = 21
    tst.s = 45
    RA2Rad = TimeToRad(tst)
    RA2Deg = RadToDeg(RA2Rad)
    DEC2Rad = DegToRad(89.222)
    TelHorizAngle2 = DegToRad(310.98)
    TelElevAngle2 = DegToRad(35.04)


    Dim lmn_Tel_1 As Vector     ' Telescope coordinates
    Dim lmn_Tel_2 As Vector
    Dim lmn_Tel_3 As Vector
    Dim LMN_Equ_1 As Vector
    Dim LMN_Equ_2 As Vector
    Dim LMN_Equ_3 As Vector
    Dim k As Double         ' Umrechnung Sonnenzeit in siderische Zeit 1.00273790935
    k = 1.00273790935

    'Equation (5.4-5)
    lmn_Tel_1.x = Cos(TelElevAngle1) * Cos(TelHorizAngle1)
    lmn_Tel_1.Y = Cos(TelElevAngle1) * Sin(TelHorizAngle1)
    lmn_Tel_1.z = Sin(TelElevAngle1)

    'Equation (5.4-6)
    LMN_Equ_1.x = Cos(DEC1Rad) * Cos(RA1Rad - k * (ObservTime1Rad - InitTimerad))
    LMN_Equ_1.Y = Cos(DEC1Rad) * Sin(RA1Rad - k * (ObservTime1Rad - InitTimerad))
    LMN_Equ_1.z = Sin(DEC1Rad)

    'Equation (5.4-7)
    lmn_Tel_2.x = Cos(TelElevAngle2) * Cos(TelHorizAngle2)
    lmn_Tel_2.Y = Cos(TelElevAngle2) * Sin(TelHorizAngle2)
    lmn_Tel_2.z = Sin(TelElevAngle2)

    'Equation (5.4-8)
    LMN_Equ_2.x = Cos(DEC2Rad) * Cos(RA2Rad - k * (ObservTime2Rad - InitTimerad))
    LMN_Equ_2.Y = Cos(DEC2Rad) * Sin(RA2Rad - k * (ObservTime2Rad - InitTimerad))
    LMN_Equ_2.z = Sin(DEC2Rad)

    Dim V1_cross_V2 As Vector
    Dim Len_V1_cross_V2 As Double

    'Equation (5.4-13)
    V1_cross_V2 = CrossProduct(lmn_Tel_1, lmn_Tel_2)
    Len_V1_cross_V2 = LenghtVector(V1_cross_V2)
    lmn_Tel_3 = ScalarProduct((1 / Len_V1_cross_V2), V1_cross_V2)

    'Equation (5.4-14)
    V1_cross_V2 = CrossProduct(LMN_Equ_1, LMN_Equ_2)
    Len_V1_cross_V2 = LenghtVector(V1_cross_V2)
    LMN_Equ_3 = ScalarProduct((1 / Len_V1_cross_V2), V1_cross_V2)


    'From equation(5.4 - 11)
    Dim LMN_Equ_Matrix(10, 10) As Double
    Dim LMN_Equ_MatrixInvers(10, 10) As Double
    Dim lmn_Tel_Matrix(10, 10) As Double
    Dim TransformationMatrix(10, 10) As Double

    LMN_Equ_Matrix(0, 0) = LMN_Equ_1.x: LMN_Equ_Matrix(0, 1) = LMN_Equ_2.x: LMN_Equ_Matrix(0, 2) = LMN_Equ_3.x
    LMN_Equ_Matrix(1, 0) = LMN_Equ_1.Y: LMN_Equ_Matrix(1, 1) = LMN_Equ_2.Y: LMN_Equ_Matrix(1, 2) = LMN_Equ_3.Y
    LMN_Equ_Matrix(2, 0) = LMN_Equ_1.z: LMN_Equ_Matrix(2, 1) = LMN_Equ_2.z: LMN_Equ_Matrix(2, 2) = LMN_Equ_3.z

    Calculate_Inverse 3, LMN_Equ_Matrix, LMN_Equ_MatrixInvers
                Dim dmy As Double
                dmy = LMN_Equ_MatrixInvers(0, 0): dmy = LMN_Equ_MatrixInvers(0, 1): dmy = LMN_Equ_MatrixInvers(0, 2)
                dmy = LMN_Equ_MatrixInvers(1, 0): dmy = LMN_Equ_MatrixInvers(1, 1): dmy = LMN_Equ_MatrixInvers(1, 2)
                dmy = LMN_Equ_MatrixInvers(2, 0): dmy = LMN_Equ_MatrixInvers(2, 1): dmy = LMN_Equ_MatrixInvers(2, 2)

    lmn_Tel_Matrix(0, 0) = lmn_Tel_1.x: lmn_Tel_Matrix(0, 1) = lmn_Tel_2.x: lmn_Tel_Matrix(0, 2) = lmn_Tel_3.x
    lmn_Tel_Matrix(1, 0) = lmn_Tel_1.Y: lmn_Tel_Matrix(1, 1) = lmn_Tel_2.Y: lmn_Tel_Matrix(1, 2) = lmn_Tel_3.Y
    lmn_Tel_Matrix(2, 0) = lmn_Tel_1.z: lmn_Tel_Matrix(2, 1) = lmn_Tel_2.z: lmn_Tel_Matrix(2, 2) = lmn_Tel_3.z

    '==================================================================================================
    'This is the TransformationMatrix which transforms a vector from eqatorial to telescope coordinates
    '==================================================================================================
    MatrixProduct lmn_Tel_Matrix, 3, 3, LMN_Equ_MatrixInvers, 3, 3, TransformationMatrix
                dmy = TransformationMatrix(0, 0): dmy = TransformationMatrix(0, 1): dmy = TransformationMatrix(0, 2)
                dmy = TransformationMatrix(1, 0): dmy = TransformationMatrix(1, 1): dmy = TransformationMatrix(1, 2)
                dmy = TransformationMatrix(2, 0): dmy = TransformationMatrix(2, 1): dmy = TransformationMatrix(2, 2)




    '=================================
    ' Example: Beta Cet: Deneb Kaitos
    '=================================

    Dim RA_BetaCet As MyTime
    Dim DEC_BetaCet As Double
    Dim AimTime As MyTime
    Dim RA_BetaCetRad As Double
    Dim DEC_BetaCetRad As Double
    Dim AimTimeRad As Double



    ' if you want to aim the telescope at Beta Cet (RA = 0h43m07s, DEC = -18.038°) at 21h52m12s
    RA_BetaCet.H = 0: RA_BetaCet.M = 43: RA_BetaCet.s = 7
    RA_BetaCetRad = TimeToRad(RA_BetaCet)
    DEC_BetaCetRad = DegToRad(-18.038)
    AimTime.H = 21: AimTime.M = 52: AimTime.s = 12
    AimTimeRad = TimeToRad(AimTime)

    'point to Alpheratz
'    RA_BetaCet.H = 0: RA_BetaCet.M = 7: RA_BetaCet.s = 54
'    RA_BetaCetRad = TimeToRad(RA_BetaCet)
'    DEC_BetaCetRad = DegToRad(29.038)
'    AimTime.H = 21: AimTime.M = 27: AimTime.s = 56
'    AimTimeRad = TimeToRad(AimTime)

 
    'LMN_Equ_Result: Vector points to Deneb in equatorial coordinats
    Dim LMN_Equ_Result  As Vector
    LMN_Equ_Result.x = Cos(DEC_BetaCetRad) * Cos(RA_BetaCetRad - k * (AimTimeRad - InitTimerad))
    LMN_Equ_Result.Y = Cos(DEC_BetaCetRad) * Sin(RA_BetaCetRad - k * (AimTimeRad - InitTimerad))
    LMN_Equ_Result.z = Sin(DEC_BetaCetRad)


    Dim LMN_Equ_ResultMatrix(10, 10) As Double
    Dim lmn_Tel_ResultMatrix(10, 10) As Double
    LMN_Equ_ResultMatrix(0, 0) = LMN_Equ_Result.x
    LMN_Equ_ResultMatrix(1, 0) = LMN_Equ_Result.Y
    LMN_Equ_ResultMatrix(2, 0) = LMN_Equ_Result.z

    MatrixProduct TransformationMatrix, 3, 3, LMN_Equ_ResultMatrix, 3, 1, lmn_Tel_ResultMatrix

    'lmn_Tel__Matrix: Vector points to Beta Cet in equatorial coordinats

    Dim lmn_Tel_Result  As Vector
    lmn_Tel_Result.x = lmn_Tel_ResultMatrix(0, 0)
    lmn_Tel_Result.Y = lmn_Tel_ResultMatrix(1, 0)
    lmn_Tel_Result.z = lmn_Tel_ResultMatrix(2, 0)

    Dim AzAlt_BetaCet As AzAlt
    Dim Az_BetaCetRad As Double
    Dim Alt_BetaCetRad As Double
    Dim Az_BetaCet As Double
    Dim Az_BetaCet_corrected_1 As Double
    Dim Az_BetaCet_corrected_2 As Double
    Dim Alt_BetaCet As Double

    AzAlt_BetaCet = VectorToAzAlt(lmn_Tel_Result)
    Az_BetaCetRad = AzAlt_BetaCet.Az
    Alt_BetaCetRad = AzAlt_BetaCet.Alt

    Az_BetaCet = RadToDeg(Az_BetaCetRad)
    
    ' !!! hier muß möglicherweise noch 180° addiert werden !!!
    Az_BetaCet_corrected_1 = 180 - Az_BetaCet
    Az_BetaCet_corrected_2 = Az_BetaCet_corrected_1 + 180

    Alt_BetaCet = RadToDeg(Alt_BetaCetRad)


End Sub

Private Sub C_TestKalibrierung_2_Click()
    ' matrix_method_rev_d.pdf Seite 37
    Dim tmp As Vector

    Dim tst As MyTime
    Dim InitTimerad As Double
    Dim ObservTime1Rad As Double
    Dim ObservTime2Rad As Double
    Dim RA1Rad As Double
    Dim RA2Rad As Double
    Dim RA1Deg As Double
    Dim RA2Deg As Double
    Dim DEC1Rad As Double
    Dim DEC2Rad As Double
    Dim TelHorizAngle1 As Double
    Dim TelHorizAngle2 As Double
    Dim TelElevAngle1 As Double
    Dim TelElevAngle2 As Double

    'Observation date/time 1.8.2000 22:00:00 UT
    'Observation location: Munich: 48°08'00"N (-)11°34'00"E

    'Initial time
    tst.H = 22
    tst.M = 0
    tst.s = 0
    InitTimerad = TimeToRad(tst)

    'time pointing to 1. reference star
    tst.H = 22
    tst.M = 0
    tst.s = 0
    ObservTime1Rad = TimeToRad(tst)

    'RA DEC of 1. star (Alpheratz oder Sirrah - Alpha And)
    tst.H = 0
    tst.M = 8
    tst.s = 24
    RA1Rad = TimeToRad(tst)
    RA1Deg = RadToDeg(RA1Rad)   '? 1.975°
    DEC1Rad = DegToRad(29.0906)
    'Telescope cooridinates 1. star
    TelHorizAngle1 = -DegToRad(83.119)      'Danger! Telescope angle is CCW
    TelElevAngle1 = DegToRad(34.3481)

    'time pointing to 2. reference star
    tst.H = 22
    tst.M = 0
    tst.s = 0
    ObservTime2Rad = TimeToRad(tst)

    'RA DEC of 2. star (Polaris - Alpha Umi)
    tst.H = 2
    tst.M = 31
    tst.s = 48
    RA2Rad = TimeToRad(tst)
    RA2Deg = RadToDeg(RA2Rad)   '? 35,4375°
    DEC2Rad = DegToRad(89.2642)
    'Telescope cooridinates 2. star
    TelHorizAngle2 = -DegToRad(1.058)       'Danger! Telescope angle is CCW
    TelElevAngle2 = DegToRad(47.931)


    Dim lmn_Tel_1 As Vector     ' Telescope coordinates
    Dim lmn_Tel_2 As Vector
    Dim lmn_Tel_3 As Vector
    Dim LMN_Equ_1 As Vector
    Dim LMN_Equ_2 As Vector
    Dim LMN_Equ_3 As Vector
    Dim k As Double         ' Umrechnung Sonnenzeit in siderische Zeit 1.00273790935
    k = 1.00273790935

    'Equation (5.4-5)
    'Telescope coordinates star 1
    lmn_Tel_1.x = Cos(TelElevAngle1) * Cos(TelHorizAngle1)  '0.099
    lmn_Tel_1.Y = Cos(TelElevAngle1) * Sin(TelHorizAngle1)  '0.8207
    lmn_Tel_1.z = Sin(TelElevAngle1)                        '0.5628


    'Equation (5.4-6)
    'RA DEC star 1
    LMN_Equ_1.x = Cos(DEC1Rad) * Cos(RA1Rad - k * (ObservTime1Rad - InitTimerad))   '0.8738
    LMN_Equ_1.Y = Cos(DEC1Rad) * Sin(RA1Rad - k * (ObservTime1Rad - InitTimerad))   '0.0301
    LMN_Equ_1.z = Sin(DEC1Rad)                                                      '0.4854

    'Equation (5.4-7)
    'Telescope coordinates star 2
    lmn_Tel_2.x = Cos(TelElevAngle2) * Cos(TelHorizAngle2)  '0.6699
    lmn_Tel_2.Y = Cos(TelElevAngle2) * Sin(TelHorizAngle2)  '0.0124
    lmn_Tel_2.z = Sin(TelElevAngle2)                        '0.7423

    'Equation (5.4-8)
    'RA DEC star 2
    LMN_Equ_2.x = Cos(DEC2Rad) * Cos(RA2Rad - k * (ObservTime2Rad - InitTimerad))   '0.0111
    LMN_Equ_2.Y = Cos(DEC2Rad) * Sin(RA2Rad - k * (ObservTime2Rad - InitTimerad))   '0.0079
    LMN_Equ_2.z = Sin(DEC2Rad)                                                      '0.9999

    Dim V1_cross_V2 As Vector
    Dim Len_V1_cross_V2 As Double

    'Equation (5.4-13)
    V1_cross_V2 = CrossProduct(lmn_Tel_1, lmn_Tel_2)
    Len_V1_cross_V2 = LenghtVector(V1_cross_V2)
    lmn_Tel_3 = ScalarProduct((1 / Len_V1_cross_V2), V1_cross_V2)

    'Equation (5.4-14)
    V1_cross_V2 = CrossProduct(LMN_Equ_1, LMN_Equ_2)
    Len_V1_cross_V2 = LenghtVector(V1_cross_V2)
    LMN_Equ_3 = ScalarProduct((1 / Len_V1_cross_V2), V1_cross_V2)


    'From equation(5.4 - 11)
    Dim LMN_Equ_Matrix(10, 10) As Double
    Dim LMN_Equ_MatrixInvers(10, 10) As Double
    Dim lmn_Tel_Matrix(10, 10) As Double
    Dim TransformationMatrix(10, 10) As Double

    LMN_Equ_Matrix(0, 0) = LMN_Equ_1.x
    LMN_Equ_Matrix(0, 1) = LMN_Equ_2.x
    LMN_Equ_Matrix(0, 2) = LMN_Equ_3.x
    LMN_Equ_Matrix(1, 0) = LMN_Equ_1.Y
    LMN_Equ_Matrix(1, 1) = LMN_Equ_2.Y
    LMN_Equ_Matrix(1, 2) = LMN_Equ_3.Y
    LMN_Equ_Matrix(2, 0) = LMN_Equ_1.z
    LMN_Equ_Matrix(2, 1) = LMN_Equ_2.z
    LMN_Equ_Matrix(2, 2) = LMN_Equ_3.z

    Calculate_Inverse 3, LMN_Equ_Matrix, LMN_Equ_MatrixInvers

    lmn_Tel_Matrix(0, 0) = lmn_Tel_1.x
    lmn_Tel_Matrix(0, 1) = lmn_Tel_2.x
    lmn_Tel_Matrix(0, 2) = lmn_Tel_3.x
    lmn_Tel_Matrix(1, 0) = lmn_Tel_1.Y
    lmn_Tel_Matrix(1, 1) = lmn_Tel_2.Y
    lmn_Tel_Matrix(1, 2) = lmn_Tel_3.Y
    lmn_Tel_Matrix(2, 0) = lmn_Tel_1.z
    lmn_Tel_Matrix(2, 1) = lmn_Tel_2.z
    lmn_Tel_Matrix(2, 2) = lmn_Tel_3.z

    '==================================================================================================
    'This is the TransformationMatrix which transforms a vector from eqatorial to telescope coordinates
    '==================================================================================================
    MatrixProduct lmn_Tel_Matrix, 3, 3, LMN_Equ_MatrixInvers, 3, 3, TransformationMatrix
        Dim dmy As Double
        dmy = TransformationMatrix(0, 0): dmy = TransformationMatrix(0, 1): dmy = TransformationMatrix(0, 2)
        dmy = TransformationMatrix(1, 0): dmy = TransformationMatrix(1, 1): dmy = TransformationMatrix(1, 2)
        dmy = TransformationMatrix(2, 0): dmy = TransformationMatrix(2, 1): dmy = TransformationMatrix(2, 2)



    '=================================
    ' Example:Deneb
    '=================================

    Dim RA_BetaCet As MyTime
    Dim DEC_BetaCet As Double
    Dim AimTime As MyTime
    Dim RA_BetaCetRad As Double
    Dim DEC_BetaCetRad As Double
    Dim AimTimeRad As Double

    ' if you want to aim the telescope at Deneb (RA = 20h41m24s, DEC = 45.28°) at 22h0m0s
    RA_BetaCet.H = 20: RA_BetaCet.M = 41: RA_BetaCet.s = 24
    RA_BetaCetRad = TimeToRad(RA_BetaCet)
    DEC_BetaCetRad = DegToRad(45.2803)
    AimTime.H = 22: AimTime.M = 30: AimTime.s = 0
    AimTimeRad = TimeToRad(AimTime)

    'LMN_Equ_Result: Vector points to Deneb in equatorial coordinats
    Dim LMN_Equ_Result  As Vector
    LMN_Equ_Result.x = Cos(DEC_BetaCetRad) * Cos(RA_BetaCetRad - k * (AimTimeRad - InitTimerad))
    LMN_Equ_Result.Y = Cos(DEC_BetaCetRad) * Sin(RA_BetaCetRad - k * (AimTimeRad - InitTimerad))
    LMN_Equ_Result.z = Sin(DEC_BetaCetRad)


    Dim LMN_Equ_ResultMatrix(10, 10) As Double
    Dim lmn_Tel_ResultMatrix(10, 10) As Double
    LMN_Equ_ResultMatrix(0, 0) = LMN_Equ_Result.x
    LMN_Equ_ResultMatrix(1, 0) = LMN_Equ_Result.Y
    LMN_Equ_ResultMatrix(2, 0) = LMN_Equ_Result.z

    MatrixProduct TransformationMatrix, 3, 3, LMN_Equ_ResultMatrix, 3, 1, lmn_Tel_ResultMatrix
        dmy = lmn_Tel_ResultMatrix(0, 0): dmy = lmn_Tel_ResultMatrix(0, 1): dmy = lmn_Tel_ResultMatrix(0, 2)
        dmy = lmn_Tel_ResultMatrix(1, 0): dmy = lmn_Tel_ResultMatrix(1, 1): dmy = lmn_Tel_ResultMatrix(1, 2)
        dmy = lmn_Tel_ResultMatrix(2, 0): dmy = lmn_Tel_ResultMatrix(2, 1): dmy = lmn_Tel_ResultMatrix(2, 2)

    'lmn_Tel__Matrix: Vector points to Beta Cet in equatorial coordinats

    Dim lmn_Tel_Result  As Vector
    lmn_Tel_Result.x = lmn_Tel_ResultMatrix(0, 0)
    lmn_Tel_Result.Y = lmn_Tel_ResultMatrix(1, 0)
    lmn_Tel_Result.z = lmn_Tel_ResultMatrix(2, 0)

    Dim AzAlt_BetaCet As AzAlt
    Dim Az_BetaCetRad As Double
    Dim Alt_BetaCetRad As Double
    Dim Az_BetaCet As Double
    Dim Az_BetaCet_corrected_1 As Double
    Dim Az_BetaCet_corrected_2 As Double
    Dim Alt_BetaCet As Double

    AzAlt_BetaCet = VectorToAzAlt(lmn_Tel_Result)
    Az_BetaCetRad = AzAlt_BetaCet.Az
    Alt_BetaCetRad = AzAlt_BetaCet.Alt

    Az_BetaCet = RadToDeg(Az_BetaCetRad)
    
    ' !!! hier muß möglicherweise noch 180° addiert werden !!!
    Az_BetaCet_corrected_1 = 180 - Az_BetaCet
    Az_BetaCet_corrected_2 = Az_BetaCet_corrected_1 + 180

    Alt_BetaCet = RadToDeg(Alt_BetaCetRad)


End Sub

Private Sub C_TestKalibrierung_3_Click()
    ' matrix_method_rev_d.pdf Seite 37
    Dim tst As MyTime
    Dim InitTimerad As Double
    Dim ObservTime1Rad As Double
    Dim ObservTime2Rad As Double
    Dim RA1Rad As Double
    Dim RA2Rad As Double
    Dim RA1Deg As Double
    Dim RA2Deg As Double
    Dim DEC1Rad As Double
    Dim DEC2Rad As Double
    Dim TelHorizAngle1 As Double
    Dim TelHorizAngle2 As Double
    Dim TelElevAngle1 As Double
    Dim TelElevAngle2 As Double


    'Observation date/time: unknown
    'Observation location: unknown

    'Initial time
     tst.H = 21: tst.M = 0: tst.s = 0
    InitTimerad = TimeToRad(tst)

    'time pointing to 1. reference star
    tst.H = 21: tst.M = 27: tst.s = 56
    ObservTime1Rad = TimeToRad(tst)

    'RA DEC of 1. star (Alpheratz oder Sirrah - Alpha And)
    tst.H = 0: tst.M = 7: tst.s = 54
    RA1Rad = TimeToRad(tst)
    RA1Deg = RadToDeg(RA1Rad)
    DEC1Rad = DegToRad(29.038)
    TelHorizAngle1 = DegToRad(99.25)      'Danger! Telescope angle is CCW
    TelElevAngle1 = DegToRad(83.87)

    'time pointing to 2. reference star
    tst.H = 21: tst.M = 37: tst.s = 2
    ObservTime2Rad = TimeToRad(tst)

    'RA DEC of 2. star (Polaris - Alpha Umi)
    tst.H = 2: tst.M = 21: tst.s = 45
    RA2Rad = TimeToRad(tst)
    RA2Deg = RadToDeg(RA2Rad)
    DEC2Rad = DegToRad(89.222)
    TelHorizAngle2 = DegToRad(310.98)      'Danger! Telescope angle is CCW
    TelElevAngle2 = DegToRad(35.04)

    CalibrateTelescope InitTimerad, _
                       RA1Rad, DEC1Rad, TelHorizAngle1, TelElevAngle1, ObservTime1Rad, _
                       RA2Rad, DEC2Rad, TelHorizAngle2, TelElevAngle2, ObservTime2Rad, _
                       TransformationMatrix


    '=================================
    ' Example: Beta Cet: Deneb Kaitos
    '=================================
    Dim RA_BetaCet As MyTime
    Dim DEC_BetaCet As Double
    Dim AimTime As MyTime
    Dim RA_BetaCetRad As Double
    Dim DEC_BetaCetRad As Double
    Dim AimTimeRad As Double
    Dim AzAlt_BetaCet As AzAlt
    
    ' if you want to aim the telescope at Beta Cet (RA = 0h43m07s, DEC = -18.038°) at 21h52m12s
    RA_BetaCet.H = 0: RA_BetaCet.M = 43: RA_BetaCet.s = 7
    RA_BetaCetRad = TimeToRad(RA_BetaCet)
    DEC_BetaCetRad = DegToRad(-18.038)
    AimTime.H = 21: AimTime.M = 52: AimTime.s = 12
    AimTimeRad = TimeToRad(AimTime)

    'point to Alpheratz
'    RA_BetaCet.H = 0: RA_BetaCet.M = 7: RA_BetaCet.s = 54
'    RA_BetaCetRad = TimeToRad(RA_BetaCet)
'    DEC_BetaCetRad = DegToRad(29.038)
'    AimTime.H = 21: AimTime.M = 27: AimTime.s = 56
'    AimTimeRad = TimeToRad(AimTime)

    CalculateTelescopeCoordinates InitTimerad, _
                                  RA_BetaCetRad, DEC_BetaCetRad, AimTimeRad, TransformationMatrix, _
                                  AzAlt_BetaCet

 
    Dim Az_BetaCetRad As Double
    Dim Alt_BetaCetRad As Double
    Dim Az_BetaCet_corrected_1 As Double
    Dim Az_BetaCet_corrected_2 As Double
    Dim Az_BetaCet As Double
    Dim Alt_BetaCet As Double
    
    Az_BetaCetRad = AzAlt_BetaCet.Az
    Alt_BetaCetRad = AzAlt_BetaCet.Alt

    Az_BetaCet = RadToDeg(Az_BetaCetRad)
    
    ' !!! hier muß möglicherweise noch 180° addiert werden !!!
    Az_BetaCet_corrected_1 = 180 - Az_BetaCet
    Az_BetaCet_corrected_2 = Az_BetaCet_corrected_1 + 180

    Alt_BetaCet = RadToDeg(Alt_BetaCetRad)



End Sub

' Test siderial time
' https://de.wikibooks.org/wiki/Astronomische_Berechnungen_f%C3%BCr_Amateure/_Zeit/_Zeitrechnungen
' Welchen Wert hatte die mittlere Sternzeit?
' Berlin (Länge = +13.5°) am 25. Dezember 2007 um 20 h UT (entspricht 21 MEZ in Berlin)?
' Ergebnis: 3,1634161794371 = 3h 09m 48,3s
Private Sub C_TestSiderialTime_Click()

    Dim DemoDate As MyDate
    Dim DemoTime As MyTime
    Dim SiderialTime As MyTime
    Dim SiderialTimeGreenwich As MyTime
    Dim s As String

    DemoDate.YY = 2007
    DemoDate.MM = 12
    DemoDate.DD = 25
    DemoTime.H = 20
    DemoTime.M = 0
    DemoTime.s = 0

    SiderialTimeGreenwich = GMST(DemoDate, DemoTime)
    SiderialTime = TimeDezToHMS(SiderialTimeGreenwich.TimeDec + 13.5 / 15)
    L_SiderialTime = SiderialTime.TimeDec
    L_SiderialTimeHMS = SiderialTime.H & ":" & SiderialTime.M & ":" & Format(SiderialTime.s, "00.00")

End Sub

Private Sub Command1_Click()
    Dim ut As Date
    Dim AnyDateTime As Date

    Dim tst1 As Integer
    Dim tst2 As Integer
    Dim tst3 As Integer
    Dim tst4 As Integer
    Dim tst5 As Integer
    Dim tst6 As Integer

    ut = UtcTime(Now)

    AnyDateTime = "18.2.2019 1:0:0"
    ut = UtcTime(AnyDateTime)

    tst1 = Day(ut)
    tst2 = Month(ut)
    tst3 = Year(ut)
    tst4 = Hour(ut)
    tst5 = Minute(ut)
    tst6 = Second(ut)

    'Achtung: "2019.11.4 1:0:00" liefert nur "4.11.2019"

End Sub

Private Sub Command2_Click()
    TestJulianischesDatum.Show
End Sub

Private Sub Command3_Click()
    Dim a As String
    Dim B As String
    Dim erg As Long

    a = SetNexStarPosition(1234567)

    B = Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H11) & Chr$(&H24) & Chr$(&H80)
'    b = Chr$(&H0) & Chr$(&H3) & Chr$(&HE8)

    erg = GetNexStarPosition(a)

End Sub


Private Sub Command4_Click()
    Dim Jetzt As String
    Dim Lont As MyTime

    ' Datensatz Saturn Demo aus dem Sript
    Dim SaturnDateTime As Date
    SaturnDateTime = "13.11.1978 4:34:0"

    Dim RA_Saturn As MyTime
    RA_Saturn.H = 10
    RA_Saturn.M = 57              '57
    RA_Saturn.s = 35.681

    Dim DEC_Saturn As GeoCoord
    DEC_Saturn.deg = 8
    DEC_Saturn.Min = 25
    DEC_Saturn.Sec = 58.1
    DEC_Saturn.Sign = "+"

    Lont = TimeDezToHMS(4.35808335) '  -4.358°              ' Observer’s longitude

    Dim Longitude As GeoCoord
    Longitude.deg = Lont.H
    Longitude.Min = Lont.M
    Longitude.Sec = Lont.s
    Longitude.Sign = "E"

    Dim Latitude As GeoCoord    '  50°47'55''                 ' Observer’s latitude
    Latitude.deg = 50
    Latitude.Min = 47
    Latitude.Sec = 55
    Latitude.Sign = "N"

    Dim Az As Double
    Dim Alt As Double
    Dim LocalHourAngleRad As Double
    Dim HourAngle As MyTime

    Dim RA_Saturn_Rad As Double
    RA_Saturn_Rad = TimeToRad(RA_Saturn)
    Dim DEC_Saturn_Rad As Double
    DEC_Saturn_Rad = DegToRad(GeoToDez(DEC_Saturn))

    RA_DEC_to_AZ_ALT_radian RA_Saturn_Rad, DEC_Saturn_Rad, Longitude, Latitude, SaturnDateTime, Az, Alt, LocalHourAngleRad

    If Mainform.O_OrientationNorth.Value Then Az = Az + Pi
    L_AzStar = CutAngle(RadToDeg(Az))
    L_AltStar = RadToDeg(Alt)

    HourAngle = RadToTime(LocalHourAngleRad)
    L_HourAngle = HourAngle.H & ":" & HourAngle.M & ":" & Format(HourAngle.s, "00.00")


 ' Capella Kassel
'''    Dim CapellaDateTime As Date
'''    CapellaDateTime = "2.2.2019 19:00:00"
'''
'''    Dim RA_Capella As MyTime
'''    RA_Capella.H = 5
'''    RA_Capella.M = 18
'''    RA_Capella.s = 6
'''
'''    Dim DEC_Capella As GeoCoord
'''    DEC_Capella.Deg = 46
'''    DEC_Capella.Min = 1
'''    DEC_Capella.Sec = 0
'''    DEC_Capella.Sign = "+"
'''
'''    Dim Longitude As GeoCoord                     ' Observer’s longitude
'''    Longitude.Deg = 9
'''    Longitude.Min = 18
'''    Longitude.Sec = 3
'''    Longitude.Sign = "E"
'''
'''    Dim Latitude As GeoCoord                     ' Observer’s latitude
'''    Latitude.Deg = 51
'''    Latitude.Min = 11
'''    Latitude.Sec = 27
'''    Latitude.Sign = "N"
'''
'''    Dim Az As Double
'''    Dim Alt As Double
'''    Dim LocalHourAngleRad As Double
'''    Dim HourAngle As MyTime
'''
'''    Dim RA_Capella_Rad As Double
'''    RA_Capella_Rad = TimeToRad(RA_Capella)
'''    Dim DEC_Capella_Rad As Double
'''    DEC_Capella_Rad = DegToRad(GeoToDez(DEC_Capella))
'''
'''    RA_DEC_to_AZ_ALT_radian RA_Capella_Rad, DEC_Capella_Rad, Longitude, Latitude, CapellaDateTime, Az, Alt, LocalHourAngleRad
'''
'''    If O_OrientationNorth.Value Then Az = Az + Pi
'''    L_AzStar = CutAngle(RadToDeg(Az))
'''    L_AltStar = RadToDeg(Alt)
'''
'''    HourAngle = RadToTime(LocalHourAngleRad)
'''    L_HourAngle = HourAngle.H & ":" & HourAngle.M & ":" & Format(HourAngle.s, "00.00")

 ' Deneb München
'''    Dim DenebDateTime As Date
'''    DenebDateTime = "2.2.2019 19:00:00"
'''
'''    Dim RA_Deneb As MyTime
'''    RA_Deneb.H = 20
'''    RA_Deneb.M = 42
'''    RA_Deneb.s = 4
'''
'''    Dim DEC_Deneb As GeoCoord
'''    DEC_Deneb.Deg = 45
'''    DEC_Deneb.Min = 21
'''    DEC_Deneb.Sec = 0
'''    DEC_Deneb.Sign = "+"
'''
'''    Dim Longitude As GeoCoord                     ' Observer’s longitude
'''    Longitude.Deg = 11
'''    Longitude.Min = 34
'''    Longitude.Sec = 0
'''    Longitude.Sign = "E"
'''
'''    Dim Latitude As GeoCoord                     ' Observer’s latitude
'''    Latitude.Deg = 48
'''    Latitude.Min = 8
'''    Latitude.Sec = 0
'''    Latitude.Sign = "N"
'''
'''    Dim Az As Double
'''    Dim Alt As Double
'''    Dim LocalHourAngleRad As Double
'''    Dim HourAngle As MyTime
'''
'''    Dim RA_Deneb_Rad As Double
'''    RA_Deneb_Rad = TimeToRad(RA_Deneb)
'''    Dim DEC_Deneb_Rad As Double
'''    DEC_Deneb_Rad = DegToRad(GeoToDez(DEC_Deneb))
'''
'''     RA_DEC_to_AZ_ALT_radian RA_Deneb_Rad, DEC_Deneb_Rad, Longitude, Latitude, DenebDateTime, Az, Alt, LocalHourAngleRad
'''
'''    If O_OrientationNorth.value Then Az = Az + Pi
'''    L_AzStar = CutAngle(RadToDeg(Az))
'''    L_AltStar = RadToDeg(Alt)
'''
'''    HourAngle = RadToTime(LocalHourAngleRad)
'''    L_HourAngle = HourAngle.H & ":" & HourAngle.M & ":" & Format(HourAngle.s, "00.00")



End Sub


Private Sub Command5_Click()
    Dim tst As Double
    Dim t1 As Double
    Dim t2 As Double
    
  
    
    tst = MotorIncrSystem_to_MatrixSystem(EncoderResolution / 4)
End Sub

Private Sub Command7_Click()
    Dim i As Long
    
    Dim dummy As Double
    Label1 = "1"
        For i = 1 To 10000000
            dummy = dummy * Pi
        Next i
    Label1 = "2"
        For i = 1 To 10000000
            dummy = dummy * Pi
        Next i
    Label1 = "3"
End Sub
