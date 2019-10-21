Attribute VB_Name = "Math_old"
Option Explicit

Public Sub RA_DEC_to_AZ_ALT_radian(RA_Star_Rad As Double, DEC_Star_Rad As Double, Longitude As GeoDegMinSec, Latitude As GeoDegMinSec, LocalDateTime As Date, Az As Double, Alt As Double, LocalHourAngleRad As Double)
    ' matrix_method_rev_d.pdf Seite 15
    Dim SiderialTime As MyTime
    Dim SiderialTimeRad As Double
    Dim LongitudeDeg As Double
    Dim LongitudeRad As Double
    Dim LocalDate As MyDate
    Dim LocalTimeUT As MyTime

    LongitudeDeg = GeoToDez(Longitude)
    LongitudeRad = DegToRad(LongitudeDeg)
       
    'Calculate siderial time at Greenwich
    LocalDate.YY = Year(LocalDateTime)
    LocalDate.MM = Month(LocalDateTime)
    LocalDate.DD = Day(LocalDateTime)
    LocalTimeUT.H = Hour(LocalDateTime)
    LocalTimeUT.M = Minute(LocalDateTime)
    LocalTimeUT.s = Second(LocalDateTime)

    SiderialTime = GMST(LocalDate, LocalTimeUT)
    SiderialTimeRad = TimeToRad(SiderialTime)
    
    ' Calculate local hour angle
    LocalHourAngleRad = SiderialTimeRad - RA_Star_Rad
    LocalHourAngleRad = LocalHourAngleRad - LongitudeRad
    
    ' Calculate star position in rectangular equatorial coordinate system
    Dim LMN_Equatorial As Vector         ' Rectangular equatorial coordinate system
    Dim LMN_EquaMatrix(10, 10) As Double

    LMN_Equatorial = PolarKarthesisch(LocalHourAngleRad, DEC_Star_Rad)
   
    LMN_EquaMatrix(0, 0) = LMN_Equatorial.X
    LMN_EquaMatrix(1, 0) = LMN_Equatorial.Y
    LMN_EquaMatrix(2, 0) = LMN_Equatorial.z
   
    'Calculate star position in rectangular horizontal coordinate system
    Dim LMN_Horizontal As Vector                ' Rectangular horizontal coordinate system
    Dim LMN_HorizMatrix(10, 10) As Double
    Dim TransformationMatrix(10, 10) As Double  ' Transformation-Matrix from equatorial Coordinates to horizontal Coordinates
    Dim LatitudeDeg As Double                   ' Observer’s latitude
    Dim LatitudeRad As Double
   
    LatitudeDeg = GeoToDez(Latitude)
    LatitudeRad = DegToRad(LatitudeDeg)
   
    TransformationMatrix(0, 0) = Cos(LatitudeRad - Pi / 2):   TransformationMatrix(0, 1) = 0:   TransformationMatrix(0, 2) = Sin(LatitudeRad - Pi / 2)
    TransformationMatrix(1, 0) = 0:                           TransformationMatrix(1, 1) = 1:   TransformationMatrix(1, 2) = 0
    TransformationMatrix(2, 0) = -Sin(LatitudeRad - Pi / 2):  TransformationMatrix(2, 1) = 0:   TransformationMatrix(2, 2) = Cos(LatitudeRad - Pi / 2)

    MatrixProduct TransformationMatrix, 3, 3, LMN_EquaMatrix, 3, 1, LMN_HorizMatrix

    Dim Lh As Double
    Dim Mh As Double
    Dim Nh As Double
    
    Lh = LMN_HorizMatrix(0, 0)                  ' Rectangular horizontal  coordinate system
    Mh = LMN_HorizMatrix(1, 0)
    Nh = LMN_HorizMatrix(2, 0)
    
    'Calculate Star position in Altazimuth horizontal coordinate system
    Dim sin_h As Double
'    When Lh >= 0, (-A) is in the 1st quadrant or the 4th quadrant.
'    When Lh < 0, (-A) is in the 2nd quadrant or the 3rd quadrant.
    Az = -Atn(Mh / Lh)
    If Lh < 0 Then
        Az = Az + Pi
    End If
    
    ' Standard: North = 0° so add 180°
     Az = Az + Pi
    
    
    'geht möglicherweise einfacher: sin(h) = Nh
    sin_h = Cos(LatitudeRad) * Cos(LocalHourAngleRad) * Cos(DEC_Star_Rad) + Sin(LatitudeRad) * Sin(DEC_Star_Rad)
    Alt = arcsin(sin_h)

End Sub

