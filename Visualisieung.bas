Attribute VB_Name = "Visualisieung"
Dim ColorCoord As Long


Public Sub DispInit()
    
    GlbCx = 1
    GlbCY = -1
    GlbScale = 3000
End Sub


Public Sub DispCoordinateSystem()
    Dim sx As Double
    Dim sy As Double
    Dim ex As Double
    Dim ey As Double

    ColorCoord = vbRed

    mx = GlbCx * GlbScale
    my = GlbCY * -GlbScale
    Vis.Pic.Circle (mx, my), (90 / 90) * GlbScale, ColorCoord
    Vis.Pic.Circle (mx, my), (70 / 90) * GlbScale, ColorCoord
    Vis.Pic.Circle (mx, my), (50 / 90) * GlbScale, ColorCoord
    Vis.Pic.Circle (mx, my), (30 / 90) * GlbScale, ColorCoord
    Vis.Pic.Circle (mx, my), (10 / 90) * GlbScale, ColorCoord

    sx = (GlbCx - 0.05) * GlbScale
    sy = GlbCY * -GlbScale
    ex = (GlbCx + 0.05) * GlbScale
    ey = GlbCY * -GlbScale
    Vis.Pic.Line (sx, sy)-(ex, ey), ColorCoord

    sx = GlbCx * GlbScale
    sy = (GlbCY - 0.05) * -GlbScale
    ex = GlbCx * GlbScale
    ey = (GlbCY + 0.05) * -GlbScale
    Vis.Pic.Line (sx, sy)-(ex, ey), ColorCoord

    ' Draw Polaris
    Dim Polaris As AzAlt
    Polaris.Az = 0
    Polaris.Alt = 0.5458
    
    ' Rotate -90°: North is below
    Polaris.Az = Polaris.Az - Pi / 2
   
    mx = (Cos(Polaris.Az) * (1 - Polaris.Alt) + GlbCx) * GlbScale
    my = (Sin(Polaris.Az) * (1 - Polaris.Alt) + GlbCY) * -GlbScale
    Vis.Pic.Circle (mx, my), 50, ColorCoord
    
End Sub


Public Sub DispTelescopePos(Polar As AzAlt)
'    For i = 100 To 1000
'        Vis.Pic.PSet (i, 4560), vbRed
'    Next i

    Dim mx As Double
    Dim my As Double
    Dim Center As AzAlt
    
    Static LastCenter As AzAlt
    ' Overwrite last circle with white
    Vis.Pic.Circle (LastCenter.Az, LastCenter.Alt), 20, vbGreen
    
    
    ' Rotate -90°: North is below
    Polar.Az = Polar.Az - Pi / 2
    
    ' convert 0..Pi/2 to 0..1
    Polar.Alt = Polar.Alt * (1 / (Pi / 2))
    
    Center.Az = (Cos(Polar.Az) * (1 - Polar.Alt) + GlbCx) * GlbScale
    Center.Alt = (Sin(Polar.Az) * (1 - Polar.Alt) + GlbCY) * -GlbScale
    Vis.Pic.Circle (Center.Az, Center.Alt), 20, vbBlue

    LastCenter = Center

End Sub

Public Sub DispAlignmentStar(Polar As AzAlt)
    Dim mx As Double
    Dim my As Double
    Dim Center As AzAlt
    
    Static LastCenter As AzAlt
    ' Overwrite last circle with white
    Vis.Pic.Circle (LastCenter.Az, LastCenter.Alt), 50, vbWhite
     
    ' Rotate -90°: North is below
    Polar.Az = Polar.Az - Pi / 2
    
     ' convert 0..Pi/2 to 0..1
    Polar.Alt = Polar.Alt * (1 / (Pi / 2))
    
   
    Center.Az = (Cos(Polar.Az) * (1 - Polar.Alt) + GlbCx) * GlbScale
    Center.Alt = (Sin(Polar.Az) * (1 - Polar.Alt) + GlbCY) * -GlbScale
    Vis.Pic.Circle (Center.Az, Center.Alt), 50, vbCyan

    LastCenter = Center

End Sub
