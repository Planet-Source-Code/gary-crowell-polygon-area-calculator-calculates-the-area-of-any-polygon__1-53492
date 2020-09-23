Attribute VB_Name = "Module"
Option Explicit
Public X      As Integer    'Our counter variables
Public Type Coordinate
    X           As Single
    Y           As Single
End Type
Public c()      As Coordinate
Public centroid As Coordinate

Public Function CalculateArea(GridSpacing As Integer, _
                              ByVal GridOffset As Integer) As Double
  Dim nextX    As Integer
  Dim Area     As Double

    'original area calculation - limited to convex polygon only
    'Calculates the area of the polygon
    '    'First, break up the polygon into triangles.  To find the length of the sides
    '    'of the triangle, use the distance formula.  To find the area of the triangles, use
    '    'the triangle area formula.  Add all of the areas of the triangles together to get the
    '    'area of the polygon.
    '    Dim Side1 As Double, Side2 As Double, Side3 As Double 'Sides of the triangle
    '    Dim Temp As Double
    '    Dim p As Double 'Perimeter of the triangle
    '
    '    'Loop through all of the sides of the triangle, except the first one and the last one
    '    For X = 2 To UBound(c) - 1
    '        Side1 = Sqr((c(1).X - c(X).X) ^ 2 + (c(1).Y - c(X).Y) ^ 2) / GridSpacing
    '        Side2 = Sqr((c(X).X - c(X + 1).X) ^ 2 + (c(X).Y - c(X + 1).Y) ^ 2) / GridSpacing
    '        Side3 = Sqr((c(X + 1).X - c(1).X) ^ 2 + (c(X + 1).Y - c(1).Y) ^ 2) / GridSpacing
    '
    '        p = Side1 + Side2 + Side3
    '        Area = Area + 0.25 * Sqr(p * (p - 2 * Side1) * (p - 2 * Side2) * (p - 2 * Side3))
    '    Next
    '    CalculateArea = Round(Area, 10)
    
    
    'see the following page for a good explanation of this area calculation
    'http://astronomy.swin.edu.au/~pbourke/geometry/polyarea/
    '   the polygon must be closed,
    '   it must be non self-intersecting.
    'loop through all vertices of the polygon
    Area = 0
    With centroid
        .X = 0
        .Y = 0
        'traverse the vertices in order and calculate the area.
        For X = 1 To UBound(c)
            nextX = X + 1
            If nextX > UBound(c) Then
                nextX = 1
            End If
            Area = Area + 0.5 * (c(X).X * c(nextX).Y - c(nextX).X * c(X).Y)
        Next X
        'traverse the vertices again to calculate the centroid.
        For X = 1 To UBound(c)
            nextX = X + 1
            If nextX > UBound(c) Then
                nextX = 1
            End If
            'if the area is negative, the polygon was traversed CCW, if positive it was
            'traversed CW, the appropriate centroid equations depend on the traverse order
            If Area < 0 Then
                .X = .X - (c(X).X + c(nextX).X) * (c(X).X * c(nextX).Y - c(nextX).X * c(X).Y)
                .Y = .Y - (c(X).Y + c(nextX).Y) * (c(X).X * c(nextX).Y - c(nextX).X * c(X).Y)
            Else
                .X = .X + (c(X).X + c(nextX).X) * (c(X).X * c(nextX).Y - c(nextX).X * c(X).Y)
                .Y = .Y + (c(X).Y + c(nextX).Y) * (c(X).X * c(nextX).Y - c(nextX).X * c(X).Y)
            End If
        Next X
        Area = Abs(Area)
        'convert the area into the grid units
        CalculateArea = Round(Area / (GridSpacing * GridSpacing), 10)
        If Area > 0 Then
            .X = .X / (6 * Area)
            .Y = .Y / (6 * Area)
        End If
        'convert to the grid coordinates (just leave them in the form coordinates for display)
        '.X = .X / GridSpacing - GridOffset
        '.Y = -(.Y / GridSpacing - GridOffset)
        Debug.Print .X, .Y
    End With 'centroid
End Function


