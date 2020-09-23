VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Polygon Area Calculator"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4515
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   301
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCoordinateList 
      Caption         =   "Coordinate LIst"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   4755
      Width           =   1455
   End
   Begin VB.OptionButton o4 
      Caption         =   "1/5"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   4920
      Width           =   615
   End
   Begin VB.OptionButton o3 
      Caption         =   "1/4"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   4920
      Width           =   615
   End
   Begin VB.OptionButton o2 
      Caption         =   "1/2"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   4920
      Width           =   615
   End
   Begin VB.OptionButton o1 
      Caption         =   "1"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   4920
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   15
      Left            =   2160
      ScaleHeight     =   15
      ScaleWidth      =   375
      TabIndex        =   1
      Top             =   4920
      Width           =   375
   End
   Begin VB.Line l 
      BorderColor     =   &H00C000C0&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   208
      X2              =   160
      Y1              =   168
      Y2              =   168
   End
   Begin VB.Label Label1 
      Caption         =   "Snap to"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label lblArea 
      AutoSize        =   -1  'True
      Caption         =   "Area:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   660
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   304
      Y1              =   301
      Y2              =   301
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const GridSpacing      As Integer = 30


Private Sub AlignToGrid(X As Single, _
                        Y As Single, _
                        ByVal GridSpacing As Single)
  'Align X-Coordinate

    X = X / GridSpacing
    X = Round(X)
    X = X * GridSpacing
    'Align Y-Coordinate
    Y = Y / GridSpacing
    Y = Round(Y)
    Y = Y * GridSpacing
End Sub

Private Sub cmdCoordinateList_Click()
    frmCoordinates.Show vbModal, Me
End Sub

Private Sub Form_Load()

    Me.DrawWidth = 1
    'Draws the grid
    For X = 0 To 300 Step GridSpacing
        'Draws horizontal lines
        Line (0, X)-(300, X), RGB(128, 128, 128) 'RGB(128, 128, 128) is a 50% gray
        'Draws vertical lines
        Line (X, 0)-(X, 300), RGB(128, 128, 128)
    Next '  X
    'Darkens the axis
    Line (0, 150)-(300, 150), RGB(0, 0, 0) 'RGB(0, 0, 0) is black
    Line (150, 0)-(150, 300), RGB(0, 0, 0) 'RGB(0, 0, 0) is black
    'Draws the arrows
    Line (140, 10)-(150, 0), RGB(0, 0, 0) '}__Up
    Line (150, 0)-(160, 10), RGB(0, 0, 0) '}  Arrow
    Line (10, 140)-(0, 150), RGB(0, 0, 0) '}__Left
    Line (0, 150)-(10, 160), RGB(0, 0, 0) '}  Arrow
    Line (140, 290)-(150, 300), RGB(0, 0, 0) '}__Down
    Line (150, 300)-(160, 290), RGB(0, 0, 0) '}  Arrow
    Line (290, 140)-(300, 150), RGB(0, 0, 0) '}__Right
    Line (300, 150)-(290, 160), RGB(0, 0, 0) '}  Arrow
    'Labels the axis
    Me.CurrentX = 293
    Me.CurrentY = 155
    Print "x"      '}
    Me.CurrentX = 155
    Me.CurrentY = 7
    Print "y"      '}
    'Labels the origin
    Me.CurrentX = 142
    Me.CurrentY = 150
    Print "0"
    'Labels the intervals
    Me.CurrentX = 172
    Me.CurrentY = 150
    Print "1"
    Me.CurrentX = 142
    Me.CurrentY = 120
    Print "1"
    'Resets the coordinate array
    ReDim c(0)
    Me.DrawWidth = 2

End Sub

Private Sub Form_MouseDown(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)

  'Aligns the mouse coordinates to the grid
  
  Dim Area As Double
  Dim Snap As Single
    Select Case True
     Case o1.Value
        Snap = 1
     Case o2.Value
        Snap = 0.5
     Case o3.Value
        Snap = 0.25
     Case o4.Value
        Snap = 0.2
    End Select
    AlignToGrid X, Y, GridSpacing * Snap
    If Y > 300 Then
        Exit Sub
    End If
    If lblArea.Caption <> "Area:" Then
        'Clears the grid
        Cls
        Form_Load
        lblArea.Caption = "Area:"
     ElseIf UBound(c) >= 3 Then 'NOT LBLAREA.CAPTION...
        If X = c(1).X Then
            If Y = c(1).Y Then
                'Highlight the shape in blue
                For X = 1 To UBound(c)
                    If X = UBound(c) Then
                        Line (c(1).X, c(1).Y)-(c(X).X, c(X).Y), RGB(0, 0, 255) 'RGB(0, 0, 192) is blue
                     Else 'NOT X...
                        Line (c(X).X, c(X).Y)-(c(X + 1).X, c(X + 1).Y), RGB(0, 0, 255) 'RGB(0, 0, 192) is blue
                    End If
                Next X
                'Calculates the area
                Area = CalculateArea(30, 5)
                'Displays the area
                lblArea.Caption = "Area: " & Area
                'draw an 'x' at the centroid
                Line (centroid.X - 10, centroid.Y)-(centroid.X + 10, centroid.Y), RGB(0, 0, 255)
                Line (centroid.X, centroid.Y - 10)-(centroid.X, centroid.Y + 10), RGB(0, 0, 255)
                Exit Sub
            End If
        End If
    End If
    If UBound(c) = 0 Then
        PSet (X, Y), RGB(192, 0, 0)
    End If
    ReDim Preserve c(UBound(c) + 1)
    c(UBound(c)).X = X
    c(UBound(c)).Y = Y
    If UBound(c) > 1 Then
        Line (c(UBound(c) - 1).X, c(UBound(c) - 1).Y)-(c(UBound(c)).X, c(UBound(c)).Y), RGB(192, 0, 0)
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)

  'Aligns the mouse coordinates to the grid
  
  Dim Snap As Single

    Select Case True
     Case o1.Value
        Snap = 1
     Case o2.Value
        Snap = 0.5
     Case o3.Value
        Snap = 0.25
     Case o4.Value
        Snap = 0.2
    End Select
    AlignToGrid X, Y, GridSpacing * Snap
    If Y > 300 Then
        Exit Sub
    End If
    With l
        .Visible = False
        .X2 = X
        .Y2 = Y
        .BorderColor = RGB(192, 0, 192) 'Purple line
    End With 'l
    If lblArea.Caption <> "Area:" Or UBound(c) = 0 Then
        Exit Sub
     ElseIf UBound(c) >= 3 Then 'NOT LBLAREA.CAPTION...
        If X = c(1).X Then
            If Y = c(1).Y Then
                'Changes the line color to blue
                l.BorderColor = RGB(0, 0, 192)
            End If
        End If
    End If
    'Draws a purple line
    With l
        .Visible = True
        .X1 = c(UBound(c)).X
        .Y1 = c(UBound(c)).Y
    End With 'l

End Sub


