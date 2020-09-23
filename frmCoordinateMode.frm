VERSION 5.00
Begin VB.Form frmCoordinateMode 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Area Calculator"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   185
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Clear"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calculate Area"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0C0C0&
      Height          =   1425
      ItemData        =   "frmCoordinateMode.frx":0000
      Left            =   0
      List            =   "frmCoordinateMode.frx":0002
      TabIndex        =   8
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Y 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox X 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblArea 
      AutoSize        =   -1  'True
      Caption         =   "Area:"
      Height          =   195
      Left            =   0
      TabIndex        =   10
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Coordinates:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   ")"
      Height          =   195
      Left            =   1320
      TabIndex        =   5
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   ","
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "("
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label1 
      Caption         =   "New Coordinate:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmCoordinateMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error GoTo Handler
    ReDim Preserve c(UBound(c) + 1) As Coordinate
    c(UBound(c)).X = X
    c(UBound(c)).Y = Y
    List1.AddItem "(" & c(UBound(c)).X & ", " & c(UBound(c)).Y & ")"
    Exit Sub
Handler:
End Sub

Private Sub Command2_Click()
    lblArea = "Area: " & CalculateArea(1)
End Sub

Private Sub Command3_Click()
    ReDim c(0)
    List1.Clear
End Sub

