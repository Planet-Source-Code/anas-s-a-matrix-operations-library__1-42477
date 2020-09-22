VERSION 5.00
Begin VB.Form mnFrm 
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7800
   Icon            =   "mnFrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSolution 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   4800
      Width           =   5895
   End
   Begin VB.CommandButton CalButton 
      Caption         =   "Calculate"
      Height          =   1095
      Left            =   6240
      TabIndex        =   1
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox txtDisplay 
      Height          =   4575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
   Begin VB.Frame FrmSelect 
      Height          =   6615
      Left            =   6120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   25
         Top             =   4080
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   13
         Top             =   3720
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   12
         Top             =   3360
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   " Det (A)"
         Height          =   375
         Left            =   480
         TabIndex        =   24
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   " A x Tr (B) + Inv(A)"
         Height          =   375
         Left            =   480
         TabIndex        =   23
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   " B x Inv (B)"
         Height          =   375
         Left            =   480
         TabIndex        =   22
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "  V2 / 3 "
         Height          =   375
         Left            =   480
         TabIndex        =   21
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   " 5 A"
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   " | V1 |"
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   " V1 x V2"
         Height          =   375
         Left            =   480
         TabIndex        =   18
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Transpose(B)"
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Inverse(A)"
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   " A - B"
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   " A + B"
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "mnFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''------------------------------------------------------------------------
''
'' Author      : Anas S. A.
'' Email       : e106714@metu.edu.tr
'' Date        : 18 Jan 2002
'' Version     : 1.0
'' Description : Matrix Operations Library
''
''------------------------------------------------------------------------

'' The cMathLib class contains many operations related to matrix calculations
'' like add, subtract,inverse, det,etc... Inorder to have this library in any project
'' just add the cMathLib class and then define it with a variable in your program
'' The matrices dimensions can be anything. The code is highly optimized for fast
'' operation.
'' The following shows a demonstration on how to use the library.


Option Explicit

Dim Mat As New cMathLib 'Define a variable as the matrix library class

Dim A(3, 3) As Double 'define a matrix A with dimensions (4x4)
Dim B(3, 3) As Double 'define a matrix B with dimensions (4x4)

Dim V1(2) As Double 'define a vector V1
Dim V2(2) As Double 'define a vector V2

Private Sub CalButton_Click()
 'Assign a dynamic matrix C
 Dim C() As Double
 
 If Option1(0).Value = True Then
  'Addition Case
  C = Mat.Add(A, B) 'C = Addition of A and B
  
  txtSolution = "Answer A + B = " & vbCrLf & vbCrLf
  txtSolution = txtSolution + Mat.PrintMat(C)   'Print C
 
 ElseIf Option1(1).Value = True Then
  'Subtraction Case
  
  C = Mat.Subtract(A, B) 'C = Subtration of A from B
  'Print C
  txtSolution = "Answer A - B = " & vbCrLf & vbCrLf
  txtSolution = txtSolution + Mat.PrintMat(C)   'Print C
 
 ElseIf Option1(2).Value = True Then
  'Determinant Case of Matrix A
  Dim Determinant As Double
  Determinant = Mat.Det(A)
  
  txtSolution = "Answer Determinant of A = " & vbCrLf & vbCrLf
  txtSolution = txtSolution & Determinant
  
 ElseIf Option1(3).Value = True Then
  'Inverse Case of Matrix B
  C = Mat.Inv(A)
  
  txtSolution = "Answer Inverse of A = " & vbCrLf & vbCrLf
  txtSolution = txtSolution & Mat.PrintMat(C)

 ElseIf Option1(4).Value = True Then
  'Transpose Case of Matrix B
  C = Mat.Transpose(B)
  
  txtSolution = "Answer Transpose of B = " & vbCrLf & vbCrLf
  txtSolution = txtSolution & Mat.PrintMat(C)
  
 ElseIf Option1(5).Value = True Then
  'Mutiply Vectors V1 and V2 Case
  C = Mat.MultiplyVectors(V1, V2)
  
  txtSolution = "Answer V1 x V2 = " & vbCrLf & vbCrLf
  txtSolution = txtSolution & Mat.PrintMat(C)
  
 ElseIf Option1(6).Value = True Then
  'Magnitude of Vector V1 Case
  Dim Magnitude As Double
  Magnitude = Mat.VectorMagnitude(V1)
  
  txtSolution = "Answer |V1| = " & vbCrLf & vbCrLf
  txtSolution = txtSolution & Magnitude
  
 ElseIf Option1(7).Value = True Then
  'Scalar Multiply 5*A Case
  C = Mat.ScalarMultiply(5, A)
  txtSolution = "Answer 5A = " & vbCrLf & vbCrLf
  txtSolution = txtSolution & Mat.PrintMat(C)
  
 ElseIf Option1(8).Value = True Then
  'Scalar Divide V2/3 Case
  C = Mat.ScalarDivide(3, V2)
  txtSolution = "Answer V2 / 3 = " & vbCrLf & vbCrLf
  txtSolution = txtSolution & Mat.PrintMat(C)

 ElseIf Option1(9).Value = True Then
  'Case Bxinv(B)
  C = Mat.Multiply(A, Mat.Inv(A))
  txtSolution = "Answer B x Inverse(B) = " & vbCrLf & vbCrLf
  txtSolution = txtSolution & Mat.PrintMat(C)

 ElseIf Option1(10).Value = True Then
  'Case AxTranspose(B)+inv(A)
  C = Mat.Add(Mat.Multiply(A, Mat.Transpose(B)), Mat.Inv(A))
  txtSolution = "Answer A x Transpose(B)+Inverse(A) = " & vbCrLf & vbCrLf
  txtSolution = txtSolution & Mat.PrintMat(C)
 End If
End Sub

Private Sub Form_Load()
 Dim i As Integer, j As Integer
 
 'Assign some values to matrix A and B
 A(0, 0) = 1: A(0, 1) = 2: A(0, 2) = 3: A(0, 3) = 4
 A(1, 0) = 5: A(1, 1) = 6: A(1, 2) = 7: A(1, 3) = 8
 A(2, 0) = 9: A(2, 1) = 10: A(2, 2) = 1: A(2, 3) = 12
 A(3, 0) = 13: A(3, 1) = -14: A(3, 2) = 15: A(3, 3) = 16

 For i = 0 To 3
  For j = 0 To 3
   B(i, j) = 2 * Rnd(1)
  Next j
 Next i
 
 'Assign some values to vectors V1 and V2
 V1(0) = 1: V1(1) = 2: V1(2) = 3
 V2(0) = 4: V2(1) = 5: V2(2) = 6
 
'Print Matrices And Vectors
 txtDisplay = "Matix A = " & vbCrLf
 txtDisplay = txtDisplay + Mat.PrintMat(A) & vbCrLf & vbCrLf
   
 txtDisplay = txtDisplay + "Matix B = " & vbCrLf
 txtDisplay = txtDisplay + Mat.PrintMat(B) & vbCrLf & vbCrLf
   
 txtDisplay = txtDisplay + "Vector V1 = " & vbCrLf
 txtDisplay = txtDisplay + Mat.PrintMat(V1) & vbCrLf & vbCrLf

 txtDisplay = txtDisplay + "Vector V2 = " & vbCrLf
 txtDisplay = txtDisplay + Mat.PrintMat(V2)
End Sub

