VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMathLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' The dimensions of the matrix are checked
' Here
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function Find_R_C(Mat() As Double) As Double()
 Dim Rows As Integer, Columns As Integer
 Dim i As Integer, j As Integer
 Dim Result() As Double
 Columns = 0
 If Mat_1D(Mat, Rows) Then
  ReDim Result(Rows, 1)
  Result(0, 0) = Rows
  Result(0, 1) = Columns + 1
 
  For i = 1 To Rows
   Result(i, 1) = Mat(i - 1)
  Next i
 Else
  Call Mat_2D(Mat, Rows, Columns)
  ReDim Result(Rows, Columns)
  Result(0, 0) = Rows
  Result(0, 1) = Columns
 
  For i = 1 To Rows
   For j = 1 To Columns '- 1
    Result(i, j) = Mat(i - 1, j - 1)
   Next j
  Next i
 End If
 Find_R_C = Result
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Check if matrix has only one column
' shift the matrix one level and keep
' its dimensions details in Mat(0,0) and Mat(0,1)
' Mat(0,0)= no of rows
' Mat(0,1)= no of columns
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function Mat_1D(Mat() As Double, m As Integer) As Boolean
 Dim Temp_MAT As Double
 On Error GoTo Error_Handler
 Temp_MAT = Mat(0, 0)
 Mat_1D = False
 Exit Function
Error_Handler:
 Mat_1D = True
 m = UBound(Mat) + 1
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Check if matrix has more than one column
' if so return the dimension as described above
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Mat_2D(Mat() As Double, m As Integer, n As Integer)
 Dim Temp_MAT As Double, i As Integer
 i = 0
 m = UBound(Mat) + 1
 On Error GoTo Error_Handler
 Do Until i < -1
  Temp_MAT = Mat(0, i)
  i = i + 1
 Loop
Error_Handler:
  n = i
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Add two matrices, their dimensions should be compatible!
' Function returns the summation or errors due to
' dimensions incompatibility
' Example:
'  Check Main Form !!
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Add(Mat_1() As Double, Mat_2() As Double) As Double()
 Dim Mat1() As Double, Mat2() As Double
 Dim sol() As Double
 Dim i As Integer, j As Integer
 
 On Error GoTo Error_Handler
 
 Mat1 = Find_R_C(Mat_1)
 Mat2 = Find_R_C(Mat_2)
 
 If Mat1(0, 0) <> Mat2(0, 0) Or Mat1(0, 1) <> Mat2(0, 1) Then
  GoTo Error_Dimension
 End If
 
 ReDim sol(Mat1(0, 0) - 1, Mat1(0, 1) - 1)
  For i = 1 To Mat1(0, 0)
   For j = 1 To Mat1(0, 1)
    sol(i - 1, j - 1) = Mat1(i, j) + Mat2(i, j)
   Next j
  Next i

  Add = sol
  Erase sol
 
 Exit Function
  
Error_Dimension:
 Err.Raise "5005", , "Dimensions of the two matrices do not match !"

Error_Handler:
 If Err.Number = 5005 Then
  Err.Raise "5005", , "Dimensions of the two matrices do not match !"
 Else
  Err.Raise "5022", , "One or both of the matrices are null, this operation cannot be done !!"
 End If

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Subtracts two matrices from each other, their
' dimensions should be compatible!
' Function returns the solution or errors due to
' dimensions incompatibility
' Example:
'  Check Main Form !!
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Subtract(Mat_1() As Double, Mat_2() As Double) As Double()
 Dim Mat1() As Double, Mat2() As Double
 Dim i As Integer, j As Integer, sol() As Double
 
 On Error GoTo Error_Handler
  
 Mat1 = Find_R_C(Mat_1)
 Mat2 = Find_R_C(Mat_2)
  
 If Mat1(0, 0) <> Mat2(0, 0) Or Mat1(0, 1) <> Mat2(0, 1) Then
  GoTo Error_Dimension
 End If
 
 ReDim sol(Mat1(0, 0) - 1, Mat1(0, 1) - 1)
   
 For i = 1 To Mat1(0, 0)
  For j = 1 To Mat1(0, 1)
   sol(i - 1, j - 1) = Mat1(i, j) - Mat2(i, j)
  Next j
 Next i

 Subtract = sol
 Erase sol
 Exit Function
  
Error_Dimension:
 Err.Raise "5007", , "Dimensions of the two matrices do not match !"
  
Error_Handler:
 If Err.Number = 5007 Then
  Err.Raise "5007", , "Dimensions of the two matrices do not match !"
 Else
  Err.Raise "5022", , "One or both of the matrices are null, this operation cannot be done !!"
 End If

 End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Multiply two matrices, their dimensions should be compatible!
' Function returns the solution or errors due to
' dimensions incompatibility
' Example:
'  Check Main Form !!
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Multiply(Mat_1() As Double, Mat_2() As Double) As Double()
 Dim Mat1() As Double, Mat2() As Double, l As Integer
 Dim i As Integer, j As Integer, OptiString As String
 Dim sol() As Double, MulAdd As Double
 
 On Error GoTo Error_Handler
 
 MulAdd = 0

 Mat1 = Find_R_C(Mat_1)
 Mat2 = Find_R_C(Mat_2)
 
 If Mat1(0, 1) <> Mat2(0, 0) Then
  GoTo Error_Dimension
 End If
 
 ReDim sol(Mat1(0, 0) - 1, Mat2(0, 1) - 1)
 
 For i = 1 To Mat1(0, 0)
  For j = 1 To Mat2(0, 1)
   For l = 1 To Mat1(0, 1)
    MulAdd = MulAdd + Mat1(i, l) * Mat2(l, j)
   Next l
   sol(i - 1, j - 1) = MulAdd
   MulAdd = 0
  Next j
 Next i
  
 Multiply = sol
 Erase sol
 Exit Function
  
Error_Dimension:
 Err.Raise "5009", , "Dimensions of the two matrices not suitable for multiplication !"

Error_Handler:
 If Err.Number = 5009 Then
  Err.Raise "5009", , "Dimensions of the two matrices not suitable for multiplication !"
 Else
  Err.Raise "5022", , "One or both of the matrices are null, this operation cannot be done !!"
 End If

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Determinant of a matrix should be (nxn)
' Function returns the solution or errors due to
' dimensions incompatibility
' Example:
'  Check Main Form !!
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Det(Mat() As Double) As Double
 Dim DArray() As Double, S As Integer
 Dim k As Integer, i As Integer, j As Integer
 Dim save As Double, ArrayK As Double, k1 As Integer
 Dim M1 As String, Mat1() As Double
 
 On Error GoTo Error_Handler
 
 Mat1 = Find_R_C(Mat)
 
 If Mat1(0, 0) <> Mat1(0, 1) Then GoTo Error_Dimension
 
 S = Mat1(0, 0)
 Det = 1
 DArray = Mat1()
  
For k = 1 To S
 If DArray(k, k) = 0 Then
  j = k
  Do While ((j < S) And (DArray(k, j) = 0))
    j = j + 1
  Loop
    If DArray(k, j) = 0 Then
     Det = 0
     Exit Function
    Else
     For i = k To S
      save = DArray(i, j)
      DArray(i, j) = DArray(i, k)
      DArray(i, k) = save
     Next i
    End If
    
    Det = -Det
   End If
  ArrayK = DArray(k, k)
  Det = Det * ArrayK
  If k < S Then
   k1 = k + 1
   For i = k1 To S
    For j = k1 To S
     DArray(i, j) = DArray(i, j) - DArray(i, k) * (DArray(k, j) / ArrayK)
    Next j
   Next i
  End If
  Next
  
Exit Function
   
Error_Dimension:
 Err.Raise "5011", , "Matrix should be a square matrix !"
  
Error_Handler:
 If Err.Number = 5011 Then
  Err.Raise "5011", , "Matrix should be a square matrix !"
 Else
  Err.Raise "5022", , "In order to do this operation values must be assigned to the matrix !!"
 End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Inverse of a matrix, should be (nxn) and det(Mat)<>0
' Function returns the solution or errors due to
' dimensions incompatibility
' Example:
'  Check Main Form !!
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Inv(Mat() As Double) As Double()
  Dim sol() As Double
  Dim AI() As Double, AIN As Double, AF As Double, _
           Mat1() As Double
  Dim LL As Integer, LLM As Integer, L1 As Integer, _
      L2 As Integer, LC As Integer, LCA As Integer, _
      LCB As Integer, i As Integer, j As Integer
 
  On Error GoTo Error_Handler
  
  Mat1 = Find_R_C(Mat)
  If Mat1(0, 0) <> Mat1(0, 1) Then GoTo Error_Dimension
 
  If Det(Mat1) = 0 Then GoTo Error_Zero
  ReDim sol(Mat1(0, 0) - 1, Mat1(0, 0) - 1)

  LL = Mat1(0, 0)
  LLM = Mat1(0, 1)
  ReDim AI(LL, LL)

  For L2 = 1 To LL
   For L1 = 1 To LL
    AI(L1, L2) = 0
   Next
   AI(L2, L2) = 1
  Next

  For LC = 1 To LL
   If Abs(Mat1(LC, LC)) < 0.0000000001 Then
    For LCA = LC + 1 To LL
     If LCA = LC Then GoTo 1090
     If Abs(Mat1(LC, LCA)) > 0.0000000001 Then
      For LCB = 1 To LL
       Mat1(LCB, LC) = Mat1(LCB, LC) + Mat1(LCB, LCA)
       AI(LCB, LC) = AI(LCB, LC) + AI(LCB, LCA)
      Next
      GoTo 1100
     End If
1090 Next
    End If

1100
    AIN = 1 / Mat1(LC, LC)
    For LCA = 1 To LL
     Mat1(LCA, LC) = AIN * Mat1(LCA, LC)
     AI(LCA, LC) = AIN * AI(LCA, LC)
    Next

    For LCA = 1 To LL
     If LCA = LC Then GoTo 1150
     AF = Mat1(LC, LCA)
     For LCB = 1 To LL
      Mat1(LCB, LCA) = Mat1(LCB, LCA) - AF * Mat1(LCB, LC)
      AI(LCB, LCA) = AI(LCB, LCA) - AF * AI(LCB, LC)
     Next
1150 Next

  Next
   
   For i = 1 To LL
    For j = 1 To LL
     sol(i - 1, j - 1) = AI(i, j)
    Next j
   Next i
  
 Inv = sol
 Erase sol
 
 Exit Function
  
Error_Zero:
 Err.Raise "5012", , "Determinent equals zero, inverse can't be found !"
 
Error_Dimension:
 Err.Raise "5014", , "Matrix should be a square matrix !"
 
Error_Handler:
If Err.Number = 5012 Then
  Err.Raise "5012", , "Determinent equals zero, inverse can't be found !"
ElseIf Err.Number = 5014 Then
 Err.Raise "5014", , "Matrix should be a square matrix !"
End If
 
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Multiply two vectors, dimensions should be (3x1)
' Function returns the solution or errors due to
' dimensions incompatibility
' Example:
'  Check Main Form !!
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MultiplyVectors(Mat_1() As Double, Mat_2() As Double) As Double()
 Dim Mat1() As Double, Mat2() As Double
 Dim i As Double, j As Double, k As Double
 Dim sol(2) As Double
 
 On Error GoTo Error_Handler
 
 Mat1 = Find_R_C(Mat_1)
 Mat2 = Find_R_C(Mat_2)
 
 If Mat1(0, 0) <> 3 Or Mat1(0, 1) <> 1 Then
  GoTo Error_Dimension
 End If
 
 If Mat2(0, 0) <> 3 Or Mat2(0, 1) <> 1 Then
  GoTo Error_Dimension
 End If
 
 i = Mat1(2, 1) * Mat2(3, 1) - Mat1(3, 1) * Mat2(2, 1)
 j = Mat1(3, 1) * Mat2(1, 1) - Mat1(1, 1) * Mat2(3, 1)
 k = Mat1(1, 1) * Mat2(2, 1) - Mat1(2, 1) * Mat2(1, 1)
 
 sol(0) = i: sol(1) = j: sol(2) = k
    
 MultiplyVectors = sol
 
 Exit Function
 
Error_Dimension:
 Err.Raise "5016", , "Dimension should be (3 x 1) for both matrices in order to do cross multiplication !"

Error_Handler:

 If Err.Number = 5016 Then
  Err.Raise "5016", , "Dimension should be (3 x 1) for both matrices in order to do cross multiplication !"
 Else
  Err.Raise "5022", , "One or both of the matrices are null, this operation cannot be done !!"
 End If

 End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Magnitude of a Vector, vector should be (3x1)
' Function returns the solution or errors due to
' dimensions incompatibility
' Example:
'  Check Main Form !!
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function VectorMagnitude(Mat() As Double) As Double

 Dim Mat1() As Double
 
 On Error GoTo Error_Handler
 
 Mat1 = Find_R_C(Mat)

 If Mat1(0, 0) <> 3 Or Mat1(0, 1) <> 1 Then
  GoTo Error_Dimension
 End If
    
 VectorMagnitude = Sqr(Mat1(1, 1) * Mat1(1, 1) + Mat1(2, 1) * Mat1(2, 1) + Mat1(3, 1) * Mat1(3, 1))
 
 Exit Function
  
Error_Dimension:
 Err.Raise "5018", , "Dimension of the matrix should be (1 x 3) in order to find the vector's norm !"

Error_Handler:
 If Err.Number = 5018 Then
  Err.Raise "5018", , "Dimension of the matrix should be (3 x 1) in order to find the vector's magnitude !"
 Else
  Err.Raise "5022", , "In order to do this operation values must be assigned to the matrix !!"
 End If

 End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Transpose of a matrix
' Function returns the solution or errors
' Example:
'  Check Main Form !!
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Transpose(Mat() As Double) As Double()
 Dim Mat1() As Double, Tr_Mat() As Double
 Dim i As Integer, j As Integer, sol() As Double
 
 On Error GoTo Error_Handler
 
 Mat1() = Find_R_C(Mat())
   
  ReDim Tr_Mat(0 To Mat1(0, 1), 0 To Mat1(0, 0))
  ReDim sol(Mat1(0, 1) - 1, Mat1(0, 0) - 1)
  Tr_Mat(0, 0) = Mat1(0, 1)
  Tr_Mat(0, 1) = Mat1(0, 0)
  
 For i = 1 To Mat1(0, 0)
  For j = 1 To Mat1(0, 1)
   Tr_Mat(j, i) = Mat1(i, j)
  Next j
 Next i
  
 For i = 1 To Tr_Mat(0, 0)
   For j = 1 To Tr_Mat(0, 1)
     sol(i - 1, j - 1) = Tr_Mat(i, j)
   Next j
 Next i
  
    Transpose = sol
    Erase sol
 Exit Function
  
Error_Handler:
   Err.Raise "5028", , "In order to do this operation values must be assigned to the matrix !!"

 End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Multiply a matrix or a vector with a scalar quantity
' Function returns the solution or errors
' Example:
'  Check Main Form !!
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Public Function ScalarMultiply(Value As Double, Mat() As Double) As Double()
 Dim i As Integer, j As Integer
 Dim Mat1() As Double, sol() As Double
 
 On Error GoTo Error_Handler
  
 Mat1 = Find_R_C(Mat)
 ReDim sol(Mat1(0, 0) - 1, Mat1(0, 1) - 1)
   
 For i = 1 To Mat1(0, 0)
  For j = 1 To Mat1(0, 1)
   sol(i - 1, j - 1) = Mat1(i, j) * Value
  Next j
 Next i
 
 ScalarMultiply = sol
 
 Exit Function
  
Error_Handler:
  Err.Raise "5022", , "Matrix was not assigned"
 End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Divide matrix elements or a vector by a scalar quantity
' Function returns the solution or errors
' Example:
'  Check Main Form !!
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ScalarDivide(Value As Double, Mat() As Double) As Double()
 Dim i As Integer, j As Integer
 Dim Mat1() As Double, sol() As Double
 
 On Error GoTo Error_Handler
  
 Mat1 = Find_R_C(Mat)
 ReDim sol(Mat1(0, 0) - 1, Mat1(0, 1) - 1)
   
 For i = 1 To Mat1(0, 0)
  For j = 1 To Mat1(0, 1)
   sol(i - 1, j - 1) = Mat1(i, j) / Value
  Next j
 Next i
 
 ScalarDivide = sol
 
 Exit Function
  
Error_Handler:
  Err.Raise "5022", , "Matrix was not assigned"
 End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Print a matrix to multitext text box
' Function returns the solution or errors
' Example:
'  Check Main Form !!
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMat(Mat_1() As Double) As String
 Dim N_Rows As Integer, N_Columns, k As Integer, _
     i As Integer, j As Integer, m As Integer
 Dim StrElem As String, StrLen As Long, _
     Greatest() As Integer, LarString As String
 Dim OptiString As String, sol As String
 
 Dim Mat1() As Double
 
 Mat1 = Find_R_C(Mat_1)
    
 sol = ""
 OptiString = ""
    
  N_Rows = Mat1(0, 0)
  N_Columns = Mat1(0, 1)

 ReDim Greatest(N_Columns)
 
   For i = 1 To N_Rows
    For j = 1 To N_Columns
     If i = 1 Then
      Greatest(j) = 0
      For m = 1 To N_Rows
       StrElem = Format$(Mat1(m, j), "0.0000")
       StrLen = Len(StrElem)
       If Greatest(j) < StrLen Then
        Greatest(j) = StrLen
        LarString = StrElem
       End If
      Next m
       If Mid$(LarString, 1, 1) = "-" Then Greatest(j) = Greatest(j) + 1
     End If
     StrElem = Format$(Mat1(i, j), "0.0000")
     If Mid$(StrElem, 1, 1) = "-" Then
      StrLen = Len(StrElem)
      If Greatest(j) >= StrLen Then
      For k = 1 To (Greatest(j) - StrLen)
       OptiString = OptiString & "  "
      Next k
       OptiString = OptiString & " "
     End If
     Else
      StrLen = Len(StrElem)
      If Greatest(j) > StrLen Then
      For k = 1 To (Greatest(j) - StrLen)
       OptiString = OptiString & "  "
      Next k
     End If
     End If
     OptiString = OptiString & "  " & Format$(Mat1(i, j), "0.0000")
    Next j
    If i <> N_Rows Then
     sol = sol & OptiString & vbCrLf
     OptiString = ""
   End If
     sol = sol & OptiString
     OptiString = ""
   Next i
     
  PrintMat = sol
 
  Exit Function
  
End Function

Private Function Cutting(M_L As Integer)
 Dim Num As Integer
 Num = 0
 Num = M_L \ 20
  If M_L Mod 20 <> 0 Then Num = Num + 1
 Cutting = Num
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Return the maximum of two numbers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function max(ByVal X As Double, Y As Double) As Double
 If X >= Y Then
  max = X
 ElseIf X < Y Then
  max = Y
 End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Return the minimum of two numbers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function min(ByVal X As Double, Y As Double) As Double
 If X >= Y Then
  min = Y
 ElseIf X < Y Then
  min = X
 End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''
' This routine finds the
' atan2(y,x) vlaue
'''''''''''''''''''''''''''''''''''''''''''''''
Public Function atan2(ByVal Y As Double, ByVal X As Double) As Double
 Dim yy As Double, xx As Double
 yy = Abs(Y)
 xx = Abs(X)
 If Y = 0 And X > 0 Then
  atan2 = 0
 ElseIf Y > 0 And X > 0 Then
  atan2 = Atn(yy / xx)
 ElseIf Y > 0 And X = 0 Then
  atan2 = PI / 2
 ElseIf Y > 0 And X < 0 Then
  atan2 = PI - Atn(yy / xx)
 ElseIf Y = 0 And X < 0 Then
  atan2 = PI
 ElseIf Y < 0 And X < 0 Then
  atan2 = Atn(yy / xx) - PI
 ElseIf Y < 0 And X = 0 Then
  atan2 = PI / 2
 ElseIf Y < 0 And X > 0 Then
  atan2 = -Atn(yy / xx)
 End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''
' This routine returns Pi value
'''''''''''''''''''''''''''''''''''''''''''''''
Public Function PI() As Double
 PI = 4 * Atn(1)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''
' This routine returns the
' arc sin vlaue of an angle
'''''''''''''''''''''''''''''''''''''''''''''''
Public Function asin(ByVal X As Double) As Double
 asin = Atn(X / (Sqr(Abs(1 - X * X)) + 1E-200))
End Function

'''''''''''''''''''''''''''''''''''''''''''''''
' This routine returns the
' arc cos vlaue of an angle
'''''''''''''''''''''''''''''''''''''''''''''''
Public Function acos(ByVal X As Double) As Double
 acos = Atn(-X / (Sqr(1 - X * X))) + 2 * Atn(1)
End Function






