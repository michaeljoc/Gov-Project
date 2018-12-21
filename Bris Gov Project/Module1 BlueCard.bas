Attribute VB_Name = "Module1"
Option Explicit

Sub Button1_Click()
    Main
End Sub


'!!!!!!!!!!!!!!!!!!!!!!!!All methods related to SAP  has been redacted due to request by employer!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Sub Main()

        Dim fRow As Integer
        Dim fColumn As Integer
        Dim empNumber As String
        
        Dim maxTest1 As Integer
        
        Dim shName As String
        
         
        
        
        maxTest1 = maxRowTest1()
        
        Dim payName As String
        
        Dim payNameCol As Integer
        
        payName = "Employee No."
        
        payNameCol = payNameColGet(payName)
        
        
        
        
        fRow = 2 
        fColumn = payNameCol 
        
        Dim iPro As Integer 
        
        Dim colMock(11) As Variant 
        
        
    
    
        
        
            
        colMock(0) = 18
        colMock(1) = 19
        colMock(2) = 20
        colMock(3) = 21
        colMock(4) = 22
        
        colMock(5) = 23
        colMock(6) = 24
        colMock(7) = 25
        colMock(8) = 26
        colMock(9) = 27
        colMock(10) = 28
        colMock(11) = 29
        
        
        Dim colArrayData As Variant
        
        Dim checkValid As Boolean
        
        Dim mockEmpNumbRow As Integer 
        
        
        iPro = maxRow() 
        
        
        
        
        
        Do While fRow < maxTest1 + 1 
        
            
            
            
            checkValid = Validate(fRow, fColumn)
            
            If (checkValid) Then 
                
                empNumber = getPayrollNum(fRow, fColumn)
                
                
            End If
            
            mockEmpNumbRow = getEmployeeNumberPos(empNumber) 
            
            
            
            If mockEmpNumbRow = 0 Then
                mockEmpNumbRow = 87
            End If
            
            
            colArrayData = currentArray(mockEmpNumbRow, colMock) 
            
            
            
            Call addData(colArrayData, fRow) 
        
            fRow = fRow + 1
        
        Loop
        
        MsgBox ("The process has been completed")
        
End Sub

Public Function payNameColGet(payStri As String) As Integer
    
    
    Dim rFind As Range

    With Range("A:AB")
        Set rFind = .Find(What:=payStri, LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False)
        If Not rFind Is Nothing Then
           
            
            payNameColGet = rFind.Column
            
        End If
    End With
    
   
    
    
    
    
    
    
    
End Function



Public Function Validate(fr As Integer, fc As Integer) As Boolean 'Purpose of this function is to check if the payroll number is empty in the provided row and column
    
    If Not IsEmpty(Cells(fr, fc)) Then
    
        
        
        Validate = True
    Else
        
        Validate = False
    End If
    
End Function

Public Function getPayrollNum(payRow As Integer, payCol As Integer) As String 
    
    getPayrollNum = Cells(payRow, payCol)
    
    
    
End Function

Public Function getEmployeeNumberPos(payNumb As String) As Integer 
    
    Dim testMock As Variant
    
    Dim Gen As String
    Gen = Worksheets(2).Name
    
    Dim wb As Workbook: Set wb = ThisWorkbook
        
    Dim Mock1 As Worksheet
    Set Mock1 = ThisWorkbook.Worksheets(Gen)
    
    
    
    Dim mockColumnNumber As Integer
    
    Dim rFind2 As Range

    With Mock1.Range("A:AB")
        Set rFind2 = .Find(What:="Employee No.", LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False)
        If Not rFind2 Is Nothing Then
            
            
            mockColumnNumber = rFind2.Column

        End If
    End With
    
    Dim clNo As Integer
    
    Dim clLet As String
    
    clNo = mockColumnNumber
    
    clLet = Split(Cells(1, clNo).Address, "$")(1)
    
    
    
    Dim test9 As String
    
    test9 = clLet & ":" & clLet
    
    
    
    If Application.WorksheetFunction.IsNA(Application.Match(CLng(payNumb), Mock1.Range(test9), 0)) Then
        testMock = 870
    Else
        testMock = Application.Match(CLng(payNumb), Mock1.Range(test9), 0)
    End If
    
    
    
    
    getEmployeeNumberPos = testMock
    
 
    
    
    
    
    
    
    
    
    
End Function

Public Function maxRow() As Integer 
    
    Dim wb2 As Workbook: Set wb2 = ThisWorkbook
    
    Dim Gen As String
    Gen = Worksheets(2).Name
        
    Dim Mock2 As Worksheet
    Set Mock2 = ThisWorkbook.Worksheets(Gen)
    
    Dim k As Long
    
    
    
    k = Mock2.Range("M1048576").End(xlUp).Row
    
    
    
    
    
    
    maxRow = k
    


End Function

Public Function maxRowTest1() As Integer 


    
    Dim l As Long
    
    
    
    l = Range("H1048576").End(xlUp).Row
    
    
    
    
    
    
    maxRowTest1 = l





End Function

Public Function currentArray(rowData As Integer, rcolArray() As Variant) As Variant 
    
    Dim wb3 As Workbook: Set wb3 = ThisWorkbook
    
    Dim Gen As String
    Gen = Worksheets(2).Name
        
    Dim Mock3 As Worksheet
    Set Mock3 = ThisWorkbook.Worksheets(Gen)
    
    
    
    
    Dim iArr As Integer
    Dim i As Integer
    i = 0
    
    
    iArr = UBound(rcolArray) 
    
    Dim strArray() As Variant
    ReDim strArray(iArr)
    
    
    
    Do While i < iArr + 1
    
        If IsEmpty(Mock3.Cells(rowData, rcolArray(i))) Then
            strArray(i) = "Empty"
            
        Else
            
            strArray(i) = Mock3.Cells(rowData, rcolArray(i))
        End If
        
        
        
        
        i = i + 1
        
    Loop
    
    
    
    
    
 
    
   
    
    Dim j As Integer
    
    j = 0
    
    Do While j < iArr + 1
    
        
      
        
        
        j = j + 1
        
    Loop
    
   
    
    
    currentArray = strArray
    
    
    
    
    
    
End Function

Sub addData(subArray As Variant, tRow As Integer)
    
    
  
    
    Dim colInt As Integer
    
    Dim maxCol As Integer 
    
    colInt = 18 
    
    Dim arrayCounter As Integer
    
    arrayCounter = 0
    
    maxCol = 28 
    
    Do While colInt < maxCol + 1 
    
        
        
        Cells(tRow, colInt) = subArray(arrayCounter) 
        
        arrayCounter = arrayCounter + 1
        colInt = colInt + 1
        
        
        
    Loop
    
    
    
    
    
    
    
    
End Sub


















