Attribute VB_Name = "ArrayHelper"
'---------------------------------------------------------------------------------------
' Module    : ArrayHelper
' Author    : KRISH J
' Date      : 23/09/2019
' Purpose   :
' Returns   :
'           : 23-09-2019: Remove duplicates added
'             ArrayStringArrayPointer and arrayUtrArrayPointer function moved from system module to arrayhelper module
'             arrayExists reneamed to arrayExistsByMemRef
'           07/12/2019: some functions taken from https://github.com/todar/VBA-Arrays/blob/master/ArrayFunctions.bas
'           31/01/2020: FnArrayHasItemAt added. FnArrayGetSize improved
'           02/06/2020: FnArrayGetSize bug corrected regarding 0 index with data
'                       ArrayPush and pushTop can add empty items
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

'ERROR CODES CONSTANTS
Public Const ARRAY_NOT_PASSED_IN        As Integer = 5000
Public Const ARRAY_DIMENSION_INCORRECT  As Integer = 5001

Public Function ArrayAddString(ByRef stringArray As Variant, stringValue As String)
    'Adds a string to an existing string array
    
    On Error Resume Next
    If IsBlank(stringArray) Then Exit Function
    'If stringValue = "" Then Exit Function 'Allow empty string
    
    Dim L As Long
    L = FnArrayGetSize(stringArray)
    
    INC L
    ReDim Preserve stringArray(L)
    
    stringArray(L - 1) = stringValue
    
End Function

Public Function ArrayContainsEmpties(ByVal sourceArray As Variant) As Boolean
    'CHECK TO SEE IF SINGLE DIM ARRAY CONTAINS ANY EMPTY INDEXES
    
    'THIS FUNCTION IS FOR SINGLE DIMS ONLY
    If ArrayDimensionLength(sourceArray) <> 1 Then
        err.Raise ARRAY_DIMENSION_INCORRECT, , "SourceArray must be a single dimensional array."
    End If
    
    Dim index As Integer
    For index = LBound(sourceArray, 1) To UBound(sourceArray, 1)
        If IsEmpty(sourceArray(index)) Then
            ArrayContainsEmpties = True
            Exit Function
        End If
    Next index
    
End Function

Function ArrayConvertVariantToStringArray(iVariantArray As Variant) As String()
'---------------------------------------------------------------------------------------
' Procedure : ArrayConvertVariantToStringArray
' Author    : KRISH J
' Date      : 23/09/2019
' Purpose   : converts a variant array to string array
' Returns   :
' Usage     :
'---------------------------------------------------------------------------------------
'
    'Does source array has any entries?
    If ArrayHelper.ArrayIsEmpty(iVariantArray) Then Exit Function
    
    Dim stringArray() As String
    Dim arraySize As Long
    arraySize = ArrayHelper.FnArrayGetSize(iVariantArray)
    
    'Make new holder
    ReDim stringArray(arraySize) As String
    
    Dim I As Long
    
    For I = 0 To arraySize - 1
        stringArray(I) = CStr(iVariantArray(I))
    Next
    
    ArrayConvertVariantToStringArray = stringArray
    
End Function
Public Function ArrayDimensionLength(ByVal sourceArray As Variant) As Long
' Returns the length of the dimension of an array.
    
    On Error GoTo Catch
    Do
        Dim boundIndex As Long
        boundIndex = boundIndex + 1
        
        ' Loop until this line errors out.
        Dim test As Long
        test = UBound(sourceArray, boundIndex)
    Loop
Catch:
    ' Must remove one, this gives the proper dimension length.
    ArrayDimensionLength = boundIndex - 1

End Function

Public Function ArrayExistsByMemRef(ByVal ppArray As Long) As Long
  GetMem4 ppArray, VarPtr(ArrayExistsByMemRef)
  
    'Print ArrayExists(ArrPtr(someArray))
    'Print ArrayExists(StrArrPtr(someArrayOfStrings))
    'Print ArrayExists(UDTArrPtr(someArrayOfUDTs))
End Function

Public Function ArrayExtractColumn(ByVal sourceArray As Variant, ByVal ColumnIndex As Integer) As Variant
    'GET A COLUMN FROM A TWO DIM ARRAY, AND RETURN A SINLGE DIM ARRAY
    'SINGLE DIM ARRAYS ONLY
    If ArrayDimensionLength(sourceArray) <> 2 Then
        err.Raise ARRAY_DIMENSION_INCORRECT, , "SourceArray must be a two dimensional array."
    End If
    
    Dim Temp As Variant
    ReDim Temp(LBound(sourceArray, 1) To UBound(sourceArray, 1))
    
    Dim RowIndex As Integer
    For RowIndex = LBound(sourceArray, 1) To UBound(sourceArray, 1)
        Temp(RowIndex) = sourceArray(RowIndex, ColumnIndex)
    Next RowIndex
    
    ArrayExtractColumn = Temp
    
End Function

Public Function ArrayExtractRow(ByVal sourceArray As Variant, ByVal RowIndex As Long) As Variant
    'GET A ROW FROM A TWO DIM ARRAY, AND RETURN A SINLGE DIM ARRAY
    
    Dim Temp As Variant
    ReDim Temp(LBound(sourceArray, 2) To UBound(sourceArray, 2))
    
    Dim ColIndex As Integer
    For ColIndex = LBound(sourceArray, 2) To UBound(sourceArray, 2)
        Temp(ColIndex) = sourceArray(RowIndex, ColIndex)
    Next ColIndex
    
    ArrayExtractRow = Temp
    
End Function

Public Function ArrayFromRecordset(rs As Object, Optional IncludeHeaders As Boolean = True) As Variant
'RETURNS A 2D ARRAY FROM A RECORDSET, OPTIONALLY INCLUDING HEADERS, AND IT TRANSPOSES TO KEEP
'ORIGINAL OPTION BASE. (TRANSPOSE WILL SET IT TO BASE 1 AUTOMATICALLY.)
    
    '@NOTE: -Int(IncludeHeaders) RETURNS A BOOLEAN TO AN INT (0 OR 1)
    Dim HeadingIncrement As Integer
    HeadingIncrement = -Int(IncludeHeaders)
    
    'CHECK TO MAKE SURE THERE ARE RECORDS TO PULL FROM
    If rs.BOF Or rs.EOF Then
        Exit Function
    End If
    
    'STORE RS DATA
    Dim rsData As Variant
    rsData = rs.GetRows
    
    'REDIM TEMP TO ALLOW FOR HEADINGS AS WELL AS DATA
    Dim Temp As Variant
    ReDim Temp(LBound(rsData, 2) To UBound(rsData, 2) + HeadingIncrement, LBound(rsData, 1) To UBound(rsData, 1))
        
    If IncludeHeaders = True Then
        'GET HEADERS
        Dim headerIndex As Long
        For headerIndex = 0 To rs.Fields.Count - 1
            Temp(LBound(Temp, 1), headerIndex) = rs.Fields(headerIndex).name
        Next headerIndex
    End If
    
    'GET DATA
    Dim RowIndex As Long
    Dim ColIndex As Long
    For RowIndex = LBound(Temp, 1) + HeadingIncrement To UBound(Temp, 1)
        
        For ColIndex = LBound(Temp, 2) To UBound(Temp, 2)
            Temp(RowIndex, ColIndex) = rsData(ColIndex, RowIndex - HeadingIncrement)
        Next ColIndex
        
    Next RowIndex
    
    'RETURN
    ArrayFromRecordset = Temp
    
End Function

Public Function ArrayGetValues(ByRef stringArray As Variant, Optional delimitter As String = ",")
    'returns array values in a single line
    Dim I As Long
    For I = 0 To UBound(stringArray) - 1
        If (I = 0) Then
            ArrayGetValues = ArrayGetValues & stringArray(I)
        Else
            ArrayGetValues = ArrayGetValues & "," & stringArray(I)
        End If
    Next I
End Function

Public Function ArrayIndexOf(ByVal sourceArray As Variant, ByVal SearchElement As Variant) As Integer

'RETURNS INDEX OF A SINGLE DIM ARRAY ELEMENT
    If ArrayIsEmpty(sourceArray) Then Exit Function
    
    Dim index As Long
    For index = LBound(sourceArray, 1) To UBound(sourceArray, 1)
        If sourceArray(index) = SearchElement Then
            ArrayIndexOf = index
            Exit Function
        End If
    Next index
    index = -1
End Function

Public Function ArrayIsEmpty(ByRef sourceArray) As Boolean
    'a modified version of cpearsons code <http://www.cpearson.com/excel/VBAArrays.htm>

    ' Array was not passed in.
    If Not IsArray(sourceArray) Then
        ArrayIsEmpty = True
        Exit Function
    End If

    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    err.Clear
    On Error Resume Next
    If Not IsNumeric(UBound(sourceArray)) And (err.number <> 0) Then
        ArrayIsEmpty = True
    Else
        ' On rare occasion Err.Number will be 0 for an unallocated, empty array.
        ' On these occasions, LBound is 0 and UBound is -1.
        ' To accommodate the weird behavior, test to see if LB > UB. If so, the array is not allocated.
        err.Clear
        If LBound(sourceArray) > UBound(sourceArray) Then
            ArrayIsEmpty = True
        End If
    End If

End Function

Function ArrayRemoveDuplicates(iArray As Variant) As Variant
    'DESCRIPTION: Removes duplicates from your array using the dictionary method.
    'NOTES: (1.a) You must add a reference to the Microsoft Scripting Runtime library via
    ' the Tools > References menu.
    ' (1.b) This is necessary because I use Early Binding in this function.
    ' Early Binding greatly enhances the speed of the function.
    ' (2) The scripting dictionary will not work on the Mac OS.
    'SOURCE: https://wellsr.com
    '-----------------------------------------------------------------------

    If ArrayHelper.ArrayIsEmpty(iArray) Then Exit Function

    Dim I As Long
    Dim d As Scripting.Dictionary
    Set d = New Scripting.Dictionary
    With d
        For I = LBound(iArray) To UBound(iArray)
            If IsMissing(iArray(I)) = False Then
                .Item(iArray(I)) = 1
            End If
        Next
        ArrayRemoveDuplicates = .Keys
    End With
End Function

Public Function FnArrayRemoveDuplicates(ByVal sourceArray As Variant)
'Removes duplicates in an array
'taken from : https://stackoverflow.com/questions/11870095/remove-duplicates-from-array-using-vba

    Dim poArrNoDup()
    Dim dupArrIndex As Long
    Dim I           As Long
    Dim J           As Long
    Dim dupBool     As Boolean

    dupArrIndex = -1
    For I = LBound(sourceArray) To UBound(sourceArray)
        dupBool = False

        For J = LBound(sourceArray) To I
            If sourceArray(I) = sourceArray(J) And Not I = J Then
                dupBool = True
                Exit For
            End If
        Next J

        If dupBool = False Then
            dupArrIndex = dupArrIndex + 1
            ReDim Preserve poArrNoDup(dupArrIndex)
            poArrNoDup(dupArrIndex) = sourceArray(I)
        End If
    Next I

    FnArrayRemoveDuplicates = poArrNoDup
End Function

Public Function ArraySort(sourceArray As Variant) As Variant
    'SORT AN ARRAY [SINGLE DIMENSION]
    'SORT ARRAY A-Z
    Dim OuterIndex As Long
    For OuterIndex = LBound(sourceArray) To UBound(sourceArray) - 1
        
        Dim InnerIndex As Long
        For InnerIndex = OuterIndex + 1 To UBound(sourceArray)
            
            If sourceArray(OuterIndex) > sourceArray(InnerIndex) Then
                Dim Temp As Variant
                Temp = sourceArray(InnerIndex)
                sourceArray(InnerIndex) = sourceArray(OuterIndex)
                sourceArray(OuterIndex) = Temp
            End If
            
        Next InnerIndex
    Next OuterIndex
    
    ArraySort = sourceArray

End Function

Public Function FnArrayQuickSort(ByVal vArray As Variant, inLow As Long, inHi As Long)
'Taken from https://stackoverflow.com/questions/152319/vba-array-sort-function
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend

  If (inLow < tmpHi) Then vArray = FnArrayQuickSort(vArray, inLow, tmpHi)
  If (tmpLow < inHi) Then vArray = FnArrayQuickSort(vArray, tmpLow, inHi)
End Function

Public Function ArrayStringArrayPointer(arr() As String, Optional ByVal IgnoreMe As Long = 0) As Long
  GetMem4 VarPtr(IgnoreMe) - 4, VarPtr(ArrayStringArrayPointer)
End Function

Public Sub ArrayToTextFile(arr As Variant, filePath As String, Optional delimeter As String = ",")
    'SENDS AN ARRAY TO A TEXTFILE
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim Ts As Object
    Set Ts = fso.OpenTextFile(filePath, 2, True) '2=WRITEABLE
    Ts.Write FnArrayToString(arr, delimeter)
    
    Set Ts = Nothing
    Set fso = Nothing

End Sub

Public Function ArrayTranspose(sourceArray As Variant) As Variant
'APPLICATION.TRANSPOSE HAS A LIMIT ON THE SIZE OF THE ARRAY, AND IS LIMITED TO THE 1ST DIM
    Dim Temp As Variant

    Select Case ArrayDimensionLength(sourceArray)
        
        Case 2:
        
            ReDim Temp(LBound(sourceArray, 2) To UBound(sourceArray, 2), LBound(sourceArray, 1) To UBound(sourceArray, 1))
            
            Dim I As Long
            Dim J As Long
            For I = LBound(sourceArray, 2) To UBound(sourceArray, 2)
                For J = LBound(sourceArray, 1) To UBound(sourceArray, 1)
                    Temp(I, J) = sourceArray(J, I)
                Next
            Next
    
    End Select
    
    ArrayTranspose = Temp
    sourceArray = Temp

End Function

Public Function ArrayUDTArraryPointer(ByRef arr As Variant) As Long
  If varType(arr) Or vbArray Then
    GetMem4 VarPtr(arr) + 8, VarPtr(ArrayUDTArraryPointer)
  Else
    err.Raise 5, , "Variant must contain array of user defined type"
  End If
End Function

Public Function ArrayUnShift(ByRef byrefArray, ByVal Item)
    'Adds new item to the top of the array
    ArrayHelper.FnArrayPushToTop byrefArray, Item
    
    ArrayUnShift = byrefArray
End Function

Public Function Assign(ByRef Variable As Variant, ByVal value As Variant) As String

    ' Quick tool to either set or let depending on if the element is an object
    
    If IsObject(value) Then
        Set Variable = value
    Else
        Let Variable = value
    End If
    
    Assign = TypeName(value)
End Function

Public Function ConvertToArray(ByRef Val As Variant) As Variant
    'CONVERT OTHER LIST OBJECTS TO AN ARRAY
    Select Case TypeName(Val)
    
        Case "Collection":
            Dim index As Integer
            For index = 1 To Val.Count
                FnArrayPush ConvertToArray, Val(index)
            Next index
        
        Case "Dictionary":
            ConvertToArray = Val.items()
        
        Case Else
             
            If IsArray(Val) Then
                ConvertToArray = Val
            Else
                FnArrayPush ConvertToArray, Val
            End If
            
    End Select
    
End Function

Public Function FnArrayAddItem(ByRef byrefArray, ByVal iItem)
    'Adds an item to the end of the array
    FnArrayPush byrefArray, iItem
    FnArrayAddItem = byrefArray
End Function

Public Function FnArrayContainsItem(ByRef iArray, Item As Variant) As Boolean
    'Returns true if the array contains the item provided, otherwise false
    
    'Array has any item at all?
    If Not ArrayHelper.FnArrayHasitem(iArray) Then
        Exit Function
    End If
    
    'loop through array and find a match
    Dim v As Variant
    For Each v In iArray
        If v = Item Then
            FnArrayContainsItem = True
            Exit For
        End If
    Next v
    
End Function

Public Function FnArrayGetSize(ByRef byrefArray) As Long
    'Returns the size of an array or 0
    'if the first or last item contains vbNullstring, that item is discarted
    '16/05/2020: Redim will add blank item at the end of array. Check for data at index = 0
    
    Dim T           As Long
    Dim v           As Variant
    Dim hasValue    As Boolean
    
    On Error Resume Next
    T = UBound(byrefArray, 1) ' - LBound(byrefArray, 1) + 1
    
'    'Inline arrays have ubound = 0 but will have an item at 0 index
'    If (t = 0) Then
'        hasValue = ((Not byrefArray(0) Is Nothing) And byrefArray(0) <> vbNullString)
'        If (hasValue) Then
'            t = 1
'        End If
'    End If
'
    'Some array have array size of 0 but contains an item at the 0 index. Check if array size returned from ubound contains a value. if yes increase the size by one
    If T = 0 And ArrayHelper.FnArrayHasItemAt(byrefArray, T) Then

        Select Case varType(byrefArray(T))
            Case vbArray
                ' Array contains item?
                hasValue = FnArrayHasItemAt(byrefArray, T)
            Case vbBoolean
                hasValue = Not IsBlank(byrefArray(T))
            Case vbString
                hasValue = byrefArray(T) <> vbNullString
            Case vbObject
                hasValue = Not byrefArray(T) Is Nothing
            Case Else
                hasValue = False
        End Select

        
'        hasValue = (ArrayHelper.FnArrayHasItemAt(byrefArray, t) _
'                    And Not (byrefArray(t) = vbNullString) _
'                    And Not ((varType(byrefArray(t)) = vbArray) And (FnArrayHasItemAt(byrefArray, t))) _
'                    And Not ((varType(byrefArray(t)) = vbBoolean) And (byrefArray(t) = False)) _
'                    And Not ((varType(byrefArray(t)) = vbByte) And (byrefArray(t) = 0)) _
'                    And Not ((varType(byrefArray(t)) = vbCurrency) And (byrefArray(t) = 0)) _
'                    And Not ((varType(byrefArray(t)) = vbDataObject) And (byrefArray(t) Is Nothing)) _
'                    And Not ((varType(byrefArray(t)) = vbDate) And (byrefArray(t) = 0)) _
'                    And Not ((varType(byrefArray(t)) = vbDecimal) And (byrefArray(t) = 0)) _
'                    And Not ((varType(byrefArray(t)) = vbDouble) And (byrefArray(t) = 0)) _
'                    And Not ((varType(byrefArray(t)) = vbEmpty) And (byrefArray(t) = vbEmpty)) _
'                    And Not ((varType(byrefArray(t)) = vbError) And (byrefArray(t) = 0)) _
'                    And Not ((varType(byrefArray(t)) = vbInteger) And (byrefArray(t) = 0)) _
'                    And Not ((varType(byrefArray(t)) = vbLong) And (byrefArray(t) = 0)) _
'                    And Not ((varType(byrefArray(t)) = vbNull) And (byrefArray(t) Is Null)) _
'                    And Not ((varType(byrefArray(t)) = vbObject) And (byrefArray(t) Is Nothing)) _
'                    And Not ((varType(byrefArray(t)) = vbSingle) And (byrefArray(t) = 0)) _
'                    And Not ((varType(byrefArray(t)) = vbString) And (byrefArray(t) = vbNullString)) _
'                    And Not ((varType(byrefArray(t)) = vbUserDefinedType) And (byrefArray(t) Is Nothing)) _
'                    And Not ((varType(byrefArray(t)) = vbVariant) And (IsBlank(byrefArray(t)))) _
'                    )
        If (hasValue) Then
            INC T
        End If
    End If
'
    
    If T < 0 Then
        T = 0
    End If
    
    FnArrayGetSize = T
End Function
Public Function FnArrayHasitem(ByRef iArray) As Boolean
    'Has this array contain any item?
    ' todo: rename this function
    On Error Resume Next
    FnArrayHasitem = Not IsBlank(iArray(0))
End Function

Public Function FnArrayHasItemAt(ByRef iArray, ByVal index As Long) As Boolean
    'Does this array has an item at given index? item can be blank.
    On Error GoTo IndexError
    
    Dim v As Variant
    
    If (varType(iArray(index)) = vbObject) Then
        Set v = iArray(index)
    Else
        v = iArray(index)
    End If
    
    FnArrayHasItemAt = True
    
ExitRoutine:
    On Error Resume Next
    Exit Function
    
IndexError:
    FnArrayHasItemAt = False
    Resume ExitRoutine
End Function

Public Function FnArrayCorrectItemIndex(ByVal anArray) As Variant()
    'Some arrays received from .NET comes with 0 as array size but will contain an item on the index 0
    'To get the correct array item, we will remake tha array using vba
    
    Dim a() As Variant
    On Error Resume Next
    
    Dim v As Variant
    For Each v In anArray
        If Not v Is Nothing Then
            FnArrayAddItem a, v
        End If
    Next v
    
    FnArrayCorrectItemIndex = a
    
End Function

Public Sub FnArrayMerge(ByRef addToArray, ByVal addFromArray)
    'Merges two arrays. Arrays must be of the same kind
    
    On Error Resume Next
    If Not FnArrayHasitem(addFromArray) Then Exit Sub
        
    Dim v As Variant
    For Each v In addFromArray
        FnArrayAddItem addToArray, v
    Next v
    
End Sub

Public Function FnArrayMultiDimentioned(ByRef iArr) As Boolean
    'Returns true if an array is multi dimentioned. i.e. Array(1,1) => true. Oterwise false
    On Error GoTo NOT_MULTI_DIMENTIONED
    Dim AW As Long
    
    AW = UBound(iArr, 2) - 1
    
    FnArrayMultiDimentioned = True
    Exit Function
    
NOT_MULTI_DIMENTIONED:
    FnArrayMultiDimentioned = False
End Function

Public Function FnArrayPop(ByRef byrefArray)
    'Removes the last item from an array
    
    If (Not ArrayHelper.FnArrayHasitem(byrefArray)) Then Exit Function
    
    'Resize the array
    On Error GoTo REVERSE_STATIC_ARRAY
    ReDim Preserve byrefArray(UBound(byrefArray) - 1)
    
    Exit Function
    
REVERSE_STATIC_ARRAY:
    'Static arrays cannot be re-dimentioned.
    
    gDll.ShowDialog "Array cannot be resized", vbCritical
    
End Function

Public Function FnArrayPrintValues(ByRef stringArray As Variant)
    'Prints array content
    Dim I As Long
    For I = 0 To UBound(stringArray) - 1
        DebugPrint stringArray(I)
    Next I
End Function

Public Function FnArrayPush(ByRef byrefArray, ByVal iItem)
'---------------------------------------------------------------------------------------
' Procedure : FnArrayPush
' Author    : KRISH
' Date      : 13/03/2018
' Purpose   : Adds an item to the end of an array
' Returns   :
'---------------------------------------------------------------------------------------
'

    'Even empty items should be added
    'If IsBlank(iItem) Then Exit Function
    
    On Error Resume Next
    Dim T As Long
    
    'Get array sieze
    T = FnArrayGetSize(byrefArray)
    '+1
    INC T
    
    ReDim Preserve byrefArray(0 To T)
    
    If varType(iItem) = vbObject Then
        Set byrefArray(T - 1) = iItem
    Else
        byrefArray(T - 1) = iItem
    End If
    
End Function

Private Function FnArrayAddItemAt(ByRef byrefArray, ByVal Item, index As Long)

    On Error Resume Next
    If varType(Item) = vbObject Then
        Set byrefArray(index) = Item
    Else
        byrefArray(index) = Item
    End If
        
End Function

Public Function FnArrayPushToTop(ByRef byrefArray, ByVal iItem)
'---------------------------------------------------------------------------------------
' Procedure : FnArrayPushToTop
' Author    : KRISH
' Date      : 13/03/2018
' Purpose   : Adds an item to the start of an array.  Only dynamic arrays are allowed
' Returns   :
'---------------------------------------------------------------------------------------
'

    'Even empty item should be added
    'If IsBlank(iItem) Then Exit Function
    
    If Not IsArray(byrefArray) Then Exit Function
    
    Dim L As Long
    'Get size
    L = FnArrayGetSize(byrefArray)
    
    'Size +1
    INC L
    'Re allocate spaces
    ReDim Preserve byrefArray(L)
    
    'Shift all items from right
    Dim I As Long
    For I = (L - 1) To 1 Step -1
        FnArrayAddItemAt byrefArray, byrefArray(I - 1), I
    Next I
    
    'Add the newest item to the first
    FnArrayAddItemAt byrefArray, iItem, 0
    
    'Return the array
    FnArrayPushToTop = byrefArray
    
End Function

Public Function FnArrayRemove(ByRef sourceArray, Optional elementAt As Long = 0)
    'Removes an element using its index from an array.
    
    If ArrayIsEmpty(sourceArray) Then Exit Function
    Dim L As Long
    
    L = ArrayHelper.FnArrayGetSize(sourceArray)
    
    If elementAt > L Then Exit Function 'Invalid element index
    
    FnArrayRemove = sourceArray(elementAt)
    
    If ((L - 1) = elementAt) Then 'Requested to remove the last item.
        ReDim Preserve sourceArray(UBound(sourceArray, 1) - 1)
    Else
        'Requested to remove something in between
        Dim I As Long
        
        For I = elementAt To L - 2
            sourceArray(I) = sourceArray(I + 1)
        Next I
        
        ReDim Preserve sourceArray(UBound(sourceArray, 1) - 1)
        
    End If
    
    FnArrayRemove = sourceArray

End Function

Public Function FnArrayReverse(ByVal byvalArray) As Variant()
    'Reverses the current array. Return reversed array or null
    
    If (Not FnArrayHasitem(byvalArray)) Then Exit Function
    
    Dim tempArray() As Variant
    Dim I As Long
    Dim J As Long
    Dim C As Long
    
    On Error Resume Next
    C = UBound(byvalArray)
    
    'c cannot be 0 at this time because fnArrayHasItem returned true. => it must have at least one item
    If (C = 0) Then C = 1
    
    'Init temp array
    ReDim Preserve tempArray(C)
    
    'We are going in reverse
    J = 0
    For I = (C - 1) To 0 Step -1
        
        tempArray(J) = byvalArray(I)
        INC J
        
    Next I
    
    FnArrayReverse = tempArray
End Function

Public Function FnArrayToString(ByVal iArr, Optional delimiter As Variant = ";", Optional delimiterForNewRow As Variant = vbNewLine) As String
    'Converts an array to a long delimited string
    'iArr       = source array
    'delimiter  = seperator for elements i.e. one, two,three
    
    On Error Resume Next
    Dim row As Integer
    Dim Column As Integer
    
    Dim AL As Integer 'Array length
    Dim AW As Integer 'array width
    
    If (Not ArrayHelper.FnArrayHasitem(iArr)) Then Exit Function
    
    Select Case ArrayDimensionLength(iArr)
        Case 1
            FnArrayToString = VBA.Join(iArr, delimiter)
        Case 2
            AL = ArrayHelper.FnArrayGetSize(iArr)
            If AL = 0 Then Exit Function
            
            AW = UBound(iArr, 2)
        
            For row = 0 To AL - 1
                For Column = 0 To AW - 1
                    FnArrayToString = FnArrayToString & iArr(row, Column) & delimiter
                Next Column
                
                FnArrayToString = FnArrayToString & delimiterForNewRow
                
            Next row
        
        Case Else
        
    End Select
    
EXIT_ROUTINE:
    Exit Function

End Function
