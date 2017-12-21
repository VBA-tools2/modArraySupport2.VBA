Attribute VB_Name = "modArraySupport"

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'2do:
'- refactor
'     If ... Then
'        Exit Function
'     End If
'  to
'     If ... Then Exit Function
'- create unit tests for these functions
'  (get example arrays from web sites referring to array stuff)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


Option Explicit
Option Compare Text

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'modArraySupport
'By Chip Pearson, chip@cpearson.com, www.cpearson.com
'
'This module contains procedures that provide information about and manipulate
'VB/VBA arrays. NOTE: These functions call one another. It is strongly suggested
'that you Import this entire module to a VBProject rather then copy/pasting
'individual procedures.
'
'For details on these functions, see www.cpearson.com/excel/VBAArrays.htm
'
'This module contains the following functions:
'     AreDataTypesCompatible
'     ChangeBoundsOfArray
'     CombineTwoDArrays
'     CompareArrays
'     ConcatenateArrays
'     CopyArray
'     CopyArraySubSetToArray
'     CopyNonNothingObjectsToArray
'     DataTypeOfArray
'     DeleteArrayElement
'     ExpandArray
'     FirstNonEmptyStringIndexInArray
'     GetColumn
'     GetRow
'     InsertElementIntoArray
'     IsArrayAllDefault
'     IsArrayAllNumeric
'     IsArrayAllocated
'     IsArrayDynamic
'     IsArrayEmpty
'     IsArrayObjects
'     IsArraySorted
'     IsNumericDataType
'     IsVariantArrayConsistent
'     IsVariantArrayNumeric
'     MoveEmptyStringsToEndOfArray
'     NumberOfArrayDimensions
'     NumElements
'     ResetVariantArrayToDefaults
'     ReverseArrayInPlace
'     ReverseArrayOfObjectsInPlace
'     SetObjectArrayToNothing
'     SetVariableToDefault
'     SwapArrayColumns
'     SwapArrayRows
'     TransposeArray
'     VectorsToArray
'
'Function documentation is above each function.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Error Number Constants
Public Const C_ERR_NO_ERROR = 0&
Public Const C_ERR_SUBSCRIPT_OUT_OF_RANGE = 9&
Public Const C_ERR_ARRAY_IS_FIXED_OR_LOCKED = 10&


'------------------------------------------------------------------------------
Sub AddUDFToCustomCategory()

   '===========================================================================
   'how should the category be named?
   Const sCategory As String = "Array Support"
   '===========================================================================

   With Application
      .MacroOptions Category:=sCategory, Macro:="AreDataTypesCompatible"
      .MacroOptions Category:=sCategory, Macro:="ChangeBoundsOfArray"
      .MacroOptions Category:=sCategory, Macro:="CombineTwoDArrays"
      .MacroOptions Category:=sCategory, Macro:="CompareArrays"
      .MacroOptions Category:=sCategory, Macro:="ConcatenateArrays"
      .MacroOptions Category:=sCategory, Macro:="CopyArray"
      .MacroOptions Category:=sCategory, Macro:="CopyArraySubSetToArray"
      .MacroOptions Category:=sCategory, Macro:="CopyNonNothingObjectsToArray"
      .MacroOptions Category:=sCategory, Macro:="DataTypeOfArray"
      .MacroOptions Category:=sCategory, Macro:="DeleteArrayElement"
      .MacroOptions Category:=sCategory, Macro:="ExpandArray"
      .MacroOptions Category:=sCategory, Macro:="FirstNonEmptyStringIndexInArray"
      .MacroOptions Category:=sCategory, Macro:="GetColumn"
      .MacroOptions Category:=sCategory, Macro:="GetRow"
      .MacroOptions Category:=sCategory, Macro:="InsertElementIntoArray"
      .MacroOptions Category:=sCategory, Macro:="IsArrayAllDefault"
      .MacroOptions Category:=sCategory, Macro:="IsArrayAllNumeric"
      .MacroOptions Category:=sCategory, Macro:="IsArrayAllocated"
      .MacroOptions Category:=sCategory, Macro:="IsArrayDynamic"
      .MacroOptions Category:=sCategory, Macro:="IsArrayEmpty"
      .MacroOptions Category:=sCategory, Macro:="IsArrayObjects"
      .MacroOptions Category:=sCategory, Macro:="IsArraySorted"
      .MacroOptions Category:=sCategory, Macro:="IsNumericDataType"
      .MacroOptions Category:=sCategory, Macro:="IsVariantArrayConsistent"
      .MacroOptions Category:=sCategory, Macro:="IsVariantArrayNumeric"
      .MacroOptions Category:=sCategory, Macro:="MoveEmptyStringsToEndOfArray"
      .MacroOptions Category:=sCategory, Macro:="NumberOfArrayDimensions"
      .MacroOptions Category:=sCategory, Macro:="NumElements"
      .MacroOptions Category:=sCategory, Macro:="ResetVariantArrayToDefaults"
      .MacroOptions Category:=sCategory, Macro:="ReverseArrayInPlace"
      .MacroOptions Category:=sCategory, Macro:="ReverseArrayOfObjectsInPlace"
      .MacroOptions Category:=sCategory, Macro:="SetObjectArrayToNothing"
      .MacroOptions Category:=sCategory, Macro:="SetVariableToDefault"
      .MacroOptions Category:=sCategory, Macro:="TransposeArray"
      .MacroOptions Category:=sCategory, Macro:="VectorsToArray"
   End With
End Sub




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CompareArrays
'This function compares two arrays, Array1 and Array2, element by element, and puts the results of
'the comparisons in ResultArray. Each element of ResultArray will be -1, 0, or +1. A -1 indicates that
'the element in Array1 was less than the corresponding element in Array2. A 0 indicates that the
'elements are equal, and +1 indicates that the element in Array1 is greater than Array2. Both
'Array1 and Array2 must be allocated single-dimensional arrays, and ResultArray must be dynamic array
'of a numeric data type (typically Longs). Array1 and Array2 must contain the same number of elements,
'and have the same lower bound. The LBound of ResultArray will be the same as the data arrays.
'
'An error will occur if Array1 or Array2 contains an Object or User Defined Type.
'
'When comparing elements, the procedure does the following:
'If both elements are numeric data types, they are compared arithmetically.

'If one element is a numeric data type and the other is a string and that string is numeric,
'then both elements are converted to Doubles and compared arithmetically. If the string is not
'numeric, both elements are converted to strings and compared using StrComp, with the
'compare mode set by CompareMode.
'
'If both elements are numeric strings, they are converted to Doubles and compared arithmetically.
'
'If either element is not a numeric string, the elements are converted and compared with StrComp.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CompareArrays( _
   Array1 As Variant, _
   Array2 As Variant, _
   ResultArray As Variant, _
   Optional CompareMode As VbCompareMethod = vbTextCompare _
      ) As Boolean
Attribute CompareArrays.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim Ndx1 As LongPtr
   Dim Ndx2 As LongPtr
   Dim ResNdx As LongPtr
   Dim S1 As String
   Dim S2 As String
   Dim D1 As Double
   Dim D2 As Double
   Dim Done As Boolean
   Dim Compare As VbCompareMethod
   Dim LB As LongPtr
   
   
   'Set the default return value
   CompareArrays = False
   
   'Ensure we have a Compare mode value
   If CompareMode = vbBinaryCompare Then
      Compare = vbBinaryCompare
   Else
      Compare = vbTextCompare
   End If
   
   If Not IsArray(Array1) Then Exit Function
   If Not IsArray(Array2) Then Exit Function
   If Not IsArray(ResultArray) Then Exit Function
   If Not IsArrayDynamic(ResultArray) Then Exit Function
   If NumberOfArrayDimensions(Array1) <> 1 Then Exit Function
   If NumberOfArrayDimensions(Array2) <> 1 Then Exit Function

'---
'2do: this does not make sense, because it was already tested above
'---
   'allow 0 indicating non-allocated array
   If NumberOfArrayDimensions(Array1) > 1 Then Exit Function
   
   'Ensure the LBounds are the same
   If LBound(Array1) <> LBound(Array2) Then Exit Function
   
   'Ensure the arrays are the same size
   If (UBound(Array1) - LBound(Array1)) <> (UBound(Array2) - LBound(Array2)) Then
      Exit Function
   End If
   
   'Redim ResultArray to the numbr of elements in Array1
   ReDim ResultArray(LBound(Array1) To UBound(Array1))
   
   Ndx1 = LBound(Array1)
   Ndx2 = LBound(Array2)
   
   'Scan each array to see if it contains objects or User-Defined Types
   'If found, exit with False
   For Ndx1 = LBound(Array1) To UBound(Array1)
      If IsObject(Array1(Ndx1)) Then Exit Function
      If VarType(Array1(Ndx1)) >= vbArray Then Exit Function
      If VarType(Array1(Ndx1)) = vbUserDefinedType Then Exit Function
   Next
   
   For Ndx1 = LBound(Array2) To UBound(Array2)
      If IsObject(Array2(Ndx1)) Then Exit Function
      If VarType(Array2(Ndx1)) >= vbArray Then Exit Function
      If VarType(Array2(Ndx1)) = vbUserDefinedType Then Exit Function
   Next
   
   Ndx1 = LBound(Array1)
   Ndx2 = Ndx1
   ResNdx = LBound(ResultArray)
   Done = False
   
   'Loop until we reach the end of the array
   Do Until Done = True
      If IsNumeric(Array1(Ndx1)) And IsNumeric(Array2(Ndx2)) Then
         D1 = CDbl(Array1(Ndx1))
         D2 = CDbl(Array2(Ndx2))
         If D1 = D2 Then
            ResultArray(ResNdx) = 0
         ElseIf D1 < D2 Then
            ResultArray(ResNdx) = -1
         Else
            ResultArray(ResNdx) = 1
         End If
      Else
         S1 = CStr(Array1(Ndx1))
         S2 = CStr(Array2(Ndx1))
         ResultArray(ResNdx) = StrComp(S1, S2, Compare)
      End If
           
      ResNdx = ResNdx + 1
      Ndx1 = Ndx1 + 1
      Ndx2 = Ndx2 + 1
      'If Ndx1 is greater than UBound(Array1) we've hit the end of the arrays
      If Ndx1 > UBound(Array1) Then
         Done = True
      End If
   Loop
   
   CompareArrays = True
   
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ConcatenateArrays
'This function appends ArrayToAppend to the end of ResultArray, increasing the size of ResultArray
'as needed. ResultArray must be a dynamic array, but it need not be allocated. ArrayToAppend
'may be either static or dynamic, and if dynamic it may be unallocted. If ArrayToAppend is
'unallocated, ResultArray is left unchanged.
'
'The data types of ResultArray and ArrayToAppend must be either the same data type or
'compatible numeric types. A compatible numeric type is a type that will not cause a loss of
'precision or cause an overflow. For example, ReturnArray may be Longs, and ArrayToAppend amy
'by Longs or Integers, but not Single or Doubles because information might be lost when
'converting from Double to Long (the decimal portion would be lost). To skip the compatability
'check and allow any variable type in ResultArray and ArrayToAppend, set the NoCompatabilityCheck
'parameter to True. If you do this, be aware that you may loose precision and you may will
'get an overflow error which will cause a result of 0 in that element of ResultArra.
'
'Both ReaultArray and ArrayToAppend must be one-dimensional arrays.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ConcatenateArrays( _
   ResultArray As Variant, _
   ArrayToAppend As Variant, _
   Optional NoCompatabilityCheck As Boolean = False _
      ) As Boolean
Attribute ConcatenateArrays.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim VTypeResult As VbVarType
   Dim Ndx As LongPtr
   Dim Res As LongPtr
   Dim NumElementsToAdd As LongPtr
   Dim AppendNdx As LongPtr
   Dim VTypeAppend As VbVarType
   Dim ResultLB As LongPtr
   Dim ResultUB As LongPtr
   Dim ResultWasAllocated As Boolean
   
   
   'Set the default result
   ConcatenateArrays = False
   
   If Not IsArray(ResultArray) Then Exit Function
   If Not IsArray(ArrayToAppend) Then Exit Function
   If Not IsArrayDynamic(ResultArray) Then Exit Function
'---
'2do: '>1' or '<>1'?
'---
   'Ensure both arrays are single dimensional
   If NumberOfArrayDimensions(ResultArray) > 1 Then Exit Function
   If NumberOfArrayDimensions(ArrayToAppend) > 1 Then Exit Function
   
   'Ensure ArrayToAppend is allocated. If ArrayToAppend is not allocated,
   'we have nothing to append, so exit with a True result.
   If Not IsArrayAllocated(ArrayToAppend) Then
      ConcatenateArrays = True
      Exit Function
   End If
   
   
   If NoCompatabilityCheck = False Then
      'Ensure the array are compatible data types
      If AreDataTypesCompatible(ArrayToAppend, ResultArray) = False Then
         'The arrays are not compatible data types
         Exit Function
      End If
       
      'If one array is an array of objects, ensure the other contains all
      'objects (or Nothing)
      If VarType(ResultArray) - vbArray = vbObject Then
         If IsArrayAllocated(ArrayToAppend) Then
            For Ndx = LBound(ArrayToAppend) To UBound(ArrayToAppend)
               If Not IsObject(ArrayToAppend(Ndx)) Then Exit Function
            Next
         End If
      End If
   End If
       
       
   'Get the number of elements in ArrrayToAppend
   NumElementsToAdd = UBound(ArrayToAppend) - LBound(ArrayToAppend) + 1
   
   'Get the bounds for resizing the ResultArray. If ResultArray is allocated
   'use the LBound and UBound+1. If ResultArray is not allocated, use the
   'LBound of ArrayToAppend for both the LBound and UBound of ResultArray.
   If IsArrayAllocated(ResultArray) Then
      ResultLB = LBound(ResultArray)
      ResultUB = UBound(ResultArray)
      ResultWasAllocated = True
      ReDim Preserve ResultArray(ResultLB To ResultUB + NumElementsToAdd)
   Else
      ResultUB = UBound(ArrayToAppend)
      ResultWasAllocated = False
      ReDim ResultArray(LBound(ArrayToAppend) To UBound(ArrayToAppend))
   End If
   
   '''Copy the data from ArrayToAppend to ResultArray.
   'If ResultArray was allocated, we have to put the data from ArrayToAppend
   'at the end of the ResultArray.
   If ResultWasAllocated = True Then
      AppendNdx = LBound(ArrayToAppend)
      For Ndx = ResultUB + 1 To UBound(ResultArray)
         If IsObject(ArrayToAppend(AppendNdx)) Then
            Set ResultArray(Ndx) = ArrayToAppend(AppendNdx)
         Else
            ResultArray(Ndx) = ArrayToAppend(AppendNdx)
         End If
         AppendNdx = AppendNdx + 1
         If AppendNdx > UBound(ArrayToAppend) Then
            Exit For
         End If
      Next
   'If ResultArray was not allocated, we simply copy element by element from
   'ArrayToAppend to ResultArray.
   Else
      For Ndx = LBound(ResultArray) To UBound(ResultArray)
         If IsObject(ArrayToAppend(Ndx)) Then
            Set ResultArray(Ndx) = ArrayToAppend(Ndx)
         Else
            ResultArray(Ndx) = ArrayToAppend(Ndx)
         End If
      Next
   
   End If
   
   ConcatenateArrays = True

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CopyArray
'This function copies the contents of SourceArray to the DestinationaArray. Both SourceArray
'and DestinationArray may be either static or dynamic and either or both may be unallocated.
'
'If DestinationArray is dynamic, it is resized to match SourceArray. The LBound and UBound
'of DestinationArray will be the same as SourceArray, and all elements of SourceArray will
'be copied to DestinationArray.
'
'If DestinationArray is static and has more elements than SourceArray, all of SourceArray
'is copied to DestinationArray and the right-most elements of DestinationArray are left
'intact.
'
'If DestinationArray is static and has fewer elements that SourceArray, only the left-most
'elements of SourceArray are copied to fill out DestinationArray.
'
'If SourceArray is an unallocated array, DestinationArray remains unchanged and the procedure
'terminates.
'
'If both SourceArray and DestinationArray are unallocated, no changes are made to either array
'and the procedure terminates.
'
'SourceArray may contain any type of data, including Objects and Objects that are Nothing
'(the procedure does not support arrays of User Defined Types since these cannot be coerced
'to Variants -- use classes instead of types).
'
'The function tests to ensure that the data types of the arrays are the same or are compatible.
'See the function AreDataTypesCompatible for information about compatible data types. To skip
'this compability checking, set the NoCompatabilityCheck parameter to True. Note that you may
'lose information during data conversion (e.g., losing decimal places when converting a Double
'to a Long) or you may get an overflow (storing a Long in an Integer) which will result in that
'element in DestinationArray having a value of 0.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SP: - changed order of arguments (to be consistent: "Source" first, then "Dest")
Public Function CopyArray( _
   SourceArray As Variant, _
   DestinationArray As Variant, _
   Optional NoCompatabilityCheck As Boolean = False _
      ) As Boolean
Attribute CopyArray.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim VTypeSource As VbVarType
   Dim VTypeDest As VbVarType
   Dim SNdx As LongPtr
   Dim DNdx As LongPtr
   
   
   'Set the default return value
   CopyArray = False
   
   If Not IsArray(DestinationArray) Then Exit Function
   If Not IsArray(SourceArray) Then Exit Function
   
   'Ensure DestinationArray and SourceArray are single-dimensional.
   '0 indicates an unallocated array, which is allowed.
   If NumberOfArrayDimensions(SourceArray) > 1 Then Exit Function
   If NumberOfArrayDimensions(DestinationArray) > 1 Then Exit Function
   
   'If SourceArray is not allocated, leave DestinationArray intact and
   'return a result of True.
   If Not IsArrayAllocated(SourceArray) Then
      CopyArray = True
      Exit Function
   End If
   
   If NoCompatabilityCheck = False Then
       'Ensure both arrays are the same type or compatible data types. See the
       'function AreDataTypesCompatible for information about compatible types.
      If Not AreDataTypesCompatible(SourceArray, DestinationArray) Then
         Exit Function
      End If
       'If one array is an array of objects, ensure the other contains all
       'objects (or Nothing)
      If VarType(DestinationArray) - vbArray = vbObject Then
         If IsArrayAllocated(SourceArray) Then
            For SNdx = LBound(SourceArray) To UBound(SourceArray)
               If Not IsObject(SourceArray(SNdx)) Then Exit Function
            Next
         End If
      End If
   End If
   
   'If both arrays are allocated, copy from SourceArray to DestinationArray.
   'If SourceArray is smaller that DesetinationArray, the right-most elements
   'of DestinationArray are left unchanged. If SourceArray is larger than
   'DestinationArray, the right most elements of SourceArray are not copied.
   If IsArrayAllocated(DestinationArray) Then
      If IsArrayAllocated(SourceArray) Then
         DNdx = LBound(DestinationArray)
         On Error Resume Next
         For SNdx = LBound(SourceArray) To UBound(SourceArray)
            If IsObject(SourceArray(SNdx)) Then
               Set DestinationArray(DNdx) = SourceArray(DNdx)
            Else
               DestinationArray(DNdx) = SourceArray(DNdx)
            End If
            DNdx = DNdx + 1
            If DNdx > UBound(DestinationArray) Then
               Exit For
            End If
         Next
         On Error GoTo 0
      'If SourceArray is not allocated, so we have nothing to copy.
      'Exit with a result of True. Leave DestinationArray intact.
      Else
         CopyArray = True
         Exit Function
      End If
   'If Destination array is not allocated and SourceArray is allocated,
   'Redim DestinationArray to the same size as SourceArray and copy
   'the elements from SourceArray to DestinationArray.
   Else
      If IsArrayAllocated(SourceArray) Then
         On Error Resume Next
         ReDim DestinationArray(LBound(SourceArray) To UBound(SourceArray))
         For SNdx = LBound(SourceArray) To UBound(SourceArray)
            If IsObject(SourceArray(SNdx)) Then
               Set DestinationArray(SNdx) = SourceArray(SNdx)
            Else
               DestinationArray(SNdx) = SourceArray(SNdx)
            End If
         Next
         On Error GoTo 0
      'If both SourceArray and DestinationArray are unallocated, we have
      'nothing to copy (this condition is actually detected above, but
      'included here for consistancy), so get out with a result of True.
      Else
         CopyArray = True
         Exit Function
      End If
   End If
   
   CopyArray = True

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CopyArraySubSetToArray
'This function copies elements of InputArray to ResultArray. It takes the elements
'from FirstElementToCopy to LastElementToCopy (inclusive) from InputArray and
'copies them to ResultArray, starting at DestinationElement. Existing data in
'ResultArray will be overwrittten. If ResultArray is a dynamic array, it will
'be resized if needed. If ResultArray is a static array and it is not large
'enough to copy all the elements, no elements are copied and the function
'returns False.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CopyArraySubSetToArray( _
   InputArray As Variant, _
   ResultArray As Variant, _
   FirstElementToCopy As LongPtr, _
   LastElementToCopy As LongPtr, _
   DestinationElement As LongPtr _
      ) As Boolean
Attribute CopyArraySubSetToArray.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim SrcNdx As LongPtr
   Dim DestNdx As LongPtr
   Dim NumElementsToCopy As LongPtr
   
   
   'Set the default return value
   CopyArraySubSetToArray = False
   
   If Not IsArray(InputArray) Then Exit Function
   If Not IsArray(ResultArray) Then Exit Function
   If NumberOfArrayDimensions(InputArray) <> 1 Then Exit Function
   'Ensure ResultArray is unallocated or single dimensional
   If NumberOfArrayDimensions(ResultArray) > 1 Then Exit Function
   
   'Ensure the bounds and indexes are valid
   If FirstElementToCopy < LBound(InputArray) Then Exit Function
   If LastElementToCopy > UBound(InputArray) Then Exit Function
   If FirstElementToCopy > LastElementToCopy Then Exit Function
   
   
   'Calculate the number of elements we'll copy from InputArray to ResultArray
   NumElementsToCopy = LastElementToCopy - FirstElementToCopy + 1
   
   If Not IsArrayDynamic(ResultArray) Then
      If (DestinationElement + NumElementsToCopy - 1) > UBound(ResultArray) Then
         'ResultArray is static and can't be resized.
         'There is not enough room in the array to copy all the data.
         Exit Function
      End If
   'ResultArray is dynamic and can be resized
   Else
      'Test whether we need to resize the array, and resize it if required
      If IsArrayEmpty(ResultArray) Then
         'ResultArray is unallocated. Resize it to
         'DestinationElement + NumElementsToCopy - 1.
         'This provides empty elements to the left of the DestinationElement
         'and room to copy NumElementsToCopy.
         ReDim ResultArray(1 To DestinationElement + NumElementsToCopy - 1)
      'ResultArray is allocated.
      Else
         'If there isn't room enough in ResultArray to hold NumElementsToCopy
         'starting at DestinationElement, we need to resize the array.
         If (DestinationElement + NumElementsToCopy - 1) > UBound(ResultArray) Then
            If DestinationElement + NumElementsToCopy > UBound(ResultArray) Then
               'Resize the ResultArray.
               If NumElementsToCopy + DestinationElement > UBound(ResultArray) Then
                  ReDim Preserve ResultArray(LBound(ResultArray) To UBound(ResultArray) + DestinationElement - 1)
               Else
                  ReDim Preserve ResultArray(LBound(ResultArray) To UBound(ResultArray) + NumElementsToCopy)
               End If
            Else
               'Resize the array to hold NumElementsToCopy starting at
               'DestinationElement.
               ReDim Preserve ResultArray(LBound(ResultArray) To UBound(ResultArray) + NumElementsToCopy - DestinationElement + 2)
            End If
         Else
            'The ResultArray is large enough to hold NumberOfElementToCopy
            'starting at DestinationElement. No need to resize the array.
         End If
      End If
   End If
   
   'Copy the elements from InputArray to ResultArray.
   'Note that there is no type compatibility checking when copying the elements.
   DestNdx = DestinationElement
   For SrcNdx = FirstElementToCopy To LastElementToCopy
      If IsObject(InputArray(SrcNdx)) Then
         Set ResultArray(DestNdx) = InputArray(DestNdx)
      Else
         On Error Resume Next
         ResultArray(DestNdx) = InputArray(SrcNdx)
         On Error GoTo 0
      End If
      DestNdx = DestNdx + 1
   Next
   
   CopyArraySubSetToArray = True

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CopyNonNothingObjectsToArray
'This function copies all objects that are not Nothing from SourceArray
'to ResultArray. ResultArray MUST be a dynamic array of type Object or Variant.
'E.g.,
'      Dim ResultArray() As Object 'Or
'      Dim ResultArray() as Variant
'
'ResultArray will be Erased and then resized to hold the non-Nothing elements
'from SourceArray. The LBound of ResultArray will be the same as the LBound
'of SourceArray, regardless of what its LBound was prior to calling this
'procedure.
'
'This function returns True if the operation was successful or False if an
'an error occurs. If an error occurs, a message box is displayed indicating
'the error. To suppress the message boxes, set the NoAlerts parameter to
'True.
'
'This function uses the following procedures.
'      IsArrayDynamic
'      IsArrayEmpty
'      NumberOfArrayDimensions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CopyNonNothingObjectsToArray( _
   ByRef SourceArray As Variant, _
   ByRef ResultArray As Variant, _
   Optional NoAlerts As Boolean = False _
      ) As Boolean
Attribute CopyNonNothingObjectsToArray.VB_ProcData.VB_Invoke_Func = " \n20"
   
   Dim ResNdx As LongPtr
   Dim InNdx  As LongPtr
   
   
   'Set the default return value
   CopyNonNothingObjectsToArray = False
   
   'Ensure SourceArray is an array
   If Not IsArray(SourceArray) Then
      If NoAlerts = False Then
         MsgBox "SourceArray is not an array."
      End If
      Exit Function
   End If
   'Ensure SourceArray is a single dimensional array
   Select Case NumberOfArrayDimensions(SourceArray)
      Case 0
         'Unallocated dynamic array. Not Allowed.
         If NoAlerts = False Then
            MsgBox "SourceArray is an unallocated array."
         End If
         Exit Function
      Case 1
         'Single-dimensional array. This is OK.
      Case Else
         'Multi-dimensional array. This is not allowed.
         If NoAlerts = False Then
            MsgBox "SourceArray is a multi-dimensional array. This is not allowed."
         End If
         Exit Function
   End Select
   'Ensure ResultArray is an array
   If Not IsArray(ResultArray) Then
      If NoAlerts = False Then
         MsgBox "ResultArray is not an array."
      End If
      Exit Function
   End If
   'Ensure ResultArray is an dynamic
   If Not IsArrayDynamic(ResultArray) Then
      If NoAlerts = False Then
         MsgBox "ResultArray is not a dynamic array."
      End If
      Exit Function
   End If
   'Ensure ResultArray is a single dimensional array
   Select Case NumberOfArrayDimensions(ResultArray)
      Case 0
         'Unallocated dynamic array. This is OK.
      Case 1
         'Single-dimensional array. This is OK.
      Case Else
         'Multi-dimensional array. This is not allowed.
         If NoAlerts = False Then
            MsgBox "SourceArray is a multi-dimensional array. This is not allowed."
         End If
         Exit Function
   End Select
   
   'Ensure that all the elements of SourceArray are in fact objects
   For InNdx = LBound(SourceArray) To UBound(SourceArray)
      If Not IsObject(SourceArray(InNdx)) Then
         If NoAlerts = False Then
            MsgBox "Element " & CStr(InNdx) & " of SourceArray is not an object."
         End If
         Exit Function
      End If
   Next
   
   'Erase the ResultArray. Since ResultArray is dynamic, this will relase the
   'memory used by ResultArray and return the array to an unallocated state.
   Erase ResultArray
   'Now, size ResultArray to the size of SourceArray. After moving all the
   'non-Nothing elements, we'll do another resize to get ResultArray to the
   'used size. This method allows us to avoid Redim Preserve for every element.
   ReDim ResultArray(LBound(SourceArray) To UBound(SourceArray))
   
   ResNdx = LBound(SourceArray)
   For InNdx = LBound(SourceArray) To UBound(SourceArray)
      If Not SourceArray(InNdx) Is Nothing Then
         Set ResultArray(ResNdx) = SourceArray(InNdx)
         ResNdx = ResNdx + 1
      End If
   Next
   'Now that we've copied all the non-Nothing elements from SourceArray to
   'ResultArray, we call Redim Preserve to resize the ResultArray to the size
   'actually used. Test ResNdx to see if we actually copied any elements.
   '
   'If ResNdx > LBound(SourceArray) then we copied at least one element out
   'of SourceArray.
   If ResNdx > LBound(SourceArray) Then
      ReDim Preserve ResultArray(LBound(ResultArray) To ResNdx - 1)
   'Otherwise, we didn't copy any elements from SourceArray
   '(all elements in SourceArray were Nothing). In this case, Erase ResultArray.
   Else
      Erase ResultArray
   End If
   
   CopyNonNothingObjectsToArray = True

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DataTypeOfArray
'
'Returns a VbVarType value indicating data type of the elements of
'Arr.
'
'The VarType of an array is the value vbArray plus the VbVarType value of the
'data type of the array. For example the VarType of an array of Longs is 8195,
'which equal to vbArray + vbLong. This code subtracts the value of vbArray to
'return the native data type.
'
'If Arr is a simple array, either single- or mulit-
'dimensional, the function returns the data type of the array. Arr
'may be an unallocated array. We can still get the data type of an unallocated
'array.
'
'If Arr is an array of arrays, the function returns vbArray. To retrieve
'the data type of a subarray, pass into the function one of the sub-arrays. E.g.,
'Dim R As VbVarType
'R = DataTypeOfArray(A(LBound(A)))
'
'This function support single and multidimensional arrays. It does not
'support user-defined types. If Arr is an array of empty variants (vbEmpty)
'it returns vbVariant
'
'Returns -1 if Arr is not an array.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DataTypeOfArray( _
   Arr As Variant _
      ) As VbVarType
Attribute DataTypeOfArray.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim Element As Variant
   Dim NumDimensions As LongPtr
   
   
   If Not IsArray(Arr) Then
      DataTypeOfArray = -1
      Exit Function
   End If
   
   'If the array is unallocated, we can still get its data type.
   'The result of VarType of an array is vbArray + the VarType of elements of
   'the array (e.g., the VarType of an array of Longs is 8195, which is
   'vbArray + vbLong). Thus, to get the basic data type of the array, we
   'subtract the value vbArray.
   If IsArrayEmpty(Arr) Then
      DataTypeOfArray = VarType(Arr) - vbArray
   Else
      'get the number of dimensions in the array
      NumDimensions = NumberOfArrayDimensions(Arr)
       'set variable Element to first element of the first dimension of the array
      If NumDimensions = 1 Then
         If IsObject(Arr(LBound(Arr))) Then
            DataTypeOfArray = vbObject
            Exit Function
         End If
         Element = Arr(LBound(Arr))
      Else
         If IsObject(Arr(LBound(Arr), 1)) Then
            DataTypeOfArray = vbObject
            Exit Function
         End If
         Element = Arr(LBound(Arr), 1)
      End If
      'if we were passed an array of arrays, IsArray(Element) will be true.
      'Therefore, return vbArray. If IsArray(Element) is false, we weren't
      'passed an array of arrays, so simply return the data type of Element.
      If IsArray(Element) Then
         DataTypeOfArray = vbArray
      Else
         If VarType(Element) = vbEmpty Then
            DataTypeOfArray = vbVariant
         Else
            DataTypeOfArray = VarType(Element)
         End If
      End If
   End If

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DeleteArrayElement
'This function deletes an element from InputArray, and shifts elements that are to the
'right of the deleted element to the left. If InputArray is a dynamic array, and the
'ResizeDynamic parameter is True, the array will be resized one element smaller. Otherwise,
'the right-most entry in the array is set to the default value appropriate to the data
'type of the array (0, vbNullString, Empty, or Nothing). If the array is an array of Variant
'types, the default data type is the data type of the last element in the array.
'The function returns True if the elememt was successfully deleted, or False if an error
'occurrred. This procedure works only on single-dimensional
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DeleteArrayElement( _
   InputArray As Variant, _
   ElementNumber As LongPtr, _
   Optional ResizeDynamic As Boolean = False _
      ) As Boolean
Attribute DeleteArrayElement.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim Ndx As LongPtr
   Dim VType As VbVarType
   
   
   'Set the default return value
   DeleteArrayElement = False
   
   If Not IsArray(InputArray) Then Exit Function
   If NumberOfArrayDimensions(InputArray) <> 1 Then Exit Function
   
   'Ensure we have a valid ElementNumber
   If ElementNumber < LBound(InputArray) Then Exit Function
   If ElementNumber > UBound(InputArray) Then Exit Function
   
   'Get the variable data type of the element we're deleting
   VType = VarType(InputArray(UBound(InputArray)))
   If VType >= vbArray Then
      VType = VType - vbArray
   End If
   'Shift everything to the left
   For Ndx = ElementNumber To UBound(InputArray) - 1
      InputArray(Ndx) = InputArray(Ndx + 1)
   Next
   'If ResizeDynamic is True, resize the array if it is dynamic
   If IsArrayDynamic(InputArray) Then
      If ResizeDynamic = True Then
         'Resize the array and get out.
         ReDim Preserve InputArray(LBound(InputArray) To UBound(InputArray) - 1)
         DeleteArrayElement = True
         Exit Function
      End If
   End If
   'Set the last element of the InputArray to the proper default value
   Select Case VType
      Case vbByte, vbInteger, vbLong, vbLongLong, vbSingle, vbDouble, vbDate, vbCurrency, vbDecimal
         InputArray(UBound(InputArray)) = 0
      Case vbString
         InputArray(UBound(InputArray)) = vbNullString
      Case vbArray, vbVariant, vbEmpty, vbError, vbNull, vbUserDefinedType
         InputArray(UBound(InputArray)) = Empty
      Case vbBoolean
         InputArray(UBound(InputArray)) = False
      Case vbObject
         Set InputArray(UBound(InputArray)) = Nothing
      Case Else
         InputArray(UBound(InputArray)) = 0
   End Select
   
   DeleteArrayElement = True

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FirstNonEmptyStringIndexInArray
'This returns the index into InputArray of the first non-empty string.
'This is generally used when InputArray is the result of a sort operation,
'which puts empty strings at the beginning of the array.
'Returns -1 if an error occurred or if the entire array is empty strings.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FirstNonEmptyStringIndexInArray( _
   InputArray As Variant _
      ) As LongPtr
Attribute FirstNonEmptyStringIndexInArray.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim Ndx As LongPtr
   
   
   'Set the default return value
   FirstNonEmptyStringIndexInArray = -1
   
   If Not IsArray(InputArray) Then Exit Function
   Select Case NumberOfArrayDimensions(InputArray)
      Case 0
         'indicates an unallocated dynamic array
         Exit Function
      Case 1
         'single dimensional array. OK.
      Case Else
         'multidimensional array. Invalid.
         Exit Function
   End Select
   
   For Ndx = LBound(InputArray) To UBound(InputArray)
      If InputArray(Ndx) <> vbNullString Then
         FirstNonEmptyStringIndexInArray = Ndx
         Exit Function
      End If
   Next
   
   FirstNonEmptyStringIndexInArray = -1

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'InsertElementIntoArray
'This function inserts an element with a value of Value into InputArray at locatation Index.
'InputArray must be a dynamic array. The Value is stored in location Index, and everything
'to the right of Index is shifted to the right. The array is resized to make room for
'the new element. The value of Index must be greater than or equal to the LBound of
'InputArray and less than or equal to UBound+1. If Index is UBound+1, the Value is
'placed at the end of the array.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function InsertElementIntoArray( _
   InputArray As Variant, _
   Index As LongPtr, _
   Value As Variant _
      ) As Boolean
Attribute InsertElementIntoArray.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim Ndx As LongPtr
   
   
   'Set the default return value
   InsertElementIntoArray = False
   
   If Not IsArray(InputArray) Then Exit Function
   If Not IsArrayDynamic(InputArray) Then Exit Function
   If Not IsArrayAllocated(InputArray) Then Exit Function
   If NumberOfArrayDimensions(InputArray) <> 1 Then Exit Function
   
   'Ensure Index is a valid element index. We allow Index to be equal to
   'UBound + 1 to facilitate inserting a value at the end of the array. E.g.,
   'InsertElementIntoArray(Arr,UBound(Arr)+1,123) will insert 123 at the end
   'of the array.
   If (Index < LBound(InputArray)) Or (Index > UBound(InputArray) + 1) Then
      Exit Function
   End If
   
   'Resize the array
   ReDim Preserve InputArray(LBound(InputArray) To UBound(InputArray) + 1)
   'First, we set the newly created last element of InputArray to Value.
   'This is done to trap an error 13, type mismatch. This last entry will be
   'overwritten when we shift elements to the right, and the Value will be
   'inserted at Index.
   On Error Resume Next
   err.Clear
   InputArray(UBound(InputArray)) = Value
   If err.Number <> 0 Then
      'An error occurred, most likely an error 13, type mismatch.
      'Redim the array back to its original size and exit the function.
      ReDim Preserve InputArray(LBound(InputArray) To UBound(InputArray) - 1)
      Exit Function
   End If
   'Shift everything to the right
   For Ndx = UBound(InputArray) To Index + 1 Step -1
      InputArray(Ndx) = InputArray(Ndx - 1)
   Next
   
   'Insert Value at Index
   InputArray(Index) = Value
       
   InsertElementIntoArray = True

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IsArrayAllEmpty
'Returns True if the array contains all default values for its
'data type:
'  Variable Type           Value
'  -------------           -------------------
'  Variant                 Empty
'  String                  vbNullString
'  Numeric                 0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsArrayAllDefault( _
   InputArray As Variant _
      ) As Boolean
Attribute IsArrayAllDefault.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim Ndx As LongPtr
   Dim DefaultValue As Variant
   
   
   'Set the default return value
   IsArrayAllDefault = False
   
   If Not IsArray(InputArray) Then Exit Function
   'Ensure array is allocated. An unallocated is considered to be all the same
   'type. Return True.
   If Not IsArrayAllocated(InputArray) Then
      IsArrayAllDefault = True
      Exit Function
   End If
       
   'Test the type of variable
   Select Case VarType(InputArray)
      Case vbArray + vbVariant
         DefaultValue = Empty
      Case vbArray + vbString
         DefaultValue = vbNullString
      Case Is > vbArray
         DefaultValue = 0
   End Select
   For Ndx = LBound(InputArray) To UBound(InputArray)
      If IsObject(InputArray(Ndx)) Then
         If Not InputArray(Ndx) Is Nothing Then Exit Function
      Else
         If VarType(InputArray(Ndx)) <> vbEmpty Then
            If InputArray(Ndx) <> DefaultValue Then Exit Function
         End If
      End If
   Next
   
   'If we make it up to here, the array is all defaults.
   IsArrayAllDefault = True

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IsArrayAllNumeric
'This function returns True if Arr is entirely numeric. False otherwise. The AllowNumericStrings
'parameter indicates whether strings containing numeric data are considered numeric. If this
'parameter is True, a numeric string is considered a numeric variable. If this parameter is
'omitted or False, a numeric string is not considered a numeric variable.
'Variants that are numeric or Empty are allowed. Variants that are arrays, objects, or
'non-numeric data are not allowed.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'2do:
'- This function only tests a vector
'  --> so a better name would be 'IsVectorAllNumeric'
'- There is no test for the 1-dimensionality
'- What is the benefit of this function over 'IsVariantArrayNumeric'?
Public Function IsArrayAllNumeric( _
   Arr As Variant, _
   Optional AllowNumericStrings As Boolean = False _
      ) As Boolean
Attribute IsArrayAllNumeric.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim Ndx As LongPtr
   
   
   'Set the default return value
   IsArrayAllNumeric = False
   
   If Not IsArray(Arr) Then Exit Function
   'Ensure Arr is allocated (non-empty)
'---
'2do: what is really needed: 'IsArrayEmpty' or 'IsArrayAllocated'?
'---
   If IsArrayEmpty(Arr) Then Exit Function
   
   'Loop through the array
   For Ndx = LBound(Arr) To UBound(Arr)
      Select Case VarType(Arr(Ndx))
         Case vbInteger, vbLong, vbLongLong, vbDouble, vbSingle, vbCurrency, vbDecimal, vbEmpty
            'all valid numeric types
         Case vbString
            'For strings, check the AllowNumericStrings parameter.
            'If True and the element is a numeric string, allow it.
            'If it is a non-numeric string, exit with False.
            'If AllowNumericStrings is False, all strings, even
            'numeric strings, will cause a result of False.
            If AllowNumericStrings = True Then
               'Allow numeric strings.
               If Not IsNumeric(Arr(Ndx)) Then Exit Function
            Else
               Exit Function
            End If
         Case vbVariant
            'For Variants, disallow Arrays and Objects.
            'If the element is not an array or an object, test whether it is
            'numeric. Allow numeric Variants.
            If IsArray(Arr(Ndx)) Then Exit Function
            If IsObject(Arr(Ndx)) Then Exit Function
            If Not IsNumeric(Arr(Ndx)) Then Exit Function
         Case Else
            'any other data type returns False
            Exit Function
      End Select
   Next
   
   IsArrayAllNumeric = True

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IsArrayAllocated
'Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
'sized with Redim) or FALSE if the array is not allocated (a dynamic that has not yet
'been sized with Redim, or a dynamic array that has been Erased). Static arrays are always
'allocated.
'
'The VBA IsArray function indicates whether a variable is an array, but it does not
'distinguish between allocated and unallocated arrays. It will return TRUE for both
'allocated and unallocated arrays. This function tests whether the array has actually
'been allocated.
'
'This function is just the reverse of IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsArrayAllocated( _
   Arr As Variant _
      ) As Boolean
Attribute IsArrayAllocated.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim N As LongPtr
   
   
   'Set the default return value
   IsArrayAllocated = False
   
   On Error Resume Next
   
   If Not IsArray(Arr) Then Exit Function
   
   'Attempt to get the UBound of the array. If the array has not been allocated,
   'an error will occur. Test Err.Number to see if an error occurred.
   N = UBound(Arr, 1)
   If (err.Number = 0) Then
       'Under some circumstances, if an array is not allocated, Err.Number
       'will be 0. To acccomodate this case, we test whether LBound <= Ubound.
       'If this is True, the array is allocated. Otherwise, the array is not
       'allocated.
      If LBound(Arr) <= UBound(Arr) Then
         'no error. array has been allocated
         IsArrayAllocated = True
      End If
   Else
      'error. unallocated array
   End If

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IsArrayDynamic
'This function returns TRUE or FALSE indicating whether Arr is a dynamic array.
'Note that if you attempt to ReDim a static array in the same procedure in which it is
'declared, you'll get a compiler error and your code won't run at all.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsArrayDynamic( _
   ByRef Arr As Variant _
      ) As Boolean
Attribute IsArrayDynamic.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim LUBound As LongPtr
   
   
   'Set the default return value
   IsArrayDynamic = False
   
   If Not IsArray(Arr) Then Exit Function
   
   'If the array is empty, it hasn't been allocated yet, so we know it must be
   'a dynamic array
   If IsArrayEmpty(Arr) Then
      IsArrayDynamic = True
      Exit Function
   End If
   
   'Save the UBound of Arr.
   'This value will be used to restore the original UBound if Arr is a
   'single-dimensional dynamic array. Unused if Arr is multi-dimensional,
   'or if Arr is a static array.
   LUBound = UBound(Arr)
   
   On Error Resume Next
   err.Clear
   
   'Attempt to increase the UBound of Arr and test the value of Err.Number.
   'If Arr is a static array, either single- or multi-dimensional, we'll get a
   'C_ERR_ARRAY_IS_FIXED_OR_LOCKED error. In this case, return FALSE.
   '
   'If Arr is a single-dimensional dynamic array, we'll get C_ERR_NO_ERROR error.
   '
   'If Arr is a multi-dimensional dynamic array, we'll get a
   'C_ERR_SUBSCRIPT_OUT_OF_RANGE error.
   '
   'For either C_NO_ERROR or C_ERR_SUBSCRIPT_OUT_OF_RANGE, return TRUE.
   'For C_ERR_ARRAY_IS_FIXED_OR_LOCKED, return FALSE.
   ReDim Preserve Arr(LBound(Arr) To LUBound + 1)
   Select Case err.Number
      Case C_ERR_NO_ERROR
         'We successfully increased the UBound of Arr.
         'Do a ReDim Preserve to restore the original UBound.
         ReDim Preserve Arr(LBound(Arr) To LUBound)
         IsArrayDynamic = True
      Case C_ERR_SUBSCRIPT_OUT_OF_RANGE
         'Arr is a multi-dimensional dynamic array.
         'Return True.
         IsArrayDynamic = True
      Case C_ERR_ARRAY_IS_FIXED_OR_LOCKED
         'Arr is a static single- or multi-dimensional array.
         'Return False
         IsArrayDynamic = False
      Case Else
         'We should never get here.
         'Some unexpected error occurred. Be safe and return False.
         IsArrayDynamic = False
   End Select

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IsArrayEmpty
'This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
'
'The VBA IsArray function indicates whether a variable is an array, but it does not
'distinguish between allocated and unallocated arrays. It will return TRUE for both
'allocated and unallocated arrays. This function tests whether the array has actually
'been allocated.
'
'This function is really the reverse of IsArrayAllocated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsArrayEmpty( _
   Arr As Variant _
      ) As Boolean
Attribute IsArrayEmpty.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim LB As LongPtr
   Dim UB As LongPtr
   
   
   err.Clear
   On Error Resume Next
   If Not IsArray(Arr) Then
      'we weren't passed an array, return True
      IsArrayEmpty = True
   End If
   
   'Attempt to get the UBound of the array. If the array is
   'unallocated, an error will occur.
   UB = UBound(Arr, 1)
   If (err.Number <> 0) Then
      IsArrayEmpty = True
   Else
      'On rare occassion, under circumstances I cannot reliably replictate,
      'Err.Number will be 0 for an unallocated, empty array.
      'On these occassions, LBound is 0 and UBoung is -1.
      'To accomodate the weird behavior, test to see if LB > UB.
      'If so, the array is not allocated.
      err.Clear
      LB = LBound(Arr)
      If LB > UB Then
         IsArrayEmpty = True
      Else
         IsArrayEmpty = False
      End If
   End If

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IsArrayObjects
'Returns True if InputArray is entirely objects (Nothing objects are
'optionally allowed -- default it true, allow Nothing objects). Set the
'AllowNothing to true or false to indicate whether Nothing objects
'are allowed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsArrayObjects( _
   InputArray As Variant, _
   Optional AllowNothing As Boolean = True _
      ) As Boolean
Attribute IsArrayObjects.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim Ndx As LongPtr
   
   
   'Set the default return value
   IsArrayObjects = False
   
   If Not IsArray(InputArray) Then Exit Function
   
   'Ensure we have a single dimensional array
   Select Case NumberOfArrayDimensions(InputArray)
      Case 0
         'Unallocated dynamic array. Not allowed.
         Exit Function
      Case 1
         'OK
      Case Else
         'Multi-dimensional array. Not allowed.
         Exit Function
   End Select
   
   For Ndx = LBound(InputArray) To UBound(InputArray)
      If Not IsObject(InputArray(Ndx)) Then Exit Function
      If InputArray(Ndx) Is Nothing Then
         If AllowNothing = False Then Exit Function
      End If
   Next
   
   IsArrayObjects = True

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IsNumericDataType
'
'This function returns TRUE or FALSE indicating whether the data
'type of a variable is a numeric data type. It will return TRUE
'for all of the following data types:
'      vbCurrency
'      vbDecimal
'      vbDouble
'      vbInteger
'      vbLong, vbLongLong
'      vbSingle
'
'It will return FALSE for any other data type, including empty Variants and objects.
'If TestVar is an allocated array, it will test data type of the array
'and return TRUE or FALSE for that data type. If TestVar is an allocated
'array, it tests the data type of the first element of the array. If
'TestVar is an array of Variants, the function will indicate only whether
'the first element of the array is numeric. Other elements of the array
'may not be numeric data types. To test an entire array of variants
'to ensure they are all numeric data types, use the IsVariantArrayNumeric
'function.
'
'It will return FALSE for any other data type. Use this procedure
'instead of VBA's IsNumeric function because IsNumeric will return
'TRUE if the variable is a string containing numeric data. This
'will cause problems with code like
'       Dim V1 As Variant
'       Dim V2 As Variant
'       V1 = "1"
'       V2 = "2"
'       If IsNumeric(V1) Then
'           If IsNumeric(V2) Then
'               Debug.Print V1 + V2
'           End If
'       End If
'
'The output of the Debug.Print statement will be "12", not 3,
'because V1 and V2 are strings and the '+'operator acts like
'the '&'operator when used with strings. This can lead to
'unexpected results.
'
'IsNumeric should only be used to test strings for numeric content
'when converting a string value to a numeric variable.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsNumericDataType( _
   TestVar As Variant _
      ) As Boolean
Attribute IsNumericDataType.VB_ProcData.VB_Invoke_Func = " \n20"
   
   Dim Element As Variant
   Dim NumDims As LongPtr
   
   
   'Set the default return value
   IsNumericDataType = False
   
   If IsArray(TestVar) Then
      NumDims = NumberOfArrayDimensions(TestVar)
'---
'2do:
'- is a change needed here? First test, if 'IsVariantArrayNumeric' is supposed
'  to handle this!
'---
      If NumDims > 1 Then
         'this procedure does not support multi-dimensional arrays
         Exit Function
      End If
      If IsArrayAllocated(TestVar) Then
'---
'2do:
'- is it intentional to test only the first element of 'TestVar'?
'  --> according to the functions description yes ...
'---
         Element = TestVar(LBound(TestVar))
         Select Case VarType(Element)
            Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbLongLong, vbSingle
               IsNumericDataType = True
               Exit Function
            Case Else
               Exit Function
         End Select
      Else
         Select Case VarType(TestVar) - vbArray
            Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbLongLong, vbSingle
               IsNumericDataType = True
               Exit Function
            Case Else
               Exit Function
         End Select
      End If
   End If
   
   Select Case VarType(TestVar)
      Case vbCurrency, vbDecimal, vbDouble, vbInteger, vbLong, vbLongLong, vbSingle
         IsNumericDataType = True
      Case Else
         IsNumericDataType = False
   End Select

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IsVariantArrayConsistent
'
'This returns TRUE or FALSE indicating whether an array of variants
'contains all the same data types. Returns FALSE under the following
'circumstances:
'      Arr is not an array
'      Arr is an array but is unallocated
'      Arr is a multidimensional array
'      Arr is allocated but does not contain consistant data types.
'
'If Arr is an array of objects, objects that are Nothing are ignored.
'As long as all non-Nothing objects are the same object type, the
'function returns True.
'
'It returns TRUE if all the elements of the array have the same
'data type. If Arr is an array of a specific data types, not variants,
'(E.g., Dim V(1 To 3) As LongPtr), the function will return True. If
'an array of variants contains an uninitialized element (VarType =
'vbEmpty) that element is skipped and not used in the comparison. The
'reasoning behind this is that an empty variable will return the
'data type of the variable to which it is assigned (e.g., it will
'return vbNullString to a String and 0 to a Double).
'
'The function does not support arrays of User Defined Types.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsVariantArrayConsistent( _
   Arr As Variant _
      ) As Boolean
Attribute IsVariantArrayConsistent.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim FirstDataType As VbVarType
   Dim Ndx As LongPtr
   
   
   'Set the default return value
   IsVariantArrayConsistent = False
   
   If Not IsArray(Arr) Then Exit Function
   If Not IsArrayAllocated(Arr) Then Exit Function

   'Exit with false on multi-dimensional arrays
'---
'2do: can this be changed if still true?
'---
   If NumberOfArrayDimensions(Arr) <> 1 Then Exit Function
   
   'Test if we have an array of a specific type rather than Variants. If so,
   'return TRUE and get out.
   If (VarType(Arr) <= vbArray) And _
       (VarType(Arr) <> vbVariant) Then
      IsVariantArrayConsistent = True
      Exit Function
   End If
   
   'Get the data type of the first element
   FirstDataType = VarType(Arr(LBound(Arr)))
   'Loop through the array and exit if a differing data type if found.
   For Ndx = LBound(Arr) + 1 To UBound(Arr)
      If VarType(Arr(Ndx)) <> vbEmpty Then
         If IsObject(Arr(Ndx)) Then
            If Not Arr(Ndx) Is Nothing Then
               If VarType(Arr(Ndx)) <> FirstDataType Then Exit Function
            End If
         Else
            If VarType(Arr(Ndx)) <> FirstDataType Then Exit Function
         End If
      End If
   Next
   
   'If we make it up to here, then the array is consistent
   IsVariantArrayConsistent = True

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IsVariantArrayNumeric
'
'This function returns TRUE if all the elements of an array of
'variants are numeric data types. They need not all be the same data
'type. You can have a mix of Integer, Longs, Doubles, and Singles.
'As long as they are all numeric data types, the function will
'return TRUE. If a non-numeric data type is encountered, the
'function will return FALSE. Also, it will return FALSE if
'TestArray is not an array, or if TestArray has not been
'allocated. TestArray may be a multi-dimensional array. This
'procedure uses the IsNumericDataType function to determine whether
'a variable is a numeric data type. If there is an uninitialized
'variant (VarType = vbEmpty) in the array, it is skipped and not
'used in the comparison (i.e., Empty is considered a valid numeric
'data type since you can assign a number to it).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'2do:
'- add "Optional AllowNumericStrings As Boolean = False" from
'  'IsArrayAllNumeric'?
Public Function IsVariantArrayNumeric( _
   TestArray As Variant _
      ) As Boolean
Attribute IsVariantArrayNumeric.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim Ndx As LongPtr
   Dim DimNdx As LongPtr
   Dim NumDims As LongPtr
   
   
   'Set the default return value
   IsVariantArrayNumeric = False
   
   If Not IsArray(TestArray) Then Exit Function
   If Not IsArrayAllocated(TestArray) Then Exit Function
   
'---
'2do:
'- can this be simplified with a simple
'     Dim item as Variant
'     For Each item in TestArray
'        ...
'     Next
'  ?
'  (see <https://excelmacromastery.com/excel-vba-array/>)
'---
   NumDims = NumberOfArrayDimensions(TestArray)
   Select Case NumDims
      Case 1
         For Ndx = LBound(TestArray) To UBound(TestArray)
            If IsObject(TestArray(Ndx)) Then Exit Function
              
            If VarType(TestArray(Ndx)) <> vbEmpty Then
               If Not IsNumericDataType(TestArray(Ndx)) Then
                  Exit Function
               End If
            End If
         Next
      Case 2
         For DimNdx = LBound(TestArray, 2) To UBound(TestArray, 2)
            For Ndx = LBound(TestArray, DimNdx) To UBound(TestArray, DimNdx)
               If VarType(TestArray(Ndx, DimNdx)) <> vbEmpty Then
                  If Not IsNumericDataType(TestArray(Ndx, DimNdx)) Then
                     Exit Function
                  End If
               End If
            Next
         Next
      Case Else
         'currently there is no handler for "higher"-dimensional arrays
         Exit Function
   End Select
   
   'If we made it up to here, then the array is entirely numeric
   IsVariantArrayNumeric = True

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This procedure takes the SORTED array InputArray, which, if sorted in
'ascending order, will have all empty strings at the front of the array.
'This procedure moves those strings to the end of the array, shifting
'the non-empty strings forward in the array.
'Note that InputArray MUST be sorted in ascending order.
'Returns True if the array was correctly shifted (if necessary) and False
'if an error occurred.
'
'This function uses the following functions.
'      FirstNonEmptyStringIndexInArray
'      NumberOfArrayDimensions
'      IsArrayAllocated
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MoveEmptyStringsToEndOfArray( _
   InputArray As Variant _
      ) As Boolean
Attribute MoveEmptyStringsToEndOfArray.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim Temp As String
   Dim Ndx As LongPtr
   Dim Ndx2 As LongPtr
   Dim NonEmptyNdx As LongPtr
   Dim FirstNonEmptyNdx As LongPtr
   
   
   'Set the default return value
   MoveEmptyStringsToEndOfArray = False

   If Not IsArray(InputArray) Then Exit Function
   If Not IsArrayAllocated(InputArray) Then Exit Function
   
   
   FirstNonEmptyNdx = FirstNonEmptyStringIndexInArray(InputArray)
   If FirstNonEmptyNdx <= LBound(InputArray) Then
      'No empty strings at the beginning of the array. Get out now.
      MoveEmptyStringsToEndOfArray = True
      Exit Function
   End If
   
   
   'Loop through the array, swapping vbNullStrings at the beginning with
   'values at the end.
   NonEmptyNdx = FirstNonEmptyNdx
   For Ndx = LBound(InputArray) To UBound(InputArray)
      If InputArray(Ndx) = vbNullString Then
         InputArray(Ndx) = InputArray(NonEmptyNdx)
         InputArray(NonEmptyNdx) = vbNullString
         NonEmptyNdx = NonEmptyNdx + 1
         If NonEmptyNdx > UBound(InputArray) Then
            Exit For
         End If
      End If
   Next
   'Set entires (Ndx+1) to UBound(InputArray) to vbNullStrings
   For Ndx2 = Ndx + 1 To UBound(InputArray)
      InputArray(Ndx2) = vbNullString
   Next
   
   MoveEmptyStringsToEndOfArray = True

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'NumberOfArrayDimensions
'This function returns the number of dimensions of an array. An unallocated dynamic array
'has 0 dimensions. This condition can also be tested with IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NumberOfArrayDimensions( _
   Arr As Variant _
      ) As Integer
Attribute NumberOfArrayDimensions.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim Ndx As Integer
   Dim Res As Integer
   
   
   On Error Resume Next
   'Loop, increasing the dimension index Ndx, until an error occurs.
   'An error will occur when Ndx exceeds the number of dimension in the array.
   'Return Ndx - 1.
   Do
      Ndx = Ndx + 1
      Res = UBound(Arr, Ndx)
   Loop Until err.Number <> 0
   
   NumberOfArrayDimensions = Ndx - 1

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'NumElements
'Returns the number of elements in the specified dimension (Dimension) of the array in
'Arr. If you omit Dimension, the first dimension is used. The function will return
'0 under the following circumstances:
'    Arr is not an array, or
'    Arr is an unallocated array, or
'    Dimension is greater than the number of dimension of Arr, or
'    Dimension is less than 1.
'
'This function does not support arrays of user-defined Type variables.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NumElements( _
   Arr As Variant, _
   Optional Dimension As Integer = 1 _
      ) As LongPtr
Attribute NumElements.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim NumDimensions As LongPtr
   
   
   'Set the default return value
   NumElements = 0
   
   If Not IsArray(Arr) Then Exit Function
   If IsArrayEmpty(Arr) = True Then Exit Function
   
   'ensure that Dimension is at least 1
   If Dimension < 1 Then Exit Function
   
   'check if 'Dimension is not larger than 'NumDimensions'
   NumDimensions = NumberOfArrayDimensions(Arr)
   If NumDimensions < Dimension Then Exit Function
   
   'returns the number of elements in the array
   NumElements = UBound(Arr, Dimension) - LBound(Arr, Dimension) + 1

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ResetVariantArrayToDefaults
'This resets all the elements of an array of Variants back to their appropriate
'default values. The elements of the array may be of mixed types (e.g., some Longs,
'some Objects, some Strings, etc). Each data type will be set to the appropriate
'default value (0, vbNullString, Empty, or Nothing). It returns True if the
'array was set to defautls, or False if an error occurred. InputArray must be
'an allocated single-dimensional array. This function differs from the Erase
'function in that it preserves the original data types, while Erase sets every
'element to Empty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ResetVariantArrayToDefaults( _
   InputArray As Variant _
      ) As Boolean
Attribute ResetVariantArrayToDefaults.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim Ndx As LongPtr
   
   'Set the default return value
   ResetVariantArrayToDefaults = False
   
   If Not IsArray(InputArray) Then Exit Function
   If NumberOfArrayDimensions(InputArray) <> 1 Then Exit Function
   
   For Ndx = LBound(InputArray) To UBound(InputArray)
      SetVariableToDefault InputArray(Ndx)
   Next
   
   ResetVariantArrayToDefaults = True

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ReverseArrayInPlace
'This procedure reverses the order of an array in place -- this is, the array variable
'in the calling procedure is reversed. This works only on single-dimensional arrays
'of simple data types (String, Single, Double, Integer, Long). It will not work
'on arrays of objects. Use ReverseArrayOfObjectsInPlace to reverse an array of objects.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ReverseArrayInPlace( _
   InputArray As Variant, _
   Optional NoAlerts As Boolean = False _
      ) As Boolean
Attribute ReverseArrayInPlace.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim Temp As Variant
   Dim Ndx As LongPtr
   Dim Ndx2 As LongPtr
   
   
   'Set the default return value
   ReverseArrayInPlace = False
   
   'ensure we have an array
   If Not IsArray(InputArray) Then
      If NoAlerts = False Then
         MsgBox "The InputArray parameter is not an array."
      End If
      Exit Function
   End If
   
   'Test the number of dimensions of the InputArray. If 0, we have an empty,
   'unallocated array. Get out with an error message. If greater than one, we
   'have a multi-dimensional array, which is not allowed. Only an allocated
   '1-dimensional array is allowed.
   Select Case NumberOfArrayDimensions(InputArray)
      Case 0
         If NoAlerts = False Then
            MsgBox "The input array is an empty, unallocated array."
         End If
         Exit Function
      Case 1
         'ok
      Case Else
         If NoAlerts = False Then
            MsgBox "The input array is multi-dimensional. ReverseArrayInPlace works only " & _
                      "on single-dimensional arrays."
         End If
         Exit Function
   End Select
   
   Ndx2 = UBound(InputArray)
   
   'loop from the LBound of InputArray to the midpoint of InputArray
   For Ndx = LBound(InputArray) To ((UBound(InputArray) - LBound(InputArray) + 1) \ 2) - 1
      'swap the elements
      Temp = InputArray(Ndx)
      InputArray(Ndx) = InputArray(Ndx2)
      InputArray(Ndx2) = Temp
      'decrement the upper index
      Ndx2 = Ndx2 - 1
   Next
   
   ReverseArrayInPlace = True

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ReverseArrayOfObjectsInPlace
'This procedure reverses the order of an array in place -- this is, the array variable
'in the calling procedure is reversed. This works only with arrays of objects. It does
'not work on simple variables. Use ReverseArrayInPlace for simple variables. An error
'will occur if an element of the array is not an object.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ReverseArrayOfObjectsInPlace( _
   InputArray As Variant, _
   Optional NoAlerts As Boolean = False _
      ) As Boolean
Attribute ReverseArrayOfObjectsInPlace.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim Temp As Variant
   Dim Ndx As LongPtr
   Dim Ndx2 As LongPtr
   
   
   'Set the default return value
   ReverseArrayOfObjectsInPlace = False
   
   'ensure we have an array
   If Not IsArray(InputArray) Then
      If NoAlerts = False Then
         MsgBox "The InputArray parameter is not an array."
      End If
      Exit Function
   End If
   
   'Test the number of dimensions of the InputArray. If 0, we have an empty,
   'unallocated array. Get out with an error message. If greater than one, we
   'have a multi-dimensional array, which is not allowed. Only an allocated
   '1-dimensional array is allowed.
   Select Case NumberOfArrayDimensions(InputArray)
      Case 0
         If NoAlerts = False Then
            MsgBox "The input array is an empty, unallocated array."
         End If
         Exit Function
      Case 1
         'ok
      Case Else
         If NoAlerts = False Then
            MsgBox "The input array is multi-dimensional. " & _
               "ReverseArrayInPlace works only on single-dimensional arrays."
         End If
         Exit Function
   End Select
   
   Ndx2 = UBound(InputArray)
   
   'ensure the entire array consists of objects (Nothing objects are allowed)
   For Ndx = LBound(InputArray) To UBound(InputArray)
      If Not IsObject(InputArray(Ndx)) Then
         If NoAlerts = False Then
            MsgBox "Array item " & CStr(Ndx) & " is not an object."
         End If
         Exit Function
      End If
   Next
   
   'loop from the LBound of InputArray to the midpoint of InputArray
   For Ndx = LBound(InputArray) To ((UBound(InputArray) - LBound(InputArray) + 1) \ 2)
      Set Temp = InputArray(Ndx)
      Set InputArray(Ndx) = InputArray(Ndx2)
      Set InputArray(Ndx2) = Temp
      'decrement the upper index
      Ndx2 = Ndx2 - 1
   Next
   
   ReverseArrayOfObjectsInPlace = True

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SetObjectArrrayToNothing
'This sets all the elements of InputArray to Nothing. Use this function
'rather than Erase because if InputArray is an array of Variants, Erase
'will set each element to Empty, not Nothing, and the element will cease
'to be an object.
'
'The function returns True if successful, False otherwise.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SetObjectArrayToNothing( _
   InputArray As Variant _
      ) As Boolean
Attribute SetObjectArrayToNothing.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim N As LongPtr
   
   
   'Set the default return value
   SetObjectArrayToNothing = False
   
   If Not IsArray(InputArray) Then Exit Function
   If NumberOfArrayDimensions(InputArray) <> 1 Then Exit Function
   
   'Ensure the array is allocated and that each element is an object (or Nothing).
   'If the array is not allocated, return True. We do this test before setting
   'any element to Nothing so we don't end up with an array that is a mix of
   'Empty and Nothing values. This means looping through the array twice, but
   'it ensures all or none of the elements get set to Nothing.
   If IsArrayAllocated(InputArray) Then
      For N = LBound(InputArray) To UBound(InputArray)
         If Not IsObject(InputArray(N)) Then Exit Function
      Next
   Else
      SetObjectArrayToNothing = True
      Exit Function
   End If
   
   'Set each element of InputArray to Nothing
   For N = LBound(InputArray) To UBound(InputArray)
      Set InputArray(N) = Nothing
   Next
   
   SetObjectArrayToNothing = True

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'AreDataTypesCompatible
'This function determines if SourceVar is compatiable with DestVar. If the two
'data types are the same, they are compatible. If the value of SourceVar can
'be stored in DestVar with no loss of precision or an overflow, they are compatible.
'For example, if DestVar is a Long and SourceVar is an Integer, they are compatible
'because an integer can be stored in a Long with no loss of information. If DestVar
'is a Long and SourceVar is a Double, they are not compatible because information
'will be lost converting from a Double to a Long (the decimal portion will be lost).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SP: - changed order of arguments (to be consistent: "Source" first, then "Dest")
Public Function AreDataTypesCompatible( _
   SourceVar As Variant, _
   DestVar As Variant _
      ) As Boolean
Attribute AreDataTypesCompatible.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim SVType As VbVarType
   Dim DVType As VbVarType
   
   
   'Set the default return value
   AreDataTypesCompatible = False
   
   'If DestVar is an array, get the type of array. If it is an array its
   'VarType is vbArray + VarType(element) so we subtract vbArray to get then
   'data type of the aray. E.g., the VarType of an array of Longs is
   '8195 = vbArray + vbLong,
   '8195 - vbArray = vbLong (=3).
   If IsArray(DestVar) Then
      DVType = VarType(DestVar) - vbArray
   Else
      DVType = VarType(DestVar)
   End If
   'If SourceVar is an array, get the type of array
   If IsArray(SourceVar) Then
      SVType = VarType(SourceVar) - vbArray
   Else
      SVType = VarType(SourceVar)
   End If
   
   'If one variable is an array and the other is not an array, they are incompatible
   If ((IsArray(DestVar) = True) And (IsArray(SourceVar) = False) Or _
       (IsArray(DestVar) = False) And (IsArray(SourceVar) = True)) Then
      Exit Function
   End If
   
'---
'2do:
'- there is no loop, so can't all "Exit Function"s be safely removed, because
'  the function would be exited anyway after the corresponding line?
'---
   '''Test the data type of DestVar and return a result if SourceVar is
   '''compatible with that type.
   'The the variable types are the same, they are compatible
   If SVType = DVType Then
      AreDataTypesCompatible = True
      Exit Function
   'If the data types are not the same, determine whether they are compatible
   Else
      Select Case DVType
         Case vbInteger
            Select Case SVType
               Case vbInteger
                  AreDataTypesCompatible = True
                  Exit Function
               Case Else
                  Exit Function
            End Select
         Case vbLong, vbLongLong
            Select Case SVType
               Case vbInteger, vbLong, vbLongLong
                  AreDataTypesCompatible = True
                  Exit Function
               Case Else
                  Exit Function
            End Select
         Case vbSingle
            Select Case SVType
               Case vbInteger, vbLong, vbLongLong, vbSingle
                  AreDataTypesCompatible = True
                  Exit Function
               Case Else
                  Exit Function
            End Select
         Case vbDouble
            Select Case SVType
               Case vbInteger, vbLong, vbLongLong, vbSingle, vbDouble
                  AreDataTypesCompatible = True
                  Exit Function
               Case Else
                  Exit Function
            End Select
         Case vbString
            Select Case SVType
               Case vbString
                  AreDataTypesCompatible = True
                  Exit Function
               Case Else
                  Exit Function
            End Select
         Case vbObject
            Select Case SVType
               Case vbObject
                  AreDataTypesCompatible = True
                  Exit Function
               Case Else
                  Exit Function
            End Select
         Case vbBoolean
            Select Case SVType
               Case vbBoolean, vbInteger
                  AreDataTypesCompatible = True
                  Exit Function
               Case Else
                  Exit Function
            End Select
         Case vbByte
            Select Case SVType
               Case vbByte
                  AreDataTypesCompatible = True
                  Exit Function
               Case Else
                  Exit Function
            End Select
         Case vbCurrency
            Select Case SVType
               Case vbInteger, vbLong, vbLongLong, vbSingle, vbDouble
                  AreDataTypesCompatible = True
                  Exit Function
               Case Else
                  Exit Function
            End Select
         Case vbDecimal
            Select Case SVType
               Case vbInteger, vbLong, vbLongLong, vbSingle, vbDouble
                  AreDataTypesCompatible = True
                  Exit Function
               Case Else
                  Exit Function
            End Select
         Case vbDate
            Select Case SVType
               Case vbLong, vbLongLong, vbSingle, vbDouble
                  AreDataTypesCompatible = True
                  Exit Function
               Case Else
                  Exit Function
            End Select
         Case vbEmpty
            Select Case SVType
               Case vbVariant
                  AreDataTypesCompatible = True
                  Exit Function
               Case Else
                  Exit Function
            End Select
         Case vbError
            AreDataTypesCompatible = False
            Exit Function
         Case vbNull
            Exit Function
         Case vbObject
            Select Case SVType
               Case vbObject
                  AreDataTypesCompatible = True
                  Exit Function
               Case Else
                  Exit Function
            End Select
          Case vbVariant
            AreDataTypesCompatible = True
            Exit Function
      End Select
   End If

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SetVariableToDefault
'This procedure sets Variable to the appropriate default
'value for its data type. Note that it cannot change User-Defined
'Types.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetVariableToDefault( _
   ByRef Variable As Variant _
)
Attribute SetVariableToDefault.VB_ProcData.VB_Invoke_Func = " \n20"

   'We test with IsObject here so that the object itself, not the default
   'property of the object, is evaluated.
   If IsObject(Variable) Then
      Set Variable = Nothing
   Else
      Select Case VarType(Variable)
         Case Is >= vbArray
            'The VarType of an array is equal to vbArray + VarType(ArrayElement).
            'Here we check for anything >= vbArray
            Erase Variable
         Case vbBoolean
            Variable = False
         Case vbByte
            Variable = CByte(0)
         Case vbCurrency
            Variable = CCur(0)
         Case vbDataObject
            Set Variable = Nothing
         Case vbDate
            Variable = CDate(0)
         Case vbDecimal
            Variable = CDec(0)
         Case vbDouble
            Variable = CDbl(0)
         Case vbEmpty
            Variable = Empty
         Case vbError
            Variable = Empty
         Case vbInteger
            Variable = CInt(0)
         Case vbLong, vbLongLong
            Variable = CLngPtr(0)
         Case vbNull
            Variable = Empty
         Case vbObject
            Set Variable = Nothing
         Case vbSingle
            Variable = CSng(0)
         Case vbString
            Variable = vbNullString
         Case vbUserDefinedType
            'User-Defined-Types cannot be set to a general default value.
            'Each element must be explicitly set to its default value. No
            'assignment takes place in this procedure.
         Case vbVariant
            'This case is included for constistancy, but we will never get
            'here. If the Variant contains data, VarType returns the type of
            'that data. An Empty Variant is type vbEmpty.
            Variable = Empty
      End Select
   End If

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'TransposeArray
'This transposes a two-dimensional array. It returns True if successful or
'False if an error occurs. InputArr must be two-dimensions. OutputArr must be
'a dynamic array. It will be Erased and resized, so any existing content will
'be destroyed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TransposeArray( _
   InputArr As Variant, _
   OutputArr As Variant _
      ) As Boolean
Attribute TransposeArray.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim RowNdx As LongPtr
   Dim ColNdx As LongPtr
   Dim LB1 As LongPtr
   Dim LB2 As LongPtr
   Dim UB1 As LongPtr
   Dim UB2 As LongPtr
   
   
   'Set the default return value
   TransposeArray = False
   
   If Not IsArray(InputArr) Then Exit Function
   If Not IsArray(OutputArr) Then Exit Function
   If Not IsArrayDynamic(OutputArr) Then Exit Function
   If NumberOfArrayDimensions(InputArr) <> 2 Then Exit Function
   
   'Get the Lower and Upper bounds of InputArr
   LB1 = LBound(InputArr, 1)
   LB2 = LBound(InputArr, 2)
   UB1 = UBound(InputArr, 1)
   UB2 = UBound(InputArr, 2)
   
   'Erase and ReDim OutputArr
   Erase OutputArr
   'Redim the Output array. Not the that the LBound and UBound values are preserved.
   ReDim OutputArr(LB2 To LB2 + UB2 - LB2, LB1 To LB1 + UB1 - LB1)
   'Loop through the elemetns of InputArr and put each value in the proper
   'element of the tranposed array
   For RowNdx = LBound(InputArr, 2) To UBound(InputArr, 2)
      For ColNdx = LBound(InputArr, 1) To UBound(InputArr, 1)
         OutputArr(RowNdx, ColNdx) = InputArr(ColNdx, RowNdx)
      Next
   Next
   
   TransposeArray = True

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VectorsToArray
'This function takes 1 or more single-dimensional arrays and converts
'them into a single multi-dimensional array. Each array in Vectors
'comprises one row of the new array. The number of columns in the
'new array is the maximum of the number of elements in each vector.
'Arr MUST be a dynamic array of a data type compatible with ALL the
'elements in each Vector. The code does NOT trap for an error
'13 - Type Mismatch.
'
'If the Vectors are of differing sizes, Arr is sized to hold the
'maximum number of elements in a Vector. The procedure Erases the
'Arr array, so when it is reallocated with Redim, all elements will
'be the reset to their default value (0 or vbNullString or Empty).
'Unused elements in the new array will remain the default value for
'that data type.
'
'Each Vector in Vectors must be a single dimensional array, but
'the Vectors may be of different sizes and LBounds.
'
'Each element in each Vector must be a simple data type. The elements
'may NOT be Object, Arrays, or User-Defined Types.
'
'The rows and columns of the result array are 0-based, regardless of
'the LBound of each vector and regardless of the Option Base statement.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function VectorsToArray( _
   Arr As Variant, _
   ParamArray Vectors() _
      ) As Boolean
Attribute VectorsToArray.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim Vector As Variant
   Dim VectorNdx As LongPtr
   Dim NumElements As LongPtr
   Dim NumRows As LongPtr
   Dim NumCols As LongPtr
   Dim RowNdx As LongPtr
   Dim ColNdx As LongPtr
   Dim VType As VbVarType
   
   
   'Set the default return value
   VectorsToArray = False
   
   If Not IsArray(Arr) Then Exit Function
   If Not IsArrayDynamic(Arr) Then Exit Function

   'Ensure that at least one vector was passed in Vectors
   If IsMissing(Vectors) Then Exit Function
   
   'Loop through Vectors to determine the size of the result array. We do this
   'loop first to prevent having to do a Redim Preserve. This requires looping
   'through Vectors a second time, but this is still faster than doing
   'Redim Preserves.
   For Each Vector In Vectors
       'Ensure Vector is single dimensional array. This will take care of the
       'case if Vector is an unallocated array (NumberOfArrayDimensions = 0
       'for an unallocated array).
      If NumberOfArrayDimensions(Vector) <> 1 Then Exit Function
'---
'2do: this test is a bit late, right?
'---
      'Ensure that Vector is not an array
      If Not IsArray(Vector) Then Exit Function
      'Increment the number of rows. Each Vector is one row or the result array.
      'Test the size of Vector. If it is larger than the existing value of
      'NumCols, set NumCols to the new, larger, value.
      NumRows = NumRows + 1
      If NumCols < UBound(Vector) - LBound(Vector) + 1 Then
         NumCols = UBound(Vector) - LBound(Vector) + 1
      End If
   Next
   'Redim Arr to the appropriate size. Arr is 0-based in both directions,
   'regardless of the LBound of the original Arr and regardless of the
   'LBounds of the Vectors.
   ReDim Arr(0 To NumRows - 1, 0 To NumCols - 1)
   
   'Loop through the rows
   For RowNdx = 0 To NumRows - 1
      'Loop through the columns
      For ColNdx = 0 To NumCols - 1
         'Set Vector (a Variant) to the Vectors(RowNdx) array. We declare
         'Vector as a variant so it can take an  array of any simple data type.
         Vector = Vectors(RowNdx)
         'The vectors need not ber
         If ColNdx < UBound(Vector) - LBound(Vector) + 1 Then
            VType = VarType(Vector(LBound(Vector) + ColNdx))
            If VType >= vbArray Then
               'Test for VType >= vbArray. The VarType of an array is
               'vbArray + VarType(element of array). E.g., the  VarType of an
               'array of Longs equal vbArray + vbLong.  Anything greater than
               'or equal to vbArray is an array of some time.
               Exit Function
            End If
            If VType = vbObject Then
               Exit Function
            End If
            'Vector(LBound(Vector) + ColNdx) is  a simple data type.
            'If Vector(LBound(Vector) + ColNdx) is not a compatible data type
            'with Arr, then a Type Mismatch error will occur. We do NOT trap
            'this error.
            Arr(RowNdx, ColNdx) = Vector(LBound(Vector) + ColNdx)
         End If
      Next
   Next
   
   VectorsToArray = True

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ChangeBoundsOfArray
'This function changes the upper and lower bounds of the specified
'array. InputArr MUST be a single-dimensional dynamic array.
'If the new size of the array (NewUpperBound - NewLowerBound + 1)
'is greater than the original array, the unused elements on
'right side of the array are the default values for the data type
'of the array. If the new size is less than the original size,
'only the first (left-most) N elements are included in the new array.
'The elements of the array may be simple variables (Strings, Longs, etc.),
'Object, or Arrays. User-Defined Types are not supported.
'
'The function returns True if successful, False otherwise.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'2do: make 'NewUpperBound' optional and use size of 'InputArr' to calculate
'     'NewUpperBound' if not given
'2do: better name would be 'ChangeBoundsOfVector', because 'InputArr' has to be
'     a single dimensional array
Public Function ChangeBoundsOfArray( _
   ByRef InputArr As Variant, _
   ByVal NewLowerBound As LongPtr, _
   ByVal NewUpperBound As LongPtr _
      ) As Boolean
Attribute ChangeBoundsOfArray.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim TempArr() As Variant
   Dim InNdx As LongPtr
   Dim OutNdx As LongPtr
   Dim TempNdx As LongPtr
   Dim FirstIsObject As Boolean
   
   
   'Set the default return value
   ChangeBoundsOfArray = False
   
   If NewLowerBound > NewUpperBound Then Exit Function
   If Not IsArray(InputArr) Then Exit Function
   If Not IsArrayDynamic(InputArr) Then Exit Function
   If Not IsArrayAllocated(InputArr) Then Exit Function
   If NumberOfArrayDimensions(InputArr) <> 1 Then Exit Function
   
   'We need to save the IsObject status of the first element of the InputArr
   'to properly handle the Empty variables is we are making the array larger
   'than it was before.
   FirstIsObject = IsObject(InputArr(LBound(InputArr)))
   
   
   'Resize TempArr and save the values in InputArr in TempArr. TempArr will
   'have an LBound of 1 and a UBound of the size of
   '(NewUpperBound - NewLowerBound +1)
   ReDim TempArr(1 To (NewUpperBound - NewLowerBound + 1))
   'Load up TempArr
   TempNdx = 0
   For InNdx = LBound(InputArr) To UBound(InputArr)
      TempNdx = TempNdx + 1
      If TempNdx > UBound(TempArr) Then
         Exit For
      End If
       
      If (IsObject(InputArr(InNdx)) = True) Then
         If InputArr(InNdx) Is Nothing Then
            Set TempArr(TempNdx) = Nothing
         Else
            Set TempArr(TempNdx) = InputArr(InNdx)
         End If
      Else
         TempArr(TempNdx) = InputArr(InNdx)
      End If
   Next
   
   'Now, Erase InputArr, resize it to the new bounds, and load up the values
   'from TempArr to the new InputArr
   Erase InputArr
   ReDim InputArr(NewLowerBound To NewUpperBound)
   OutNdx = LBound(InputArr)
   For TempNdx = LBound(TempArr) To UBound(TempArr)
      If OutNdx <= UBound(InputArr) Then
         If IsObject(TempArr(TempNdx)) Then
            Set InputArr(OutNdx) = TempArr(TempNdx)
         Else
            If FirstIsObject = True Then
               If IsEmpty(TempArr(TempNdx)) Then
                  Set InputArr(OutNdx) = Nothing
               Else
                  Set InputArr(OutNdx) = TempArr(TempNdx)
               End If
            Else
               InputArr(OutNdx) = TempArr(TempNdx)
            End If
         End If
      Else
         Exit For
      End If
      OutNdx = OutNdx + 1
   Next
   
   ChangeBoundsOfArray = True

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'IsArraySorted
'This function determines whether a single-dimensional array is sorted. Because
'sorting is an expensive operation, especially so on large array of Variants,
'you may want to determine if an array is already in sorted order prior to
'doing an actual sort.
'This function returns True if an array is in sorted order (either ascending or
'descending order, depending on the value of the Descending parameter -- default
'is false = Ascending). The decision to do a string comparison (with StrComp) or
'a numeric comparison (with < or >) is based on the data type of the first
'element of the array.
'If TestArray is not an array, is an unallocated dynamic array, or has more than
'one dimension, or the VarType of TestArray is not compatible, the function
'returns NULL.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsArraySorted( _
   TestArray As Variant, _
   Optional Descending As Boolean = False _
      ) As Variant
Attribute IsArraySorted.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim StrCompResultFail As LongPtr
   Dim NumericResultFail As Boolean
   Dim Ndx As LongPtr
   Dim NumCompareResult As Boolean
   Dim StrCompResult As LongPtr
   
   Dim IsString As Boolean
   Dim VType As VbVarType
   
   
   'Set the default return value
   IsArraySorted = Null
   
   If Not IsArray(TestArray) Then Exit Function
   If NumberOfArrayDimensions(TestArray) <> 1 Then Exit Function
   
   'The following code sets the values of comparison that will indicate that
   'the array is unsorted. It the result of StrComp (for strings) or ">="
   '(for numerics) equals the value specified below, we know that the array is
   'unsorted.
   If Descending = True Then
      StrCompResultFail = -1
      NumericResultFail = False
   Else
      StrCompResultFail = 1
      NumericResultFail = True
   End If
   
   'Determine whether we are going to do a string comparison or a numeric
   'comparison
   VType = VarType(TestArray(LBound(TestArray)))
   Select Case VType
      Case vbArray, vbDataObject, vbEmpty, vbError, vbNull, vbObject, vbUserDefinedType
         'Unsupported types. Reutrn Null.
         IsArraySorted = Null
         Exit Function
      Case vbString, vbVariant
         'Compare as string
         IsString = True
      Case Else
         'Compare as numeric
         IsString = False
   End Select
   
   For Ndx = LBound(TestArray) To UBound(TestArray) - 1
      If IsString Then
         StrCompResult = StrComp(TestArray(Ndx), TestArray(Ndx + 1))
         If StrCompResult = StrCompResultFail Then Exit Function
      Else
         NumCompareResult = (TestArray(Ndx) >= TestArray(Ndx + 1))
         If NumCompareResult = NumericResultFail Then Exit Function
      End If
   Next
   
   'If we made it up to here, then the array is in sorted order.
   IsArraySorted = True

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'TwoArraysToOneArray
'This takes two 2-dimensional arrays, Arr1 and Arr2, and
'returns an array combining the two. The number of Rows
'in the result is NumRows(Arr1) + NumRows(Arr2). Arr1 and
'Arr2 must have the same number of columns, and the result
'array will have that many columns. All the LBounds must
'be the same. E.g.,
'The following arrays are legal:
'       Dim Arr1(0 To 4, 0 To 10)
'       Dim Arr2(0 To 3, 0 To 10)
'
'The following arrays are illegal
'       Dim Arr1(0 To 4, 1 To 10)
'       Dim Arr2(0 To 3, 0 To 10)
'
'The returned result array is Arr1 with additional rows
'appended from Arr2. For example, the arrays
'   a    b        and     e    f
'   c    d                g    h
'become
'   a    b
'   c    d
'   e    f
'   g    h
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CombineTwoDArrays( _
   Arr1 As Variant, _
   Arr2 As Variant _
      ) As Variant
Attribute CombineTwoDArrays.VB_ProcData.VB_Invoke_Func = " \n20"

   'Upper and lower bounds of Arr1
   Dim LBoundRow1 As LongPtr
   Dim UBoundRow1 As LongPtr
   Dim LBoundCol1 As LongPtr
   Dim UBoundCol1 As LongPtr
   
   'Upper and lower bounds of Arr2
   Dim LBoundRow2 As LongPtr
   Dim UBoundRow2 As LongPtr
   Dim LBoundCol2 As LongPtr
   Dim UBoundCol2 As LongPtr
   
   'Upper and lower bounds of Result
   Dim LBoundRowResult As LongPtr
   Dim UBoundRowResult As LongPtr
   Dim LBoundColResult As LongPtr
   Dim UBoundColResult As LongPtr
   
   'Index Variables
   Dim RowNdx1 As LongPtr
   Dim ColNdx1 As LongPtr
   Dim RowNdx2 As LongPtr
   Dim ColNdx2 As LongPtr
   Dim RowNdxResult As LongPtr
   Dim ColNdxResult As LongPtr
   
   'Array Sizes
   Dim NumRows1 As LongPtr
   Dim NumCols1 As LongPtr
   
   Dim NumRows2 As LongPtr
   Dim NumCols2 As LongPtr
   
   Dim NumRowsResult As LongPtr
   Dim NumColsResult As LongPtr
   
   Dim Done As Boolean
   Dim Result() As Variant
   Dim ResultTrans() As Variant
   
   Dim V As Variant
   
   
   'Set the default return value
   CombineTwoDArrays = Null
   
   If Not IsArray(Arr1) Then Exit Function
   If Not IsArray(Arr2) Then Exit Function
   If NumberOfArrayDimensions(Arr1) <> 2 Then Exit Function
   If NumberOfArrayDimensions(Arr2) <> 2 Then Exit Function
   
   '''Ensure that the LBound and UBounds of the second dimension are the
   '''same for both Arr1 and Arr2
   'Get the existing bounds
   LBoundRow1 = LBound(Arr1, 1)
   UBoundRow1 = UBound(Arr1, 1)
   
   LBoundCol1 = LBound(Arr1, 2)
   UBoundCol1 = UBound(Arr1, 2)
   
   LBoundRow2 = LBound(Arr2, 1)
   UBoundRow2 = UBound(Arr2, 1)
   
   LBoundCol2 = LBound(Arr2, 2)
   UBoundCol2 = UBound(Arr2, 2)
   
   'Get the total number of rows for the result array
   NumRows1 = UBoundRow1 - LBoundRow1 + 1
   NumCols1 = UBoundCol1 - LBoundCol1 + 1
   NumRows2 = UBoundRow2 - LBoundRow2 + 1
   NumCols2 = UBoundCol2 - LBoundCol2 + 1
   
   'Ensure the number of columns are equal
   If NumCols1 <> NumCols2 Then Exit Function
   
   NumRowsResult = NumRows1 + NumRows2
   
   'Ensure that ALL the LBounds are equal
   If (LBoundRow1 <> LBoundRow2) Or _
      (LBoundRow1 <> LBoundCol1) Or _
      (LBoundRow1 <> LBoundCol2) Then _
         Exit Function
   
   'Get the LBound of the columns of the result array
   LBoundColResult = LBoundRow1
   'Get the UBound of the columns of the result array
   UBoundColResult = UBoundCol1
   
   UBoundRowResult = LBound(Arr1, 1) + NumRows1 + NumRows2 - 1
   'Redim the Result array to have number of rows equal to
   'number-of-rows(Arr1) + number-of-rows(Arr2)
   'and number-of-columns equal to number-of-columns(Arr1)
   ReDim Result(LBoundRow1 To UBoundRowResult, LBoundColResult To UBoundColResult)
   
   RowNdxResult = LBound(Result, 1) - 1
   
   Done = False
   Do Until Done
      'Copy elements of Arr1 to Result
      For RowNdx1 = LBound(Arr1, 1) To UBound(Arr1, 1)
         RowNdxResult = RowNdxResult + 1
         For ColNdx1 = LBound(Arr1, 2) To UBound(Arr1, 2)
            V = Arr1(RowNdx1, ColNdx1)
            Result(RowNdxResult, ColNdx1) = V
         Next
      Next
   
      'Copy elements of Arr2 to Result
      For RowNdx2 = LBound(Arr2, 1) To UBound(Arr2, 1)
         RowNdxResult = RowNdxResult + 1
         For ColNdx2 = LBound(Arr2, 2) To UBound(Arr2, 2)
            V = Arr2(RowNdx2, ColNdx2)
            Result(RowNdxResult, ColNdx2) = V
         Next
      Next
       
      If RowNdxResult >= UBound(Result, 1) + (LBoundColResult = 1) Then
         Done = True
      End If
   Loop
   
   CombineTwoDArrays = Result

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ExpandArray
'This expands a two-dimensional array in either dimension. It returns the result
'array if successful, or NULL if an error occurred. The original array is never
'changed.
'Paramters:
'--------------------
'Arr                   is the array to be expanded.
'
'WhichDim              is either 1 for additional rows or 2 for
'                      additional columns.
'
'AdditionalElements    is the number of additional rows or columns
'                      to create.
'
'FillValue             is the value to which the new array elements should be
'                      initialized.
'
'You can nest calls to Expand array to expand both the number of rows and
'columns. E.g.,
'
'C = ExpandArray( _
'        ExpandArray( _
'           Arr:=A, _
'           WhichDim:=1, _
'           AdditionalElements:=3, _
'           FillValue:="R") _
'        , _
'        WhichDim:=2, _
'        AdditionalElements:=4, _
'        FillValue:="C")
'
'This first adds three rows at the bottom of the array, and then adds four
'columns on the right of the array.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ExpandArray( _
   Arr As Variant, _
   WhichDim As LongPtr, _
   AdditionalElements As LongPtr, _
   FillValue As Variant _
      ) As Variant
Attribute ExpandArray.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim Result As Variant
   Dim RowNdx As LongPtr
   Dim ColNdx As LongPtr
   Dim ResultRowNdx As LongPtr
   Dim ResultColNdx As LongPtr
   Dim NumRows As LongPtr
   Dim NumCols As LongPtr
   Dim NewUBound As LongPtr
   
   '===========================================================================
   Const ROWS_ As LongPtr = 1
   '===========================================================================
   
   
   'Set the default return value
   ExpandArray = Null
   
   If Not IsArray(Arr) Then Exit Function
   If NumberOfArrayDimensions(Arr) <> 2 Then Exit Function
   
   'Ensure the dimension is 1 or 2
   Select Case WhichDim
      Case 1, 2
      Case Else
         Exit Function
   End Select
   
   'Ensure AdditionalElements is > 0.
   'If AdditionalElements  = 0, return Arr.
   If AdditionalElements < 0 Then
      Exit Function
   ElseIf AdditionalElements = 0 Then
      ExpandArray = Arr
      Exit Function
   End If
   
   NumRows = UBound(Arr, 1) - LBound(Arr, 1) + 1
   NumCols = UBound(Arr, 2) - LBound(Arr, 2) + 1
   
   If WhichDim = ROWS_ Then
      'Redim Result
      ReDim Result(LBound(Arr, 1) To UBound(Arr, 1) + AdditionalElements, LBound(Arr, 2) To UBound(Arr, 2))
      'Transfer Arr array to Result
      For RowNdx = LBound(Arr, 1) To UBound(Arr, 1)
         For ColNdx = LBound(Arr, 2) To UBound(Arr, 2)
            Result(RowNdx, ColNdx) = Arr(RowNdx, ColNdx)
         Next
      Next
      'Fill the rest of the result array with FillValue
      For RowNdx = UBound(Arr, 1) + 1 To UBound(Result, 1)
         For ColNdx = LBound(Arr, 2) To UBound(Arr, 2)
            Result(RowNdx, ColNdx) = FillValue
         Next
      Next
   Else
      'Redim Result
      ReDim Result(LBound(Arr, 1) To UBound(Arr, 1), UBound(Arr, 2) + AdditionalElements)
      'Transfer Arr array to Result
      For RowNdx = LBound(Arr, 1) To UBound(Arr, 1)
         For ColNdx = LBound(Arr, 2) To UBound(Arr, 2)
            Result(RowNdx, ColNdx) = Arr(RowNdx, ColNdx)
         Next
      Next
      'Fill the rest of the result array with FillValue
      For RowNdx = LBound(Arr, 1) To UBound(Arr, 1)
         For ColNdx = UBound(Arr, 2) + 1 To UBound(Result, 2)
            Result(RowNdx, ColNdx) = FillValue
         Next
      Next
   End If
   
   ExpandArray = Result

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SwapArrayRows
'This function returns an array based on Arr with Row1 and Row2 swapped.
'It returns the result array or NULL if an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SwapArrayRows( _
   Arr As Variant, _
   Row1 As LongPtr, _
   Row2 As LongPtr _
      ) As Variant
Attribute SwapArrayRows.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim V As Variant
   Dim Result As Variant
   Dim RowNdx As LongPtr
   Dim ColNdx As LongPtr
   
   
   'Set the default return value
   SwapArrayRows = Null
   
   If Not IsArray(Arr) Then Exit Function
   If NumberOfArrayDimensions(Arr) <> 2 Then Exit Function
   
   'Ensure Row1 and Row2 are less than or equal to the number of rows
   If Row1 > UBound(Arr, 1) Then Exit Function
   If Row2 > UBound(Arr, 1) Then Exit Function
   
   'If Row1 = Row2, just return the array and exit. Nothing to do.
   If Row1 = Row2 Then
      SwapArrayRows = Arr
      Exit Function
   End If
   
   'Set Result to Arr
   Result = Arr
   
   'Redim V to the number of columns
   ReDim V(LBound(Arr, 2) To UBound(Arr, 2))
   'Put Row1 in V
   For ColNdx = LBound(Arr, 2) To UBound(Arr, 2)
      V(ColNdx) = Arr(Row1, ColNdx)
      Result(Row1, ColNdx) = Arr(Row2, ColNdx)
      Result(Row2, ColNdx) = V(ColNdx)
   Next
   
   SwapArrayRows = Result

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'SwapArrayColumns
'This function returns an array based on Arr with Col1 and Col2 swapped.
'It returns the result array or NULL if an error occurred.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SwapArrayColumns( _
   Arr As Variant, _
   Col1 As LongPtr, _
   Col2 As LongPtr _
      ) As Variant
Attribute SwapArrayColumns.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim V As Variant
   Dim Result As Variant
   Dim RowNdx As LongPtr
   Dim ColNdx As LongPtr
   
   
   'Set the default return value
   SwapArrayColumns = Null
   
   If Not IsArray(Arr) Then Exit Function
   If NumberOfArrayDimensions(Arr) <> 2 Then Exit Function
   
   'Ensure Col1 and Col2 are less than or equal to the number of columns
   If Col1 > UBound(Arr, 2) Then Exit Function
   If Col2 > UBound(Arr, 2) Then Exit Function
       
   'If Col1 = Col2, just return the array and exit. Nothing to do.
   If Col1 = Col2 Then
      SwapArrayColumns = Arr
      Exit Function
   End If
   
   'Set Result to Arr
   Result = Arr
   
   'Redim V to the number of columns
   ReDim V(LBound(Arr, 1) To UBound(Arr, 1))
   'Put Col2 in V
   For RowNdx = LBound(Arr, 1) To UBound(Arr, 1)
      V(RowNdx) = Arr(RowNdx, Col1)
      Result(RowNdx, Col1) = Arr(RowNdx, Col2)
      Result(RowNdx, Col2) = V(RowNdx)
   Next
   
   SwapArrayColumns = Result

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GetColumn
'This populates ResultArr with a one-dimensional array that is the
'specified column of Arr. The existing contents of ResultArr are
'destroyed. ResultArr must be a dynamic array.
'Returns True or False indicating success.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetColumn( _
   Arr As Variant, _
   ResultArr As Variant, _
   ColumnNumber As LongPtr _
      ) As Boolean
Attribute GetColumn.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim RowNdx As LongPtr
   
   
   'Set the default return value
   GetColumn = False
   
   If Not IsArray(Arr) Then Exit Function
   If NumberOfArrayDimensions(Arr) <> 2 Then Exit Function
   If Not IsArrayDynamic(ResultArr) Then Exit Function
   
   'Ensure ColumnNumber is less than or equal to the number of columns
   If UBound(Arr, 2) < ColumnNumber Then Exit Function
   If LBound(Arr, 2) > ColumnNumber Then Exit Function
   
   Erase ResultArr
   ReDim ResultArr(LBound(Arr, 1) To UBound(Arr, 1))
   For RowNdx = LBound(ResultArr) To UBound(ResultArr)
      ResultArr(RowNdx) = Arr(RowNdx, ColumnNumber)
   Next
   
   GetColumn = True

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'GetRow
'This populates ResultArr with a one-dimensional array that is the
'specified row of Arr. The existing contents of ResultArr are
'destroyed. ResultArr must be a dynamic array.
'Returns True or False indicating success.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetRow( _
   Arr As Variant, _
   ResultArr As Variant, _
   RowNumber As LongPtr _
      ) As Boolean
Attribute GetRow.VB_ProcData.VB_Invoke_Func = " \n20"

   Dim ColNdx As LongPtr
   
   
   'Set the default return value
   GetRow = False
   
   If Not IsArray(Arr) Then Exit Function
   If NumberOfArrayDimensions(Arr) <> 2 Then Exit Function
   If Not IsArrayDynamic(ResultArr) Then Exit Function
   
   'Ensure ColumnNumber is less than or equal to the number of columns
   If UBound(Arr, 1) < RowNumber Then Exit Function
   If LBound(Arr, 1) > RowNumber Then Exit Function
   
   Erase ResultArr
   ReDim ResultArr(LBound(Arr, 2) To UBound(Arr, 2))
   For ColNdx = LBound(ResultArr) To UBound(ResultArr)
      ResultArr(ColNdx) = Arr(RowNumber, ColNdx)
   Next
   
   GetRow = True

End Function

'------------------------------------------------------------------------------

'2do:
'- add to upper list
'- add to 'AddUDFToCustomCategory'
'- add some parameter checking
Public Function VectorTo1DArray( _
   InputVector As Variant, _
   Optional LowerBoundOfSecondDimension As Integer = 0 _
      ) As Variant
   
   Dim ResultArray() As Variant
   Dim i As Integer
   
   
   ReDim ResultArray(LBound(InputVector) To UBound(InputVector), LowerBoundOfSecondDimension To LowerBoundOfSecondDimension)
   For i = LBound(InputVector) To UBound(InputVector)
      ResultArray(i, LowerBoundOfSecondDimension) = InputVector(i)
   Next
   
   VectorTo1DArray = ResultArray
   
End Function
