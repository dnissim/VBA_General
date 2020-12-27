Attribute VB_Name = "mdlGeneral"
'Author: David Nissim
Option Explicit

'This module contains general code for the following categories:
'   List Objects (tables)
'   Chart Objects
'   Arrays
'   Worksheets & Ranges
'   Collections
'   Strings
'   SQL implementation
'   Numbers
'   Shapes
'===========================LIST OBJECTS=================================
    Function GetTable(strTableName As String) As ListObject
        'This function searches the workbook for a table with the exact name strTableName
        'Returns the table object
        'If nothing found then display message
        On Error Resume Next

        Dim sht As Worksheet
        Dim tbl As ListObject

        '# Modified so that it doesn't have to go through each table in the workbook (more efficient) 1/26/2020
        For Each sht In ThisWorkbook.Sheets
            For Each tbl In sht.ListObjects
                'Debug.Print sht.Name & " " & tbl.Name      'uncomment to print all table names
        
                If LCase(tbl.Name) = LCase(strTableName) Then
                    Set GetTable = tbl
                    Exit Function
                End If
            Next tbl
        Next sht

        'For Each sht In ThisWorkbook.Worksheets
        '
        '    Set GetTable = sht.ListObjects(strTableName)
            
        '    If Err.Number = 0 Then Exit Function
            
        '    Err.Clear

        'Next sht



        'If the code reaches this point it means the table wasn't found.
        'This may have negative implications depending on where this function is called.
        'This message gives the user an out
        Dim ans As Byte
        ans = MsgBox("Could not find table with name '" & strTableName & "'." & vbNewLine & vbNewLine & _
            "Would you like to abort code?", vbCritical + vbYesNo, "Table not found")
            
        If ans = vbYes Then End

        'Set GetTable = Nothing   '#This is redundant

    End Function

    Function FindValueIndex(rng As Range, strValue As String) As Integer
        'This function returns the first index of an exact match within a specified range.
        'If nothing is found, then return 0.

        Dim i   As Integer
        Dim out As Integer

        For i = 1 To rng.count
            If LCase(rng(i).Value) = LCase(strValue) Then
                out = i
                Exit For
            End If
        Next i

        FindValueIndex = out

    End Function

    Function TableColumnIndex(tbl As ListObject, strHeader As String) As Integer
        'This function returns the column number with the header:strHeader in Table:tbl.
        'If the header isn't found then an error is generated
        Dim i   As Integer
        Dim out As Integer

        out = FindValueIndex(tbl.HeaderRowRange, strHeader)

        If out = 0 Then
            MsgBox "Column header " & strHeader & " not found in table " & tbl.Name & ".", vbCritical + vbOKCancel, "Error"
            End
        End If

        TableColumnIndex = out
    End Function

    Function TableRowIndex(tbl As ListObject, strHeader As String, strValue As String) As Integer
        'This function returns the row number (local to Databodyrange) for a specific value:strValue.
        'It searches in table:tbl, column:strHeader

        Dim i         As Integer
        Dim out       As Integer
        Dim colHeader As Integer

        colHeader = TableColumnIndex(tbl, strHeader)

        TableRowIndex = FindValueIndex(tbl.DataBodyRange.Columns(colHeader).Rows, strValue)

    End Function

    Function TableLookup(tbl As ListObject, strSearchValue As String, strSearchColumn As String, strReturnColumn As String) As Variant
        'This function acts like Vlookup.
        'If it exists, it will find strSearchValue in the column:strSearchColumn of Table:tbl
        'Then it will return the corresponding value in the table from column:strReturnColumn

        Dim colReturn As Integer
        Dim rowReturn As Integer

        colReturn = TableColumnIndex(tbl, strReturnColumn)
        rowReturn = TableRowIndex(tbl, strSearchColumn, strSearchValue)

        If rowReturn = 0 Then
            TableLookup = 0
        Else
            TableLookup = tbl.DataBodyRange(rowReturn, colReturn)
        End If

        End Function

        Sub EmptyTable(tbl As ListObject)
        'If there is any data in the table then clear it.

        If Not (tbl.DataBodyRange Is Nothing) Then tbl.DataBodyRange.Rows.Delete

        tbl.Range.Cells(2, 1) = "1" 'Enter a value so Databodyrange exists

    End Sub

    Sub UpdateTableWithRecordset(rs As ADODB.Recordset, tbl As ListObject)
        'Inserts the contents of a recordset into a pre-existing table

        EmptyTable tbl
        RemoveFilteringFromSheet tbl.Parent

        With tbl.DataBodyRange
            
            If Not rs.EOF Then
                Dim arrSQL() As Variant
                arrSQL = TransposeArray(rs.GetRows())

                Dim rowEnd As Long

                rowEnd = UBound(arrSQL, 1) + 1

                Range(.Cells(1, 1), .Cells(rowEnd, rs.Fields.count)).Value = arrSQL

            End If

        End With

    End Sub

    Sub UpdateTableWithArray(arr As Variant, tbl As ListObject)
        'Inserts the contents of an array into a pre-existing table
        'For the provided array, dimensions should be (row, column)
        Application.ScreenUpdating = False

        Dim arrDimensions As Integer
        Dim rowEnd        As Long
        Dim colEnd        As Integer
        Dim tblColumns    As Integer

        arrDimensions = GetArrayDimension(arr)

        'Check for valid array sizing
        If arrDimensions > 2 Or arrDimensions = 0 Then
            MsgBox "Array must be 1D or 2D to update a table." & vbNewLine & vbNewLine & _
                "Provided array has " & arrDimensions & " dimensions.", _
                vbCritical + vbOKOnly, _
                "Invalid dimensions"
                
            Exit Sub
        End If

        rowEnd = UBound(arr, 1) - LBound(arr, 1) + 1
        If arrDimensions = 1 Then
            colEnd = 1
        Else
            colEnd = UBound(arr, 2) - LBound(arr, 2) + 1
        End If

        tblColumns = tbl.DataBodyRange.Columns.count

        If colEnd <> tblColumns Then
            MsgBox "# of table columns (" & tblColumns & ") do not match array columns (" & colEnd & ".", _
            vbCritical + vbOKOnly, _
            "Error updating table with array"
            Exit Sub
        End If

        EmptyTable tbl
        RemoveFilteringFromSheet tbl.Parent

        With tbl.DataBodyRange
            Range(.Cells(1, 1), .Cells(rowEnd, colEnd)) = arr
        End With

        Application.ScreenUpdating = True

    End Sub

    Function ListOfTables(Optional wb As Workbook) As Variant
        'Outputs an array of all the tables in the workbook, and their
        'parent worksheets.  If no workbook is provided to the function it assumes This Workbook
        Dim sht As Worksheet
        Dim tbl As ListObject
        Dim tblCount As Integer
        Dim outArray() As String
        tblCount = 0

        If wb Is Nothing Then Set wb = ThisWorkbook

        For Each sht In wb.Worksheets
            For Each tbl In sht.ListObjects
                tblCount = tblCount + 1
                ReDim Preserve outArray(1, 1 To tblCount) As String
                outArray(0, tblCount) = tbl.Name
                outArray(1, tblCount) = sht.Name
            Next tbl
        Next sht

        ListOfTables = TransposeArray(outArray)

        End Function

        Sub FocusTable(tbl As ListObject)
        'This sub activates the provided table
        On Error Resume Next

        tbl.Parent.Activate
        tbl.Range.Select

    End Sub


    '---TABLE FUNCTIONS RETURNING OBJECTS INSTEAD OF INDICES---
    Function FindValueRange(rngSearch As Range, strValue As String) As Range
        'This function returns the first range of an exact match within a specified range.
        'If nothing is found, then return 0.

        Dim i   As Integer
        Dim out As Range

        For i = 1 To rngSearch.count
            If LCase(rngSearch(i).Value) = LCase(strValue) Then
                Set out = rngSearch(i)
                Exit For
            End If
        Next i

        Set out = rngSearch.Find(strValue)

        Set FindValueRange = out
        Set out = Nothing

    End Function

    Function TableColumnRange(tbl As ListObject, strHeader As String) As Range
        'This function returns the column Range with the header:strHeader in Table:tbl.

        Dim i        As Integer
        Dim colIndex As Integer
        Dim out      As Range

        colIndex = TableColumnIndex(tbl, strHeader)
        Set out = tbl.DataBodyRange.Columns(colIndex)

        Set TableColumnRange = out
        End Function

        Function TableRowRange(tbl As ListObject, strHeader As String, strValue As String) As Range
        'This function returns the row Range with the header:strHeader in Table:tbl for a specific value:strValue.
        'It searches in table:tbl, column:strHeader

        Dim out As Range

        Set out = tbl.DataBodyRange(TableRowIndex(tbl, strHeader, strValue))

        Set TableRowRange = out
    End Function


'==========================CHART OBJECTS==================================
    Function GetChart(strChartName As String) As Chart
        'This function searches the workbook for a chart with the exact name strChartName
        'Returns the chart object
        'If nothing found then display message

        Dim sht As Worksheet
        Dim cht As ChartObject

        For Each sht In ThisWorkbook.Sheets
            For Each cht In sht.ChartObjects
                'debug.print cht.name       'Uncomment to print all chart names
            
                If LCase(cht.Name) = LCase(strChartName) Then
                    Set GetChart = cht.Chart
                    Exit Function
                End If
            Next cht
        Next sht

        MsgBox "Could not find chart with name '" & strChartName & "'.", vbCritical + vbOKOnly, "Chart not found"
        End Function


'==========================ARRAYS==================================
    Function TransposeArray(InArray As Variant) As Variant
        'Produces a transposed 2D array (works when Application.WorksheetFunction.Transpose does not)

        'Only the last dimension in an array can be modified by ReDim in VBA
        'However this means you are generally working with row vectors.
        'Typically you'll want to write column vectors to Excel which is what makes this function handy

        Select Case GetArrayDimension(InArray)
            Case 0
                TransposeArray = InArray
            Case 1
                TransposeArray = Application.Transpose(InArray)
            Case 2

                Dim iStart As Long, iEnd As Long, i As Long
                Dim jStart As Long, jEnd As Long, j As Long
                
                'Define lower and upper bounds for both dimensions
                iStart = LBound(InArray, 1)
                iEnd = UBound(InArray, 1)
                
                jStart = LBound(InArray, 2)
                jEnd = UBound(InArray, 2)
                
                'Size that shit accordingly
                Dim outArray() As Variant
                ReDim outArray(jStart To jEnd, iStart To iEnd)
                
                'Zhu Li, do the thing
                For i = iStart To iEnd
                    For j = jStart To jEnd
                    
                        outArray(j, i) = InArray(i, j)
                    
                    Next j
                Next i
                
                TransposeArray = outArray
                
            Case Else
                TransposeArray = 0
                MsgBox "TransposeArray Function can only accept arrays of dimensions 1 or 2.", _
                    vbOKOnly + vbExclamation, "Transposition Failed"
        End Select
    End Function

    Function FindIndexinArray(SearchValue As Variant, InArray As Variant) As Coordinates2D
        'Returns the first index of a value (SearchValue) occurring in an array (InArray) if a match is found

        Dim r As Long, c As Long

        For r = LBound(InArray, 1) To UBound(InArray, 1)
            For c = LBound(InArray, 2) To UBound(InArray, 2)
                If InArray(r, c) = SearchValue Then
                    FindIndexinArray.Row = r
                    FindIndexinArray.Column = c
                    FindIndexinArray.Found = True
                    Exit Function
                End If
            Next c
        Next r

        FindIndexinArray.Found = False

    End Function

    Function RangeToArray(rng As Range, _
                    Optional blnRemoveSpecialChars As Boolean = True) _
                As Variant
        'This function returns an array with the values from a specified range
        'Values must be called from array by using the following indices (Row, Column)****
        'Array indicies start at 1

        Dim arrOut As Variant
        Dim r As Long, c As Long
        Dim strCurrent As String

        '#Delete timer?
        Dim start As Single
        start = Timer

        'If the importing range is a single row then arrOut won't be an array. The if statement is to catch that case and prevent an error.
        If TypeName(rng.Value2) = "Variant()" Then
            arrOut = rng.Value2
        Else
            ReDim arrOut(1 To 1, 1 To 1)
            arrOut(1, 1) = rng.Value2
        End If

        'Replace special characters with spaces to increase search string parts, for better guesses
        If blnRemoveSpecialChars Then
            For r = LBound(arrOut, 1) To UBound(arrOut, 1)
                For c = LBound(arrOut, 2) To UBound(arrOut, 2)
                    'If InStr(1, arrOut(r, c), "#", vbBinaryCompare) <> 0 Then arrOut(r, c) = Replace(arrOut(r, c), "#", " ")
                    'If InStr(1, arrOut(r, c), "-", vbBinaryCompare) <> 0 Then arrOut(r, c) = Replace(arrOut(r, c), "-", " ")
                    strCurrent = arrOut(r, c)
                    arrOut(r, c) = CascadeString(RemoveSpecialCharacters(strCurrent), 1)(1) 'added to make the guess functions work together properly
                                                                                            'Removes special characters then replaces with spaces
                                                                                            'Then ensures there's only 1 space per... space
                    Next c
            Next r
        End If

        RangeToArray = arrOut

    End Function

    Function GetArrayDimension(arr As Variant) As Long
        'This function returns the number of dimensions in an array
        'Note this is not the length of those dimensions.
        'The count starts at 1. A zero is an empty array

        On Error GoTo Err
        Dim i As Long
        Dim tmp As Long
        i = 0
        Do While True
            i = i + 1
            tmp = UBound(arr, i)
        Loop

Err:
            GetArrayDimension = i - 1
    End Function


'===============WORKSHEETS & RANGES==================================
    Sub RemoveFilteringFromSheet(sht As Worksheet)
        'The update table sub gets messed up if there's any filtering happening on the sheet.
        'This sub makes sure there's nothing like that
        'On Error statement required because if there is no filtering then it generates an error

        On Error Resume Next

        'Make sure no tables have filters
        Dim LO As ListObject

        For Each LO In sht.ListObjects
            LO.AutoFilter.ShowAllData
        Next LO

        'Make sure no non-table filters have been applied
        sht.ShowAllData

    End Sub

    Public Function LettertoNumber(ByVal ColumnLetter As String) As Integer
        'Gives the column number for a given letter.  Doesn't work past XFD

        LettertoNumber = Range(ColumnLetter & 1).Column

    End Function


    Public Function LastRow(rng As Range)
        'Returns the last row in the first column of a given range
        Dim sht As Worksheet
        Set sht = rng.Parent
        With sht
            LastRow = .Cells(.Rows.count, rng.Column).End(xlUp).Row
        End With

    End Function

    Public Function LastColumn(rng As Range)
        'Returns the last column in the first row of a given range
        Dim sht As Worksheet
        Set sht = rng.Parent
        With sht
            LastColumn = .Cells(rng.Row, .Columns.count).End(xlToLeft).Row
        End With

    End Function

'==========================COLLECTIONS==================================
    Function FindInCollection(SearchValue As Variant, coll As Collection) As CollectionItem
        'Returns the Index and Value of a matching item in a collection to SearchValue
        'Currently written to compare text

        Dim i As Long

        For i = 1 To coll.count
            If LCase(coll(i)) = LCase(SearchValue) Then
                FindInCollection.Value = coll(i)
                FindInCollection.Index = i
                FindInCollection.Found = True
                Exit Function
            End If
        Next i

        FindInCollection.Found = False

    End Function

'=============================STRINGS=====================================
    Public Function IsLetter(ByVal inputString As String) As Boolean
        'Returns whether a single character is a letter or not.  If input length is longer than one, it's truncated to the first letter

        If Len(inputString) > 1 Then inputString = Left(inputString, 1)

        If LenB(inputString) <> 0 Then
            Dim ascNum As Integer
            ascNum = Asc(UCase(inputString))
            
            IsLetter = ascNum >= 65 And ascNum <= 90
        Else
            IsLetter = False
        End If

        End Function

        Function RemoveSpecialCharacters(strDirty As String, Optional strReplace As String = " ") As String
        'Replaces special characters with spaces
        Dim iChar As Integer
        Dim strClean As String

        For iChar = 1 To Len(strDirty)
            If Mid(strDirty, iChar, 1) Like "[0-9a-zA-Z]" Then
                strClean = strClean & Mid(strDirty, iChar, 1)
            Else
                strClean = strClean & strReplace
            End If
        Next iChar

        RemoveSpecialCharacters = strClean
    End Function



    Function SplitMultiDelims(ByRef Text As String, ByRef DelimChars As String, _
            Optional ByVal IgnoreConsecutiveDelimiters As Boolean = False, _
            Optional ByVal Limit As Long = -1) As String()
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' SplitMultiDelims by alainbryden
        ' This function splits Text into an array of substrings, each substring
        ' delimited by any character in DelimChars. Only a single character
        ' may be a delimiter between two substrings, but DelimChars may
        ' contain any number of delimiter characters. It returns a single element
        ' array containing all of text if DelimChars is empty, or a 1 or greater
        ' element array if the Text is successfully split into substrings.
        ' If IgnoreConsecutiveDelimiters is true, empty array elements will not occur.
        ' If Limit greater than 0, the function will only split Text into 'Limit'
        ' array elements or less. The last element will contain the rest of Text.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'https://www.experts-exchange.com/articles/1480/How-to-Split-a-String-with-Multiple-Delimiters-in-VBA.html

        Dim ElemStart As Long, N As Long, M As Long, Elements As Long
        Dim lDelims As Long, lText As Long
        Dim arr() As String
        
        lText = Len(Text)
        lDelims = Len(DelimChars)
        If lDelims = 0 Or lText = 0 Or Limit = 1 Then
            ReDim arr(0 To 0)
            arr(0) = Text
            SplitMultiDelims = arr
            Exit Function
        End If
        ReDim arr(0 To IIf(Limit = -1, lText - 1, Limit))
        
        Elements = 0: ElemStart = 1
        
        
        For N = 1 To lText
            If InStr(DelimChars, Mid(Text, N, 1)) Then
                arr(Elements) = Mid(Text, ElemStart, N - ElemStart)
                If IgnoreConsecutiveDelimiters Then
                    If Len(arr(Elements)) > 0 Then Elements = Elements + 1
                Else
                    Elements = Elements + 1
                End If
                ElemStart = N + 1
                If Elements + 1 = Limit Then Exit For
            End If
        Next N
        
        'Get the last token terminated by the end of the string into the array
        If ElemStart <= lText Then arr(Elements) = Mid(Text, ElemStart)
        'Since the end of string counts as the terminating delimiter, if the last character
        'was also a delimiter, we treat the two as consecutive, and so ignore the last elemnent
        If IgnoreConsecutiveDelimiters Then If Len(arr(Elements)) = 0 Then Elements = Elements - 1
        
        ReDim Preserve arr(0 To Elements) 'Chop off unused array elements
        SplitMultiDelims = arr
    End Function

    Function CascadeString(strOriginal As String, NumPieces As Integer) As Collection
        'Delimits an entered string by spaces. Creates NumPieces smaller strings combining each segment in a cascading manner.
        'E.g.  CascadeString("Hello world foobar", 1) -> "Hello world foobar"
        'E.g.  CascadeString("Hello world foobar", 2) -> "Hello world" and "world foobar"
        'E.g.  CascadeString("Hello world foobar", 3) -> "Hello", "world", and "foobar"

        Dim strList As Collection 'List of all string segments

        Dim NewStrSize As Integer   'Size of the new output strings
        Dim segment As Variant

        Set strList = New Collection

        'Break string into segments and store in strList
        For Each segment In SplitMultiDelims(strOriginal, " ", True)
            strList.Add segment
        Next

        'Cannot have more pieces than segments.  If input NumPieces is more than # of segments, then adjust NumPieces
        If NumPieces > strList.count Then NumPieces = strList.count

        'How many string segments will be included in each output string
        NewStrSize = strList.count - NumPieces + 1

        Dim StartSeg As Integer, CurSeg As Integer
        Dim subString As String
        Dim collOut As Collection 'Output collection

        Set collOut = New Collection

        For StartSeg = 1 To NumPieces
            subString = vbNullString
            For CurSeg = StartSeg To StartSeg + (NewStrSize - 1)
                subString = subString & strList(CurSeg) & " "
            Next CurSeg
            
            'Remove last space
            subString = Left(subString, Len(subString) - 1)
            collOut.Add subString
        Next StartSeg

        Set CascadeString = collOut
        Set strList = Nothing
        Set collOut = Nothing

    End Function

'=============================SQL=====================================

    Function OpenRST(strSQL As String) As ADODB.Recordset
        'Returns an open recordset object

        Dim cn As ADODB.Connection
        Dim strProvider As String, strExtendedProperties As String
        Dim strFile As String, strCon As String

        strFile = ThisWorkbook.FullName
        strFile = Replace(strFile, "/", "\")


        'The ADODB connection doesn't work if the file is being accessed from a cloud drive.  It has to be saved locally.
        'Check whether the file is saved locally to proceed. Otherwise display message and exit sub.
        If InStr(strFile, "http") = 1 Then
        'Workbook is saved on a cloud drive, and SQL updates won't work.
            MsgBox "Cloud storage detected.  Cannot update background budget tables unless this file is saved locally." & vbNewLine & vbNewLine & _
            "Please save this file to your computer then click the Update Budget Tables button. After this is done, you can save to the cloud drive again.", _
            vbOKOnly + vbCritical, "Cannot update background budget tables"

            End
        End If

        strProvider = "Microsoft.ACE.OLEDB.12.0"
        strExtendedProperties = """Excel 12.0;HDR=Yes;IMEX=1"";"


        strCon = "Provider=" & strProvider & _
                ";Data Source=" & strFile & _
                ";Extended Properties=" & strExtendedProperties

        Set cn = CreateObject("ADODB.Connection")
        Set OpenRST = CreateObject("ADODB.Recordset")

        cn.Open strCon

        OpenRST.Open strSQL, cn

    End Function

    Function SQLTableAddress(lstTable As ListObject) As String
        'Returns SQL string text for referencing a specific table
        'Format returned is [sheetName!$Range]
        Dim strAddress As String
        Dim strSheet As String

        With lstTable
            
            strAddress = Replace(.Range.Address, "$", "")
            strSheet = .Range.Parent.Name
            
        End With

            SQLTableAddress = "[" & strSheet & "$" & strAddress & "]"
    End Function

'=========================SLICERS=====================================
    Sub ClearSlicers()
        'This clears selections from all slicers in the workbook
        Dim slc As SlicerCache

        For Each slc In ThisWorkbook.SlicerCaches
            slc.ClearManualFilter
        Next slc

    End Sub

'=========================NUMBERS=====================================
    Function IsInteger(ByVal Value As Variant) As Boolean
        IsInteger = Int(Value) = Value
    End Function

    Function ConvertDecimalToAnotherBase(ByVal ValueToConvert As Long, ByVal newBase As Integer) As String
        'This function takes ValueToConvert in decimal and converts it to a new base.
        'The return is a string of base 10 numbers separated by spaces representing each digit in the new base.
        'If ValueToConvert is a negative, a "-" will be added to the front of the output string
        'E.g. 22 in base 12 would return "1 10"
        'This function can be used in tandem with split to get each digit from the string

        'Check for positive bases
        If newBase < 1 Then
            MsgBox "ConvertToAnotherBase was only designed for positive bases. Cannot use a non-positive integer as a base.", vbCritical + vbOKOnly, "Error using ConvertToAnotherBase Function"
            Exit Function
        End If

        Dim strOUT As String        'Placeholder for the output string
        Dim leftover As Long        'The leftover whole number after you remove the remainder from the division
        Dim remainder As Integer    'The remainder after dividing by the newBase
        Dim blnNegative As Boolean  'Stores whether ValueToConvert was negative

        'If input is 0 then just output 0 regardless of base
        'Otherwise, determine each digit based on the remainder of division by the new base.


        If ValueToConvert = 0 Then
            strOUT = "0"
        Else
            blnNegative = ValueToConvert < 0

            leftover = Abs(ValueToConvert)
            
            Do While leftover >= newBase
                remainder = newBase * ((leftover / newBase) - Int(leftover / newBase))
                leftover = (leftover - remainder) / newBase
                
                strOUT = remainder & " " & strOUT
            Loop
            
            If leftover > 0 Then strOUT = leftover & " " & strOUT
            
            strOUT = Left(strOUT, Len(strOUT) - 1)      'Remove extra space from end of string
            
            If blnNegative Then strOUT = "-" & strOUT   'Add negative in front if the original value was negative
        End If

        ConvertDecimalToAnotherBase = strOUT

    End Function

''=========================SHAPES=====================================
    '#This is pretty inefficient with a lot of shapes
    'Function GetShape(strShapeName As String) As Shape
    ''This function searches the workbook for a shape with the exact name strShapeName
    ''Returns the shape object
    ''If nothing found then display message
    '
    'Dim sht As Worksheet
    'Dim shp As Shape
    '
    'For Each sht In ThisWorkbook.Sheets
    '    For Each shp In sht.Shapes
    '        'debug.print cht.name       'Uncomment to print all chart names
    '
    '        If LCase(shp.Name) = LCase(strShapeName) Then
    '            Set GetShape = shp
    '            Exit Function
    '        End If
    '    Next shp
    'Next sht
    '
    'MsgBox "Could not find shape with name '" & strShapeName & "'.", vbCritical + vbOKOnly, "Shape not found"
    'End Function
