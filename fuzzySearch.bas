Option Explicit

Sub FindReplaceList()
'
' Macro designed to run a find/replace routine on a long list of paired items.
' The paired find/replace items are place in a separate file named "list.docx" in a single column table.
' The table should not include a header row.
'
Dim main As Document
Dim mainstr As String 'variable for holding the path of the main document
' Dim dlg As dialog 'optional for creating a dialog box to let the user select the "list.docx" but it is not included in this macro
' Dim fname As String 'optional for capturing the path string of the file selected via the dlg dialog.
Dim liststr As String 'variable holding the path string for "list.docx"
Dim list As Document
Dim tbl As Table 'the table in "list.docx"
Dim col1 As Range 'the first column of the "list.docx" table
Dim rng As Range
Dim i As Long 'generic counter for loops
Dim find As Variant 'final array holding all the find terms
Dim findstr As String 'intermediary string holding all the find terms
Dim replacestr As String 'intermediary string holding all the replace terms
Dim replace As Variant 'final array for holding all the replace terms
Dim temparray As Variant 'intermediary array holding the contents of the list.docx table
Dim temp As Variant 'second temporary array holding the contents of the list.docx table


Set main = ActiveDocument
mainstr = ActiveDocument.Path
'set the path for the find/replace list.  It should be a single, two-column table
Set list = Application.Documents.Open("list.docx") 'two column table, first column is find terms, second is replace terms
liststr = list.Path 'put list path into string
Set tbl = list.Tables(1) 'there should only be one table, no header, no blank cells in "find" column
Set rng = tbl.ConvertToText(Separator:=vbTab) 'temporarily converts the table to text to push into an array

temparray = Split(rng.Text, vbCr) 'pulls in rows as elements in the array
list.Undo 'put the converted text back into a table
ReDim temp(UBound(temparray))

For i = 0 To UBound(temparray) - 1
    temp(i) = Split(temparray(i), vbTab) 'split temparray into a Ubound(temparray) by 2 array
Next

i = 0 'reset the counter to zero

For i = 0 To UBound(temp) - 1 'create two separate strings one for find and one for replace,
'split these strings in the next step to create two arrays

    findstr = findstr & temp(i)(0) & vbCr
    replacestr = replacestr & temp(i)(1) & vbCr
Next

find = Split(findstr, vbCr)  'create the find array
replace = Split(replacestr, vbCr) 'create the replace array

ReDim Preserve find(UBound(find) - 1) 'get rid of the final empty item
ReDim Preserve replace(UBound(replace) - 1) 'remove the final, empty item

main.Activate 'go back to the main document
Selection.HomeKey unit:=wdStory 'place the cursor at the top of the document

i = 0 'reset the counter to zero


'loop through the find array, replace with the corresponding item from the replace array

For i = 0 To UBound(find)
    Selection.find.ClearFormatting
        Selection.find.Replacement.ClearFormatting
        With Selection.find
            .Text = find(i)
            .Replacement.Text = replace(i)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.find.Execute replace:=wdReplaceAll
Next i


End Sub



