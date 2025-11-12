VERSION 5.00
Begin VB.UserForm TableCopyForm 
   Caption         =   "テーブル行をコピー"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblTableName
      Caption         =   "テーブル:"
      Height          =   255
      Left            =   150
      TabIndex        =   4
      Top             =   120
      Width           =   6300
   End
   Begin VB.Label lblInstruction
      Caption         =   "コピーしたい行を選択してください (複数選択可)。"
      Height          =   255
      Left            =   150
      TabIndex        =   3
      Top             =   360
      Width           =   6300
   End
   Begin VB.ListBox lstRows
      Height          =   2520
      IntegralHeight  =   0
      Left            =   150
      MultiSelect     =   1  'fmMultiSelectMulti
      TabIndex        =   0
      Top             =   660
      Width           =   6300
   End
   Begin VB.CheckBox chkTranspose
      Caption         =   "行と列を入れ換える"
      Height          =   255
      Left            =   150
      TabIndex        =   1
      Top             =   3300
      Width           =   2400
   End
   Begin VB.CommandButton btnCopy
      Caption         =   "クリップボードにコピー"
      Default         =   -1  'True
      Height          =   360
      Left            =   3420
      TabIndex        =   2
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel
      Cancel          =   -1  'True
      Caption         =   "キャンセル"
      Height          =   360
      Left            =   5010
      TabIndex        =   5
      Top             =   3720
      Width           =   1440
   End
End
Attribute VB_Name = "TableCopyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tableData As Variant
Private headerData As Variant

Public Sub InitializeFromListObject(ByVal lo As ListObject, Optional ByVal defaultRowIndex As Long = 0)
    lblTableName.Caption = "テーブル: " & lo.Name & " (" & lo.ListColumns.Count & " 列)"

    If lo.DataBodyRange Is Nothing Then
        lstRows.Clear
        btnCopy.Enabled = False
        headerData = EnsureTwoDimensionalArray(lo.HeaderRowRange.Value)
        Exit Sub
    End If

    headerData = EnsureTwoDimensionalArray(lo.HeaderRowRange.Value)
    tableData = EnsureTwoDimensionalArray(lo.DataBodyRange.Value)

    lstRows.ColumnCount = UBound(tableData, 2)
    lstRows.List = tableData
    lstRows.ListStyle = fmListStyleOption
    btnCopy.Enabled = (lstRows.ListCount > 0)

    Dim i As Long
    For i = 0 To lstRows.ListCount - 1
        lstRows.Selected(i) = False
    Next i

    If lstRows.ListCount > 0 Then
        Dim targetIndex As Long
        targetIndex = defaultRowIndex

        If targetIndex < 1 Or targetIndex > lstRows.ListCount Then
            targetIndex = 1
        End If

        lstRows.Selected(targetIndex - 1) = True
    End If
End Sub

Private Function EnsureTwoDimensionalArray(ByVal source As Variant) As Variant
    If Not IsArray(source) Then
        Dim tmp() As Variant
        ReDim tmp(1 To 1, 1 To 1)
        tmp(1, 1) = source
        EnsureTwoDimensionalArray = tmp
        Exit Function
    End If

    On Error GoTo ConvertSingleRow
    Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long
    lb1 = LBound(source, 1)
    ub1 = UBound(source, 1)
    lb2 = LBound(source, 2)
    ub2 = UBound(source, 2)
    EnsureTwoDimensionalArray = source
    Exit Function

ConvertSingleRow:
    Dim cols As Long
    cols = UBound(source) - LBound(source) + 1
    Dim arr() As Variant
    ReDim arr(1 To 1, 1 To cols)
    Dim i As Long
    For i = 1 To cols
        arr(1, i) = source(LBound(source) + i - 1)
    Next i
    EnsureTwoDimensionalArray = arr
End Function

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnCopy_Click()
    If lstRows.ListCount = 0 Then
        MsgBox "コピーできる行がありません。", vbInformation
        Exit Sub
    End If

    Dim selectedIndexes As Collection
    Set selectedIndexes = GetSelectedRowIndexes()

    If selectedIndexes.Count = 0 Then
        MsgBox "少なくとも1行選択してください。", vbInformation
        Exit Sub
    End If

    Dim dataToCopy As Variant
    dataToCopy = BuildSelectedData(selectedIndexes)

    If chkTranspose.Value = True Then
        dataToCopy = TransposeMatrix(dataToCopy)
    End If

    Dim clipboardText As String
    clipboardText = ConvertMatrixToTabSeparatedText(dataToCopy)

    CopyTextToClipboard clipboardText

    MsgBox "選択したデータをクリップボードにコピーしました。Word に貼り付けできます。", vbInformation
    Unload Me
End Sub

Private Function GetSelectedRowIndexes() As Collection
    Dim results As New Collection

    Dim i As Long
    For i = 0 To lstRows.ListCount - 1
        If lstRows.Selected(i) Then
            results.Add i + 1 ' 1-based index for tableData
        End If
    Next i

    Set GetSelectedRowIndexes = results
End Function

Private Function BuildSelectedData(ByRef selectedIndexes As Collection) As Variant
    Dim rowCount As Long
    rowCount = selectedIndexes.Count

    Dim colCount As Long
    colCount = UBound(tableData, 2) - LBound(tableData, 2) + 1

    Dim result() As Variant
    ReDim result(1 To rowCount + 1, 1 To colCount)

    Dim c As Long
    For c = 1 To colCount
        result(1, c) = headerData(1, c)
    Next c

    Dim r As Long
    For r = 1 To rowCount
        Dim srcRow As Long
        srcRow = selectedIndexes(r)

        For c = 1 To colCount
            result(r + 1, c) = tableData(srcRow, c)
        Next c
    Next r

    BuildSelectedData = result
End Function

Private Function TransposeMatrix(ByVal source As Variant) As Variant
    Dim rows As Long
    rows = UBound(source, 1) - LBound(source, 1) + 1

    Dim cols As Long
    cols = UBound(source, 2) - LBound(source, 2) + 1

    Dim result() As Variant
    ReDim result(1 To cols, 1 To rows)

    Dim r As Long, c As Long
    For r = 1 To rows
        For c = 1 To cols
            result(c, r) = source(LBound(source, 1) + r - 1, LBound(source, 2) + c - 1)
        Next c
    Next r

    TransposeMatrix = result
End Function

Private Function ConvertMatrixToTabSeparatedText(ByVal source As Variant) As String
    Dim rows As Long
    rows = UBound(source, 1) - LBound(source, 1) + 1

    Dim cols As Long
    cols = UBound(source, 2) - LBound(source, 2) + 1

    Dim sb As String
    Dim r As Long, c As Long
    For r = 1 To rows
        Dim rowValues() As String
        ReDim rowValues(1 To cols)

        For c = 1 To cols
            rowValues(c) = CStr(source(LBound(source, 1) + r - 1, LBound(source, 2) + c - 1))
        Next c

        sb = sb & Join(rowValues, vbTab)
        If r < rows Then
            sb = sb & vbCrLf
        End If
    Next r

    ConvertMatrixToTabSeparatedText = sb
End Function

Private Sub CopyTextToClipboard(ByVal textToCopy As String)
    Dim dataObj As Object
    Set dataObj = CreateObject("MSForms.DataObject")
    dataObj.SetText textToCopy
    dataObj.PutInClipboard
End Sub
