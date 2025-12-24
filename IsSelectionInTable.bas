Option Explicit

Public Function IsSelectionInTable(Optional ByVal target As Range) As Boolean
    Dim rng As Range
    Dim lo As ListObject

    ' 対象の決定（Selection が Range でない場合は False）
    If target Is Nothing Then
        If TypeName(Selection) <> "Range" Then Exit Function
        Set rng = Selection
    Else
        Set rng = target
    End If
    If rng Is Nothing Then Exit Function

    ' 同じシート上にテーブルが無ければ False
    If rng.Parent.ListObjects.Count = 0 Then Exit Function

    ' 各テーブルと交差判定（複数領域選択も Intersect が処理）
    For Each lo In rng.Parent.ListObjects
        If Not Application.Intersect(rng, lo.Range) Is Nothing Then
            IsSelectionInTable = True
            Exit Function
        End If
    Next lo
End Function
