Attribute VB_Name = "TableCopyContextMenu"
Option Explicit

Private Const MENU_TAG As String = "TableCopyContextMenuItem"
Private Const MENU_CAPTION As String = "テーブル行をWord用にコピー..."
Private Const TABLE_COMMAND_BAR As String = "List Range Popup"
Private Const CELL_COMMAND_BAR As String = "Cell"

' エントリーポイント: ブックを開いた時に呼び出されます。
Public Sub Auto_Open()
    InstallTableCopyMenu
End Sub

' エントリーポイント: ブックを閉じる時に呼び出されます。
Public Sub Auto_Close()
    RemoveTableCopyMenu
End Sub

' コンテキストメニューにコマンドを追加します。
Public Sub InstallTableCopyMenu()
    On Error Resume Next
    RemoveTableCopyMenu
    On Error GoTo 0

    AddMenuToCommandBar TABLE_COMMAND_BAR
    AddMenuToCommandBar CELL_COMMAND_BAR
End Sub

' コンテキストメニューからコマンドを削除します。
Public Sub RemoveTableCopyMenu()
    RemoveMenuFromCommandBar TABLE_COMMAND_BAR
    RemoveMenuFromCommandBar CELL_COMMAND_BAR
End Sub

' メニューから呼び出されるエントリーポイントです。
Public Sub ShowTableCopyDialog()
    Dim lo As ListObject

    On Error Resume Next
    Set lo = ActiveCell.ListObject
    On Error GoTo 0

    If lo Is Nothing Then
        MsgBox "Excel テーブル内のセルを選択した状態で実行してください。", vbExclamation
        Exit Sub
    End If

    Dim frm As TableCopyForm
    Set frm = New TableCopyForm

    Dim defaultRowIndex As Long
    defaultRowIndex = GetActiveCellRowIndex(lo)

    frm.InitializeFromListObject lo, defaultRowIndex
    frm.Show
End Sub

Private Function GetActiveCellRowIndex(ByVal lo As ListObject) As Long
    If lo.DataBodyRange Is Nothing Then
        GetActiveCellRowIndex = 0
        Exit Function
    End If

    If Intersect(ActiveCell, lo.DataBodyRange) Is Nothing Then
        GetActiveCellRowIndex = 0
        Exit Function
    End If

    GetActiveCellRowIndex = ActiveCell.Row - lo.DataBodyRange.Row + 1
End Function

Private Sub AddMenuToCommandBar(ByVal commandBarName As String)
    Dim targetBar As CommandBar

    On Error Resume Next
    Set targetBar = Application.CommandBars(commandBarName)
    On Error GoTo 0

    If targetBar Is Nothing Then Exit Sub

    Dim control As CommandBarControl
    Set control = targetBar.Controls.Add(Type:=msoControlButton, Temporary:=True)

    With control
        .Caption = MENU_CAPTION
        .Tag = MENU_TAG
        .BeginGroup = True
        .OnAction = "ShowTableCopyDialog"
        .FaceId = 130
    End With
End Sub

Private Sub RemoveMenuFromCommandBar(ByVal commandBarName As String)
    Dim targetBar As CommandBar

    On Error Resume Next
    Set targetBar = Application.CommandBars(commandBarName)
    On Error GoTo 0

    If targetBar Is Nothing Then Exit Sub

    Dim ctrl As CommandBarControl
    For Each ctrl In targetBar.Controls
        If ctrl.Tag = MENU_TAG Then
            ctrl.Delete
        End If
    Next ctrl
End Sub
