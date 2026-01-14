# xlsx

このリポジトリには、Excel テーブル（ListObject）のセルを右クリックした際に表示されるコンテキストメニューへ「テーブル行をWord用にコピー...」というメニューを追加し、選択した行を Word に貼り付け可能な形式でクリップボードへコピーするための VBA モジュール一式を収録しています。メニューを実行すると、対象テーブルの行を複数選択できるリストボックスと、行列の入れ換え（転置）を指定できるチェックボックスを備えたフォームが表示されます。通常のセルメニューにも同じコマンドを追加するため、テーブル以外からのアクセスも可能です。

## 構成ファイル

| パス | 説明 |
| ---- | ---- |
| `VBA/TableCopyContextMenu.bas` | コンテキストメニュー項目の追加・削除、およびフォーム表示を担う標準モジュール。 |
| `VBA/TableCopyForm.frm` | 行選択・転置指定・コピー処理を行うユーザーフォーム。 |

## 使い方

1. Excel の VBA エディターを開き、任意のブックへ本リポジトリの `VBA/TableCopyContextMenu.bas` と `VBA/TableCopyForm.frm` をそれぞれインポートします。
2. 参照設定で **Microsoft Forms 2.0 Object Library** を有効にします（`MSForms.DataObject` を利用するため）。
3. `ThisWorkbook` モジュールに次のコードを追加し、ブックのオープン／クローズ時にコンテキストメニューを更新します。

   ```vb
   Private Sub Workbook_Open()
       TableCopyContextMenu.InstallTableCopyMenu
   End Sub

   Private Sub Workbook_BeforeClose(ByVal Cancel As Boolean)
       TableCopyContextMenu.RemoveTableCopyMenu
   End Sub
   ```

   Auto_Open / Auto_Close マクロにも対応しているため、個別の呼び出しが難しい場合はマクロ有効ブックとして保存するだけでも動作します。

4. 対象となる Excel テーブル内のセルを右クリックすると、新しいメニューが表示されます。メニューをクリックするとフォームが開き、右クリックしたセルを含む行が既定で選択された状態になります。そのまま、または複数の行を追加で選択して「クリップボードにコピー」を実行できます。必要に応じて「行と列を入れ換える」にチェックを入れると、選択内容を転置した形でコピーできます。
5. コピー後は Word などにそのまま貼り付けると、タブ区切りテキストが表として貼り付けられます。

## 注意事項

- コンテキストメニューは Excel アプリケーション全体に影響するため、不要になった場合は `TableCopyContextMenu.RemoveTableCopyMenu` を実行して削除してください。
- クリップボードへのコピーには `MSForms.DataObject` を利用します。参照設定が行われていない場合はランタイムエラーになります。
- テーブルのデータ部分が空の場合は、フォーム上でコピー用ボタンが無効化されます。
