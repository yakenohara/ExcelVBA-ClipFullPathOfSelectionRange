Attribute VB_Name = "ClipFullPathOfSelectionRange"
'<License>------------------------------------------------------------
'
' Copyright (c) 2019 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>

'
'選択範囲(セル範囲)を、ファイルのパスを含めて文字列にしてクリップボードにコピーする
'
Sub ClipFullPathOfSelectionRange()

    '<Settings>--------------------------------------------------
    
    'styleStr_DirStart = "<file://"
    styleStr_DirStart = "<"
    styleStr_DirEnd = ">"
    styleStr_EndOfList = "┗"
    styleStr_PrefOfSheetName = "Sheet:`"
    styleStr_SuffOfSheetName = "`"
    styleStr_PrefOfAddressName = "Address:`"
    styleStr_SuffOfAddressName = "`"
    styleStr_Indent = "  "
    
    '-------------------------------------------------</Settings>

    dirStr = ActiveWorkbook.Path 'ファイルディレクトリの取得
    If dirStr = "" Then '未保存の場合
        ret = MsgBox("ActiveWorkbook.Path stores null string.(Maybe this workbook is not saved yet)", vbExclamation)
        Exit Sub
    End If
    filenameStr = ActiveWorkbook.Name 'ファイル名の取得
    
'    Debug.Print ("SelectionType:" + TypeName(Selection))
    
    shtnameStr = ActiveWorkbook.ActiveSheet.Name 'シート名の取得
    
    Set selectingRange = getSelectingRange() '選択範囲の取得
    
    If selectingRange Is Nothing Then '位置取得エラーの場合
        ret = MsgBox("Cannot find selecting cell range", vbExclamation)
        Exit Sub
    End If
    
'    Debug.Print ("dirStr:" + dirStr)
'    Debug.Print ("filenameStr:" + filenameStr)
'    Debug.Print ("shtnameStr:" + shtnameStr)
'    Debug.Print ("Range:" + selectingRange.Address(RowAbsolute:=False, ColumnAbsolute:=False))
    
    '文字列生成
    pathStr = ""
    pathStr = styleStr_DirStart + dirStr + styleStr_DirEnd 'ファイル格納ディレクトリ
    pathStr = pathStr + vbCrLf + styleStr_EndOfList + filenameStr 'ファイル名
    pathStr = pathStr + vbCrLf + WorksheetFunction.Rept(styleStr_Indent, 1) + styleStr_EndOfList + styleStr_PrefOfSheetName + shtnameStr + styleStr_SuffOfSheetName 'シート名
    pathStr = pathStr + vbCrLf + WorksheetFunction.Rept(styleStr_Indent, 2) + styleStr_EndOfList + styleStr_PrefOfAddressName + selectingRange.Address(RowAbsolute:=False, ColumnAbsolute:=False) + styleStr_SuffOfAddressName '選択範囲のアドレス
    
'    Debug.Print pathStr
    
    'クリップボードにコピー
    SetCB (pathStr)
    
End Sub

'
'現在選択しているセルの左上部分を返す
'
'エラー発生時は Nothing を返す
'
Private Function getSelectingRange() As Variant

    Dim retRange As Variant
    
    On Error GoTo TopLeftCell_is_not_defined
    
    Select Case (TypeName(Selection))
        Case "Range"
            Set retRange = Selection '選択範囲をそのまま返す
        
        Case Else
            'オブジェクトが配置されている範囲のセルを返す
            Set retRange = ActiveWorkbook.ActiveSheet.Range(Selection.TopLeftCell, Selection.BottomRightCell)
            
    End Select
    
    Set getSelectingRange = retRange
    Exit Function
    
TopLeftCell_is_not_defined: 'オブジェクトに .TopLeftCell or .BottomRightCell プロパティが存在しなかった場合
    
    Set retRange = Nothing
    
    Set getSelectingRange = retRange
    Exit Function
    
End Function

'<クリップボード操作>-------------------------------------------

'クリップボードに文字列を格納
Private Sub SetCB(ByVal str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    .Text = str
    .SelStart = 0
    .SelLength = .TextLength
    .Copy
  End With
End Sub

'クリップボードから文字列を取得
Private Sub GetCB(ByRef str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    If .CanPaste = True Then .Paste
    str = .Text
  End With
End Sub

'------------------------------------------</クリップボード操作>

