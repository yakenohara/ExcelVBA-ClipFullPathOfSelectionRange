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
'�I��͈�(�Z���͈�)���A�t�@�C���̃p�X���܂߂ĕ�����ɂ��ăN���b�v�{�[�h�ɃR�s�[����
'
Sub ClipFullPathOfSelectionRange()

    '<Settings>--------------------------------------------------
    
    'styleStr_DirStart = "<file://"
    styleStr_DirStart = "<"
    styleStr_DirEnd = ">"
    styleStr_EndOfList = "��"
    styleStr_PrefOfSheetName = "Sheet:`"
    styleStr_SuffOfSheetName = "`"
    styleStr_PrefOfAddressName = "Address:`"
    styleStr_SuffOfAddressName = "`"
    styleStr_Indent = "  "
    
    '-------------------------------------------------</Settings>

    dirStr = ActiveWorkbook.Path '�t�@�C���f�B���N�g���̎擾
    If dirStr = "" Then '���ۑ��̏ꍇ
        ret = MsgBox("ActiveWorkbook.Path stores null string.(Maybe this workbook is not saved yet)", vbExclamation)
        Exit Sub
    End If
    filenameStr = ActiveWorkbook.Name '�t�@�C�����̎擾
    
'    Debug.Print ("SelectionType:" + TypeName(Selection))
    
    shtnameStr = ActiveWorkbook.ActiveSheet.Name '�V�[�g���̎擾
    
    Set selectingRange = getSelectingRange() '�I��͈͂̎擾
    
    If selectingRange Is Nothing Then '�ʒu�擾�G���[�̏ꍇ
        ret = MsgBox("Cannot find selecting cell range", vbExclamation)
        Exit Sub
    End If
    
'    Debug.Print ("dirStr:" + dirStr)
'    Debug.Print ("filenameStr:" + filenameStr)
'    Debug.Print ("shtnameStr:" + shtnameStr)
'    Debug.Print ("Range:" + selectingRange.Address(RowAbsolute:=False, ColumnAbsolute:=False))
    
    '�����񐶐�
    pathStr = ""
    pathStr = styleStr_DirStart + dirStr + styleStr_DirEnd '�t�@�C���i�[�f�B���N�g��
    pathStr = pathStr + vbCrLf + styleStr_EndOfList + filenameStr '�t�@�C����
    pathStr = pathStr + vbCrLf + WorksheetFunction.Rept(styleStr_Indent, 1) + styleStr_EndOfList + styleStr_PrefOfSheetName + shtnameStr + styleStr_SuffOfSheetName '�V�[�g��
    pathStr = pathStr + vbCrLf + WorksheetFunction.Rept(styleStr_Indent, 2) + styleStr_EndOfList + styleStr_PrefOfAddressName + selectingRange.Address(RowAbsolute:=False, ColumnAbsolute:=False) + styleStr_SuffOfAddressName '�I��͈͂̃A�h���X
    
'    Debug.Print pathStr
    
    '�N���b�v�{�[�h�ɃR�s�[
    SetCB (pathStr)
    
End Sub

'
'���ݑI�����Ă���Z���̍��㕔����Ԃ�
'
'�G���[�������� Nothing ��Ԃ�
'
Private Function getSelectingRange() As Variant

    Dim retRange As Variant
    
    On Error GoTo TopLeftCell_is_not_defined
    
    Select Case (TypeName(Selection))
        Case "Range"
            Set retRange = Selection '�I��͈͂����̂܂ܕԂ�
        
        Case Else
            '�I�u�W�F�N�g���z�u����Ă���͈͂̃Z����Ԃ�
            Set retRange = ActiveWorkbook.ActiveSheet.Range(Selection.TopLeftCell, Selection.BottomRightCell)
            
    End Select
    
    Set getSelectingRange = retRange
    Exit Function
    
TopLeftCell_is_not_defined: '�I�u�W�F�N�g�� .TopLeftCell or .BottomRightCell �v���p�e�B�����݂��Ȃ������ꍇ
    
    Set retRange = Nothing
    
    Set getSelectingRange = retRange
    Exit Function
    
End Function

'<�N���b�v�{�[�h����>-------------------------------------------

'�N���b�v�{�[�h�ɕ�������i�[
Private Sub SetCB(ByVal str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    .Text = str
    .SelStart = 0
    .SelLength = .TextLength
    .Copy
  End With
End Sub

'�N���b�v�{�[�h���當������擾
Private Sub GetCB(ByRef str As String)
  With CreateObject("Forms.TextBox.1")
    .MultiLine = True
    If .CanPaste = True Then .Paste
    str = .Text
  End With
End Sub

'------------------------------------------</�N���b�v�{�[�h����>

