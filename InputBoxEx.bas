Attribute VB_Name = "modInputBoxEx"
' InputBoxEx
'  VBに標準で入っているInputBox関数は挙動がおかしいので自作したもの
'  ヘルプ関係が未対応だが、
'  入力文字数の制限や、キャンセルが押されたときの戻り値の設定もできる。
Option Explicit

Dim d As String

Public Function InputBoxEx(Prompt, Optional Title, Optional Default, Optional XPos, Optional YPos, Optional HelpFile, Optional Context, Optional Length) As String
    With FormInputBoxEx
        .lblMessage.Caption = CStr(Prompt)
        If Not IsMissing(Title) Then .Caption = CStr(Title)
        If Not IsMissing(Default) Then d = CStr(Default)
        .txtInput.Text = d
        If Not IsMissing(XPos) Then .Left = CSng(XPos)
        If Not IsMissing(YPos) Then .Top = CSng(YPos)
        If Not IsMissing(Length) Then .txtInput.MaxLength = CInt(Length)
        .m = 1
        .Show vbModal
        If .m = 2 Then InputBoxEx = .txtInput.Text Else InputBoxEx = d
        .m = 0
        Unload FormInputBoxEx
    End With
End Function

