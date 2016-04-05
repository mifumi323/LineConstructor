VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStringArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_Element() As String

Private Sub Class_Initialize()
    InitSize 0
End Sub

'サイズを指定して初期化
Public Sub InitSize(Size As Long)
    ReDim m_Element(Size)   '１個要素が多いけど気にしない
End Sub

'文字をSplitした配列で初期化
Public Sub InitSplit(Expression As String, Delimiter As String)
    m_Element = Split(Expression, Delimiter)
    ChangeSize GetSize() + 1
End Sub

'内容を変えずにサイズ変更
Public Sub ChangeSize(Size As Long)
    ReDim Preserve m_Element(Size)   '１個要素が多いけど気にしない
    m_Element(Size) = ""
End Sub

'配列のサイズ
Public Function GetSize() As Long
    GetSize = UBound(m_Element)
End Function

'要素が存在するかどうか
Public Function IsExist(Index As Long) As Boolean
    IsExist = 0 <= Index And Index < UBound(m_Element)
End Function

'要素を得る(説明不要だワナ)
Public Function GetElement(Index As Long)
    If Index < 0 Then Exit Function
    If UBound(m_Element) <= Index Then Exit Function
    GetElement = m_Element(Index)
End Function

'要素を設定する(説明不要だワナ)
Public Sub SetElement(Index As Long, NewValue As String)
    If Index < 0 Then Exit Sub
    If UBound(m_Element) <= Index Then Exit Sub
    m_Element(Index) = NewValue
End Sub

'要素を最初に追加する
Public Sub Unshift(NewValue As String)
    Insert 0, NewValue
End Sub

'最初の要素を削除してその値を得る
Public Function Shift() As String
    Shift = Delete(0)
End Function

'要素を最後に追加する
Public Sub Push(NewValue As String)
    ChangeSize GetSize() + 1
    SetElement GetSize() - 1, NewValue
End Sub

'最後の要素を削除してその値を得る
Public Function Pop() As String
    If GetSize() = 0 Then Exit Function
    Pop = GetElement(GetSize() - 1)
    ChangeSize GetSize() - 1
End Function

'任意の位置に要素を追加し、それ以降の要素を一つずつ後ろにずらす
Public Sub Insert(Index As Long, NewValue As String)
    If Index < 0 Or GetSize() < Index Then Exit Sub
    ChangeSize GetSize() + 1
    Dim I As Long
    For I = GetSize() - 2 To Index Step -1
        m_Element(I + 1) = m_Element(I)
    Next
    m_Element(Index) = NewValue
End Sub

'任意の要素を削除してその値を得る
Public Function Delete(Index As Long)
    If Index < 0 Or GetSize() <= Index Then Exit Function
    Delete = m_Element(Index)
    Dim I As Long
    For I = Index To GetSize() - 2
        m_Element(I) = m_Element(I + 1)
    Next
    ChangeSize GetSize() - 1
End Function

Public Property Get Element(Index As Long) As String
    Element = GetElement(Index)
End Property

Public Property Let Element(Index As Long, ByVal NewValue As String)
    SetElement Index, NewValue
End Property

'五十音順にソートする
Public Sub Sort(Optional ByVal lngCompare As Long)
    StrSort m_Element, 0, GetSize() - 1, lngCompare
End Sub

Private Sub StrSort _
    (ByRef strArray() As String, _
     ByVal lngStart As Long, _
     ByVal lngEnd As Long, _
     Optional ByVal lngCompare As Long)

 Dim lngBaseNumber As Long                                          '中央の要素番号を格納する変数
 Dim strBaseValue As String                                         '基準値を格納する変数
 Dim lngCounter As Long                                             '格納位置カウンタ
 Dim strBuffer As String                                            '値をスワップするための作業域
 Dim I As Long                                                      'ループカウンタ
 
    If lngStart >= lngEnd Then Exit Sub                             '終了番号が開始番号以下の場合、プロシージャを抜ける
    lngBaseNumber = (lngStart + lngEnd) \ 2                         '中央の要素番号を求める
    strBaseValue = strArray(lngBaseNumber)                          '中央の値を基準値とする
    strArray(lngBaseNumber) = strArray(lngStart)                    '中央の要素に開始番号の値を格納
    lngCounter = lngStart                                           '格納位置カウンタを開始番号と同じにする
    For I = (lngStart + 1) To lngEnd Step 1                         '開始番号の次の要素から終了番号までループ
        If StrComp(strArray(I), strBaseValue, lngCompare) = -1 Then '値が基準値より小さい場合
            lngCounter = lngCounter + 1                             '格納位置カウンタをインクリメント
            strBuffer = strArray(lngCounter)                        'strArray(i) と strArray(lngCounter) の値をスワップ
            strArray(lngCounter) = strArray(I)
            strArray(I) = strBuffer
        End If
    Next I
    strArray(lngStart) = strArray(lngCounter)                       'strArray(lngCounter) を開始番号の値にする
    strArray(lngCounter) = strBaseValue                             '基準値を strArray(lngCounter) に格納
    Call StrSort(strArray(), lngStart, lngCounter - 1)              '分割された配列をクイックソート(再帰)
    Call StrSort(strArray(), lngCounter + 1, lngEnd)                '分割された配列をクイックソート(再帰)

End Sub

'五十音順にソートする
Public Sub ReverseSort(Optional ByVal lngCompare As Long)
    StrRSort m_Element, 0, GetSize() - 1, lngCompare
End Sub

Private Sub StrRSort _
    (ByRef strArray() As String, _
     ByVal lngStart As Long, _
     ByVal lngEnd As Long, _
     Optional ByVal lngCompare As Long)

 Dim lngBaseNumber As Long                                          '中央の要素番号を格納する変数
 Dim strBaseValue As String                                         '基準値を格納する変数
 Dim lngCounter As Long                                             '格納位置カウンタ
 Dim strBuffer As String                                            '値をスワップするための作業域
 Dim I As Long                                                      'ループカウンタ
 
    If lngStart >= lngEnd Then Exit Sub                             '終了番号が開始番号以下の場合、プロシージャを抜ける
    lngBaseNumber = (lngStart + lngEnd) \ 2                         '中央の要素番号を求める
    strBaseValue = strArray(lngBaseNumber)                          '中央の値を基準値とする
    strArray(lngBaseNumber) = strArray(lngStart)                    '中央の要素に開始番号の値を格納
    lngCounter = lngStart                                           '格納位置カウンタを開始番号と同じにする
    For I = (lngStart + 1) To lngEnd Step 1                         '開始番号の次の要素から終了番号までループ
        If StrComp(strArray(I), strBaseValue, lngCompare) = 1 Then  '値が基準値より大きい場合
            lngCounter = lngCounter + 1                             '格納位置カウンタをインクリメント
            strBuffer = strArray(lngCounter)                        'strArray(i) と strArray(lngCounter) の値をスワップ
            strArray(lngCounter) = strArray(I)
            strArray(I) = strBuffer
        End If
    Next I
    strArray(lngStart) = strArray(lngCounter)                       'strArray(lngCounter) を開始番号の値にする
    strArray(lngCounter) = strBaseValue                             '基準値を strArray(lngCounter) に格納
    Call StrRSort(strArray(), lngStart, lngCounter - 1)              '分割された配列をクイックソート(再帰)
    Call StrRSort(strArray(), lngCounter + 1, lngEnd)                '分割された配列をクイックソート(再帰)

End Sub

Public Sub Shuffle()
    Dim I As Long
    For I = 0 To GetSize() - 1
        Swap I, Int(Rnd() * GetSize())
    Next
End Sub

Public Sub Swap(Index1 As Long, Index2 As Long)
    If Index1 = Index2 Then Exit Sub
    If Index1 < 0 Then Exit Sub
    If UBound(m_Element) <= Index1 Then Exit Sub
    If Index2 < 0 Then Exit Sub
    If UBound(m_Element) <= Index2 Then Exit Sub
    Dim Buf As String
    Buf = m_Element(Index1)
    m_Element(Index1) = m_Element(Index2)
    m_Element(Index2) = Buf
End Sub
