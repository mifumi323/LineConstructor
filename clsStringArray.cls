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

'�T�C�Y���w�肵�ď�����
Public Sub InitSize(Size As Long)
    ReDim m_Element(Size)   '�P�v�f���������ǋC�ɂ��Ȃ�
End Sub

'������Split�����z��ŏ�����
Public Sub InitSplit(Expression As String, Delimiter As String)
    m_Element = Split(Expression, Delimiter)
    ChangeSize GetSize() + 1
End Sub

'���e��ς����ɃT�C�Y�ύX
Public Sub ChangeSize(Size As Long)
    ReDim Preserve m_Element(Size)   '�P�v�f���������ǋC�ɂ��Ȃ�
    m_Element(Size) = ""
End Sub

'�z��̃T�C�Y
Public Function GetSize() As Long
    GetSize = UBound(m_Element)
End Function

'�v�f�����݂��邩�ǂ���
Public Function IsExist(Index As Long) As Boolean
    IsExist = 0 <= Index And Index < UBound(m_Element)
End Function

'�v�f�𓾂�(�����s�v�����i)
Public Function GetElement(Index As Long)
    If Index < 0 Then Exit Function
    If UBound(m_Element) <= Index Then Exit Function
    GetElement = m_Element(Index)
End Function

'�v�f��ݒ肷��(�����s�v�����i)
Public Sub SetElement(Index As Long, NewValue As String)
    If Index < 0 Then Exit Sub
    If UBound(m_Element) <= Index Then Exit Sub
    m_Element(Index) = NewValue
End Sub

'�v�f���ŏ��ɒǉ�����
Public Sub Unshift(NewValue As String)
    Insert 0, NewValue
End Sub

'�ŏ��̗v�f���폜���Ă��̒l�𓾂�
Public Function Shift() As String
    Shift = Delete(0)
End Function

'�v�f���Ō�ɒǉ�����
Public Sub Push(NewValue As String)
    ChangeSize GetSize() + 1
    SetElement GetSize() - 1, NewValue
End Sub

'�Ō�̗v�f���폜���Ă��̒l�𓾂�
Public Function Pop() As String
    If GetSize() = 0 Then Exit Function
    Pop = GetElement(GetSize() - 1)
    ChangeSize GetSize() - 1
End Function

'�C�ӂ̈ʒu�ɗv�f��ǉ����A����ȍ~�̗v�f��������ɂ��炷
Public Sub Insert(Index As Long, NewValue As String)
    If Index < 0 Or GetSize() < Index Then Exit Sub
    ChangeSize GetSize() + 1
    Dim I As Long
    For I = GetSize() - 2 To Index Step -1
        m_Element(I + 1) = m_Element(I)
    Next
    m_Element(Index) = NewValue
End Sub

'�C�ӂ̗v�f���폜���Ă��̒l�𓾂�
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

'�܏\�����Ƀ\�[�g����
Public Sub Sort(Optional ByVal lngCompare As Long)
    StrSort m_Element, 0, GetSize() - 1, lngCompare
End Sub

Private Sub StrSort _
    (ByRef strArray() As String, _
     ByVal lngStart As Long, _
     ByVal lngEnd As Long, _
     Optional ByVal lngCompare As Long)

 Dim lngBaseNumber As Long                                          '�����̗v�f�ԍ����i�[����ϐ�
 Dim strBaseValue As String                                         '��l���i�[����ϐ�
 Dim lngCounter As Long                                             '�i�[�ʒu�J�E���^
 Dim strBuffer As String                                            '�l���X���b�v���邽�߂̍�ƈ�
 Dim I As Long                                                      '���[�v�J�E���^
 
    If lngStart >= lngEnd Then Exit Sub                             '�I���ԍ����J�n�ԍ��ȉ��̏ꍇ�A�v���V�[�W���𔲂���
    lngBaseNumber = (lngStart + lngEnd) \ 2                         '�����̗v�f�ԍ������߂�
    strBaseValue = strArray(lngBaseNumber)                          '�����̒l����l�Ƃ���
    strArray(lngBaseNumber) = strArray(lngStart)                    '�����̗v�f�ɊJ�n�ԍ��̒l���i�[
    lngCounter = lngStart                                           '�i�[�ʒu�J�E���^���J�n�ԍ��Ɠ����ɂ���
    For I = (lngStart + 1) To lngEnd Step 1                         '�J�n�ԍ��̎��̗v�f����I���ԍ��܂Ń��[�v
        If StrComp(strArray(I), strBaseValue, lngCompare) = -1 Then '�l����l��菬�����ꍇ
            lngCounter = lngCounter + 1                             '�i�[�ʒu�J�E���^���C���N�������g
            strBuffer = strArray(lngCounter)                        'strArray(i) �� strArray(lngCounter) �̒l���X���b�v
            strArray(lngCounter) = strArray(I)
            strArray(I) = strBuffer
        End If
    Next I
    strArray(lngStart) = strArray(lngCounter)                       'strArray(lngCounter) ���J�n�ԍ��̒l�ɂ���
    strArray(lngCounter) = strBaseValue                             '��l�� strArray(lngCounter) �Ɋi�[
    Call StrSort(strArray(), lngStart, lngCounter - 1)              '�������ꂽ�z����N�C�b�N�\�[�g(�ċA)
    Call StrSort(strArray(), lngCounter + 1, lngEnd)                '�������ꂽ�z����N�C�b�N�\�[�g(�ċA)

End Sub

'�܏\�����Ƀ\�[�g����
Public Sub ReverseSort(Optional ByVal lngCompare As Long)
    StrRSort m_Element, 0, GetSize() - 1, lngCompare
End Sub

Private Sub StrRSort _
    (ByRef strArray() As String, _
     ByVal lngStart As Long, _
     ByVal lngEnd As Long, _
     Optional ByVal lngCompare As Long)

 Dim lngBaseNumber As Long                                          '�����̗v�f�ԍ����i�[����ϐ�
 Dim strBaseValue As String                                         '��l���i�[����ϐ�
 Dim lngCounter As Long                                             '�i�[�ʒu�J�E���^
 Dim strBuffer As String                                            '�l���X���b�v���邽�߂̍�ƈ�
 Dim I As Long                                                      '���[�v�J�E���^
 
    If lngStart >= lngEnd Then Exit Sub                             '�I���ԍ����J�n�ԍ��ȉ��̏ꍇ�A�v���V�[�W���𔲂���
    lngBaseNumber = (lngStart + lngEnd) \ 2                         '�����̗v�f�ԍ������߂�
    strBaseValue = strArray(lngBaseNumber)                          '�����̒l����l�Ƃ���
    strArray(lngBaseNumber) = strArray(lngStart)                    '�����̗v�f�ɊJ�n�ԍ��̒l���i�[
    lngCounter = lngStart                                           '�i�[�ʒu�J�E���^���J�n�ԍ��Ɠ����ɂ���
    For I = (lngStart + 1) To lngEnd Step 1                         '�J�n�ԍ��̎��̗v�f����I���ԍ��܂Ń��[�v
        If StrComp(strArray(I), strBaseValue, lngCompare) = 1 Then  '�l����l���傫���ꍇ
            lngCounter = lngCounter + 1                             '�i�[�ʒu�J�E���^���C���N�������g
            strBuffer = strArray(lngCounter)                        'strArray(i) �� strArray(lngCounter) �̒l���X���b�v
            strArray(lngCounter) = strArray(I)
            strArray(I) = strBuffer
        End If
    Next I
    strArray(lngStart) = strArray(lngCounter)                       'strArray(lngCounter) ���J�n�ԍ��̒l�ɂ���
    strArray(lngCounter) = strBaseValue                             '��l�� strArray(lngCounter) �Ɋi�[
    Call StrRSort(strArray(), lngStart, lngCounter - 1)              '�������ꂽ�z����N�C�b�N�\�[�g(�ċA)
    Call StrRSort(strArray(), lngCounter + 1, lngEnd)                '�������ꂽ�z����N�C�b�N�\�[�g(�ċA)

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