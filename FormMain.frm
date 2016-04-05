VERSION 5.00
Begin VB.Form FormMain 
   Caption         =   "LineConstructor"
   ClientHeight    =   5310
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6705
   Icon            =   "FormMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  '�蓮
   ScaleHeight     =   5310
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.HScrollBar hscData 
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.VScrollBar vscData 
      Height          =   1215
      Left            =   4800
      TabIndex        =   6
      Top             =   1440
      Width           =   255
   End
   Begin VB.PictureBox picToolBar 
      Align           =   1  '�㑵��
      BorderStyle     =   0  '�Ȃ�
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6705
      TabIndex        =   0
      Top             =   0
      Width           =   6705
      Begin VB.CommandButton cmdDown 
         Caption         =   "��"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         ToolTipText     =   "���"
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "��"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         ToolTipText     =   "���"
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "-"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         ToolTipText     =   "�폜"
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "+"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         ToolTipText     =   "�ǉ�"
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "E"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "�ҏW"
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�t�@�C��(&F)"
      Begin VB.Menu mnuFileNew 
         Caption         =   "�V�K�쐬(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "�J��(&O)..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "�㏑���ۑ�(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "���O�����ĕۑ�(&A)..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�I��(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�ҏW(&E)"
      Begin VB.Menu mnuEditEdit 
         Caption         =   "�I�����ڂ�ҏW(&E)"
      End
      Begin VB.Menu mnuEditEditS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAddBefore 
         Caption         =   "�I�����ڂ̑O�ɒǉ�(&B)"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuEditAddAfter 
         Caption         =   "�I�����ڂ̌�ɒǉ�(&A)"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu mnuEditAddFirst 
         Caption         =   "�擪�ɒǉ�(&T)"
      End
      Begin VB.Menu mnuEditAddLast 
         Caption         =   "�Ō�ɒǉ�(&P)"
      End
      Begin VB.Menu mnuEditAddS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "�폜(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditDeleteS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditMoveUp 
         Caption         =   "��ֈړ�(&U)"
      End
      Begin VB.Menu mnuEditMoveDown 
         Caption         =   "���ֈړ�(&N)"
      End
      Begin VB.Menu mnuEditMoveFirst 
         Caption         =   "��ԏ�Ɉړ�(&F)"
      End
      Begin VB.Menu mnuEditMoveLast 
         Caption         =   "��ԉ��Ɉړ�(&L)"
      End
      Begin VB.Menu mnuEditMoves 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "����(&1)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditFindDown 
         Caption         =   "��������(2)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditFindUp 
         Caption         =   "�������(&2)"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuEditFindS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSort 
         Caption         =   "�\�[�g(&S)"
      End
      Begin VB.Menu mnuEditSortReverse 
         Caption         =   "�t�\�[�g(&R)"
      End
      Begin VB.Menu mnuSortShuffle 
         Caption         =   "�V���b�t��(&J)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "�w���v(&H)"
      Begin VB.Menu mnuHelpHowTo 
         Caption         =   "�g����(&H)"
      End
      Begin VB.Menu mnuHelpHowToS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "�o�[�W����(&A)"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Lines As New clsStringArray
Dim FileName As String
Dim SelectedLine As Long
Dim FindString As String
Dim Scroll As New clsLongScroll

Private Sub cmdAdd_Click()
    mnuEditAddAfter_Click
End Sub

Private Sub cmdDelete_Click()
    mnuEditDelete_Click
End Sub

Private Sub cmdDown_Click()
    mnuEditMoveDown_Click
End Sub

Private Sub cmdEdit_Click()
    mnuEditEdit_Click
End Sub

Private Sub cmdUp_Click()
    mnuEditMoveUp_Click
End Sub

Private Sub Form_DblClick()
    Randomize
    mnuEditEdit_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp
        If SelectedLine <= 0 Then Exit Sub
        SelectedLine = SelectedLine - 1
        Focus
        OnDraw
    Case vbKeyDown
        If SelectedLine >= Lines.GetSize() - 1 Then Exit Sub
        SelectedLine = SelectedLine + 1
        Focus
        OnDraw
    End Select
End Sub

Private Sub Form_Load()
    Scroll.ScrollBar = vscData
    Dim fn As String
    fn = Replace(Command$, """", "")
    If Dir(fn) <> "" And fn <> "" Then
        OpenFile fn
    Else
        mnuFileNew_Click
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Dim TH As Single
        Dim Offset As Long
        Dim bkColor As Long
        Dim T As Single
        Dim Sel As Long
        T = picToolBar.Height
        TH = TextHeight(" ")
        Offset = Scroll.Value
        Sel = Int((Y - T) / TH) + Offset
        If Lines.IsExist(Sel) Then SelectedLine = Sel
        OnDraw
        SetCapture hWnd
    ElseIf Button = vbRightButton Then
        PopupMenu mnuEdit, vbPopupMenuRightButton
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Dim TH As Single
        Dim Offset As Long
        Dim bkColor As Long
        Dim T As Single
        Dim Sel As Long
        T = picToolBar.Height
        TH = TextHeight(" ")
        Offset = Scroll.Value
        Sel = Int((Y - T) / TH) + Offset
        If Lines.IsExist(Sel) Then SelectedLine = Sel
        Focus
        OnDraw
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    OpenFile Data.Files(1)
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Effect = vbDropEffectCopy
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim L As Single, T As Single, W As Single, H As Single
    L = 0
    T = picToolBar.Height
    W = ScaleWidth - vscData.Width
    H = ScaleHeight - picToolBar.Height - hscData.Height
    vscData.Move L + W, T, vscData.Width, H
    hscData.Move L, T + H, W
End Sub

Private Sub hscData_Change()
    OnDraw
End Sub

Private Sub hscData_Scroll()
    OnDraw
End Sub

Private Sub mnuEditAddAfter_Click()
    Lines.Insert SelectedLine + 1, GetAddValue()
    SelectedLine = SelectedLine + 1
    Focus
    OnDraw
End Sub

Private Sub mnuEditAddBefore_Click()
    Lines.Insert SelectedLine, GetAddValue()
    Focus
    OnDraw
End Sub

Private Sub mnuEditAddFirst_Click()
    Lines.Unshift GetAddValue()
    SelectedLine = 0
    Focus
    OnDraw
End Sub

Private Sub mnuEditAddLast_Click()
    Lines.Push GetAddValue()
    SelectedLine = Lines.GetSize() - 1
    Focus
    OnDraw
End Sub

Private Sub mnuEditDelete_Click()
    Lines.Delete SelectedLine
    OnDraw
End Sub

Private Sub mnuEditEdit_Click()
    If Not Lines.IsExist(SelectedLine) Then Exit Sub
    Lines.Element(SelectedLine) = InputBoxEx("�V�����l�́H", "", Lines.Element(SelectedLine))
    OnDraw
End Sub

Private Sub mnuEditFind_Click()
    Dim Find As String
    Find = InputBoxEx("����������H", "", "")
    If Find = "" Then Exit Sub
    FindString = Find
    Dim I As Long
    For I = SelectedLine + 1 To Lines.GetSize() - 1
        If InStr(Lines.Element(I), FindString) Then
            SelectedLine = I
            Focus
            OnDraw
            Exit Sub
        End If
    Next
    MsgBox FindString & "���������Ɍ�����܂���ł����B"
End Sub

Private Sub mnuEditFindDown_Click()
    If FindString = "" Then
        Dim Find As String
        Find = InputBoxEx("����������H", "", "")
        If Find = "" Then Exit Sub
        FindString = Find
    End If
    Dim I As Long
    For I = SelectedLine + 1 To Lines.GetSize() - 1
        If InStr(Lines.Element(I), FindString) Then
            SelectedLine = I
            Focus
            OnDraw
            Exit Sub
        End If
    Next
    MsgBox FindString & "���������Ɍ�����܂���ł����B"
End Sub

Private Sub mnuEditFindUp_Click()
    If FindString = "" Then
        Dim Find As String
        Find = InputBoxEx("����������H", "", "")
        If Find = "" Then Exit Sub
        FindString = Find
    End If
    Dim I As Long
    For I = SelectedLine - 1 To 0 Step -1
        If InStr(Lines.Element(I), FindString) Then
            SelectedLine = I
            Focus
            OnDraw
            Exit Sub
        End If
    Next
    MsgBox FindString & "��������Ɍ�����܂���ł����B"
End Sub

Private Sub mnuEditMoveDown_Click()
    If SelectedLine >= Lines.GetSize() - 1 Then Exit Sub
    Lines.Swap SelectedLine, SelectedLine + 1
    SelectedLine = SelectedLine + 1
    Focus
    OnDraw
End Sub

Private Sub mnuEditMoveFirst_Click()
    Do Until SelectedLine <= 0
        Lines.Swap SelectedLine, SelectedLine - 1
        SelectedLine = SelectedLine - 1
    Loop
    Focus
    OnDraw
End Sub

Private Sub mnuEditMoveLast_Click()
    Do Until SelectedLine >= Lines.GetSize() - 1
        Lines.Swap SelectedLine, SelectedLine + 1
        SelectedLine = SelectedLine + 1
    Loop
    Focus
    OnDraw
End Sub

Private Sub mnuEditMoveUp_Click()
    If SelectedLine <= 0 Then Exit Sub
    Lines.Swap SelectedLine, SelectedLine - 1
    SelectedLine = SelectedLine - 1
    Focus
    OnDraw
End Sub

Private Sub mnuEditSort_Click()
    Lines.Sort
    OnDraw
End Sub

Private Sub mnuEditSortReverse_Click()
    Lines.ReverseSort
    OnDraw
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub Form_Paint()
    OnDraw
End Sub

Private Sub OnDraw()
    Dim I As Long, Index As Long
    Dim TW As Single, TH As Single
    Dim LineCount As Long, Offset As Long, OffsetX As Long
    Dim bkColor As Long
    Dim L As Single, T As Single, W As Single, H As Single
    L = 0
    T = picToolBar.Height
    W = ScaleWidth - vscData.Width
    H = ScaleHeight - picToolBar.Height - hscData.Height
    TW = TextWidth(" "): TH = TextHeight(" ")
    LineCount = Int(H / TH)
    Offset = Scroll.Value
    OffsetX = hscData.Value
    For I = 0 To LineCount
        Index = I + Offset
        If Lines.IsExist(Index) = False Then
            bkColor = vbApplicationWorkspace
        ElseIf SelectedLine <> Index Then
            bkColor = vbWindowBackground
        Else
            bkColor = vbHighlight
        End If
        Line (L, T + I * TH)-(L + W - 1, T + (I + 1) * TH - 1), bkColor, BF
        If Lines.IsExist(Index) Then
            If SelectedLine <> Index Then
                ForeColor = vbWindowText
            Else
                ForeColor = vbHighlightText
            End If
            CurrentX = L
            CurrentY = T + I * TH
            Print Format(Index, "0000000 : ") & Mid$(Lines.Element(Index), hscData.Value + 1)
        End If
    Next
    If Lines.GetSize Then
        Scroll.Max = Lines.GetSize() - 1
    Else
        Scroll.Max = 0
    End If
    If LineCount > 1 Then vscData.LargeChange = LineCount - 1
End Sub

Private Sub mnuFileNew_Click()
    Lines.InitSize 0
    SelectedLine = -1
    FileName = ""
    OnDraw
End Sub

Private Sub mnuFileOpen_Click()
    Dim ofn As OPENFILENAME
    SetOPENFILENAME ofn, Me.hWnd, "�S�Ẵt�@�C��" & Chr(0) & "*.*", 250, 250, "", _
        "�t�@�C�����J��", OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST, ""
    If GetOpenFileName(ofn) Then
        OpenFile ofn.lpstrFileTitle
    End If
End Sub

Private Sub mnuFileSave_Click()
    If FileName = "" Then mnuFileSaveAs_Click Else SaveFile FileName
End Sub

Private Sub mnuFileSaveAs_Click()
    Dim ofn As OPENFILENAME
    SetOPENFILENAME ofn, Me.hWnd, "�S�Ẵt�@�C��" & Chr(0) & "*.*", 250, 250, "", _
        "�t�@�C����ۑ�", OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT, ""
    If GetSaveFileName(ofn) Then
        SaveFile ofn.lpstrFileTitle
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox App.Title & " version" & App.Major & "." & App.Minor & vbNewLine & "Copyright Mifumi"
End Sub

Private Sub mnuHelpHowTo_Click()
    MsgBox "�厖�ȃt�@�C����ςɂ������Ȃ��悤�ɂ����͋C������I" & vbNewLine & "�ȏ�B"
End Sub

Private Sub mnuSortShuffle_Click()
    Lines.Shuffle
    OnDraw
End Sub

Private Sub vscData_Change()
    Scroll.Update
    OnDraw
End Sub

Private Sub vscData_Scroll()
    vscData_Change
End Sub

Public Function GetAddValue() As String
    Dim av As String
    av = InputBoxEx("�ǉ�����s�̃f�[�^����͂��Ă��������B", App.Title)
    GetAddValue = av
End Function

Public Sub OpenFile(fn As String)
    Dim N As Integer
    Dim FileData As String
    N = FreeFile
    Open fn For Binary Access Read As #N
        FileData = Input(LOF(N), #N)
    Close #N
    FileData = Replace(FileData, vbCr, "")
    Lines.InitSplit FileData, vbLf
    FileName = fn
    OnDraw
End Sub

Public Sub SaveFile(fn As String)
    If fn = "" Then Exit Sub
    Dim N As Integer
    N = FreeFile
    Open fn For Output As #N
        If Lines.GetSize() Then
            Dim I As Long
            I = 0
            Do
                If I < Lines.GetSize() - 1 Then
                    Print #1, Lines.Element(I)
                Else
                    Print #1, Lines.Element(I);
                    Exit Do
                End If
                I = I + 1
            Loop
        End If
    Close #N
End Sub

Public Sub Focus()
    Dim TH As Single
    Dim LineCount As Long
    Dim H As Single
    H = ScaleHeight - picToolBar.Height - hscData.Height
    TH = TextHeight(" ")
    LineCount = Int(H / TH)
    If Scroll.Value > SelectedLine Then
        If SelectedLine >= 0 Then
            Scroll.Value = SelectedLine
        End If
    ElseIf Scroll.Value < SelectedLine - LineCount + 1 Then
        Scroll.Value = SelectedLine - LineCount + 1
    End If
End Sub
