Attribute VB_Name = "modResponsive"
Option Compare Database
Option Explicit

' ============================================================
'  Responsives Layout fuer Access-Formulare
'
'  Buttons:  Tag = "flex"  bzw. "flex:1", "flex:2", ... (Reihenfolge)
'            -> werden nebeneinander angeordnet und brechen in die
'               naechste Zeile um, wenn die Breite nicht reicht.
'  Felder:   Tag = "stretch"
'            -> werden bis zum rechten Rand verbreitert.
'
'  Aufruf im Formular:
'     Private Sub Form_Load():   ResponsiveLayout Me: End Sub
'     Private Sub Form_Resize(): ResponsiveLayout Me: End Sub
' ============================================================

Public Const FLEX_GAP As Long = 113        ' Abstand zwischen Buttons (~2 mm, Twips)
Public Const FLEX_MARGIN As Long = 170     ' Rand (~3 mm, Twips)
Public Const STRETCH_MIN As Long = 1134    ' Mindestbreite Felder (~2 cm)

Public Enum FlexAlign
    flexLeft = 0
    flexCenter = 1
    flexRight = 2
End Enum

Private mBusy As Boolean                   ' verhindert Rekursion durch Resize

' Gesamtes Layout: Felder strecken + Buttons anordnen
Public Sub ResponsiveLayout(frm As Form, _
                            Optional ByVal buttonSection As Integer = acFooter, _
                            Optional ByVal align As FlexAlign = flexLeft)
    If mBusy Then Exit Sub
    mBusy = True
    On Error GoTo Ende

    frm.Painting = False

    ' Formular erst breit genug machen, damit nichts "aneckt"
    If frm.Width < frm.InsideWidth Then frm.Width = frm.InsideWidth

    StretchControls frm
    FlexLayout frm, buttonSection, align

    ' Formularbreite wieder auf Fensterbreite zuruecksetzen (keine Scrollbar)
    On Error Resume Next
    If frm.Width > frm.InsideWidth Then frm.Width = frm.InsideWidth
    On Error GoTo Ende

Ende:
    If Err.Number <> 0 Then Debug.Print "ResponsiveLayout: " & Err.Number & " " & Err.Description
    frm.Painting = True
    mBusy = False
End Sub

' Felder mit Tag "stretch" bis zum rechten Rand verbreitern
Public Sub StretchControls(frm As Form)
    Dim ctl As Control, w As Long
    For Each ctl In frm.Controls
        If LCase$(Left$(ctl.Tag, 7)) = "stretch" Then
            w = frm.InsideWidth - ctl.Left - FLEX_MARGIN
            If w < STRETCH_MIN Then w = STRETCH_MIN
            ctl.Width = w
        End If
    Next
End Sub

' Buttons mit Tag "flex" zeilenweise anordnen (wie CSS flex-wrap)
Public Sub FlexLayout(frm As Form, _
                      Optional ByVal sec As Integer = acFooter, _
                      Optional ByVal align As FlexAlign = flexLeft)
    Dim btns() As Control, n As Long
    n = CollectFlex(frm, sec, btns)
    If n = 0 Then Exit Sub

    Dim avail As Long
    avail = frm.InsideWidth - 2 * FLEX_MARGIN
    If avail < 1 Then Exit Sub

    ' --- 1. Positionen berechnen ---
    Dim newL() As Long, newT() As Long
    ReDim newL(n - 1): ReDim newT(n - 1)

    Dim i As Long, j As Long, rowStart As Long
    Dim rowW As Long, rowH As Long, y As Long, x As Long

    y = FLEX_MARGIN: rowStart = 0: rowW = 0: rowH = 0
    For i = 0 To n
        ' Zeile abschliessen: am Ende oder wenn naechster Button nicht passt
        If i = n Then
            GoSub ZeileSetzen
        ElseIf rowW > 0 And rowW + FLEX_GAP + btns(i).Width > avail Then
            GoSub ZeileSetzen
            y = y + rowH + FLEX_GAP
            rowStart = i: rowW = 0: rowH = 0
        End If
        If i < n Then
            If rowW > 0 Then rowW = rowW + FLEX_GAP
            rowW = rowW + btns(i).Width
            If btns(i).Height > rowH Then rowH = btns(i).Height
        End If
    Next
    Dim neededH As Long
    neededH = y + rowH + FLEX_MARGIN

    ' --- 2. Bereich gross genug machen, Buttons setzen, Hoehe anpassen ---
    If frm.Section(sec).Height < neededH Then frm.Section(sec).Height = neededH
    For i = 0 To n - 1
        btns(i).Left = newL(i)
        btns(i).Top = newT(i)
    Next
    On Error Resume Next   ' andere Steuerelemente im Bereich koennen Verkleinern verhindern
    frm.Section(sec).Height = neededH
    On Error GoTo 0
    Exit Sub

ZeileSetzen:
    Select Case align
        Case flexCenter: x = FLEX_MARGIN + (avail - rowW) \ 2
        Case flexRight:  x = FLEX_MARGIN + (avail - rowW)
        Case Else:       x = FLEX_MARGIN
    End Select
    If x < FLEX_MARGIN Then x = FLEX_MARGIN
    For j = rowStart To i - 1
        newL(j) = x
        newT(j) = y
        x = x + btns(j).Width + FLEX_GAP
    Next
    Return
End Sub

' Buttons eines Bereichs sammeln und nach "flex:<Nr>" sortieren
Private Function CollectFlex(frm As Form, ByVal sec As Integer, btns() As Control) As Long
    Dim ctl As Control, n As Long, i As Long, j As Long
    Dim ord() As Long, tmpC As Control, tmpO As Long

    For Each ctl In frm.Section(sec).Controls
        If LCase$(Left$(ctl.Tag, 4)) = "flex" Then
            ReDim Preserve btns(n): ReDim Preserve ord(n)
            Set btns(n) = ctl
            ord(n) = Val(Mid$(ctl.Tag, 6))
            n = n + 1
        End If
    Next

    ' einfache Sortierung nach Reihenfolge-Nummer
    For i = 1 To n - 1
        For j = i To 1 Step -1
            If ord(j) >= ord(j - 1) Then Exit For
            tmpO = ord(j): ord(j) = ord(j - 1): ord(j - 1) = tmpO
            Set tmpC = btns(j): Set btns(j) = btns(j - 1): Set btns(j - 1) = tmpC
        Next
    Next
    CollectFlex = n
End Function
