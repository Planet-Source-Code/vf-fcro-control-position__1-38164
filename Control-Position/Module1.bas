Attribute VB_Name = "Module1"
Public ControlColl As New Collection
Public MinHeight As Long
Public MaxHeight As Long
Public MinWidth As Long
Public MaxWidth As Long

Type POINTAPI
    x As Long
    y As Long
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long

End Type

Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Sub ReadFormControlCollenction(FRM As Form)
Dim RT(3) As String
Dim O As Object
For Each O In FRM.Controls
RT(0) = (CDbl(O.Top) / CDbl(FRM.ScaleHeight)) * CDbl(100)
RT(1) = (CDbl(O.Left) / CDbl(FRM.ScaleWidth)) * CDbl(100)
RT(2) = (CDbl(O.Height) / CDbl(FRM.ScaleHeight)) * CDbl(100)
RT(3) = (CDbl(O.Width) / CDbl(FRM.ScaleWidth)) * CDbl(100)
ControlColl.Add Join(RT, "\"), O.Name
Next
End Sub

Public Sub FormResize(FRM As Form)
Dim PA As POINTAPI
If MinWidth <> 0 Then
If FRM.Width < MinWidth Then FRM.Width = MinWidth: RelCapX1 FRM
End If
If MinHeight <> 0 Then
If FRM.Height < MinHeight Then FRM.Height = MinHeight: RelCapY1 FRM
End If
If MaxWidth <> 0 Then
If FRM.Width > MaxWidth Then FRM.Width = MaxWidth: RelCapX2 FRM
End If
If MaxHeight <> 0 Then
If FRM.Height > MaxHeight Then FRM.Height = MaxHeight: RelCapY2 FRM
End If
End Sub

Public Sub RelCapX1(FRM As Form)
Dim RCT As RECT
GetWindowRect FRM.hwnd, RCT
Dim PA As POINTAPI
GetCursorPos PA
SetCursorPos RCT.Left + (MinWidth / 15), PA.y
SetCapture FRM.hwnd
End Sub

Public Sub RelCapX2(FRM As Form)
Dim RCT As RECT
GetWindowRect FRM.hwnd, RCT
Dim PA As POINTAPI
GetCursorPos PA
SetCursorPos RCT.Left + (MaxWidth / 15), PA.y
SetCapture FRM.hwnd
End Sub
Public Sub RelCapY1(FRM As Form)
Dim RCT As RECT
GetWindowRect FRM.hwnd, RCT
Dim PA As POINTAPI
GetCursorPos PA
SetCursorPos PA.x, RCT.Top + (MinHeight / 15)
SetCapture FRM.hwnd
End Sub
Public Sub RelCapY2(FRM As Form)
Dim RCT As RECT
GetWindowRect FRM.hwnd, RCT
Dim PA As POINTAPI
GetCursorPos PA
SetCursorPos PA.x, RCT.Top + (MaxHeight / 15)
SetCapture FRM.hwnd
End Sub
Public Sub ResizeCtrl(ByVal ControlName As String, ByVal FRM As Form, Optional LeftPositionChange As Boolean = True, Optional TopPositionChange As Boolean = True, Optional WidthResize As Boolean = True, Optional HeightResize As Boolean = True, Optional ByVal ResizeWidthTillReachControl As String, Optional ByVal ResizeHeightTillReachControl As String)
Dim DBS() As Double
DBS = GetSizes(ControlName)

Dim CTSW() As Double
Dim CTSH() As Double

If ResizeWidthTillReachControl <> "" Then
CTSW = GetSizes(ResizeWidthTillReachControl)
Else
ReDim CTSW(1)
CTSW(1) = DBS(1) + DBS(3)
End If

If ResizeHeightTillReachControl <> "" Then
CTSH = GetSizes(ResizeHeightTillReachControl)
Else
ReDim CTSH(0)
CTSH(0) = DBS(0) + DBS(2)
End If


If TopPositionChange Then
FRM.Controls(ControlName).Top = (DBS(0) * CDbl(FRM.ScaleHeight)) / CDbl(100)
End If



If LeftPositionChange Then
FRM.Controls(ControlName).Left = (DBS(1) * CDbl(FRM.ScaleWidth)) / CDbl(100)
End If


If DBS(0) + DBS(2) <= CTSH(0) Then
If HeightResize Then
FRM.Controls(ControlName).Height = (DBS(2) * CDbl(FRM.ScaleHeight)) / CDbl(100)
End If
End If


If DBS(1) + DBS(3) <= CTSW(1) Then
If WidthResize Then
FRM.Controls(ControlName).Width = (DBS(3) * CDbl(FRM.ScaleWidth)) / CDbl(100)
End If
End If

End Sub

Public Function GetSizes(ByVal ControlName As String) As Double()
Dim RT As String
RT = ControlColl.Item(ControlName)
Dim DB() As String
DB = Split(RT, "\")
Dim DBS(3) As Double
DBS(0) = CDbl(DB(0))
DBS(1) = CDbl(DB(1))
DBS(2) = CDbl(DB(2))
DBS(3) = CDbl(DB(3))
GetSizes = DBS
End Function
