Attribute VB_Name = "DragNSnap_Mod"
'DragNSnap By Sean Siegel (SeanMSiegel@hotmail.com)
'The Drag and snap sub allows you drag forms in all versions of
'windows, unlike the old form drag api.
'The sub also has optional screen edge snapping.
'Enjoy and please vote :)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Dim CurXY As POINTAPI
Dim Cx As Long
Dim Cy As Long
Dim Tpx As Long
Dim Tpy As Long
Public Function Get_Mouse_X() As Long
    resp = GetCursorPos(CurXY) 'load mouse xy into user type curxy
    Get_Mouse_X = CurXY.X 'return the x
End Function
Public Function Get_Mouse_Y() As Long
    resp = GetCursorPos(CurXY) 'load mouse xy into user type curxy
    Get_Mouse_Y = CurXY.Y 'return the y
End Function
Sub DragNSnap(TheForm As Form, TheButton As Integer, X As Single, Y As Single, Optional ScreenSnapping As Boolean = True, Optional SnapPixels As Byte = 10)
    If Tpx = 0 Then Tpx = Screen.TwipsPerPixelX 'i use if so we only load the variable once to save process
    If Tpy = 0 Then Tpy = Screen.TwipsPerPixelY
    If TheButton <> 1 Then 'make sure they are left clicking to drag
        Cx = X 'save the mouse x so we can calculte the left later on if a button is pressed
        Cy = Y 'save the mouse y so we can calculte the top later on if a button is pressed
    Else
        tx = Get_Mouse_X * Tpx - Cx 'get the mousex in pixels and subtract its x to set the left of the form
        If tx / Tpx < -SnapPixels Then GoTo cnty
        If ScreenSnapping And tx / Tpx < SnapPixels Then tx = 0
        If (tx + TheForm.Width) / Tpx > (Screen.Width / Tpx) + SnapPixels Then GoTo cnty
        If ScreenSnapping And (tx + TheForm.Width) / Tpx > (Screen.Width / Tpx) - SnapPixels Then tx = Screen.Width - TheForm.Width
cnty:
        ty = Get_Mouse_Y * Tpy - Cy 'get the mousey in pixels and subtract its y to set the top of the form
        If ty / Tpy < -SnapPixels Then GoTo cnt:
        If ScreenSnapping And ty / Tpx < SnapPixels Then ty = 0
        If (ty + TheForm.Height) / Tpy > (Screen.Height / Tpy) + SnapPixels Then GoTo cnt:
        If ScreenSnapping And (ty + TheForm.Height) / Tpx > (Screen.Height / Tpx) - SnapPixels Then ty = Screen.Height - TheForm.Height
cnt:
        TheForm.Move tx, ty 'move the form to its new location
    End If
End Sub
