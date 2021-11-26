Attribute VB_Name = "gFunText"

'Modificacion de Bases: CASL 05.12.2000
'---------------------------------------'

Option Explicit
Public Const LF_FACESIZE = 32
'EJVG20120903 ***
Public Enum Alineamiento
    Izquierda = 0
    Derecha = 1
    Centro = 2
End Enum
'END EJVG *******
Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    'lfFaceName(1 To LF_FACESIZE) As Byte
    lfFaceName As String * 32
End Type
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


Public Sub DegreesToXY(CenterX As Long, CenterY As Long, degree As Double, radiusX As Long, radiusY As Long, x As Long, Y As Long)
    Dim convert As Double
    convert = 3.141593 / 180
    x = CenterX - (Sin(-degree * convert) * radiusX)
    Y = CenterY - (Sin((90 + (degree)) * convert) * radiusY)
End Sub

Public Sub RotateText(Degrees As Integer, obj As Object, fontname As String, Fontsize As Single, x As Integer, Y As Integer, Caption As String)
Dim RotateFont As LOGFONT
Dim CurFont As Long, rFont As Long, foo As Long

RotateFont.lfEscapement = Degrees * 10
RotateFont.lfFaceName = fontname & Chr$(0)

If obj.FontBold Then
    RotateFont.lfWeight = 800
Else
    RotateFont.lfWeight = 400
End If
RotateFont.lfHeight = (Fontsize * -20) / Screen.TwipsPerPixelY
rFont = CreateFontIndirect(RotateFont)
CurFont = SelectObject(obj.hdc, rFont)

obj.CurrentX = x
obj.CurrentY = Y
obj.Print Caption

'Restore
foo = SelectObject(obj.hdc, CurFont)
foo = DeleteObject(rFont)

End Sub
Public Sub TextCircle(obj As Object, txt As String, x As Long, Y As Long, radius As Long, startdegree As Double)
Dim foo As Integer, TxtX As Long, TxtY As Long, checkit As Integer
Dim twipsperdegree As Long, wrktxt As String, wrklet As String, degreexy As Double, degree As Double
twipsperdegree = (radius * 3.14159 * 2) / 360
If startdegree < 0 Then
    Select Case startdegree
    Case -1
        startdegree = Int(360 - (((obj.TextWidth(txt)) / twipsperdegree) / 2))
    Case -2
        radius = (obj.TextWidth(txt) / 2) / 3.14159
        twipsperdegree = (radius * 3.14159 * 2) / 360
    End Select
End If
For foo = 1 To Len(txt)
    wrklet = Mid$(txt, foo, 1)
    degreexy = (obj.TextWidth(wrktxt)) / twipsperdegree + startdegree
    DegreesToXY x, Y, degreexy, radius, radius, TxtX, TxtY
    degree = (obj.TextWidth(wrktxt) + 0.5 * obj.TextWidth(wrklet)) / twipsperdegree + startdegree
    RotateText 360 - degree, obj, obj.fontname, obj.Fontsize, (TxtX), (TxtY), wrklet
    wrktxt = wrktxt & wrklet
Next foo
End Sub


Public Function EliminaEnters(lsTexto As String) As String
    EliminaEnters = Replace(lsTexto, Chr(13), "")
End Function
'EJVG20120903 ***
Public Function AlinearTexto(ByVal psTexto As String, ByVal pnTamano As Integer, Optional ByVal pnAlineamiento As Alineamiento = Izquierda, Optional ByVal psCaracterDeRelleno As String = " ") As String
    Dim lsTemp As String
    Dim lnPosIni As Long
    Select Case pnAlineamiento
    Case Alineamiento.Izquierda
        AlinearTexto = Left(psTexto & String(pnTamano, psCaracterDeRelleno), pnTamano)
    Case Alineamiento.Derecha
        AlinearTexto = Right(String(pnTamano, psCaracterDeRelleno) & psTexto, pnTamano)
    Case Alineamiento.Centro
        psTexto = Trim(psTexto)
        If Len(psTexto) > pnTamano Then
            psTexto = Left(psTexto, pnTamano)
        Else
            lnPosIni = Int(pnTamano / 2) - Int(Len(psTexto) / 2)
            psTexto = String(lnPosIni, psCaracterDeRelleno) & psTexto & String(lnPosIni, psCaracterDeRelleno)
            AlinearTexto = psTexto
        End If
    End Select
End Function
'END EJVG *******
