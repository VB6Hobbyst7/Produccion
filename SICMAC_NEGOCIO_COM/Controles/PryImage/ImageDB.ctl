VERSION 5.00
Begin VB.UserControl ImageDB 
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   2475
   ScaleWidth      =   5775
   ToolboxBitmap   =   "ImageDB.ctx":0000
   Begin VB.Frame Frame1 
      Height          =   2475
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   5685
      Begin VB.PictureBox ImgFirma 
         BackColor       =   &H8000000E&
         Height          =   2025
         Left            =   30
         ScaleHeight     =   1965
         ScaleWidth      =   5310
         TabIndex        =   3
         Top             =   135
         Width           =   5370
         Begin VB.PictureBox PicTemp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1425
            Left            =   60
            ScaleHeight     =   1425
            ScaleWidth      =   5160
            TabIndex        =   4
            Top             =   75
            Width           =   5160
         End
      End
      Begin VB.VScrollBar VSVert 
         Height          =   2295
         Left            =   5400
         TabIndex        =   2
         Top             =   135
         Width           =   240
      End
      Begin VB.HScrollBar HSHoriz 
         Height          =   255
         Left            =   30
         TabIndex        =   1
         Top             =   2175
         Width           =   5400
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6195
      Top             =   1590
      Width           =   540
   End
End
Attribute VB_Name = "ImageDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private sRutaImg As String
Private nHeight As Long
Private nWidth As Long
Private nZoom  As Long
Const ZoomFactor = 0.05

Private R As ADODB.Recordset
Private RStream As ADODB.Stream
Private bEnabled As Boolean
Public Conexion As ADODB.Connection
Dim nHeighOrig As Long
Dim nWidthOrig As Long
Public Sub GrabarFirma(ByRef Rf As ADODB.Recordset, ByVal psPersCod As String, ByVal psUltimaAct As String)
    If Not GrabarFirmaenBD(Rf, psPersCod, psUltimaAct) Then
        MsgBox "Ha ocurrido un Error al Grabar la Firma si el error persite consulte con al Area de Sistemas", vbExclamation, "Aviso"
        Exit Sub
    End If
End Sub
Public Sub CargarFirma(ByVal pR As ADODB.Recordset)
    If Not CargarFirmadeBD(pR) Then
        UserControl.Enabled = False
        MsgBox "La persona no posee firma", vbInformation, "Aviso"
        gbTieneFirma = False 'ande 20170914
        Exit Sub
    Else
        gbTieneFirma = True 'ande 20170914
    End If
    
End Sub
Private Function GrabarFirmaenBD(ByRef Rf As ADODB.Recordset, ByVal psPersCod As String, ByVal psUltimaAct As String) As Boolean
Dim ssql As String

    On Error GoTo ErrorGrabarFirmaenBD
    If Rf.BOF And Rf.EOF Then
        Rf.AddNew
        Rf.Fields("cPersCod").value = psPersCod
        Rf.Fields("cUltimaActualizacion").value = psUltimaAct
    End If
    
    Set RStream = New ADODB.Stream
    RStream.Type = adTypeBinary
    RStream.Open
    RStream.LoadFromFile sRutaImg
    Rf.Fields("iPersFirma").value = RStream.Read
    RStream.Close
    Set RStream = Nothing
    GrabarFirmaenBD = True
    Exit Function
    
ErrorGrabarFirmaenBD:
    GrabarFirmaenBD = False
    
End Function
Private Function CargarFirmadeBD(ByVal pR As ADODB.Recordset) As Boolean
Dim sRutaTmp As String, sRutaServer As String
    On Error GoTo ErrorCargarFirmadeBD
        
    If pR Is Nothing Then
        GoTo ErrorCargarFirmadeBD
    End If
    
    UserControl.Enabled = True
    Set RStream = New ADODB.Stream
    RStream.Type = adTypeBinary
    RStream.Open
    
    RStream.Write pR.Fields("iPersFirma").value
    
    sRutaServer = "\\srvxenc\Firmas"
    sRutaTmp = sRutaServer & "\Temp_" & pR.Fields("cPersCod").value & ".bmp"
    RStream.SaveToFile sRutaTmp, adSaveCreateOverWrite
    sRutaImg = sRutaTmp
    If Len(Trim(sRutaImg)) > 0 Then
        PicTemp.Height = 1425
        PicTemp.Width = 5160
        PicTemp.Picture = LoadPicture(sRutaImg)
        Image1.Picture = LoadPicture(sRutaImg)
        PicTemp.Height = Image1.Height
        PicTemp.Width = Image1.Width
        nHeighOrig = PicTemp.Height
        nWidthOrig = PicTemp.Width
        nZoom = -1
        ZoomMas
    Else
        ImgFirma.Picture = Nothing
    End If
    Kill (sRutaTmp)
    RStream.Close
    Set RStream = Nothing
    CargarFirmadeBD = True
    Exit Function
    
ErrorCargarFirmadeBD:
    CargarFirmadeBD = False
    PicTemp.Picture = LoadPicture("")
    Image1.Picture = LoadPicture("")
End Function


Private Sub ActualizarImagen(picBox As PictureBox, ByVal sizePic As StdPicture, ByVal sizeWidth As Single, ByVal sizeHeight As Single)
  'WIOR 20120816 Cambio el tipo del segundo parametro: de ByVal sizePic As Picture a ByVal sizePic As StdPicture
  picBox.Picture = LoadPicture("")
  picBox.Width = sizeWidth
  picBox.Height = sizeHeight
  picBox.AutoRedraw = True
  picBox.PaintPicture sizePic, 0, 0, sizeWidth, sizeHeight
  picBox.Picture = picBox.Image
  picBox.AutoRedraw = False
End Sub
Private Sub ScrollBar(ByVal picWidth As Integer, ByVal picHeight As Integer, ByVal areaWidth As Integer, ByVal areaHeight As Integer)
    
  If picWidth - areaWidth > 0 Then
     HSHoriz.Max = picWidth - areaWidth
     'HSHoriz.Visible = True
     'HSHoriz.Enabled = True
  Else
     HSHoriz.Max = 0
     'HSHoriz.Visible = False
     'HSHoriz.Enabled = False
  End If
  
  If picHeight - areaHeight > 0 Then
     VSVert.Max = picHeight - areaHeight
     'VSVert.Visible = True
     'VSVert.Enabled = True
  Else
     VSVert.Max = 0
     'VSVert.Visible = False
     'VSVert.Enabled = False
  End If

End Sub
Private Sub TamañoOriginal()
        PicTemp.Height = nHeighOrig
        PicTemp.Width = nWidthOrig
        nZoom = -1
        ZoomMas
End Sub
Private Sub ZoomMas()
  
  If nZoom < 38 Then
      nHeight = PicTemp.Height + (PicTemp.Height * ZoomFactor)
      nWidth = PicTemp.Width + (PicTemp.Width * ZoomFactor)
      Image1.Stretch = True
      Image1.Width = nWidth
      Image1.Height = nHeight
      Image1.Stretch = False
    
      nZoom = nZoom + 1
    'WIOR 20120816 ********************************************
    Dim stdPic As StdPicture
    Set stdPic = Image1.Picture
    'ActualizarImagen PicTemp, Image1.Picture, nWidth, nHeight
    ActualizarImagen PicTemp, stdPic, nWidth, nHeight
    'WIOR FIN *************************************************
    
      ScrollBar PicTemp.Width, PicTemp.Height, ImgFirma.Width, ImgFirma.Height
   End If
End Sub
Private Sub ZoomMenos()

If nZoom > -27 Then
  nHeight = PicTemp.Height - (PicTemp.Height * ZoomFactor)
  nWidth = PicTemp.Width - (PicTemp.Width * ZoomFactor)
  
  Image1.Stretch = True
  Image1.Width = nWidth
  Image1.Height = nHeight
  Image1.Stretch = False
  
  nZoom = nZoom - 1
    'WIOR 20120816 ********************************************
    Dim stdPic As StdPicture
    Set stdPic = Image1.Picture
    'ActualizarImagen PicTemp, Image1.Picture, nWidth, nHeight
    ActualizarImagen PicTemp, stdPic, nWidth, nHeight
    'WIOR FIN *************************************************

  ScrollBar PicTemp.Width, PicTemp.Height, ImgFirma.Width, ImgFirma.Height
End If
End Sub

Private Sub HSHoriz_Change()
    PicTemp.Left = -HSHoriz.value
End Sub

Private Sub HSHoriz_Scroll()
    PicTemp.Left = -HSHoriz.value
End Sub


Private Sub UserControl_Initialize()
    nZoom = 0
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 43 '+
            ZoomMas
        Case 45 '-
            ZoomMenos
        Case 42
            TamañoOriginal
    End Select
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 39 'Derecha
            If HSHoriz.value + 50 < HSHoriz.Max Then
                HSHoriz.value = HSHoriz.value + 50
            Else
                HSHoriz.value = HSHoriz.Max
            End If
        Case 37 'Izquierda
            If HSHoriz.value - 50 > HSHoriz.Min Then
                HSHoriz.value = HSHoriz.value - 50
            Else
                HSHoriz.value = HSHoriz.Min
            End If
        Case 38 'Arriba
            If VSVert.value - 50 > VSVert.Min Then
                VSVert.value = VSVert.value - 50
            Else
                VSVert.value = VSVert.Min
            End If
        Case 40 'Abajo
            If VSVert.value + 50 < VSVert.Max Then
                VSVert.value = VSVert.value + 50
            Else
                VSVert.value = VSVert.Max
            End If
    End Select
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width < 3405 Then
        UserControl.Width = 3405
    End If
    If UserControl.Height < 1680 Then
        UserControl.Height = 1680
    End If
    Frame1.Width = UserControl.Width - 5
    Frame1.Height = UserControl.Height - 5
    Frame1.top = 1
    Frame1.Left = 1
    ImgFirma.Width = Int(Frame1.Width * 0.95)
    ImgFirma.Height = Int(Frame1.Height * 0.83)
    ImgFirma.top = Frame1.top + 140
    ImgFirma.Left = Frame1.Left + 50
    VSVert.Width = Int(0.03 * Frame1.Width)
    VSVert.Height = ImgFirma.Height - 5
    VSVert.Left = ImgFirma.Left + ImgFirma.Width + 1
    VSVert.top = ImgFirma.top
    HSHoriz.Height = Int(0.08 * Frame1.Height)
    HSHoriz.Width = ImgFirma.Width - 1
    HSHoriz.top = ImgFirma.top + ImgFirma.Height + 2
    HSHoriz.Left = ImgFirma.Left
End Sub

Public Property Get RutaImagen() As String
    RutaImagen = sRutaImg
End Property

Public Property Let RutaImagen(ByVal psRuta As String)
On Error GoTo ErrorRutaImagen
    sRutaImg = psRuta
    If Len(Trim(psRuta)) > 0 Then
        PicTemp.Height = 1425
        PicTemp.Width = 5160
        PicTemp.Picture = LoadPicture(psRuta)
        Image1.Picture = LoadPicture(psRuta)
        nZoom = -1
        ZoomMas
    Else
        ImgFirma.Picture = Nothing
        PicTemp.Picture = LoadPicture("")
    End If
    Exit Property
ErrorRutaImagen:
    ImgFirma.Picture = Nothing
    PicTemp.Picture = Nothing
    MsgBox "No se pudo cargar la firma, consulte con el Area de Sistemas", vbInformation, "Aviso"
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    RutaImagen = PropBag.ReadProperty("RutaImagen", "")
    Enabled = PropBag.ReadProperty("Enabled", True)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "RutaImagen", RutaImagen, ""
    PropBag.WriteProperty "Enabled", Enabled, ""
End Sub

Private Sub VSVert_Change()
    PicTemp.top = -VSVert.value
End Sub

Private Sub VSVert_Scroll()
    PicTemp.top = -VSVert.value
End Sub

Public Property Get Enabled() As Boolean
    Enabled = bEnabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    bEnabled = vNewValue
    UserControl.Enabled = bEnabled
End Property
