VERSION 5.00
Begin VB.Form frmLogSelCotiza 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6345
   ClientLeft      =   1215
   ClientTop       =   1785
   ClientWidth     =   9645
   Icon            =   "frmLogSelCotiza.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCot 
      Caption         =   "&Generar"
      Enabled         =   0   'False
      Height          =   390
      Left            =   7710
      TabIndex        =   10
      Top             =   4800
      Width           =   1290
   End
   Begin Sicmact.FlexEdit fgePro 
      Height          =   1665
      Left            =   225
      TabIndex        =   7
      Top             =   4485
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   2937
      Cols0           =   4
      HighLight       =   2
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-Codigo-Nombre-Dirección"
      EncabezadosAnchos=   "400-0-2500-3500"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X"
      ListaControles  =   "0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "L-L-L-L"
      FormatosEdit    =   "0-0-0-0"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   0
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   285
   End
   Begin Sicmact.FlexEdit fgeSel 
      Height          =   855
      Left            =   210
      TabIndex        =   6
      Top             =   3270
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   1508
      ScrollBars      =   0
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483624
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   0
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   -1
      RowHeight0      =   240
   End
   Begin Sicmact.FlexEdit fgeCot 
      Height          =   2355
      Left            =   210
      TabIndex        =   5
      Top             =   915
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   4154
      Cols0           =   6
      HighLight       =   2
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-Codigo-Bien/Servicio-Unidad-Cantidad-Precio"
      EncabezadosAnchos=   "400-0-3000-0-0-0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "L-L-L-L-R-R"
      FormatosEdit    =   "0-0-0-0-0-0"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   0
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   285
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   7710
      TabIndex        =   0
      Top             =   5505
      Width           =   1305
   End
   Begin Sicmact.TxtBuscar txtSelNro 
      Height          =   285
      Left            =   1125
      TabIndex        =   1
      Top             =   330
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TipoBusqueda    =   2
      sTitulo         =   ""
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   15
      Top             =   5865
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label lblAdqNro 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6690
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   2760
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Bienes/Servicios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Index           =   2
      Left            =   375
      TabIndex        =   9
      Top             =   705
      Width           =   1560
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Proveedores "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Index           =   1
      Left            =   390
      TabIndex        =   8
      Top             =   4275
      Width           =   1245
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Número :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Index           =   5
      Left            =   270
      TabIndex        =   4
      Top             =   375
      Width           =   870
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Area :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Index           =   0
      Left            =   300
      TabIndex        =   3
      Top             =   90
      Width           =   555
   End
   Begin VB.Label lblAreaDes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1140
      TabIndex        =   2
      Top             =   45
      Width           =   3705
   End
End
Attribute VB_Name = "frmLogSelCotiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim psFrmTpo As String
Dim pnMinCot As Currency

Public Sub Inicio(ByVal psFormTpo As String)
psFrmTpo = psFormTpo
Me.Show 1
End Sub

Private Sub cmdCot_Click()
    Dim clsDMov As DLogMov
    Dim clsDGnral As DLogGeneral
    Dim sSelNro As String, sSelTraNro As String, sAdqNro As String, sActualiza As String
    Dim nCont As Integer, nSum As Integer, nCot As Integer, nSel As Integer, nResult As Integer
    Dim sBSCod As String, sSelCotNro As String, sPersCod As String
    
    'Verifica que siempre este por lo menos UNO
    For nSel = 6 To fgeSel.Cols - 1
        If fgeSel.TextMatrix(1, nSel) <> "" Then nSum = nSum + 1
    Next
    If nSum = 0 Then
        MsgBox "Falta seleccionar a los Provedores a enviar cotizaciones", vbInformation, " Aviso"
        Exit Sub
    End If
    If nSum < pnMinCot Then
        MsgBox "Mínimo de cotizaciones a generar debe ser por lo menos " & pnMinCot, vbInformation, " Aviso"
        Exit Sub
    End If
    
    sSelNro = txtSelNro.Text
    sAdqNro = Trim(lblAdqNro.Caption)
    If sSelNro = "" Or sAdqNro = "" Then Exit Sub
    
    If MsgBox("¿ Estás seguro de generar las cotizaciones del proceso de selección " & sSelNro & " ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
        Set clsDGnral = New DLogGeneral
        sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
        Set clsDGnral = Nothing
        sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
        Set clsDMov = New DLogMov
        
        'Grabación de MOV -MOVREF
        clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", Trim(Str(gLogSelEstadoCotizacion))
        clsDMov.InsertaMovRef sSelTraNro, sSelNro
        
        'Actualiza LogSeleccion
        clsDMov.ActualizaSeleccion sSelNro, gdFecSis, "", _
            "", "", sActualiza, gLogSelEstadoCotizacion
        
        'Modifica  LogAdquisición
        clsDMov.ActualizaAdquisicion sAdqNro, gLogAdqEstadoCotiza, sActualiza
        
        nSum = 0: nCont = 0
        For nSel = 6 To fgeSel.Cols - 1
            nSum = nSum + 1
            If fgeSel.TextMatrix(1, nSel) <> "" Then
                nCont = nCont + 1
                sSelCotNro = GeneraCotiza(sSelNro, nCont)
                sPersCod = fgePro.TextMatrix(nSum, 1)
                'Inserta LogSelCotiza
                clsDMov.InsertaSelCotiza sSelNro, sSelCotNro, sPersCod, sActualiza
                For nCot = 1 To fgeCot.Rows - 1
                    sBSCod = fgeCot.TextMatrix(nCot, 1)
                    'Inserta LogSelCotDetalle
                    clsDMov.InsertaSelCotDetalle sSelCotNro, sBSCod, 0, 0, sActualiza
                Next
            End If
        Next
        'Ejecuta todos los querys en una transacción
        'nResult = clsDMov.EjecutaBatch
        Set clsDMov = Nothing
        
        If nResult = 0 Then
            fgeSel.Enabled = False
            cmdCot.Enabled = False
            Call CargaTxtSelNro
        Else
            MsgBox "Error al grabar la información", vbInformation, " Aviso "
        End If
    End If

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub fgeSel_Click()
Dim nCol As Integer
Dim nFil As Integer
Dim nPos As Integer
Dim i As Integer
Dim pColCot As Integer

Dim phCab As ColorConstants

pColCot = 5
phCab = vbBlue         '&H00008000&
nPos = fgeSel.Col - pColCot
nCol = fgeSel.Col
nFil = 1

If nCol >= 3 Then
    fgeSel.Col = nCol
    fgeSel.CellForeColor = vbBlue
   
    fgePro.Row = nPos
    fgeCot.Col = nCol
    If Len(fgeSel.TextMatrix(nFil, nCol)) = 0 Then
       fgeSel.TextMatrix(nFil, nCol) = "X"
       fgePro.ForeColorRow (&HFF0000)
       For i = 1 To fgeCot.Rows - 1
           fgeCot.Row = i
           fgeCot.CellForeColor = &HFF0000
       Next
    Else
       fgeSel.TextMatrix(nFil, nCol) = ""
       fgePro.ForeColorRow (&H5)
       For i = 1 To fgeCot.Rows - 1
           fgeCot.Row = i
           fgeCot.CellForeColor = "&H00000005"
       Next
    End If
End If
End Sub

Private Sub Form_Load()
    Dim clsDGnral As DLogGeneral
    Call CentraForm(Me)
    'Carga información de la relación usuario-area
    Usuario.Inicio gsCodUser
    If Len(Usuario.AreaCod) = 0 Then
        MsgBox "Usuario no determinado", vbInformation, "Aviso"
        Exit Sub
    End If
    lblAreaDes.Caption = Usuario.AreaNom
    
    Me.Caption = "Solicitudes de Cotización"
    'OJO. En cargado de valor debe utilizarse las variables
    'Valor de máxima suma de parámetros de
    Set clsDGnral = New DLogGeneral
    pnMinCot = clsDGnral.CargaParametro(5000, 1002)
    Set clsDGnral = Nothing
    
    
    Call CargaTxtSelNro
End Sub

Private Sub CargaTxtSelNro()
    Dim clsDAdq As DLogAdquisi
    Dim rs As ADODB.Recordset
    Set clsDAdq = New DLogAdquisi
    Set rs = New ADODB.Recordset
    
    Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, , gLogSelEstadoPublicacion)
    If rs.RecordCount > 0 Then
        txtSelNro.rs = rs
    Else
        txtSelNro.Enabled = False
    End If
    Set rs = Nothing
    Set clsDAdq = Nothing
End Sub

Private Sub txtSelNro_EmiteDatos()
    Dim clsDAdq As DLogAdquisi
    Dim clddreq As DLogRequeri
    Dim clsDReq As DLogRequeri
    
    Dim rs As ADODB.Recordset
    Dim sSelNro As String, sAdqNro As String
    
    
    If txtSelNro.Ok = False Then Exit Sub
    
    fgeSel.Enabled = True
    cmdCot.Enabled = True
    Set clsDAdq = New DLogAdquisi
    Set clsDReq = New DLogRequeri
    Set rs = New ADODB.Recordset
    Call Limpiar
    sSelNro = txtSelNro.Text
    
    Set rs = clsDAdq.CargaSeleccion(SelUnRegistro, sSelNro)
    If rs.RecordCount > 0 Then
        sAdqNro = rs!cLogAdqNro
        lblAdqNro.Caption = sAdqNro
        'Muestra detalle de Bienes/Servicios
        'Set rs = clsDAdq.CargaAdqDetalle(AdqDetUnRegCoti, sAdqNro)
        Set rs = clsDReq.CargaReqDetalle(ReqDetUnRegistroTramite, sAdqNro)
        If rs.RecordCount > 0 Then
            Set fgeCot.Recordset = rs
            Call TransObj(rs)
        End If
    End If
End Sub

Private Sub TransObj(ByVal poRS As ADODB.Recordset)
Dim nCont As Integer, k As Integer, n As Integer
Dim rp As New ADODB.Recordset
Dim sObj As String

Dim clsDProv As DLogProveedor
Dim rs As ADODB.Recordset
Set clsDProv = New DLogProveedor
Set rs = New ADODB.Recordset

'nNumObj = poRS.RecordCount
Dim pColCot As Integer
pColCot = 5

fgeSel.Cols = 5
fgeSel.ColWidth(0) = 400
fgeSel.ColWidth(1) = 0
fgeSel.ColWidth(2) = 3000
fgeSel.ColWidth(3) = 1
fgeSel.ColWidth(4) = 1
fgeSel.ColWidth(5) = 1
fgeSel.RowHeight(0) = 300
fgeSel.TextMatrix(1, 2) = "SELECCION DE PROVEEDORES"

'fgeCot.Rows = nNumObj + 1
k = 0
    
For nCont = 1 To fgeCot.Rows - 1
    
    sObj = fgeCot.TextMatrix(nCont, 1)
    'Carga Proveedores del producto X
    Set rs = clsDProv.CargaProveedorBS(ProBSProveedor, sObj)
    If Not rs.EOF Then
        Do While Not rs.EOF
            k = 0
            k = SeHalla(rs!cPersCod)
            If k = 0 Then
               'n = InsFlex(fgePro)
               fgePro.AdicionaFila
               n = fgePro.Rows - 1
               fgePro.TextMatrix(n, 0) = "P" & Format(n, "00")
               fgePro.TextMatrix(n, 1) = rs!cPersCod
               fgePro.TextMatrix(n, 2) = rs!cPersNombre
               fgePro.TextMatrix(n, 3) = IIf(IsNull(rs!cPersDireccDomicilio), "", rs!cPersDireccDomicilio)
               
               fgeCot.Cols = fgePro.Rows + pColCot
               fgeCot.ColWidth(fgePro.Rows + (pColCot - 1)) = 400
               
               fgeSel.Cols = fgePro.Rows + pColCot
               fgeSel.ColWidth(fgePro.Rows + (pColCot - 1)) = 400
               
               fgeCot.TextMatrix(0, fgePro.Rows + (pColCot - 1)) = "P" & Format(n, "00")
               fgeCot.TextMatrix(nCont, fgePro.Rows + (pColCot - 1)) = "X"
               'fgeCot.ColAlignment(fgePro.Rows + (pColCot - 1)) = 4
            
               fgeSel.TextMatrix(0, fgePro.Rows + (pColCot - 1)) = "P" & Format(n, "00")
               'fgeSel.ColAlignment(fgePro.Rows + (pColCot - 1)) = 4
            
            Else
               fgeCot.TextMatrix(nCont, k + pColCot) = "X"
            End If
            rs.MoveNext
        Loop
    'Else
        'MsgBox fgeCot.TextMatrix(nCont, 2) & " sin proveedores ", vbInformation, " Aviso"
    End If
Next
End Sub

Function SeHalla(vCod As String) As Integer
Dim z As Integer
SeHalla = 0
For z = 1 To fgePro.Rows - 1
    If fgePro.TextMatrix(z, 1) = vCod Then
       SeHalla = z
       Exit For
    End If
Next
End Function

Private Sub Limpiar()
    lblAdqNro.Caption = ""
    fgeCot.Clear
    fgeCot.FormaCabecera
    fgeCot.Rows = 2
    fgeSel.Clear
    fgeSel.FormaCabecera
    fgeSel.Rows = 2
    fgePro.Clear
    fgePro.FormaCabecera
    fgePro.Rows = 2
End Sub
