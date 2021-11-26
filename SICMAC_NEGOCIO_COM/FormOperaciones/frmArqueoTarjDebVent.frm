VERSION 5.00
Begin VB.Form frmArqueoTarjDebVent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Arqueo de Stock de Tarjetas de Débito - Ventanilla"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   Icon            =   "frmArqueoTarjDebVent.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbClase 
      Height          =   315
      ItemData        =   "frmArqueoTarjDebVent.frx":030A
      Left            =   840
      List            =   "frmArqueoTarjDebVent.frx":0314
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdRegArqueo 
      Caption         =   "Registrar Arqueo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   6360
      Width           =   1700
   End
   Begin VB.Frame fraGlosa 
      Caption         =   "Glosa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   7695
      Begin VB.TextBox txtGlosa 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   405
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   7455
      End
   End
   Begin SICMACT.FlexEdit feArqueoVent 
      Height          =   4695
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8281
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Cod. Sub Lote-Cant. Stock-Cant. Física-Detalle-Aux-Estado"
      EncabezadosAnchos=   "300-1500-1200-1200-1200-0-1500"
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
      ColumnasAEditar =   "X-X-X-3-4-5-X"
      ListaControles  =   "0-0-0-0-1-1-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-C-C-L-L"
      FormatosEdit    =   "0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   6
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label Label1 
      Caption         =   "Clase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblFechaValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   6600
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblFechaField 
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   4
      Top             =   165
      Width           =   615
   End
   Begin VB.Label lblUsuSupValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblUsuSupField 
      Caption         =   "Usuario Supervisa:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   165
      Width           =   1815
   End
   Begin VB.Label lblUsuArqValue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblUsuArqField 
      Caption         =   "Usuario Arqueado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   1695
   End
End
Attribute VB_Name = "frmArqueoTarjDebVent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***************************************************************************
'** Nombre : frmArqueoTarjDebVent
'** Descripción : Formulario para realizar el Arqueo de Stock de Tarjetas de Debito  - Ventanilla
'** Creación : PASI, 20151221
'** Referencia : TI-ERS069-2015
'***************************************************************************
Option Explicit
Dim bResultadoVisto As Boolean
Dim oVisto As frmVistoElectronico
Dim cUsuVisto As String
Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral
Dim nMatDetFaltante() As TDetFaltante

Dim bConforme As Boolean 'GIPO 30/09/16 ERS051-2016
Dim textoEstado As String 'GIPO 04/10/16 ERS051-2016

Private Type TDetFaltante
    cCodSubLote As String
    cNumTarjeta As String
    nFaltante As Integer
End Type
Public Sub Inicia()
    Set oVisto = New frmVistoElectronico
    bResultadoVisto = oVisto.Inicio(15)
    If Not bResultadoVisto Then
        Exit Sub
    End If
    cUsuVisto = oVisto.ObtieneUsuarioVisto
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    If oCaja.ObtieneExisteArqueoVentanilla(gdFecSis, UCase(gsCodUser)) Then
        MsgBox "El Arqueo de este día ya ha sido realizado.", vbInformation, "Mensaje"
        Exit Sub
    End If
    Me.Show 1
End Sub

'GIPO
Private Sub cmbClase_Click()
    Me.txtGlosa.Text = textoEstado & "/" & Me.cmbClase.Text
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdRegArqueo_Click()
Dim oDMov As DMov
Dim i As Integer
Dim X As Integer
Dim lsMovNro As String
Dim nIdArqueo As Integer
Dim nIdArqueoVent As Integer
Dim bTrans As Boolean

Me.cmdRegArqueo.Enabled = False 'GIPO 27-10-2016

Set oDMov = New DMov
On Error GoTo ErrorRegistra
If feArqueoVent.TextMatrix(1, 1) = "" Then
    MsgBox "No Existen Datos para realizar el Arqueo. ", vbInformation, "Mensaje"
    Exit Sub
End If

If feArqueoVent.TextMatrix(1, 1) <> "" Then
    For i = 1 To feArqueoVent.Rows - 1
        If Trim(feArqueoVent.TextMatrix(i, 4)) = "" Then
            MsgBox "En el Stock de Tarjeta en Bóveda no se ha ingresado el detalle del registro (" & i & "). Verifique.", vbInformation, "Mensaje"
            Me.cmdRegArqueo.Enabled = True  'GIPO 27-10-2016
            Exit Sub
        End If
    Next
End If
    If MsgBox("¿Está seguro de realizar el Arqueo?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Me.cmdRegArqueo.Enabled = True  'GIPO 27-10-2016
        Exit Sub
    End If
    lsMovNro = oDMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    oCaja.dBeginTrans
    bTrans = True
    'GIPO ERS051-2016
    nIdArqueo = oCaja.RegistrarArqueoTarjDebito(lblUsuArqValue.Caption, lblUsuSupValue.Caption, gdFecSis & " " & GetHoraServer, lsMovNro, Trim(Replace(Replace((txtGlosa.Text), Chr(10), ""), Chr(13), "")))
    For i = 1 To feArqueoVent.Rows - 1
        nIdArqueoVent = oCaja.RegistrarArqueoTarjDebitoEnVentanilla(nIdArqueo, feArqueoVent.TextMatrix(i, 1), feArqueoVent.TextMatrix(i, 2), feArqueoVent.TextMatrix(i, 3))
        For X = 1 To UBound(nMatDetFaltante)
            If nMatDetFaltante(X).cCodSubLote = feArqueoVent.TextMatrix(i, 1) Then
                oCaja.RegistrarArqueoTarjDebitoEnVentanillaDet nIdArqueoVent, nMatDetFaltante(X).cNumTarjeta, nMatDetFaltante(X).nFaltante
            End If
        Next
    Next
    oCaja.dCommitTrans
    'MARG ERS052-2017----
    oVisto.RegistraVistoElectronico 0, , lblUsuArqValue.Caption
    'END MARG- -------------
    MsgBox "El Arqueo ha sido realizado correctamente.", vbInformation, "Aviso"
    'ImprimePDF GIPO ERS051-2016
    generarPDFacta (nIdArqueo)
    bTrans = False
    Unload Me
Exit Sub
ErrorRegistra:
    If bTrans Then
        oCaja.dRollbackTrans
        Set oCaja = Nothing
    End If
    MsgBox err.Number & " - " & err.Description, vbInformation, "Error"
End Sub

'Created By GIPO 04-11-2016
Public Sub generarPDFacta(ByVal IdArqueo As Integer)
    Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    Dim rs As ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs = oCaja.obtenerDatosActaTarjArqueadas(IdArqueo)
    Set rs2 = oCaja.obtenerDetalleTarjArqueadasVentanilla(IdArqueo)
    
    Dim oDoc As cPDF
    Set oDoc = New cPDF
    If Not oDoc.PDFCreate(App.Path & "\Spooler\ACTA_ARQUEO_VENTANILLA_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    oDoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
    oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"
    'Tamaño de hoja A4
    oDoc.NewPage A4_Vertical
    
    oDoc.WImage 75, 40, 35, 105, "Logo"
    oDoc.WTextBox 63, 40, 15, 500, "ACTA DE ARQUEO DE TARJETAS - VENTANILLA", "F2", 12, hCenter
    oDoc.WTextBox 90, 40, 732, 510, "En la ciudad de " & rs!Ciudad & " , el día " & rs!FechaCompleta & ", " & _
                                    "a horas " & rs!Hora & " se suscribe en actas que en Bóveda de la " & rs!Agencia & "," & _
                                    "con la presencia del (la) Sr(a). " & rs!NombrePersonaArqueada & " (" & rs!CargoUserArqueado & ") " & _
                                    "y el (la) Sr(a). " & rs!NombreArqueador & " (" & rs!CargoUserSuperviza & "), se procedió a realizar " & _
                                    "el arqueo de tarjetas débito Visa en la Agencia mencionada.", "F1", 9, hjustify, vTop, vbBlack, 0, vbBlack, False, 10
                                     
    Dim h1 As Integer
    h1 = 90 + 30 'espacio después del cuadro
    
    '************************DETALLE DE TARJETAS HABILITADAS**************************
    oDoc.WTextBox h1 + 30, 50, 20, 510, "DETALLE DE TARJETAS ARQUEADAS EN VENTANILLA", "F2", 9, hjustify
    
    oDoc.WTextBox h1 + 50, 50, 20, 100, "Número Tarjeta", "F2", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
    oDoc.WTextBox h1 + 50, 150, 20, 80, "Usuario Receptor", "F2", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
    oDoc.WTextBox h1 + 50, 230, 20, 80, "Estado", "F2", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
    oDoc.WTextBox h1 + 50, 310, 20, 120, "Cliente", "F2", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
    
    h1 = h1 + 50
    Do While Not rs2.EOF
        h1 = h1 + 20
        If h1 >= 745 Then
            'imprimir pie de página antes de crear la siguiente
            oDoc.WTextBox 740 + 27, 45, 15, 510, printFooter, "F1", 7, hjustify
            oDoc.NewPage A4_Vertical
            h1 = h1 - 705
        End If
        oDoc.WTextBox h1, 50, 20, 100, rs2!cNumTarjeta, "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
        oDoc.WTextBox h1, 150, 20, 80, rs2!cUserArqueado, "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
        oDoc.WTextBox h1, 230, 20, 80, rs2!Estado, "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
        oDoc.WTextBox h1, 310, 20, 120, rs2!Cliente, "F1", 8, hCenter, vMiddle, vbBlack, 1, vbBlack, False, 1
        rs2.MoveNext
    Loop
    
    '************************PÁRRAFO FINAL**************************
    h1 = h1 + 40
     If h1 >= 700 Then
      'imprimir pie de página antes de crear la siguiente
        oDoc.WTextBox 740 + 27, 45, 15, 510, printFooter, "F1", 7, hjustify
        oDoc.NewPage A4_Vertical
        h1 = h1 - 650
    End If
    
    Dim parrafoFinal As String
    parrafoFinal = "Siendo las " & rs!Hora & " del día " & rs!FechaCompleta & ", se dió por concluido " & _
                   "el arqueo de tarjetas de débito VISA; por lo cual firman en señal de conformidad."
    oDoc.WTextBox h1, 50, 20, 510, parrafoFinal, "F1", 9, hjustify
    
    h1 = h1 + 70
    oDoc.WTextBox h1, 50, 20, 250, "_________________________________", "F1", 9, hCenter
    oDoc.WTextBox h1 + 15, 50, 20, 250, rs!NombreArqueador, "F2", 9, hCenter
    oDoc.WTextBox h1 + 25, 50, 20, 250, rs!CargoUserSuperviza, "F2", 9, hCenter
    oDoc.WTextBox h1 + 35, 50, 20, 250, "Responsable Arqueo", "F2", 9, hCenter
    
    oDoc.WTextBox h1, 300, 20, 250, "_________________________________", "F1", 9, hCenter
    oDoc.WTextBox h1 + 15, 300, 20, 250, rs!NombrePersonaArqueada, "F2", 9, hCenter
    oDoc.WTextBox h1 + 25, 300, 20, 250, rs!CargoUserArqueado, "F2", 9, hCenter
    oDoc.WTextBox h1 + 35, 300, 20, 250, "Usuario Arqueado", "F2", 9, hCenter
    
    
    '************************PIE DE PÁGINA**************************
    Dim sPiePagina As String
    sPiePagina = printFooter
    oDoc.WTextBox 767, 45, 15, 510, sPiePagina, "F1", 7, hjustify
    
    oDoc.PDFClose
    oDoc.Show
End Sub
Private Function printFooter()
Dim footer As String
footer = "Oficina Principal: Jr. Próspero No  791 -  Iquitos ;" & _
                 "Ag. Calle Arequipa: Ca Arequipa Nº 428;  Agencia Punchana Av. 28 de Julio 829 - Iquitos;" & _
                 "Ag. Belén: Av. Grau Nº 1260 - Iquitos; Ag. San Juan Bautista- Avda. Abelardo Quiñones Nº 2670- Iquitos;" & _
                 "Ag. Pucallpa: Jr. Ucayali  No  850 - 852 ; Ag. Huánuco: Jr. General Prado No   836;" & _
                 "Ag. Yurimaguas: Ca. Simón Bolívar Nº 113; Ag. Tingo María: Av. Antonio Raymondi  Nº 246 ;" & _
                 "Ag. Tarapoto: Jr San Martín Nº 205 ;  Ag. Requena: Calle San Francisco Mz 28 Lt 07;" & _
                 "Ag. Cajamarca: Jr. Amalia Puga  Nº 417; Ag.  Aguaytía- Jr. Rio Negro Nº 259;" & _
                 "Ag. Cerro de Pasco; Plaza Carrión Nº 191; Ag. Minka: Ciudad Comercial Minka - Av. Argentina Nº 3093- Local 230 - Callao."
printFooter = footer
End Function


Private Sub feArqueoVent_OnCellChange(pnRow As Long, pnCol As Long)
    Call GetArqueoConforme 'GIPO
    Dim i As Integer
    If pnCol = 3 Then
        If Trim(feArqueoVent.TextMatrix(pnRow, 2)) = Trim(feArqueoVent.TextMatrix(pnRow, 3)) Then
            feArqueoVent.TextMatrix(pnRow, 4) = "OK"
            feArqueoVent.TextMatrix(pnRow, 6) = "CONFORME" 'GIPO
        Else
            feArqueoVent.TextMatrix(pnRow, 4) = ""
            feArqueoVent.TextMatrix(pnRow, 6) = "NO CONFORME" 'GIPO
        End If
        For i = 1 To UBound(nMatDetFaltante)
            If nMatDetFaltante(i).cCodSubLote = feArqueoVent.TextMatrix(pnRow, 1) Then
                nMatDetFaltante(i).nFaltante = 0
            End If
        Next
        SendKeys "{Tab}", True
    End If
End Sub

'GIPO
Private Sub GetArqueoConforme()
    Dim i As Integer
     For i = 1 To feArqueoVent.Rows - 1
          If feArqueoVent.TextMatrix(i, 2) <> feArqueoVent.TextMatrix(i, 3) Then
            Me.txtGlosa.Text = "NO CONFORME/" & Me.cmbClase.Text
            textoEstado = "NO CONFORME"
            Exit Sub
          ElseIf feArqueoVent.TextMatrix(i, 3) = "" Then
            Exit Sub
          End If
     Next
     Me.txtGlosa.Text = "CONFORME/" & Me.cmbClase.Text
     textoEstado = "CONFORME"
End Sub

Private Sub feArqueoVent_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
    Dim row As Integer
    Dim vMat As Variant
    Dim vMatFalt As Variant
    Dim i As Integer
    Dim bHab As Boolean
    Dim X As Long
    Dim iMat As Integer
    
    Dim bConforme As Boolean  'GIPO
    bConforme = False 'GIPO
    
    row = feArqueoVent.row
    If Trim(feArqueoVent.TextMatrix(row, 2)) = Trim(feArqueoVent.TextMatrix(row, 3)) Then
        'MsgBox "El caso no requiere registrar faltante. Continue.", vbInformation, "Mensaje"
        psCodigo = "OK"
        bConforme = True
        'Exit Sub
    End If
    If Trim(feArqueoVent.TextMatrix(row, 3)) = "" Then
        MsgBox "No se ha ingresado la Cantidad Fisica. Verifique.", vbInformation, "Mensaje"
        Exit Sub
    End If
    ReDim vMat(5, 0)
    For i = 1 To UBound(nMatDetFaltante)
        If nMatDetFaltante(i).cCodSubLote = feArqueoVent.TextMatrix(row, 1) Then
        
        iMat = UBound(vMat, 2) + 1
        ReDim Preserve vMat(5, 0 To iMat)
        vMat(1, iMat) = ""
        vMat(2, iMat) = ""
        vMat(3, iMat) = nMatDetFaltante(i).cCodSubLote 'codSubLote
        vMat(4, iMat) = nMatDetFaltante(i).cNumTarjeta 'NumTarjeta
        vMat(5, iMat) = nMatDetFaltante(i).nFaltante 'nFaltante
        End If
    Next
    If UBound(vMat, 2) >= 1 Then
        bHab = False
        vMatFalt = frmArqueoTarjDebBovedaDetFal.Inicio(vMat, CLng(Trim(feArqueoVent.TextMatrix(row, 2))) - CLng(Trim(feArqueoVent.TextMatrix(row, 3))), bConforme)
        For i = 1 To UBound(nMatDetFaltante)
            For X = 1 To UBound(vMatFalt, 2)
                If vMatFalt(3, X) = nMatDetFaltante(i).cCodSubLote And _
                    vMatFalt(4, X) = nMatDetFaltante(i).cNumTarjeta Then
                    nMatDetFaltante(i).nFaltante = vMatFalt(5, X)
                    If vMatFalt(5, X) = 1 Then
                        bHab = True
                    End If
                End If
            Next
        Next
        If bConforme = False Then
            psCodigo = IIf(bHab, "OK", "")
        End If
    End If
End Sub
Private Sub feArqueoVent_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
      If pnCol = 3 Then
        If IsNumeric(feArqueoVent.TextMatrix(feArqueoVent.row, pnCol)) = False Then
            Cancel = False
            SendKeys "{Tab}", True
            Exit Sub
        End If
        If feArqueoVent.TextMatrix(feArqueoVent.row, pnCol) < 0 Then
            Cancel = False
            SendKeys "{Tab}", True
            Exit Sub
        End If
        If CLng(feArqueoVent.TextMatrix(feArqueoVent.row, 2)) < CLng(feArqueoVent.TextMatrix(feArqueoVent.row, pnCol)) Then
            Cancel = False
            SendKeys "{Tab}", True
            Exit Sub
        End If
    End If
End Sub
Private Sub Form_Load()
    CargaDatos
End Sub
Private Sub CargaDatos()
    Me.cmbClase.ListIndex = 0 'GIPO
    Dim rs As ADODB.Recordset
    Dim rsNTarj As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rsNTarj = New ADODB.Recordset
    
    Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral
    
    Me.lblUsuArqValue.Caption = UCase(gsCodUser)
    Me.lblUsuSupValue.Caption = UCase(cUsuVisto)
    Me.lblFechaValue.Caption = CDate(gdFecSis)
    
    ReDim Preserve nMatDetFaltante(0)
    
    Set rs = oCaja.ObtieneSubLotesxArqueoVentanilla(lblUsuArqValue.Caption, gdFecSis)
    Do While Not rs.EOF
        feArqueoVent.AdicionaFila
        feArqueoVent.TextMatrix(feArqueoVent.row, 1) = rs!cCodSubLote
        feArqueoVent.TextMatrix(feArqueoVent.row, 2) = rs!nStock
        
        Set rsNTarj = oCaja.ObtieneTarjetasxArqueoVentanilla(lblUsuArqValue.Caption, gdFecSis, rs!cCodSubLote)
        Do While Not rsNTarj.EOF
            ReDim Preserve nMatDetFaltante(UBound(nMatDetFaltante) + 1)
            nMatDetFaltante(UBound(nMatDetFaltante)).cCodSubLote = rs!cCodSubLote
            nMatDetFaltante(UBound(nMatDetFaltante)).cNumTarjeta = rsNTarj!cNumTarjeta
            nMatDetFaltante(UBound(nMatDetFaltante)).nFaltante = 0
            rsNTarj.MoveNext
        Loop
        rs.MoveNext
    Loop
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdRegArqueo.SetFocus
    End If
End Sub


