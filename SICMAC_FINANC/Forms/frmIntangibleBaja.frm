VERSION 5.00
Begin VB.Form frmIntangibleBaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Intangibles - Baja de Intangibles"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14010
   Icon            =   "frmIntangibleBaja.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   14010
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtGlosa 
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   4800
      Width           =   10095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   12480
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin Sicmact.FlexEdit feBaja 
      Height          =   3255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   5741
      Cols0           =   9
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Cod.-Descripción-Rubro-Monto Amortizado-Fecha Ult. Amort.--nmovnro-CtaCont"
      EncabezadosAnchos=   "300-1500-6000-1500-1800-1500-800-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-6-X-X"
      ListaControles  =   "0-0-0-0-0-0-4-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-C-L-R-L-C-C-C"
      FormatosEdit    =   "0-1-1-1-2-1-1-1-1"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame frmBuscar 
      Caption         =   "Buscar"
      Height          =   975
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton cmdBuscar 
         Appearance      =   0  'Flat
         Caption         =   "Mostrar"
         Height          =   345
         Left            =   3720
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cboRubro 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Rubro:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdDarBaja 
      Caption         =   "Dar Baja"
      Height          =   375
      Left            =   11280
      TabIndex        =   6
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Glosa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   735
   End
End
Attribute VB_Name = "frmIntangibleBaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**-------------------------------------------------------------------------------------**'
'** Formulario : frmIntangibleBaja                                                      **'
'** Finalidad  : Este formulario permite realizar la baja de las intangibles            **'                                                 **'
'** Programador: Paolo Hector Sinti Cabrera - PASI                                      **'
'** Fecha/Hora : 20140305 11:50 AM                                                      **'
'**-------------------------------------------------------------------------------------**'
Option Explicit
Dim oIntangible As dIntangible
Dim lsRubro As String
Dim lsNroAsientoRef() As String
Private Sub cmdBuscar_Click()
    CargarDatos
End Sub
Private Sub CargarDatos()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oIntangible = New dIntangible
    Dim row As Integer
    lsRubro = IIf(Right(cboRubro.Text, 1) = 0, "%", Right(cboRubro.Text, 1))
    
    FormateaFlex feBaja
    Set rs = oIntangible.ListaAmortizacionesCompletas(lsRubro)
    If Not rs.EOF Then
        Do While Not rs.EOF
            feBaja.AdicionaFila
            row = feBaja.row
            feBaja.TextMatrix(row, 1) = rs!Codigo
            feBaja.TextMatrix(row, 2) = rs!Descripcion
            feBaja.TextMatrix(row, 3) = rs!Rubro
            feBaja.TextMatrix(row, 4) = Format(rs!MontoAmort, "#,#0.00")
            feBaja.TextMatrix(row, 5) = rs!FechaUlt
            feBaja.TextMatrix(row, 6) = rs!Estado
            feBaja.TextMatrix(row, 7) = rs!nmovnrointg
            feBaja.TextMatrix(row, 8) = rs!CtaCont
            rs.MoveNext
        Loop
    Else
        MsgBox "No hay Datos para Mostrar", vbInformation, "Aviso!!!"
    End If
End Sub
Private Sub cmdDarBaja_Click()
    Dim nactivados As Integer
    Dim I As Integer, X As Integer
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim oMov As DMov
    Set oMov = New DMov
    Dim oIntangible As dIntangible
    Set oIntangible = New dIntangible
    Dim rsCtas As ADODB.Recordset
    Set rsCtas = New ADODB.Recordset
    Dim lsCodAmort As String
    Dim lsCodOpe As String
    Dim Dope As DOperacion
    Set Dope = New DOperacion
    Dim lsCadenaAsiento As String
    Dim oImpr As NContImprimir
    Set oImpr = New NContImprimir
    Dim oPrevio As clsPrevioFinan
    Set oPrevio = New clsPrevioFinan
    
    Dim lsCtaD As String, lsCtaH As String
    
    If feBaja.TextMatrix(1, 1) <> "" Then
        If Len(Me.txtGlosa.Text) = 0 Then
            MsgBox "No se ha ingresado una descripción válida para la glosa.", vbInformation, "Aviso!!!"
            Exit Sub
        End If
        nactivados = 0
        For I = 1 To feBaja.Rows - 1
            If feBaja.TextMatrix(I, 6) = "." Then
                nactivados = nactivados + 1
            End If
        Next I
        If nactivados = 0 Then
            MsgBox "Asegurese de haber seleccionado al menos una intangible para dar de baja", vbInformation, "Aviso!!!"
            Exit Sub
        End If
        If MsgBox("¿ Está seguro de realizar la Baja de las intangibles Seleccionadas", vbYesNo, "Atención") = vbNo Then Exit Sub
        
        oMov.BeginTrans
        lsCodAmort = feBaja.TextMatrix(1, 1)
                Select Case Mid(lsCodAmort, 7, 2)
                    Case "01"
                            lsCodOpe = gAmortizaIntangibleLicencia
                    Case "02"
                            lsCodOpe = gAmortizaIntangibleSoftware
                    Case "03"
                            lsCodOpe = gAmortizaIntangibleOtros
                End Select
        Set rsCtas = Dope.ObtenerCtasAmortIntangible(lsCodOpe)
        lsCtaD = rsCtas!cCtaContCodH
        lsCtaH = rsCtas!cCtaContCodOtroH
        ReDim Preserve lsNroAsientoRef(1 To 1, 0 To 0)
        X = 0
        For I = 1 To feBaja.Rows - 1
            If feBaja.TextMatrix(I, 6) = "." Then
                X = X + 1
                ReDim Preserve lsNroAsientoRef(1 To 1, X)
                lsMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                oMov.InsertaMov lsMovNro, gBajaIntangible, Trim(Me.txtGlosa.Text), 10
                lnMovNro = oMov.GetnMovNro(lsMovNro)
                oMov.InsertaMovIntgAmort Left(lsMovNro, 4), feBaja.TextMatrix(I, 7), 1, feBaja.TextMatrix(I, 1), 0, lnMovNro, 0, gdFecSis
                oIntangible.ActualizaBajaIntangible feBaja.TextMatrix(I, 7), feBaja.TextMatrix(I, 1), gdFecSis
                oMov.InsertaMovCta lnMovNro, 1, Replace(lsCtaD, "AG", Right(feBaja.TextMatrix(I, 8), 2)), Round(feBaja.TextMatrix(I, 4), 2)
                oMov.InsertaMovCta lnMovNro, 2, Replace(lsCtaH, "AG", Right(feBaja.TextMatrix(I, 8), 2)), Round(feBaja.TextMatrix(1, 4), 2) * -1
                lsNroAsientoRef(1, X) = lsMovNro
            End If
        Next I
        oMov.CommitTrans
        lsCadenaAsiento = ""
        For I = 1 To UBound(lsNroAsientoRef, 2)
            lsCadenaAsiento = lsCadenaAsiento + oImpr.ImprimeAsientoContable(lsNroAsientoRef(1, I), 66, 79)
        Next I
        oPrevio.Show lsCadenaAsiento, gBajaIntangible, False, 66, gImpresora
        CargarDatos
    Else
        MsgBox "No existen intangibles para dar de Baja", vbInformation, "Aviso!!!"
    End If
End Sub
Private Sub Form_Load()
    CargaCombo
End Sub
Private Sub CargaCombo()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oIntangible = New dIntangible
    Set rs = oIntangible.ListaTipoIntangible()
    If Not rs.EOF Then
        cboRubro.Clear
        Do While Not rs.EOF
            cboRubro.AddItem Trim(rs(1) & Space(100) & Trim(rs(0)))
            rs.MoveNext
        Loop
    End If
    If cboRubro.ListCount > 0 Then
        cboRubro.ListIndex = 0
    End If
End Sub
