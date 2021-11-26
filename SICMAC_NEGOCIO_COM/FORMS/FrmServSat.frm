VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmServSat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "S.A.T. - Información"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
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
      Left            =   8790
      TabIndex        =   7
      Top             =   6180
      Width           =   1200
   End
   Begin VB.CommandButton cbExporta 
      Caption         =   "Grabar BD"
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
      Left            =   7530
      TabIndex        =   6
      Top             =   6180
      Width           =   1200
   End
   Begin VB.CommandButton cbCargaDatos 
      Caption         =   "Obtener Datos"
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
      Left            =   8160
      TabIndex        =   4
      Top             =   825
      Width           =   1695
   End
   Begin VB.Frame frTipoCarga 
      Caption         =   "Tipo de Carga "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   855
      Index           =   1
      Left            =   720
      TabIndex        =   9
      Top             =   360
      Width           =   3255
      Begin VB.OptionButton opTributos 
         Caption         =   "Tributos"
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
         Left            =   255
         TabIndex        =   1
         Top             =   510
         Width           =   1215
      End
      Begin VB.OptionButton opPapeletas 
         Caption         =   "Papeletas"
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
         Left            =   255
         TabIndex        =   0
         Top             =   270
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog cdlgfile 
      Left            =   915
      Top             =   6165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frTipoCarga 
      Caption         =   "Carga de la Informacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   855
      Index           =   0
      Left            =   4560
      TabIndex        =   8
      Top             =   360
      Width           =   3255
      Begin VB.OptionButton opMensual 
         Caption         =   "Carga Mensual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   525
         Width           =   1815
      End
      Begin VB.OptionButton opDiaria 
         Caption         =   "Carga Diaria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   2
         Top             =   285
         Width           =   1650
      End
   End
   Begin SICMACT.FlexEdit feDetPapeletas 
      Height          =   4425
      Left            =   135
      TabIndex        =   12
      Top             =   1710
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   7805
      Cols0           =   20
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"FrmServSat.frx":0000
      EncabezadosAnchos=   "600-300-300-500-1000-300-400-900-900-900-900-1050-1050-1050-1050-900-400-400-400-1000"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "L-C-C-C-C-C-C-R-C-C-C-L-L-L-L-L-L-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-2-2-2-2-5-5-0-0-0-0-0-0-0"
      TextArray0      =   "nitem"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   600
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin SICMACT.FlexEdit feDetTributo 
      Height          =   4425
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   7805
      Cols0           =   19
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"FrmServSat.frx":00B5
      EncabezadosAnchos=   "600-300-300-500-800-300-400-850-850-850-850-1000-1500-1100-1000-850-850-600-1200"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-C-C-C-R-C-R-R-L-L-L-C-R-C-L-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-2-2-2-2-5-0-0-0-2-2-0-0"
      TextArray0      =   "nItem"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   600
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label txtTitulo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   450
      TabIndex        =   11
      Top             =   1320
      Width           =   9210
   End
   Begin VB.Label Label1 
      Caption         =   "CARGA DE INFORMACION DE TRIBUTOS/PAPELETAS(DIARIA - MENSUAL)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   330
      TabIndex        =   10
      Top             =   15
      Width           =   9510
   End
End
Attribute VB_Name = "FrmServSat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cbCargaDatos_Click()
Dim sArchivo As String
On Local Error Resume Next

cdlgfile.CancelError = True
'Especificar las extensiones a usar
cdlgfile.DefaultExt = "*.txt"
cdlgfile.Filter = "Textos (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
'MuestraTituloValid
cdlgfile.ShowOpen
If Err Then
    sArchivo = "" 'Cancelada la operación de abrir
Else
    sArchivo = cdlgfile.FileName
    ObtieneDatosSatTrib sArchivo
End If
End Sub
Private Sub MuestraTituloValid()
    If Me.opPapeletas.value = True Then
        Me.txtTitulo.Caption = "CARGA PAPELETAS - "
    Else
        Me.txtTitulo.Caption = "CARGA TRIBUTOS  - "
    End If
    If Me.opDiaria.value = True Then
        Me.txtTitulo.Caption = Me.txtTitulo.Caption & " DIARIO"
    Else
        Me.txtTitulo.Caption = Me.txtTitulo.Caption & " MENSUAL"
    End If
        
End Sub
Private Sub ObtieneDatosSatTrib(ByVal sArchivo As String)
Dim sCad As String
Dim sCTributo As String, Stributo As String
Dim sRecibos As String, sDigCheq As String
Dim sPeriodo As String, sPlaca As String
Dim nValImp As Double, nValReinc As Double
Dim sAno As String
Dim nItem As Long
Dim dFecvP As Date, dFecInf As Date
Dim dFecNoti As String, dFecLect  As String
Dim nValDeu As Double, nValDerE As Double, nValAjuste As Double, nValIntM As Double
Dim ntotal As Double
Dim nValGastos As Double, nValCostas As Double
Dim dFecv As Date
Dim sNombre As String, sContri As String, sCuenta As String, sCodSis As String
Dim sPapeletas As String, sTipo As String, sFalta As String, sCodAux As String, sLicCon As String
On Error GoTo ErrFileOpen
Open sArchivo For Input As #1
'bPrimeraLinea = True
'format(cadena, "mm/dd/yyyy")
'EN LA BASE DE DATO ESTA MM/DD/YYYY
nItem = 1
'T RIBUTOS - INSERTA REGISTROS
If Me.opTributos.value = True Then
      Me.feDetTributo.Clear
      Me.feDetTributo.Rows = 2
      Me.feDetTributo.FormaCabecera
      Do While Not EOF(1)
            Line Input #1, sCad
            Me.feDetTributo.AdicionaFila
            'CAPTURA DEL TEXTO
            sCTributo = Mid(sCad, 1, 1)
            Stributo = Mid(sCad, 2, 1)
            sAno = Mid(sCad, 133, 4)
            sRecibos = Mid(sCad, 4, 7)
            sDigCheq = Mid(sCad, 11, 1)
            sPeriodo = Mid(sCad, 12, 2)
            nValDeu = CDbl(Mid(sCad, 14, 11))
            nValDerE = CDbl(Mid(sCad, 25, 11))
            nValAjuste = CDbl(Mid(sCad, 36, 11))
            nValIntM = CDbl(Mid(sCad, 47, 11))
            dFecv = CDate(Mid(sCad, 64, 2) & "/" & Mid(sCad, 62, 2) & "/" & Mid(sCad, 58, 4))
            sNombre = Mid(sCad, 66, 24)
            sCuenta = Mid(sCad, 90, 13)
            sContri = Mid(sCad, 103, 10)
            nValGastos = CDbl(Mid(sCad, 113, 8))
            nValCostas = CDbl(Mid(sCad, 121, 8))
            sCodSis = Mid(sCad, 129, 4)
            ntotal = 0
            'INSERTA EN EL FLEXEDIT
            Me.feDetTributo.TextMatrix(nItem, 0) = nItem
            Me.feDetTributo.TextMatrix(nItem, 1) = sCTributo
            Me.feDetTributo.TextMatrix(nItem, 2) = Stributo
            Me.feDetTributo.TextMatrix(nItem, 3) = sAno
            Me.feDetTributo.TextMatrix(nItem, 4) = sRecibos
            Me.feDetTributo.TextMatrix(nItem, 5) = sDigCheq
            Me.feDetTributo.TextMatrix(nItem, 6) = sPeriodo
            Me.feDetTributo.TextMatrix(nItem, 7) = CDec(nValDeu)
            Me.feDetTributo.TextMatrix(nItem, 8) = CDec(nValDerE)
            Me.feDetTributo.TextMatrix(nItem, 9) = CDec(nValAjuste)
            Me.feDetTributo.TextMatrix(nItem, 10) = CDec(nValIntM)
            Me.feDetTributo.TextMatrix(nItem, 11) = CStr(dFecv)
            Me.feDetTributo.TextMatrix(nItem, 12) = sNombre
            Me.feDetTributo.TextMatrix(nItem, 13) = sCuenta
            Me.feDetTributo.TextMatrix(nItem, 14) = sContri
            Me.feDetTributo.TextMatrix(nItem, 15) = CDec(nValGastos)
            Me.feDetTributo.TextMatrix(nItem, 16) = CDec(nValCostas)
            Me.feDetTributo.TextMatrix(nItem, 17) = sCodSis
            Me.feDetTributo.TextMatrix(nItem, 18) = ntotal
            nItem = nItem + 1
    Loop
Else ' PAPELETAS  - INSERTA REGISTRO
    Me.feDetPapeletas.Clear
    Me.feDetPapeletas.Rows = 2
    Me.feDetPapeletas.FormaCabecera
    Do While Not EOF(1)
         Line Input #1, sCad
         Me.feDetPapeletas.AdicionaFila
         'CAPTURA DEL TEXTO
          sCTributo = Mid(sCad, 1, 1)
          Stributo = Mid(sCad, 2, 1)
          sAno = Mid(sCad, 3, 1)
          sPlaca = Mid(sCad, 4, 7)
          sDigCheq = Mid(sCad, 11, 1)
          sPeriodo = Mid(sCad, 12, 2)
          nValImp = CDbl(Mid(sCad, 14, 11))
          nValReinc = CDbl(Mid(sCad, 25, 11))
          If IsNull(Mid(sCad, 36, 11)) Or Trim(Mid(sCad, 36, 11)) = "" Then
                nValGastos = 0
          Else
                  nValGastos = CDbl(Mid(sCad, 36, 11))
          End If
    
          dFecvP = Mid(sCad, 62, 2) & "/" & Mid(sCad, 64, 2) & "/" & Mid(sCad, 58, 4)
          dFecInf = Mid(sCad, 70, 2) & "/" & Mid(sCad, 72, 2) & "/" & Mid(sCad, 66, 4)
          If IsNull(Mid(sCad, 74, 8)) Or Trim(Mid(sCad, 74, 8)) = "" Then
                dFecLect = ""
           Else
                dFecLect = Mid(sCad, 78, 2) & "/" & Mid(sCad, 80, 2) & "/" & Mid(sCad, 74, 4)
           End If
           If IsNull(Mid(sCad, 82, 8)) Or Trim(Mid(sCad, 82, 8)) = "" Then
                dFecNoti = ""
           Else
                dFecNoti = Mid(sCad, 86, 2) & "/" & Mid(sCad, 88, 2) & "/" & Mid(sCad, 82, 4)
           End If
           
           sPapeletas = Mid(sCad, 90, 13)
           sTipo = Mid(sCad, 103, 3)
           sFalta = Mid(sCad, 106, 3)
           sCodAux = Mid(sCad, 109, 3)
           sLicCon = Mid(sCad, 112, 10)
          'INSERTA EN FLEXEDIT
          Me.feDetPapeletas.TextMatrix(nItem, 0) = nItem
          Me.feDetPapeletas.TextMatrix(nItem, 1) = sCTributo
          Me.feDetPapeletas.TextMatrix(nItem, 2) = Stributo
          Me.feDetPapeletas.TextMatrix(nItem, 3) = sAno
          Me.feDetPapeletas.TextMatrix(nItem, 4) = sPlaca
          Me.feDetPapeletas.TextMatrix(nItem, 5) = sDigCheq
          Me.feDetPapeletas.TextMatrix(nItem, 6) = sPeriodo
          Me.feDetPapeletas.TextMatrix(nItem, 7) = CDec(nValImp)
          Me.feDetPapeletas.TextMatrix(nItem, 8) = CDec(nValReinc)
          Me.feDetPapeletas.TextMatrix(nItem, 9) = CDec(nValGastos)
          Me.feDetPapeletas.TextMatrix(nItem, 10) = CDec(nValCostas)
          Me.feDetPapeletas.TextMatrix(nItem, 11) = CStr(dFecvP)
          Me.feDetPapeletas.TextMatrix(nItem, 12) = CStr(dFecInf)
          Me.feDetPapeletas.TextMatrix(nItem, 13) = CStr(dFecLect)
          Me.feDetPapeletas.TextMatrix(nItem, 14) = CStr(dFecNoti)
          Me.feDetPapeletas.TextMatrix(nItem, 15) = sPapeletas
          Me.feDetPapeletas.TextMatrix(nItem, 16) = sTipo
          Me.feDetPapeletas.TextMatrix(nItem, 17) = sFalta
          Me.feDetPapeletas.TextMatrix(nItem, 18) = sCodAux
          Me.feDetPapeletas.TextMatrix(nItem, 19) = sLicCon
          nItem = nItem + 1
    Loop
End If

Exit Sub
ErrFileOpen:
    Close #1
    'CmdCancelar_Click
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub cbExporta_Click()
If MsgBox("¿Desea realizar la exportación de datos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim rsVar As Recordset
    Dim oServ As NCapServicios
    
    If opTributos.value Then 'Tributos
        Set rsVar = feDetTributo.GetRsNew()
        Set oServ = New NCapServicios
        If oServ.AgregaServSATTributo(rsVar) = 0 Then
            MsgBox "Exportación realizada con éxito", vbInformation, "Aviso"
        End If
    Else 'Papeletas
        Set rsVar = feDetPapeletas.GetRsNew()
        Set oServ = New NCapServicios
        If oServ.AgregaServSATPapeletas(rsVar) = 0 Then
            MsgBox "Exportación realizada con éxito", vbInformation, "Aviso"
        End If
    End If
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.opPapeletas.value = True
Me.opDiaria.value = True
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub opPapeletas_Click()
    seleccionFlex
 End Sub
 
 Private Sub seleccionFlex()
   If Me.opPapeletas.value = True Then
        Me.feDetTributo.Visible = False
        Me.feDetPapeletas.Visible = True
                Me.txtTitulo.Caption = "CARGA PAPELETAS  "
        'Me.opPapeletas.SetFocus
   Else
        Me.feDetTributo.Visible = True
        Me.feDetPapeletas.Visible = False
        'Me.opTributos.SetFocus
        Me.txtTitulo.Caption = "CARGA TRIBUTOS   "
  End If
      
End Sub

Private Sub opTributos_Click()
       seleccionFlex
End Sub
