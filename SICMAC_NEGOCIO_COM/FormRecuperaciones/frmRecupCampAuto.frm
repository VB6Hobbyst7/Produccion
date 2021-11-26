VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRecupCampAuto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorización  de pago con campaña - Recuperaciones"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   13605
   Icon            =   "frmRecupCampAuto.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   13605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
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
      Left            =   12360
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdAutorizar 
      Caption         =   "Autorizar"
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
      Left            =   11040
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Créditos Acogidos"
      TabPicture(0)   =   "frmRecupCampAuto.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "feCreditos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin SICMACT.FlexEdit feCreditos 
         Height          =   3255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   13290
         _ExtentX        =   23442
         _ExtentY        =   5741
         Cols0           =   14
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Cod-N° Crédito-Cliente-Atraso-Capital-Interés-Mora-Gastos-ICV-Deuda-Perdón-Pago-Pendiente"
         EncabezadosAnchos=   "400-0-1700-2300-600-800-800-800-800-800-1000-1000-1000-1000"
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
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColor       =   16777215
         EncabezadosAlineacion=   "C-L-C-L-C-R-C-R-R-R-R-R-R-C"
         FormatosEdit    =   "0-0-0-0-0-2-2-2-2-2-2-2-2-2"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   6
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         CellBackColor   =   16777215
      End
   End
End
Attribute VB_Name = "frmRecupCampAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmRecupCampAuto
'** Descripción : Formulario para Autorizar el Pago de creditos que fueron acogido por la campaña de recuperaciones
'**               Creado segun TI-ERS035-2015
'** Creación    : WIOR, 20150522 09:00:00 AM
'**********************************************************************************************
Option Explicit
Private i As Integer

Private Sub cmdAutorizar_Click()
Dim oNCredito As COMNCredito.NCOMCredito
Dim oDCredito As COMDCredito.DCOMCredito
Dim RsDatos As ADODB.Recordset
Dim nCod As Long
Dim sMovNro As String
Dim bGrabar As Boolean
nCod = 0
sMovNro = ""

If Trim(feCreditos.TextMatrix(1, 0)) <> "" Then
    Set oDCredito = New COMDCredito.DCOMCredito
    Set RsDatos = oDCredito.RecuperarCampanaRecupXCredEstadoFecha(Trim(feCreditos.TextMatrix(feCreditos.row, 2)), 1, gdFecSis)
    
    If Not (RsDatos.BOF And RsDatos.EOF) Then
        If CInt(RsDatos!nEstado) = 1 Then
            MsgBox "Crédito ya fue autorizado, Se volveran a cargar el listado de créditos.", vbInformation, "Aviso"
            LimpiarDatos
            Exit Sub
        End If
    End If
    
    If MsgBox("Esta seguro de autorizar el pago de con campaña del Crédito N° " & Trim(feCreditos.TextMatrix(feCreditos.row, 2)) & "?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        nCod = CLng(feCreditos.TextMatrix(feCreditos.row, 1))
        sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        
        Set oNCredito = New COMNCredito.NCOMCredito
        bGrabar = oNCredito.GarbarAutorizacionCredCampanaRecup(nCod, gsCodCargo, sMovNro)
        
        If bGrabar Then
            MsgBox "Se realizo la autorización correctamente.", vbInformation, "Aviso"
        Else
             MsgBox "Hubo errores al grabar la autorización, Se volveran a cargar el listado de créditos.", vbError, "Error"
        End If
        LimpiarDatos
    End If
End If
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub feCreditos_Click()
If Trim(feCreditos.TextMatrix(1, 0)) <> "" Then
    cmdAutorizar.Enabled = True
End If
End Sub

Private Sub Form_Load()
cmdAutorizar.Enabled = False
Call ListarCreditos
End Sub

Private Sub ListarCreditos()
Dim oDCredito As COMDCredito.DCOMCredito
Dim RsDatos As ADODB.Recordset

Set oDCredito = New COMDCredito.DCOMCredito
Set RsDatos = oDCredito.RecuperarCampanaRecupCreditosAAprobar(gdFecSis, gsCodCargo)

LimpiaFlex feCreditos
If Not (RsDatos.EOF And RsDatos.BOF) Then
    For i = 1 To RsDatos.RecordCount
        feCreditos.AdicionaFila
        feCreditos.TextMatrix(i, 1) = CLng(RsDatos!nId)
        feCreditos.TextMatrix(i, 2) = Trim(RsDatos!cCtaCod)
        feCreditos.TextMatrix(i, 3) = Trim(RsDatos!cPersNombre)
        feCreditos.TextMatrix(i, 4) = CInt(RsDatos!nDiasAtraso)
        feCreditos.row = i
        If CDbl(RsDatos!nCapPerd) > 0 Then
            feCreditos.Col = 5
            feCreditos.CellForeColor = vbRed
        End If
        feCreditos.TextMatrix(i, 5) = Format(CDbl(RsDatos!nCapPerd), "###," & String(15, "#") & "#0.00")
        If CDbl(RsDatos!nIntPerd) > 0 Then
            feCreditos.Col = 6
            feCreditos.CellForeColor = vbRed
        End If
        feCreditos.TextMatrix(i, 6) = Format(CDbl(RsDatos!nIntPerd), "###," & String(15, "#") & "#0.00")
        If CDbl(RsDatos!nMoraPerd) > 0 Then
            feCreditos.Col = 7
            feCreditos.CellForeColor = vbRed
        End If
        feCreditos.TextMatrix(i, 7) = Format(CDbl(RsDatos!nMoraPerd), "###," & String(15, "#") & "#0.00")
        If CDbl(RsDatos!nGastoPerd) > 0 Then
            feCreditos.Col = 8
            feCreditos.CellForeColor = vbRed
        End If
        feCreditos.TextMatrix(i, 8) = Format(CDbl(RsDatos!nGastoPerd), "###," & String(15, "#") & "#0.00")
        
        'JOEP
        feCreditos.TextMatrix(i, 9) = Format(CDbl(RsDatos!nIcvPerd), "###," & String(15, "#") & "#0.00")
        feCreditos.TextMatrix(i, 10) = Format(CDbl(RsDatos!nMontoDeuda), "###," & String(15, "#") & "#0.00")
        feCreditos.TextMatrix(i, 11) = Format(CDbl(RsDatos!nMontoPerdon), "###," & String(15, "#") & "#0.00")
        feCreditos.TextMatrix(i, 12) = Format(CDbl(RsDatos!nMontoPagar), "###," & String(15, "#") & "#0.00")
        feCreditos.TextMatrix(i, 13) = Format(CDbl(RsDatos!nMontoPend), "###," & String(15, "#") & "#0.00")
        'JOEP
        
        'feCreditos.TextMatrix(i, 9) = Format(CDbl(RsDatos!nMontoDeuda), "###," & String(15, "#") & "#0.00")
        'feCreditos.TextMatrix(i, 10) = Format(CDbl(RsDatos!nMontoPerdon), "###," & String(15, "#") & "#0.00")
        'feCreditos.TextMatrix(i, 11) = Format(CDbl(RsDatos!nMontoPagar), "###," & String(15, "#") & "#0.00")
        'feCreditos.TextMatrix(i, 12) = Format(CDbl(RsDatos!nMontoPend), "###," & String(15, "#") & "#0.00")
        RsDatos.MoveNext
    Next i
Else
    MsgBox "No hay Creditos para su aprobación respectiva", vbInformation, "Aviso"
End If
Set RsDatos = Nothing
Set oDCredito = Nothing
End Sub
Private Sub LimpiarDatos()
cmdAutorizar.Enabled = False
LimpiaFlex feCreditos

feCreditos.Col = 5
feCreditos.CellForeColor = vbBlack
feCreditos.Col = 6
feCreditos.CellForeColor = vbBlack
feCreditos.Col = 7
feCreditos.CellForeColor = vbBlack
feCreditos.Col = 8
feCreditos.CellForeColor = vbBlack
Call ListarCreditos
End Sub
