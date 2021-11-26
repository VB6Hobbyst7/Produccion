VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmIntDevTpoCred 
   Caption         =   "Intereses Diferidos"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9915
   Icon            =   "frmIntDevTpoCred.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8520
      TabIndex        =   11
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdIngresar 
      Caption         =   "Ingresar"
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Left            =   7320
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
      End
      Begin SICMACT.FlexEdit FEID 
         Height          =   3135
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   5530
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-idTipoCre-Idorden-TipoCredito-Orden-Interes Diferido"
         EncabezadosAnchos=   "400-0-0-3500-3500-2000"
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
         BackColorControl=   -2147483628
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-L-L-R"
         FormatosEdit    =   "0-0-0-4-4-4"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.TextBox txtMontoID 
         Height          =   315
         Left            =   5400
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox cboOrden 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   4935
      End
      Begin VB.ComboBox cboTipoCredito 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   4935
      End
      Begin MSMask.MaskEdBox mskPeriodo1Del 
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   1200
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Interes Diferido"
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
         Left            =   3720
         TabIndex        =   8
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Orden"
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
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Crédito"
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
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmIntDevTpoCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fnTipoCambio  As Currency
Dim fdFechaFinMes  As Date
Dim fsServerConsol As String
Dim fsServerRCC As String
Dim fsBDRCC As String
Dim lsTablaTMP As String
Dim nPosi As Integer

Private Sub cmdEliminar_Click()
    If nPosi > 0 Then
        Dim obEval As COMNCredito.NCOMColocEval
        Set obEval = New COMNCredito.NCOMColocEval
        Call obEval.EliminarIntDevTpoCred(FEID.TextMatrix(nPosi, 1), FEID.TextMatrix(nPosi, 2), mskPeriodo1Del.Text)
        MsgBox "Los datos se registraron correctamente"
        Call cmdMostrar_Click
    End If
End Sub

Private Sub cmdIngresar_Click()
Dim obEval As COMNCredito.NCOMColocEval
Set obEval = New COMNCredito.NCOMColocEval
Call obEval.InsertarIntDevTpoCred(Right(cboTipoCredito.Text, 1), Right(cboOrden.Text, 1), mskPeriodo1Del.Text, txtMontoID.Text)
MsgBox "Los datos se registraron correctamente"
Call cmdMostrar_Click
End Sub

Private Sub cmdMostrar_Click()
    Dim obEval As COMNCredito.NCOMColocEval
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set obEval = New COMNCredito.NCOMColocEval
    Set rs = obEval.ObtenerIntDevTpoCred(mskPeriodo1Del.Text)
    Call CargaDatos(rs)
End Sub
Private Sub CargaDatos(ByVal rs As ADODB.Recordset)
Dim MatCalend As Variant

    On Error GoTo ErrorCargaDatos
    LimpiaFlex FEID
'    Set oCredito = Nothing
    If rs.BOF And rs.EOF Then
        MsgBox "No se Encontraron Registros", vbInformation, "Aviso"
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    Do While Not rs.EOF
        FEID.AdicionaFila
        FEID.TextMatrix(rs.Bookmark, 0) = "."
        FEID.TextMatrix(rs.Bookmark, 1) = rs!cTpoCredCod
        FEID.TextMatrix(rs.Bookmark, 2) = rs!nOrden
        FEID.TextMatrix(rs.Bookmark, 3) = rs!cDesTC
        FEID.TextMatrix(rs.Bookmark, 4) = rs!cDesOrden
        FEID.TextMatrix(rs.Bookmark, 5) = Format(rs!nSaldo, "###,###,###0.00")
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Exit Sub

ErrorCargaDatos:
    MsgBox err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub FEID_Click()
    
    nPosi = FEID.Row
    
End Sub

Private Sub Form_Load()
Dim loConstS As COMDConstSistema.NCOMConstSistema
Dim loTipCambio As COMDConstSistema.NCOMTipoCambio
Dim loConstante As COMDConstantes.DCOMConstantes
Dim rsTipCred As ADODB.Recordset
Dim rsOrden As ADODB.Recordset

Set rsTipCred = New ADODB.Recordset
Set rsOrden = New ADODB.Recordset


    Set loConstS = New COMDConstSistema.NCOMConstSistema
        fdFechaFinMes = CDate(loConstS.LeeConstSistema(gConstSistCierreMesNegocio))
        mskPeriodo1Del.Text = fdFechaFinMes
    Set loConstante = New COMDConstantes.DCOMConstantes
    Set rsTipCred = loConstante.RecuperaConstantes(3034)
    
    If Not (rsTipCred.BOF Or rsTipCred.EOF) Then
        Do While Not rsTipCred.EOF
            If Right(rsTipCred!nConsValor, 2) = "50" Then
            cboTipoCredito.AddItem Left(rsTipCred!nConsValor, 1) & " - " & rsTipCred!cConsDescripcion & Space(150) & Left(rsTipCred!nConsValor, 1)
            End If
            rsTipCred.MoveNext
        Loop
    End If
    
    Set rsTipCred = Nothing
    Set loConstante = Nothing
    
    Set loConstante = New COMDConstantes.DCOMConstantes
    Set rsOrden = loConstante.RecuperaConstantes(3039)
    If Not (rsOrden.BOF Or rsOrden.EOF) Then
        Do While Not rsOrden.EOF
            cboOrden.AddItem rsOrden!nConsValor & " - " & rsOrden!cConsDescripcion & Space(150) & rsOrden!nConsValor
            rsOrden.MoveNext
        Loop
    End If
    Set rsOrden = Nothing
    Set loConstante = Nothing
    
End Sub
