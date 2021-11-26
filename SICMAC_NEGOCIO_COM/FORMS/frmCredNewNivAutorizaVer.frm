VERSION 5.00
Begin VB.Form frmCredNewNivAutorizaVer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorizaciones requeridas del crédito"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   Icon            =   "frmCredNewNivAutorizaVer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   8160
      TabIndex        =   15
      Top             =   8280
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Exoneraciones no contempladas"
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
      Height          =   1935
      Left            =   40
      TabIndex        =   6
      Top             =   6315
      Width           =   10335
      Begin SICMACT.FlexEdit FEExoneraciones 
         Height          =   1575
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   10095
         _extentx        =   17806
         _extenty        =   2778
         cols0           =   9
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "Nº-Exoneración-Descripción Exoneración-Nivel Aprobación-Estado-nId-nItem-cNivAprCod-nEstado"
         encabezadosanchos=   "400-2400-3900-2000-1300-0-0-0-0"
         font            =   "frmCredNewNivAutorizaVer.frx":030A
         fontfixed       =   "frmCredNewNivAutorizaVer.frx":0336
         columnasaeditar =   "X-X-X-X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0-0-0-0"
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         encabezadosalineacion=   "C-L-L-C-C-C-C-C-C"
         formatosedit    =   "0-0-0-0-0-0-0-0-0"
         textarray0      =   "Nº"
         lbultimainstancia=   -1  'True
         colwidth0       =   405
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
   End
   Begin VB.Frame frmDatos 
      Caption         =   "Datos"
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
      Height          =   1035
      Left            =   40
      TabIndex        =   2
      Top             =   0
      Width           =   10335
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   435
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   3660
         _extentx        =   6456
         _extenty        =   767
         texto           =   "   Crédito"
         enabledcmac     =   -1  'True
         enabledcta      =   -1  'True
         enabledprod     =   -1  'True
         enabledage      =   -1  'True
      End
      Begin VB.Label lblPeriodo 
         Caption         =   "Periodo:"
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
         Left            =   7200
         TabIndex        =   17
         Top             =   700
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   1080
         TabIndex        =   16
         Top             =   700
         Width           =   5415
      End
      Begin VB.Label lblTasa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   9375
         TabIndex        =   14
         Top             =   700
         Width           =   735
      End
      Begin VB.Label lblCuotas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   8085
         TabIndex        =   13
         Top             =   700
         Width           =   495
      End
      Begin VB.Label lblMoneda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   8085
         TabIndex        =   12
         Top             =   360
         Width           =   465
      End
      Begin VB.Label lblMontoSol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   8535
         TabIndex        =   11
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label Label2 
         Caption         =   "Monto Crédito:"
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
         Left            =   6720
         TabIndex        =   10
         Top             =   405
         Width           =   1335
      End
      Begin VB.Label lblNroCuotas 
         Caption         =   "Nº Cuotas:"
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
         Left            =   7035
         TabIndex        =   9
         Top             =   700
         Width           =   975
      End
      Begin VB.Label lbltas 
         Caption         =   "Tasa:"
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
         Left            =   8790
         TabIndex        =   8
         Top             =   700
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
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
         Top             =   705
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Autorizaciones"
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
      Height          =   5220
      Left            =   40
      TabIndex        =   1
      Top             =   1040
      Width           =   10335
      Begin SICMACT.FlexEdit FEAutorizaciones 
         Height          =   4935
         Left            =   45
         TabIndex        =   4
         Top             =   240
         Width           =   10215
         _extentx        =   18018
         _extenty        =   8705
         cols0           =   8
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "Nº-Descripción-Nivel Aprobación-Estado-nId-nEstado-cAutorizaID-cNivAprCod"
         encabezadosanchos=   "400-6400-2000-1300-0-0-0-0"
         font            =   "frmCredNewNivAutorizaVer.frx":0364
         fontfixed       =   "frmCredNewNivAutorizaVer.frx":0390
         columnasaeditar =   "X-X-X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0-0-0"
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         encabezadosalineacion=   "C-L-L-C-C-C-C-C"
         formatosedit    =   "0-0-0-0-0-0-0-0"
         textarray0      =   "Nº"
         lbultimainstancia=   -1  'True
         colwidth0       =   405
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   9240
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   8280
      Width           =   1000
   End
End
Attribute VB_Name = "frmCredNewNivAutorizaVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'** Nombre : frmCredNewNivAutorizaVer
'** Descripción : Formulario para ver las autorizaciones de créditos pendientes según ERS002-2016
'** Creación : EJVG, 20160205 07:14:00 PM
'************************************************************************************************
Option Explicit
Dim fnTipoAutorizacion As Integer
Dim fnTipoExoneracion As Integer
Dim fnTipoDatosCredito As Integer
Public Sub Inicio()
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    Show 1
End Sub
Public Sub Consultar(ByVal psCtaCod As String)
    If Not CargaDatos(psCtaCod) Then
        Exit Sub
    End If
    Show 1
End Sub
Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
    Dim oDNiv As New COMDCredito.DCOMNivelAprobacion
    Dim lnFila As Integer
    Dim rsCred As New ADODB.Recordset
    Dim lsTipoProd As String
    
    fnTipoAutorizacion = 1
    fnTipoExoneracion = 2
    fnTipoDatosCredito = 3
On Error GoTo ErrCargaControles
    ActxCta.NroCuenta = psCtaCod
    CargaDatos = False
    Set rsCred = oDNiv.HistAutorizacionesExoneraciones(psCtaCod, fnTipoDatosCredito)
        If Not (rsCred.BOF And rsCred.EOF) Then
            lblMoneda.Caption = IIf(Mid(psCtaCod, 9, 1) = "1", "S/", "$")
            lblMontoSol.Caption = Format(rsCred!nMontoCol, "#,##0.00")
            lblNombre.Caption = rsCred!cPersNombre
            lblTasa.Caption = Format(rsCred!nTasaInteres, "#,##0.0000")
            lblCuotas.Caption = rsCred!nCuotas
            'LUCV20171212, Según Observacion SBS
            lsTipoProd = rsCred!cTpoProdCod
            If (lsTipoProd = "514") Then
                lblPeriodo.Visible = True
                lblNroCuotas.Visible = False
                lbltas.Visible = False
                lblTasa.Visible = False
            Else
                lblPeriodo.Visible = False
                lblNroCuotas.Visible = True
                lbltas.Visible = True
                lblTasa.Visible = True
            End If
            'Fin LUCV20171212
            ActxCta.Enabled = False
        Else
            MsgBox "No se pudo Encontrar el Credito", vbInformation, "Aviso"
            Exit Function
        End If
    Set rsCred = Nothing
    
    Set rsCred = oDNiv.HistAutorizacionesExoneraciones(psCtaCod, fnTipoAutorizacion)
        lnFila = 1
        If Not (rsCred.BOF And rsCred.EOF) Then
            Do While Not (rsCred.EOF)
                FEAutorizaciones.AdicionaFila
                FEAutorizaciones.TextMatrix(lnFila, 1) = rsCred!cAutorizaDesc
                If rsCred!nId <> -1 Then
                    FEAutorizaciones.TextMatrix(lnFila, 2) = rsCred!cNivelAprobacion
                    FEAutorizaciones.TextMatrix(lnFila, 3) = rsCred!cDescripcionEstado
                    If rsCred!nEstado = 0 Then
                        FEAutorizaciones.BackColorRow (&HC0C0FF)
                    Else
                        FEAutorizaciones.BackColorRow (&HC0FFC0)
                    End If
                    FEAutorizaciones.TextMatrix(lnFila, 4) = rsCred!nId
                    FEAutorizaciones.TextMatrix(lnFila, 5) = rsCred!nEstado
                    FEAutorizaciones.TextMatrix(lnFila, 6) = rsCred!cAutorizaID
                    FEAutorizaciones.TextMatrix(lnFila, 7) = rsCred!cNivAprCod
                End If
                lnFila = lnFila + 1
                rsCred.MoveNext
            Loop
        End If
    Set rsCred = Nothing
    
    Set rsCred = oDNiv.HistAutorizacionesExoneraciones(psCtaCod, fnTipoExoneracion)
        lnFila = 1
        If Not (rsCred.BOF And rsCred.EOF) Then
            Do While Not (rsCred.EOF)
                FEExoneraciones.AdicionaFila
                FEExoneraciones.TextMatrix(lnFila, 1) = rsCred!cExoneracion
                FEExoneraciones.TextMatrix(lnFila, 2) = rsCred!cDescripcion
                FEExoneraciones.TextMatrix(lnFila, 3) = rsCred!cNivelAprobacion
                FEExoneraciones.TextMatrix(lnFila, 4) = rsCred!cDescEstado
                If rsCred!nEstado = 0 Then
                    FEExoneraciones.BackColorRow (&HC0C0FF)
                Else
                    FEExoneraciones.BackColorRow (&HC0FFC0)
                End If
                FEExoneraciones.TextMatrix(lnFila, 5) = rsCred!nId
                FEExoneraciones.TextMatrix(lnFila, 6) = rsCred!nItem
                FEExoneraciones.TextMatrix(lnFila, 7) = rsCred!cNivAprCod
                FEExoneraciones.TextMatrix(lnFila, 8) = rsCred!nEstado
                lnFila = lnFila + 1
                rsCred.MoveNext
            Loop
        End If
    Set rsCred = Nothing
    
    CargaDatos = True
    Exit Function
ErrCargaControles:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not CargaDatos(ActxCta.NroCuenta) Then
            MsgBox "No se pudo Encontrar el Credito", vbInformation, "Aviso"
            'HabilitaIngreso False
        Else
            'HabilitaIngreso True
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub cmdCancelar_Click()
    ActxCta.Enabled = True
    ActxCta.Prod = ""
    ActxCta.Cuenta = ""
    lblMontoSol.Caption = ""
    lblMoneda.Caption = "S/"
    lblCuotas.Caption = ""
    lblTasa.Caption = ""
    lblNombre.Caption = ""
    Call FormateaFlex(FEAutorizaciones)
    Call FormateaFlex(FEExoneraciones)
End Sub

