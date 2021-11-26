VERSION 5.00
Begin VB.Form frmCredGarantPol 
   Caption         =   "RELACION DE CREDITO - GARANTIA - POLIZAS"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   13305
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "CREDITOS - Solo sugeridos"
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12855
      Begin SICMACT.FlexEdit feCredito 
         Height          =   1755
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   3096
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nº Crédito"
         EncabezadosAnchos=   "400-2200"
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
         ColumnasAEditar =   "X-X"
         ListaControles  =   "0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L"
         FormatosEdit    =   "0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lblCuentaSeleccionada 
         AutoSize        =   -1  'True
         Caption         =   "lblCuentaSeleccionada"
         Height          =   195
         Left            =   3600
         TabIndex        =   2
         Top             =   600
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "CREDITO-GARANTIAS-POLIZA"
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   12975
      Begin SICMACT.FlexEdit feRelaciones 
         Height          =   2115
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   12645
         _ExtentX        =   22304
         _ExtentY        =   3731
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nº Crédito-Nº Garantía-Desc. Garantía-Nº Poliza-Nº Certificado-Tipo Poliza-Prima Total-A Pagar x Cuota"
         EncabezadosAnchos=   "400-2000-1100-3000-1100-1200-1100-1200-1300"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Label lblTotalMontoPrima 
      AutoSize        =   -1  'True
      Caption         =   "lblTotalMontoPrima"
      Height          =   195
      Left            =   7800
      TabIndex        =   3
      Top             =   5160
      Width           =   1350
   End
End
Attribute VB_Name = "frmCredGarantPol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmCredGarantPol
'***     Descripcion:      Visualiza las relaciones de Credito - Garantias y Polizas
'***     Creado por:       WIOR
'***     Maquina:          TIF-1-19
'***     Fecha-Tiempo:     12/03/2012 06:00:00 PM
'***     Ultima Modificacion: Creacion de la Opcion
'*****************************************************************************************
Option Explicit

Private Sub feCredito_Click()
Dim sCtaCod As String
    
    If Mid(lblCuentaSeleccionada.Caption, 1, 2) <> "No" Then
        sCtaCod = feCredito.TextMatrix(feCredito.Row, 1)
        lblCuentaSeleccionada.Caption = "Nº Cuenta Seleccionada: " & sCtaCod
        Call LlenarGrillaRelaciones(sCtaCod)
    End If
End Sub

Private Sub Form_Load()
Call CentraForm(Me)
Me.lblCuentaSeleccionada.Caption = ""
Me.lblTotalMontoPrima.Caption = ""
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub
Public Sub Inicio(ByVal psPersCod As String)
    Call LlenarGrillaCuenta(psPersCod)
    Me.Show 1
End Sub
Public Sub CentraForm(frmCentra As Form)
    frmCentra.Move (Screen.Width - frmCentra.Width) / 2, (Screen.Height - frmCentra.Height) / 2, frmCentra.Width, frmCentra.Height
End Sub
Private Function SumarPrima(ByVal pFg As FlexEdit, ByVal pnTamano As Integer) As Double
Dim nTotal As Double
Dim nConteo As Integer

    If pnTamano > 0 Then
        For nConteo = 1 To pnTamano
          nTotal = nTotal + CDbl(pFg.TextMatrix(nConteo, 8))
        Next
    End If
Me.lblTotalMontoPrima.Caption = "Monto de Pago por Concepto de Poliza(1231): " & nTotal
End Function
Public Sub LlenarGrillaCuenta(ByVal psPersCod As String)
Dim i As Integer
Dim rsCuenta  As ADODB.Recordset
Dim oCredito As COMDCredito.DCOMCredito
Set oCredito = New COMDCredito.DCOMCredito

    Set rsCuenta = oCredito.CargaCuentasPersona(psPersCod, "2001")
    If Not (rsCuenta.BOF And rsCuenta.EOF) Then
        Call LimpiaFlex(feCredito)
        feCredito.lbEditarFlex = True
            For i = 0 To rsCuenta.RecordCount - 1
                    feCredito.AdicionaFila
                    feCredito.TextMatrix(i + 1, 0) = i + 1
                    feCredito.TextMatrix(i + 1, 1) = rsCuenta!cCtaCod
                    rsCuenta.MoveNext
            Next i
        feCredito.lbEditarFlex = False
    Else
        lblCuentaSeleccionada.Caption = "No se encontraron cuentas activas para esta Persona."
        Call LimpiaFlex(feCredito)
    End If
End Sub

Public Sub LlenarGrillaRelaciones(ByVal psCtaCod As String)
Dim i As Integer
Dim rsCuenta  As ADODB.Recordset
Dim oCredito As COMDCredito.DCOMCredito
Set oCredito = New COMDCredito.DCOMCredito

    Set rsCuenta = oCredito.ObtenerRelacionCredGarantPol(psCtaCod)
    
    If Not (rsCuenta.BOF And rsCuenta.EOF) Then
        Call LimpiaFlex(feRelaciones)
        feRelaciones.lbEditarFlex = True
            For i = 0 To rsCuenta.RecordCount - 1
                    feRelaciones.AdicionaFila
                    feRelaciones.TextMatrix(i + 1, 0) = i + 1
                    feRelaciones.TextMatrix(i + 1, 1) = rsCuenta!cCtaCod
                    feRelaciones.TextMatrix(i + 1, 2) = rsCuenta!cNumGarant
                    feRelaciones.TextMatrix(i + 1, 3) = rsCuenta!cDescripcion
                    feRelaciones.TextMatrix(i + 1, 4) = rsCuenta!cnumpoliza
                    feRelaciones.TextMatrix(i + 1, 5) = rsCuenta!cCodPolizaAseg
                    feRelaciones.TextMatrix(i + 1, 6) = rsCuenta!cConsDescripcion
                    feRelaciones.TextMatrix(i + 1, 7) = rsCuenta!Total
                    feRelaciones.TextMatrix(i + 1, 8) = rsCuenta!MontoPagar
                    rsCuenta.MoveNext
            Next i
        feRelaciones.lbEditarFlex = False
        Call SumarPrima(feRelaciones, rsCuenta.RecordCount)
    Else
        Call LimpiaFlex(feRelaciones)
        MsgBox "No se encuentraron relacion de sus garantias y/o polizas.", vbInformation, " Aviso "
        Me.lblTotalMontoPrima.Caption = ""
    End If
End Sub


