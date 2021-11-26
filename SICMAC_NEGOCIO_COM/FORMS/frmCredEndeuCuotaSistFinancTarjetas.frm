VERSION 5.00
Begin VB.Form frmCredEndeuCuotaSistFinancTarjetas 
   Caption         =   "VB por deuda potencial"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   15255
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmCredEndeuCuotaSistFinancTarjetas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   15255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   495
      Left            =   4680
      TabIndex        =   18
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Sistema Financiero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   15180
      Begin VB.TextBox txtDeudaIFI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   11400
         TabIndex        =   13
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   495
         Left            =   1080
         TabIndex        =   15
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox txtCuotaL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   11400
         TabIndex        =   14
         Top             =   3600
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   495
         Left            =   2880
         TabIndex        =   12
         Top             =   3480
         Width           =   1695
      End
      Begin SICMACT.FlexEdit FEDeudaSF 
         Height          =   3135
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5530
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-cod_IFI-OK-IFI-Relación crédito-Tipo Deuda IFI-Deuda IFI-Cuota Est."
         EncabezadosAnchos=   "400-0-500-5000-2500-2000-2000-2000"
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-X-X-X-X-X"
         ListaControles  =   "0-0-4-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-L-L-L-L-R-R"
         FormatosEdit    =   "0-0-0-4-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label3 
         Caption         =   "Deuda IFIS:"
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
         Left            =   10080
         TabIndex        =   16
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Cuota esperada:"
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
         Left            =   9960
         TabIndex        =   17
         Top             =   3720
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15180
      Begin VB.CheckBox ChkCompra 
         Caption         =   "Compra de Deuda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   3960
         TabIndex        =   19
         Top             =   1080
         Width           =   3030
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9240
         TabIndex        =   1
         Top             =   240
         Width           =   1155
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   435
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   767
         Texto           =   "Crédito:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label LblPersCod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1440
         TabIndex        =   8
         Top             =   255
         Width           =   1755
      End
      Begin VB.Label lblDocJur 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3360
         TabIndex        =   7
         Top             =   600
         Width           =   2130
      End
      Begin VB.Label lblDocNat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1440
         TabIndex        =   6
         Top             =   630
         Width           =   1035
      End
      Begin VB.Label lblNomPers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3360
         TabIndex        =   5
         Top             =   240
         Width           =   5700
      End
      Begin VB.Label lblDocJuridico 
         AutoSize        =   -1  'True
         Caption         =   "RUC :"
         Height          =   195
         Left            =   2760
         TabIndex        =   4
         Top             =   720
         Width           =   435
      End
      Begin VB.Label lblDocNatural 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Identidad :"
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   660
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   315
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmCredEndeuCuotaSistFinancTarjetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ALPA 20160516
'Obsrevación de la SBS por Sobreendeudamiento
Option Explicit
Dim lsCtaCod As String
Dim lsPersCod As String
Dim fbAceptar As Boolean
Dim fbSalir As Boolean
Dim fbFocoGrilla As Boolean
Dim fbCheckGrilla As Boolean
Dim lnCuotaL As Currency
'Public Sub Inicio(ByVal psCtaCod As String, ByVal psPersCod As String)
'    lsCtaCod = psCtaCod
'    lsPersCod = psPersCod
'    Me.Show 1
'End Sub
Private Sub ActxCta_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    
    Call cargarDatos
 End If
End Sub


Private Sub cmdBuscar_Click()
Dim oPersona As COMDPersona.UCOMPersona
Dim sPersCod As String
Dim oCliPre As COMNCredito.NCOMCredito
Set oCliPre = New COMNCredito.NCOMCredito
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Set R = New ADODB.Recordset

    cmdCancelar_Click

    Set oPersona = frmBuscaPersona.Inicio
    If Not oPersona Is Nothing Then
        LblPersCod.Caption = oPersona.sPersCod
        lblNomPers.Caption = oPersona.sPersNombre
        lblDocNat.Caption = Trim(oPersona.sPersIdnroDNI)
        lblDocJur.Caption = Trim(oPersona.sPersIdnroRUC)
        If oPersona.sPersPersoneria = "1" Then

            If Trim(oPersona.sPersIdnroDNI) = "" Then
                If Not Trim(oPersona.sPersIdnroOtro) = "" Then
                    lblDocNat.Caption = Trim(oPersona.sPersIdnroOtro)
                    'lsPersTDoc = Trim(oPersona.sPersTipoDoc)
                End If
            End If
            
            If ChkCompra = vbUnchecked Then 'ARLO20180323
                Set oCredito = New COMDCredito.DCOMCredito
                Set R = oCredito.RecuperaCreditosVigentes(oPersona.sPersCod, , Array(gColocEstSolic, gColocEstSug))
                
                If Not R.EOF Then
                    Do While Not R.EOF
                        ActxCta.NroCuenta = R!cCtaCod
                        R.MoveNext
                    Loop
                Else
                    MsgBox "El cliente no tiene créditos SOLICITADOS y/o SUGERIDOS", vbInformation, "Aviso"
                    cmdCancelar_Click
                End If
                R.Close
                Set R = Nothing
                Set oCredito = Nothing
            End If 'ARLO20180323
            Call cargarDatos
        Else
            'fbPersNatural = False
            'lsPersTDoc = "3"
        End If
    Else
        Exit Sub
    End If
End Sub
Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub CmdGrabar_Click()
    Dim oCredFor As New COMDCredito.DCOMFormatosEval
    Dim oCredForN As New COMNCredito.NCOMFormatosEval
    Dim rsCredFormEval As New ADODB.Recordset 'LUCV20160820, Según ERS004-2016
    Dim i As Integer
    Dim Prev As previo.clsprevio
    Set Prev = New previo.clsprevio
    Dim oCredDoc As COMNCredito.NCOMCredDoc
    Set oCredDoc = New COMNCredito.NCOMCredDoc
    Dim objCredito As COMDCredito.DCOMCredito
    Set objCredito = New COMDCredito.DCOMCredito
    
    Dim rsCredito As New ADODB.Recordset 'CTI3 ERS0032020
    
    Set rsCredFormEval = oCredFor.RecuperaFormatoEvaluacion(ActxCta.NroCuenta) 'LUCV20160820, Según ERS004-2016
    
    Set rsCredito = objCredito.ObtenerDatosCreditoxVB(ActxCta.NroCuenta) 'CTI3 ERS0032020

    If FEDeudaSF.TextMatrix(1, 0) = "" Then ' FEDeudaSF.Rows - 1 < 1 Then
        MsgBox "No existen registros de linea de creditos para imprimir", vbInformation, "Aviso!"
        Exit Sub
    End If
    For i = 1 To FEDeudaSF.rows - 1
        Call objCredito.ActualizaCredDeudaSistemaFinanciero(IIf(FEDeudaSF.TextMatrix(i, 2) = "", 0, 1), gdFecSis, ActxCta.NroCuenta, FEDeudaSF.TextMatrix(i, 1), FEDeudaSF.TextMatrix(i, 5), FEDeudaSF.TextMatrix(i, 6), FEDeudaSF.TextMatrix(i, 4))
    Next i
    
    MsgBox "Los datos se guardaron correctamente", vbInformation, "Aviso" 'ARLO20180328
    
    If (Len(ActxCta.NroCuenta) = 18) Then 'ARLO20180328
    'CTI3 ERS0032020
    Prev.Show oCredDoc.ImprimeCredDeudaSistemaFinanciero(LblPersCod.Caption, ActxCta.NroCuenta, lblNomPers.Caption, "", gdFecSis, 0, "CMAC Maynas", 0, "109", gsCodUser), "", True
    Set oCredDoc = Nothing
    Set Prev = Nothing
    End If
    
    'ERS0032020**********************************************
    If Not (rsCredFormEval.EOF Or rsCredFormEval.BOF) Then
        If Not (rsCredito.EOF Or rsCredito.BOF) Then
            If rsCredito!nPrdEstado = 2001 Then
                '********************************************
                Call objCredito.ActualizarEstadoxVB(ActxCta.NroCuenta, 0)
                If Len(ActxCta.NroCuenta) <> "18" Then
                    Exit Sub
                End If

                Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
                Dim nEstado As Integer
                Dim rs As ADODB.Recordset
                Dim oRS As New ADODB.Recordset
                Dim nFormEmpr As Boolean
                Dim nProducto As String
             
                Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
                Set rs = oDCOMFormatosEval.RecuperaCredFormEvalProductoEstadoExposicion(ActxCta.NroCuenta)
                nEstado = IIf(IsNull(rs!nPrdEstado), 0, rs!nPrdEstado)
                Set oRS = oDCOMFormatosEval.RecuperaFormatoEvaluacion(ActxCta.NroCuenta)
                If (oRS.EOF And oRS.BOF) Then
                    nProducto = Mid(ActxCta.Prod, 1, 1) & "00"
                End If
                nFormEmpr = False
            
                Call EvaluarCredito(ActxCta.NroCuenta, False, nEstado, CInt(Mid(rsCredito!cTpoProdCod, 1, 1) & "00"), CInt(rsCredito!cTpoProdCod), rsCredito!nExposicion, False, , nFormEmpr, , , True)

                '********************************************
            End If
        End If
    End If
    '********************************************************
    Exit Sub
End Sub

Private Sub FEDeudaSF_DblClick()
    Dim i As Integer
    Dim objCredito As COMDCredito.DCOMCredito
    Set objCredito = New COMDCredito.DCOMCredito
    Dim ldFechaFinDeMes  As Date
    If FEDeudaSF.TextMatrix(FEDeudaSF.row, 2) = "." Then
        FEDeudaSF.TextMatrix(FEDeudaSF.row, 2) = 0
    Else
        FEDeudaSF.TextMatrix(FEDeudaSF.row, 2) = 1
    End If
            lnCuotaL = 0
            For i = 1 To FEDeudaSF.rows - 1
            If FEDeudaSF.TextMatrix(i, 2) = "." Then
                 lnCuotaL = lnCuotaL + FEDeudaSF.TextMatrix(i, 6)
            End If
            Next i
            txtDeudaIFI.Text = Format(lnCuotaL, gsFormatoNumeroView)
  
End Sub

Private Sub FEDeudaSF_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
   Dim iGP As Integer
    Dim i As Integer
    Dim objCredito As COMDCredito.DCOMCredito
    Set objCredito = New COMDCredito.DCOMCredito
    Dim ldFechaFinDeMes  As Date
    If FEDeudaSF.TextMatrix(pnRow, 0) <> "" Then
        If pnCol = 2 Then
            lnCuotaL = 0
            For i = 1 To FEDeudaSF.rows - 1
            If FEDeudaSF.TextMatrix(i, 2) = "." Then
                 lnCuotaL = lnCuotaL + FEDeudaSF.TextMatrix(i, 6)
            End If
            Next i
            txtDeudaIFI.Text = Format(lnCuotaL, gsFormatoNumeroView)
        End If
    End If
End Sub

Private Sub Form_Load()
    Call InicializaControles
End Sub
Private Sub InicializaControles()
    LblPersCod.Caption = ""
    lblNomPers.Caption = ""
    lblDocNat.Caption = ""
    lblDocJur.Caption = ""
    ActxCta.Cuenta = ""
    ActxCta.Prod = ""
    LimpiaFlex FEDeudaSF
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    ActxCta.Enabled = False
    Frame1.Enabled = True
    Frame2.Enabled = False
    'txtCuotaL.Text = ""
    txtDeudaIFI.Text = ""
End Sub

Private Sub cmdCancelar_Click()
    Call InicializaControles
End Sub
Private Sub cargarDatos()
Dim objCredito As COMDCredito.DCOMCredito
Set objCredito = New COMDCredito.DCOMCredito
Dim i As Integer
Dim oRS As ADODB.Recordset
Set oRS = New ADODB.Recordset
lnCuotaL = 0

Set oRS = objCredito.RecuperaCredDeudaSistemaFinanciero(LblPersCod.Caption, ActxCta.NroCuenta)
LimpiaFlex FEDeudaSF
    If Not oRS.EOF Then
        Do While Not oRS.EOF
            FEDeudaSF.AdicionaFila
            FEDeudaSF.TextMatrix(oRS.Bookmark, 1) = IIf(IsNull(oRS!Cod_Emp), 0, oRS!Cod_Emp)
            FEDeudaSF.TextMatrix(oRS.Bookmark, 2) = IIf(IsNull(oRS!bActivo), 0, oRS!bActivo)
            FEDeudaSF.TextMatrix(oRS.Bookmark, 3) = IIf(IsNull(oRS!Nombre), 0, oRS!Nombre)
            FEDeudaSF.TextMatrix(oRS.Bookmark, 5) = IIf(IsNull(oRS!cTipoDeuda), 0, oRS!cTipoDeuda)
            FEDeudaSF.TextMatrix(oRS.Bookmark, 6) = Format(IIf(IsNull(oRS!nSaldoDeuda), 0, oRS!nSaldoDeuda), gsFormatoNumeroView)
            If IIf(IsNull(oRS!bActivo), 0, oRS!bActivo) = 1 Then
                 lnCuotaL = lnCuotaL + IIf(IsNull(oRS!nSaldoDeuda), 0, oRS!nSaldoDeuda)
            End If
            FEDeudaSF.TextMatrix(oRS.Bookmark, 4) = IIf(IsNull(oRS!cTit_titular), 0, oRS!cTit_titular)
            FEDeudaSF.TextMatrix(oRS.Bookmark, 7) = IIf(IsNull(oRS!Cuota_CE), 0, oRS!Cuota_CE)
            oRS.MoveNext
        Loop
        Frame1.Enabled = False
        Frame2.Enabled = True
    Else
        If ChkCompra = vbChecked Then 'ARLO20180323
        MsgBox "El cliente no cuenta con IFIS", vbInformation, "Aviso" 'ARLO20180328
        End If
    End If
    txtDeudaIFI.Text = Format(lnCuotaL, gsFormatoNumeroView)
End Sub
Private Sub FEDeudaSF_GotFocus()
    fbFocoGrilla = True
End Sub
Private Sub FEDeudaSF_LostFocus()
    fbFocoGrilla = False
End Sub

