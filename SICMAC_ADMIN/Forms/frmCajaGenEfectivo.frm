VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajaGenEfectivo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A Rendir en Efectivo"
   ClientHeight    =   5340
   ClientLeft      =   750
   ClientTop       =   2220
   ClientWidth     =   9540
   Icon            =   "frmCajaGenEfectivo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   Begin Sicmact.Usuario oUser 
      Left            =   3225
      Top             =   4860
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Frame FraMontosol 
      Height          =   540
      Left            =   60
      TabIndex        =   13
      Top             =   4725
      Width           =   2655
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto :"
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
         Height          =   225
         Left            =   90
         TabIndex        =   15
         Top             =   210
         Width           =   660
      End
      Begin VB.Label lblMontoSol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
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
         Height          =   315
         Left            =   870
         TabIndex        =   14
         Top             =   165
         Width           =   1605
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7890
      TabIndex        =   5
      Top             =   4875
      Width           =   1380
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6465
      TabIndex        =   4
      Top             =   4875
      Width           =   1380
   End
   Begin VB.Frame fraBilletajes 
      Caption         =   "Ingreso de Descomposición de Efectivo"
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
      Height          =   3900
      Left            =   60
      TabIndex        =   6
      Top             =   825
      Width           =   9360
      Begin Sicmact.FlexEdit fgBilletes 
         Height          =   2610
         Left            =   165
         TabIndex        =   2
         Top             =   315
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   4604
         Cols0           =   6
         FixedCols       =   2
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-Descripción-Cantidad-Monto-cEfectivoCod-nEfectivoValor"
         EncabezadosAnchos=   "350-2000-800-1200-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ColumnasAEditar =   "X-X-2-3-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-R-R-C-C"
         FormatosEdit    =   "0-0-3-4-0-0"
         AvanceCeldas    =   1
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   345
         RowHeight0      =   285
         CellBackColor   =   -2147483633
      End
      Begin Sicmact.FlexEdit fgMonedas 
         Height          =   2610
         Left            =   4725
         TabIndex        =   3
         Top             =   315
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   4604
         Cols0           =   6
         FixedCols       =   2
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-Descripción-Cantidad-Monto-cEfectivoCod-nEfectivoValor"
         EncabezadosAnchos=   "350-2000-800-1200-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ColumnasAEditar =   "X-X-2-3-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-R-R-C-C"
         FormatosEdit    =   "0-0-3-4-0-0"
         AvanceCeldas    =   1
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   345
         RowHeight0      =   285
         CellBackColor   =   -2147483633
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL BILLETES :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   150
         Left            =   1245
         TabIndex        =   12
         Top             =   3037
         Width           =   1410
      End
      Begin VB.Label lblTotalBilletes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   2685
         TabIndex        =   11
         Top             =   2962
         Width           =   1965
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   6300
         TabIndex        =   10
         Top             =   3510
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   7260
         TabIndex        =   9
         Top             =   3450
         Width           =   1965
      End
      Begin VB.Label lblTotMoneda 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   7245
         TabIndex        =   8
         Top             =   2955
         Width           =   1965
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL MONEDAS :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   165
         Left            =   5775
         TabIndex        =   7
         Top             =   3030
         Width           =   1440
      End
      Begin VB.Shape ShapeS 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   1140
         Top             =   2940
         Width           =   3525
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   5700
         Top             =   2940
         Width           =   3525
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Left            =   6120
         Top             =   3435
         Width           =   3120
      End
   End
   Begin VB.Frame fraDatosPrinc 
      Height          =   690
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   9360
      Begin Sicmact.TxtBuscar txtBuscarUser 
         Height          =   360
         Left            =   840
         TabIndex        =   17
         Top             =   195
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   8388608
      End
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   345
         Left            =   7770
         TabIndex        =   1
         Top             =   225
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblCaptionUser 
         AutoSize        =   -1  'True
         Caption         =   "Cajero :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   135
         TabIndex        =   19
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblDescUser 
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   2205
         TabIndex        =   18
         Top             =   195
         Width           =   4695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   7155
         TabIndex        =   16
         Top             =   255
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmCajaGenEfectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vbOk As Boolean
Dim rsBilletesAux As ADODB.Recordset
Dim rsMonedasAux As ADODB.Recordset
Dim lnMoneda As Moneda
Dim lnMonto As Currency
Dim lnArendirFase As ARendirFases
Dim lnDiferencia As Currency
Dim lbDiferencia As Boolean
Dim lsCaptionArendir As String
Dim lbEnableFecha As Boolean
Dim lbMuestra As Boolean
Dim lbRegistro As Boolean
Dim lsMovNro As String
Dim lnMovNro As Long
Dim lbModifica As Boolean
Dim lbSalir As Boolean
Dim oCajero As nCajero
Dim lnParamDif As Currency
Private Sub CmdAceptar_Click()
lnDiferencia = 0
If FraMontosol.Visible Then
    If CCur(lblTotal) <> CCur(lblMontoSol) Then
        If lbDiferencia Then
            If Abs(CCur(lblTotal) - CCur(lblMontoSol)) < lnParamDif Then
                If MsgBox("Monto Ingresado no cubre el Monto de Solicitado" & Chr(13) & "Existe una diferencia la cual no ha sido cubierta" & Chr(13) & "Desea continuar???", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                    fgBilletes.SetFocus
                    Exit Sub
                Else
                    lnDiferencia = CCur(lblMontoSol) - CCur(Abs(lblTotal))
                End If
            Else
                MsgBox "Monto Ingresado no cubre el Monto de Arendir", vbInformation, "Aviso"
                fgBilletes.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Monto Ingresado no cubre el Monto Solicitado", vbInformation, "Aviso"
            fgBilletes.SetFocus
            Exit Sub
        End If
    End If
Else
    If Val(lblTotal) = 0 Then
        MsgBox "Por favor Ingrese alguna denominación", vbInformation, "Aviso"
        Exit Sub
    End If
End If
Set rsBilletesAux = fgBilletes.GetRsNew
Set rsMonedasAux = fgMonedas.GetRsNew
If lbRegistro Then
    If lbModifica Then
        Dim oCont As NContFunciones
        Dim lbNuevo  As Boolean
        Set oCont = New NContFunciones
        If MsgBox("Desea Registrar el billetaje Ingresado por " & IIf(Mid(gsopecod, 3, 1) = gMonedaNacional, gcPEN_SIMBOLO, "$.") & lblTotal, vbYesNo + vbQuestion, "Aviso") = vbYes Then 'marg ers044-2016
            vbOk = True
            If lsMovNro <> "" Then
                lbNuevo = False
                'If MsgBox("Billetaje ya ha sido registro este día. Desea Proseguir??", vbYesNo + vbExclamation, "Aviso") = vbNo Then Exit Sub
            Else
                lbNuevo = True
                lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            End If
            If oCajero.GrabaRegistroEfectivo(gsFormatoFecha, lsMovNro, _
                            gsopecod, gsOpeDesc, rsBilletesAux, rsMonedasAux, txtBuscarUser, lbNuevo) = 0 Then
            
            End If
            fgBilletes.SetFocus
        End If
        Set oCont = Nothing
    Else
        vbOk = True
        Me.Hide
    End If
Else
    vbOk = True
    Me.Hide
End If
End Sub
Private Sub cmdCancelar_Click()
vbOk = False
If lbRegistro Then
    Unload Me
    Set frmCajaGenEfectivo = Nothing
Else
    Me.Hide
End If
End Sub
Public Property Get lbOk() As Variant
    lbOk = vbOk
End Property
Public Property Let lbOk(ByVal vNewValue As Variant)
    vbOk = vNewValue
End Property
Public Property Get rsBilletes() As ADODB.Recordset
    Set rsBilletes = rsBilletesAux
End Property
Public Property Get rsMonedas() As ADODB.Recordset
    Set rsMonedas = rsMonedasAux
End Property
Public Property Set rsBilletes(ByVal vNewValue As ADODB.Recordset)
    Set rsBilletes = vNewValue
End Property
Public Property Set rsMonedas(ByVal vNewValue As ADODB.Recordset)
    Set rsMonedas = vNewValue
End Property

Private Sub fgBilletes_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim lnTotal As Currency
Dim lnValor As Currency
Select Case pnCol
    Case 3
        lnValor = CCur(IIf(fgBilletes.TextMatrix(pnRow, pnCol) = "", "0", fgBilletes.TextMatrix(pnRow, pnCol)))
        If Residuo(lnValor, fgBilletes.TextMatrix(pnRow, 5)) Then
            fgBilletes.TextMatrix(pnRow, 2) = Format(Round(lnValor / fgBilletes.TextMatrix(pnRow, 5), 0), "#,#0")
        Else
            Cancel = False
            Exit Sub
        End If
    Case 2
        lnValor = CCur(IIf(fgBilletes.TextMatrix(pnRow, pnCol) = "", "0", fgBilletes.TextMatrix(pnRow, pnCol)))
        fgBilletes.TextMatrix(pnRow, 3) = Format(lnValor * CCur(IIf(fgBilletes.TextMatrix(pnRow, 5) = "", "0", fgBilletes.TextMatrix(pnRow, 5))), "#,#0.00")
End Select
lblTotalBilletes = Format(fgBilletes.SumaRow(3), "#,####0.00")
lnTotal = Format(CCur(lblTotalBilletes) + CCur(lblTotMoneda), "#,####0.00")
If FraMontosol.Visible Then
    If lnTotal > Abs(lnMonto) Then
        MsgBox "Total no sebe Superar lo solicitado", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
End If
lblTotal = Format(lnTotal, "#,####0.00")
End Sub
Private Sub fgMonedas_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim lnTotal As Currency
Dim lnValor As Currency
Select Case pnCol
    Case 3
        lnValor = CCur(IIf(fgMonedas.TextMatrix(pnRow, pnCol) = "", "0", fgMonedas.TextMatrix(pnRow, pnCol)))
        If Residuo(lnValor, fgMonedas.TextMatrix(pnRow, 5)) Then
            fgMonedas.TextMatrix(pnRow, 2) = Format(Round(lnValor / fgMonedas.TextMatrix(pnRow, 5), 0), "#,#0")
        Else
            Cancel = False
            Exit Sub
        End If
    Case 2
        lnValor = CCur(IIf(fgMonedas.TextMatrix(pnRow, pnCol) = "", "0", fgMonedas.TextMatrix(pnRow, pnCol)))
        fgMonedas.TextMatrix(pnRow, 3) = Format(lnValor * CCur(IIf(fgMonedas.TextMatrix(pnRow, 5) = "", "0", fgMonedas.TextMatrix(pnRow, 5))), "#,#0.00")
End Select
lblTotMoneda = Format(fgMonedas.SumaRow(3), "#,####0.00")
lnTotal = Format(CCur(lblTotalBilletes) + CCur(lblTotMoneda), "#,####0.00")
If FraMontosol.Visible Then
    If lnTotal > Abs(lnMonto) Then
        MsgBox "Total no sebe Superar lo solicitado", vbInformation, "Aviso"
        Cancel = False
        Exit Sub
    End If
End If
lblTotal = Format(lnTotal, "#,####0.00")

End Sub

Private Sub Form_Activate()
If lbSalir Then
    Unload Me
End If
End Sub
Private Sub Form_Load()
Dim oSubasta As DSubasta
Set oSubasta = New DSubasta

Set oCajero = New nCajero
Dim oGen As DGeneral
Set oGen = New DGeneral

lnParamDif = oGen.GetParametro(4000, 1001)
Set oGen = Nothing
CentraForm Me
lbSalir = False
CargaBilletajes
Me.Caption = gsOpeDesc
lblTotalBilletes.BackColor = IIf(lnMoneda = gMonedaNacional, vbWhite, &HC0FFC0)
lblTotMoneda.BackColor = IIf(lnMoneda = gMonedaNacional, vbWhite, &HC0FFC0)
lblTotal.BackColor = IIf(lnMoneda = gMonedaNacional, vbWhite, &HC0FFC0)
If lnMonto = 0 Then
    FraMontosol.Visible = False
End If
If lbMuestra = False Then
    txtfecha = gdFecSis
    txtfecha.Enabled = lbEnableFecha
Else
    txtfecha = Mid(lsMovNro, 7, 2) & "/" & Mid(lsMovNro, 5, 2) & "/" & Left(lsMovNro, 4)
    txtfecha.Enabled = False
    fgBilletes.lbEditarFlex = False
    fgMonedas.lbEditarFlex = False
End If
lblCaptionUser.Visible = False
txtBuscarUser.Visible = False
lblDescUser.Visible = False
If lbRegistro = True Then
    Dim oGeneral As DGeneral
    Set oGeneral = New DGeneral
    Dim lsOpeCod As String
    lblCaptionUser.Visible = True
    txtBuscarUser.Visible = True
    lblDescUser.Visible = True
    txtfecha = gdFecSis
    txtfecha.Enabled = False
    txtBuscarUser.Enabled = False
    oUser.Inicio gsCodUser
    txtBuscarUser = gsCodUser
    lblDescUser = PstaNombre(oUser.UserNom)
    If Mid(gsopecod, 3, 1) = "1" Then
        lsOpeCod = gOpeHabCajRegEfectMN
    Else
        lsOpeCod = gOpeHabCajRegEfectME
    End If
    lsMovNro = oCajero.GetMovUserBilletaje(gsopecod, gsCodUser, gdFecSis)
    lnMovNro = Val(oCajero.GetMovUserBilletaje(lsOpeCod, gsCodUser, gdFecSis, True))
    If oCajero.GetMovUserBilletajeDevuelto(lsOpeCod, gsCodUser, gdFecSis) <> "" Then
        MsgBox "Se ha realizado la Devolución del billetaje registrado ", vbInformation, "Aviso"
        lbSalir = True
    End If
    If lbModifica Then
        fgBilletes.lbEditarFlex = True
        fgMonedas.lbEditarFlex = True
    Else
        fgBilletes.lbEditarFlex = False
        fgMonedas.lbEditarFlex = False
    End If
    Set oGeneral = Nothing
End If
CargaBilletajes

If oSubasta.VerfDevSubasta(gsCodUser, gnSubCuadreCaja, gdFecSis) Then
    Me.cmdAceptar.Enabled = False
End If

End Sub


Public Sub Muestra(ByVal psMovnro As String)
lbMuestra = True
lsMovNro = psMovnro
lnMoneda = Mid(gsopecod, 3, 1)
Me.Show 1
End Sub
Public Sub RegistroEfectivo(Optional pbModifica As Boolean = True)
lbRegistro = True
lbModifica = pbModifica
lnMoneda = Mid(gsopecod, 3, 1)
Me.Show 1
End Sub

Public Sub Inicio(ByVal psOpeCod As String, ByVal psOpeDesc As String, _
            ByVal pnMontoSol As Currency, ByVal pnMoneda As Moneda, _
            Optional ByVal pbEnableFecha As Boolean = True, Optional pbCubreDif As Boolean = False)
 
lbDiferencia = pbCubreDif
lbEnableFecha = pbEnableFecha
lnMoneda = pnMoneda
lnMonto = pnMontoSol
lblMontoSol = Format(Abs(pnMontoSol), "##,###0.00")
If Val(pnMontoSol) < 0 Then
    lblMontoSol.ForeColor = &HFF&
Else
    lblMontoSol.ForeColor = &HC00000
End If
Me.Show 1
End Sub
Private Sub CargaBilletajes()
Dim sql As String
Dim rs As ADODB.Recordset
Dim oContFunct As NContFunciones
Dim oEfec As Defectivo
Dim lnFila As Long

Set oContFunct = New NContFunciones
Set oEfec = New Defectivo

Set rs = New ADODB.Recordset
If lbRegistro Then
    If lsMovNro <> "" Then
        Set rs = oCajero.GetBilletajeCajero(lsMovNro, txtBuscarUser, lnMoneda, "B")
    Else
        Set rs = oEfec.EmiteBilletajes(lnMoneda, "B")
    End If
Else
    If lbMuestra = False Then
        Set rs = oEfec.EmiteBilletajes(lnMoneda, "B")
    Else
        Set rs = oEfec.GetBilletajesMov(lsMovNro, lnMoneda, "B")
    End If
End If
fgBilletes.FontFixed.Bold = True
fgBilletes.Clear
fgBilletes.FormaCabecera
fgBilletes.Rows = 2
Do While Not rs.EOF
    fgBilletes.AdicionaFila
    lnFila = fgBilletes.row
    fgBilletes.TextMatrix(lnFila, 1) = rs!descripcion
    fgBilletes.TextMatrix(lnFila, 2) = Format(rs!cantidad, "#,#0")
    fgBilletes.TextMatrix(lnFila, 3) = Format(rs!monto, "#,#0.00")
    fgBilletes.TextMatrix(lnFila, 4) = rs!cEfectivoCod
    fgBilletes.TextMatrix(lnFila, 5) = rs!nEfectivoValor
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
fgBilletes.col = 2

Set rs = New ADODB.Recordset
If lbRegistro Then
    If lsMovNro <> "" Then
        Set rs = oCajero.GetBilletajeCajero(lsMovNro, txtBuscarUser, lnMoneda, "M")
    Else
        Set rs = oEfec.EmiteBilletajes(lnMoneda, "M")
    End If
Else
    If lbMuestra = False Then
        Set rs = oEfec.EmiteBilletajes(lnMoneda, "M")
    Else
        Set rs = oEfec.GetBilletajesMov(lsMovNro, lnMoneda, "M")
    End If
End If
fgMonedas.FontFixed.Bold = True
fgMonedas.Clear
fgMonedas.FormaCabecera
fgMonedas.Rows = 2
Do While Not rs.EOF
    fgMonedas.AdicionaFila
    lnFila = fgMonedas.row
    fgMonedas.TextMatrix(lnFila, 1) = rs!descripcion
    fgMonedas.TextMatrix(lnFila, 2) = Format(rs!cantidad, "#,#0")
    fgMonedas.TextMatrix(lnFila, 3) = Format(rs!monto, "#,#0.00")
    fgMonedas.TextMatrix(lnFila, 4) = rs!cEfectivoCod
    fgMonedas.TextMatrix(lnFila, 5) = rs!nEfectivoValor
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
Set oContFunct = Nothing
Set oEfec = Nothing
fgMonedas.col = 2
lblTotalBilletes = Format(fgBilletes.SumaRow(3), "#,#0.00")
lblTotMoneda = Format(fgMonedas.SumaRow(3), "#,#0.00")
lblTotal = Format(CCur(lblTotalBilletes) + CCur(lblTotMoneda), "#,#0.00")
End Sub
Public Function Residuo(Dividendo As Currency, Divisor As Currency) As Boolean
Dim x As Currency
x = Round(Dividendo / Divisor, 0)
Residuo = True
x = x * Divisor
If x <> Dividendo Then
   Residuo = False
End If
End Function
Public Property Get vnDiferencia() As Currency
vnDiferencia = lnDiferencia
End Property
Public Property Let vnDiferencia(ByVal vNewValue As Currency)
lnDiferencia = vNewValue
End Property
Private Sub Form_Unload(Cancel As Integer)
Set oCajero = Nothing
End Sub

Private Sub txtBuscarUser_EmiteDatos()
lblDescUser = txtBuscarUser.psDescripcion
'lsMovNro = oCajero.GetMovUserBilletaje(gsOpeCod, txtBuscarUser, gdFecSis)
'CargaBilletajes
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    fgBilletes.SetFocus
End If
End Sub
Public Property Get MovNro() As String
MovNro = lsMovNro
End Property
Public Property Get nMovNro() As Long
nMovNro = lnMovNro
End Property
Public Property Get Total() As Currency
Total = CCur(lblTotal)
End Property
Public Property Get FechaMov() As Date
FechaMov = CDate(txtfecha)
End Property

