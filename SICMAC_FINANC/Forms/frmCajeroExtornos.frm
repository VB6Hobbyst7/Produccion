VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajeroExtornos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4800
   ClientLeft      =   1050
   ClientTop       =   2265
   ClientWidth     =   9630
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajeroExtornos.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   630
      Left            =   75
      TabIndex        =   9
      Top             =   315
      Width           =   9480
      Begin Sicmact.TxtBuscar txtBuscarUser 
         Height          =   330
         Left            =   5460
         TabIndex        =   15
         Top             =   195
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   582
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
         sTitulo         =   ""
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
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
         Left            =   8010
         TabIndex        =   2
         Top             =   158
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtDesde 
         Height          =   315
         Left            =   795
         TabIndex        =   0
         Top             =   195
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   8388608
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin MSMask.MaskEdBox txthasta 
         Height          =   315
         Left            =   2655
         TabIndex        =   1
         Top             =   195
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   8388608
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Usuario :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   4605
         TabIndex        =   14
         Top             =   225
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2025
         TabIndex        =   11
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   150
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame FraLista 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   75
      TabIndex        =   8
      Top             =   960
      Width           =   9480
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7935
         TabIndex        =   7
         Top             =   3090
         Width           =   1350
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6660
         TabIndex        =   12
         Top             =   3090
         Width           =   1275
      End
      Begin VB.TextBox txtMovDesc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2940
         Width           =   6285
      End
      Begin VB.CommandButton cmdExtornar 
         Caption         =   "&Extornar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6660
         TabIndex        =   6
         Top             =   3090
         Width           =   1275
      End
      Begin VB.TextBox txtConcepto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   150
         Locked          =   -1  'True
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2220
         Width           =   9240
      End
      Begin Sicmact.FlexEdit fgListaCG 
         Height          =   1920
         Left            =   150
         TabIndex        =   3
         Top             =   240
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   3387
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-Fecha-Operación-Usuario-Persona-Importe-cPersCod-cMovNro-cOpeCod-campo2"
         EncabezadosAnchos=   "350-1500-1800-800-3200-1200-0-0-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-L-R-L-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-2-0-0-0-0"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbOrdenaCol     =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   9525
         X2              =   30
         Y1              =   2850
         Y2              =   2850
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   10275
         X2              =   0
         Y1              =   2790
         Y2              =   2790
      End
   End
   Begin Sicmact.Usuario User 
      Left            =   0
      Top             =   0
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "EXTORNOS CAJERO "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   3023
      TabIndex        =   13
      Top             =   60
      Width           =   3585
   End
End
Attribute VB_Name = "frmCajeroExtornos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCajero As nCajero
Dim lsAreaCod As String
Dim lsAgeCod As String
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Dim oGen As DGeneral

Private Sub cmdConfirmar_Click()
Dim oCon As NContFunciones
Dim lsMovNro As String
Dim lsMovNroHab As String
Dim lnImporte As Currency
Dim ldFechaMov As Date
 
Set oCon = New NContFunciones
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Concepto de Operación no válida", vbInformation, "Aviso"
    txtMovDesc.SetFocus
    Exit Sub
End If
lsMovNroHab = fgListaCG.TextMatrix(fgListaCG.Row, 7)
lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.Row, 5))
ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.Row, 1))

If MsgBox("Desea Confirmar la Habilitación respectiva??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    If oCajero.GrabaConfHabilitaAgencia(gsFormatoFecha, lsMovNro, gsOpeCod, txtMovDesc, lsMovNroHab, _
                                        lnImporte, "", lsAreaCod, lsAgeCod, "", "", gsCodUser) = 0 Then
        
        
        Dim oContImp As NContImprimir
        Dim lbOk As Boolean
        Set oContImp = New NContImprimir
        lbOk = True
        Do While lbOk
            oContImp.ImprimeBoletahabilitacion lblTitulo, "CONFIRMACION HAB. EN EFECTIVO", _
                     lsAreaCod + lsAgeCod, gsNomAge, gsCodUser, fgListaCG.TextMatrix(fgListaCG.Row, 4), Mid(gsOpeCod, 3, 1), gsOpeCod, _
                     lnImporte, gsNomAge, lsMovNro, "LPT1"
                    
            If MsgBox("Desea Reimprimir Boleta de Operación??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                lbOk = False
            End If
        Loop
        Set oContImp = Nothing
        If MsgBox("Desea realizar otra Confirmación ??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
            fgListaCG.EliminaFila fgListaCG.Row
            fgListaCG.SetFocus
            txtMovDesc = ""
            txtConcepto = ""
        Else
            Unload Me
        End If
    End If
    
End If



End Sub

Private Sub cmdExtornar_Click()
Dim oCon As NContFunciones
Dim oCaja As nCajaGeneral
Dim lsMovNro As String
Dim lsMovNroExt As String
Dim lnImporte As Currency
Dim ldFechaMov As Date

Set oCon = New NContFunciones
Set oCaja = New nCajaGeneral
If fgListaCG.TextMatrix(1, 0) = "" Then Exit Sub
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Concepto de Operación no válida", vbInformation, "Aviso"
    txtMovDesc.SetFocus
    Exit Sub
End If
'Select Case gsOpeCod
'    Case gOpeBoveAgeExtHabCajeroMN, gOpeBoveAgeExtHabCajeroME, _
         gOpeHabCajExtConfHabBovAgeMN, gOpeHabCajExtConfHabBovAgeMN, _
         gOpeHabCajExtDevABoveMN, gOpeHabCajExtDevABoveME, _
         gOpeHabCajExtTransfEfectCajerosMN, gOpeHabCajExtTransfEfectCajerosME

        lsMovNro = fgListaCG.TextMatrix(fgListaCG.Row, 7)
        lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.Row, 5))
        ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.Row, 1))
        
'End Select
If MsgBox("Desea Realizar el Extorno respectivo??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    If gdFecSis <> Format(ldFechaMov, "dd/mm/yyyy") Then
        If MsgBox("Se va a Realizar el Extorno de Movimientos de dias anteriores" & vbCrLf & " Desea Proseguir??", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If
    lsMovNroExt = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    If oCaja.GrabaExtornoMov(gdFecSis, ldFechaMov, lsMovNroExt, lsMovNro, gsOpeCod, txtMovDesc, lnImporte, lsMovNroExt, True) = 0 Then
        Set oCon = Nothing
            Dim oImp As NContImprimir
            Dim lsTexto As String
            Dim lbReimp As Boolean
            Set oImp = New NContImprimir
            
            lbReimp = True
            Do While lbReimp
                oImp.ImprimeBoletaExtornos lblTitulo, txtMovDesc, gsOpeCod, lnImporte, fgListaCG.TextMatrix(fgListaCG.Row, 3), _
                        fgListaCG.TextMatrix(fgListaCG.Row, 4), gsNomAge, lsMovNroExt, sLPT
            
                If MsgBox("Desea Reimprimir boleta de Operación", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                    lbReimp = False
                End If
            Loop
            Set oImp = Nothing
        If MsgBox("Desea realizar otro extorno de Movimiento??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
            fgListaCG.EliminaFila fgListaCG.Row
            fgListaCG.SetFocus
            txtMovDesc = ""
            txtConcepto = ""
        Else
            Unload Me
        End If
    End If
End If

End Sub

Private Sub cmdProcesar_Click()
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
Dim lsOperacion As String
If ValFecha(Me.txtDesde) = False Then Exit Sub
If ValFecha(txthasta) = False Then Exit Sub

If CDate(txtDesde) > CDate(txthasta) Then
    MsgBox "Fecha Inicial no puede ser mayor que la Final", vbInformation, "Aviso"
    Exit Sub
End If
Select Case gsOpeCod
    Case gOpeBoveAgeExtHabCajeroMN, gOpeBoveAgeExtHabCajeroME
        lsOperacion = IIf(Mid(gsOpeCod, 3, 1) = "1", gOpeBoveAgeHabCajeroMN, gOpeBoveAgeHabCajeroME)
        Set rs = oCajero.GetMovHabBovAgeCajero(lsOperacion, lsAreaCod, lsAgeCod, CDate(txtDesde), CDate(txthasta), txtBuscarUser)
    Case gOpeHabCajConfHabBovAgeMN, gOpeHabCajConfHabBovAgeME
        lsOperacion = IIf(Mid(gsOpeCod, 3, 1) = "1", gOpeBoveAgeHabCajeroMN, gOpeBoveAgeHabCajeroME)
        Set rs = oCajero.GetMovHabBovAgeCajero(lsOperacion, lsAreaCod, lsAgeCod, CDate(txtDesde), CDate(txthasta), txtBuscarUser)
    Case gOpeHabCajExtConfHabBovAgeMN, gOpeHabCajExtConfHabBovAgeME
        lsOperacion = IIf(Mid(gsOpeCod, 3, 1) = "1", gOpeBoveAgeHabCajeroMN, gOpeBoveAgeHabCajeroME)
        Set rs = oCajero.GetConfHabBovAgencia(lsOperacion, CDate(txtDesde), CDate(txthasta), lsAreaCod, lsAgeCod, txtBuscarUser)
    Case gOpeHabCajExtTransfEfectCajerosMN, gOpeHabCajExtTransfEfectCajerosME
        lsOperacion = IIf(Mid(gsOpeCod, 3, 1) = "1", gOpeHabCajTransfEfectCajerosMN, gOpeHabCajTransfEfectCajerosME)
        Set rs = oCajero.GetTransfEntreCajeros(lsOperacion, CDate(txtDesde), CDate(txthasta), txtBuscarUser)
    Case gOpeCajeroMEExtCompra, gOpeCajeroMEExtVenta
        lsOperacion = IIf(gsOpeCod = gOpeCajeroMEExtCompra, gOpeCajeroMECompra, gOpeCajeroMEVenta)
        Set rs = oCajero.GetCompraVenta(lsOperacion, CDate(txtDesde), CDate(txthasta), txtBuscarUser)
    Case gOpeCajeroVarExtSEDALIBMN, gOpeCajeroVarExtTELEFONICAMN, gOpeCajeroVarExtHIDRANDINAMN, _
        gOpeCajeroVarExtHIDRANDINAME, gOpeCajeroVarExtTELEFONICAME, gOpeCajeroVarExtSEDALIBME
        Select Case gsOpeCod
            Case gOpeCajeroVarExtSEDALIBMN
                lsOperacion = gOpeCajeroVarSEDALIBMN
            Case gOpeCajeroVarExtTELEFONICAMN
                lsOperacion = gOpeCajeroVarTELEFONICAMN
            Case gOpeCajeroVarExtHIDRANDINAMN
                lsOperacion = gOpeCajeroVarHIDRANDINAMN
            Case gOpeCajeroVarExtSEDALIBME
                lsOperacion = gOpeCajeroVarSEDALIBME
            Case gOpeCajeroVarExtTELEFONICAME
                lsOperacion = gOpeCajeroVarTELEFONICAME
            Case gOpeCajeroVarExtHIDRANDINAME
                lsOperacion = gOpeCajeroVarHIDRANDINAME
        End Select
        Set rs = oCajero.GetServicios(lsOperacion, CDate(txtDesde), CDate(txthasta), Trim(txtBuscarUser), gsCodAge)
    Case gOpeHabCajExtIngEfectRegulaFaltMN, gOpeHabCajExtIngEfectRegulaFaltME
        lsOperacion = IIf(gsOpeCod = gOpeHabCajExtIngEfectRegulaFaltMN, gOpeHabCajIngEfectRegulaFaltMN, gOpeHabCajIngEfectRegulaFaltME)
        Set rs = oCajero.GetIngEfectFalt(lsOperacion, CDate(txtDesde), CDate(txthasta), Trim(txtBuscarUser), gsCodAge)
End Select
fgListaCG.Clear
fgListaCG.FormaCabecera
fgListaCG.Rows = 2
If Not rs.EOF And Not rs.BOF Then
    Set fgListaCG.Recordset = rs
    fgListaCG.FormatoPersNom 4
    fgListaCG.SetFocus
Else
    MsgBox "Datos no encontrados", vbInformation, "Aviso"
End If
rs.Close
Set rs = Nothing
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub fgListaCG_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMovDesc.SetFocus
End If
End Sub

Private Sub fgListaCG_RowColChange()
'Select Case gsOpeCod
'    Case gOpeBoveAgeExtHabCajeroMN, gOpeBoveAgeExtHabCajeroME, _
        gOpeHabCajConfHabBovAgeMN, gOpeHabCajConfHabBovAgeME, _
        gOpeHabCajExtConfHabBovAgeMN, gOpeHabCajExtConfHabBovAgeME, _
        gOpeHabCajExtDevABoveMN, gOpeHabCajExtDevABoveME, _
        gOpeHabCajExtTransfEfectCajerosMN, gOpeHabCajExtTransfEfectCajerosME
        
        txtConcepto = fgListaCG.TextMatrix(fgListaCG.Row, 6)
'End Select
End Sub
Private Sub Form_Load()
Set oCajero = New nCajero
Dim oOpe As New DOperacion

Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
Set oOpe = New DOperacion
Set oGen = New DGeneral
CentraForm Me
Me.Caption = gsOpeDesc
txtDesde = gdFecSis
txthasta = gdFecSis
User.Inicio gsCodUser
Me.txtBuscarUser.psRaiz = "USUARIOS"
txtBuscarUser.rs = oGen.GetUserAreaAgencia(User.cAreaCodAct, gsCodAge)

cmdExtornar.Visible = False
cmdConfirmar.Visible = False
Select Case gsOpeCod
    Case gOpeBoveAgeExtHabCajeroMN, gOpeBoveAgeExtHabCajeroME
        Me.lblTitulo = gsOpeDesc
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txthasta.Enabled = False
        lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D")
        lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H")
        Set rs = GetObjetosOpeCta(gsOpeCod, "0", lsCtaDebe, "", , gsCodAge)
        If Not rs.EOF And Not rs.BOF Then
            lsAreaCod = Mid(rs(0), 1, 3)
            lsAgeCod = Mid(rs(0), 4, 2)
        End If
        rs.Close
        Set rs = Nothing
    Case gOpeHabCajConfHabBovAgeMN, gOpeHabCajConfHabBovAgeME
        cmdConfirmar.Visible = True
        txtDesde.Enabled = False
        txthasta.Enabled = False
        txtBuscarUser = gsCodUser
        txtBuscarUser.Enabled = False
        Me.lblTitulo = "CONFIRMACION DE HABILITACION"
        lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D")
        lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H")
        Set rs = GetObjetosOpeCta(gsOpeCod, "0", lsCtaDebe, "", , gsCodAge)
        If Not rs.EOF And Not rs.BOF Then
            lsAreaCod = Mid(rs(0), 1, 3)
            lsAgeCod = Mid(rs(0), 4, 2)
        End If
        rs.Close
        Set rs = Nothing
    Case gOpeHabCajExtConfHabBovAgeMN, gOpeHabCajExtConfHabBovAgeME
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txthasta.Enabled = False
        Me.lblTitulo = "EXTORNO DE CONFIRMACION DE HABILITACION DE BOVEDA"
        lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D")
        lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H")
        Set rs = GetObjetosOpeCta(gsOpeCod, "0", lsCtaDebe, "", , gsCodAge)
        If Not rs.EOF And Not rs.BOF Then
            lsAreaCod = Mid(rs(0), 1, 3)
            lsAgeCod = Mid(rs(0), 4, 2)
        End If
        rs.Close
        Set rs = Nothing
    Case gOpeHabCajExtDevABoveMN, gOpeHabCajExtDevABoveME
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txthasta.Enabled = False
        Me.lblTitulo = "EXTORNO DE DEVOLUCION A BOVEDA"
    Case gOpeHabCajExtTransfEfectCajerosMN, gOpeHabCajExtTransfEfectCajerosME
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txthasta.Enabled = False
        Me.lblTitulo = "EXTORNO TRANSFERENCIA CAJEROS"
    Case gOpeCajeroMEExtCompra, gOpeCajeroMEExtVenta
        Me.lblTitulo = gsOpeDesc
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txthasta.Enabled = False
    Case gOpeCajeroVarExtHIDRANDINAMN, gOpeCajeroVarExtTELEFONICAMN, gOpeCajeroVarExtSEDALIBMN, _
        gOpeCajeroVarExtHIDRANDINAME, gOpeCajeroVarExtTELEFONICAME, gOpeCajeroVarExtSEDALIBME
        fgListaCG.EncabezadosNombres = "N° -Fecha - Operación - Usuario - Documento/Referencia - Importe - cMovNro - cOpeCod"
        lblTitulo = gsOpeDesc
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txthasta.Enabled = False
    Case gOpeHabCajExtIngEfectRegulaFaltMN, gOpeHabCajExtIngEfectRegulaFaltME
        fgListaCG.EncabezadosNombres = "N° -Fecha - Operación - Usuario - Concepto - Importe - cMovNro - cOpeCod"
        Me.lblTitulo = gsOpeDescHijo
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txthasta.Enabled = False
End Select
Set oOpe = Nothing
End Sub
Private Sub txtDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtDesde) = False Then Exit Sub
    If CDate(txtDesde) > CDate(txthasta) Then
        MsgBox "Fecha Inicial no puede ser mayor que la Final", vbInformation, "Aviso"
        Exit Sub
    End If
    txthasta.SetFocus
End If
End Sub
Private Sub txtHasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txthasta) = False Then Exit Sub
    If CDate(txtDesde) > CDate(txthasta) Then
        MsgBox "Fecha Inicial no puede ser mayor que la Final", vbInformation, "Aviso"
        Exit Sub
    End If
    cmdProcesar.SetFocus
End If
End Sub
Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If cmdExtornar.Visible Then cmdExtornar.SetFocus
    If cmdConfirmar.Visible Then cmdConfirmar.SetFocus
End If
End Sub
