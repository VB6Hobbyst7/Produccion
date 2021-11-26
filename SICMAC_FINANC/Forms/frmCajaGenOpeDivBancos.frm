VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCajaGenOpeDivBancos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5595
   ClientLeft      =   3405
   ClientTop       =   2895
   ClientWidth     =   6945
   Icon            =   "frmCajaGenOpeDivBancos.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fradoc 
      Caption         =   "&Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   750
      Left            =   120
      TabIndex        =   23
      Top             =   4080
      Width           =   6735
      Begin VB.ComboBox cboComision 
         Height          =   315
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   300
         Visible         =   0   'False
         Width           =   3270
      End
      Begin VB.ComboBox cboDocumento 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   300
         Width           =   2310
      End
      Begin VB.TextBox txtNroDoc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2910
         TabIndex        =   25
         Top             =   300
         Width           =   1695
      End
      Begin MSMask.MaskEdBox txtFechaDoc 
         Height          =   330
         Left            =   5400
         TabIndex        =   24
         Top             =   300
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblComi 
         Caption         =   "Comisión"
         Height          =   255
         Left            =   2580
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblDocNro 
         AutoSize        =   -1  'True
         Caption         =   "Nº :"
         Height          =   195
         Left            =   2580
         TabIndex        =   28
         Top             =   360
         Width           =   270
      End
      Begin VB.Label lblDocFec 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   4830
         TabIndex        =   27
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5490
      TabIndex        =   11
      Top             =   5085
      Width           =   1275
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   5115
      Width           =   1755
   End
   Begin VB.Frame FraTipoCambio 
      Caption         =   "Tipo Cambio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   660
      Left            =   3825
      TabIndex        =   17
      Top             =   0
      Width           =   3015
      Begin VB.TextBox txtTCBanco 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1905
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   210
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Left            =   1365
         TabIndex        =   19
         Top             =   255
         Width           =   510
      End
      Begin VB.Label lblTCFijo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   465
         TabIndex        =   1
         Top             =   225
         Width           =   795
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fijo:"
         Height          =   195
         Left            =   105
         TabIndex        =   18
         Top             =   270
         Width           =   285
      End
   End
   Begin TabDlg.SSTab StabCuentas 
      Height          =   2190
      Left            =   105
      TabIndex        =   12
      Top             =   450
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3863
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   617
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Origen"
      TabPicture(0)   =   "frmCajaGenOpeDivBancos.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraOrigen"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Destino"
      TabPicture(1)   =   "frmCajaGenOpeDivBancos.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraDestino"
      Tab(1).ControlCount=   1
      Begin VB.Frame FraOrigen 
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
         Height          =   1680
         Left            =   135
         TabIndex        =   15
         Top             =   420
         Width           =   6420
         Begin Sicmact.TxtBuscar txtObjOrig 
            Height          =   330
            Left            =   180
            TabIndex        =   3
            Top             =   495
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   582
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            sTitulo         =   ""
         End
         Begin VB.Label lblDescCtabanco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   180
            TabIndex        =   5
            Top             =   885
            Width           =   6030
         End
         Begin VB.Label lblDescbanco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2265
            TabIndex        =   4
            Top             =   510
            Width           =   3945
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta Institución Finaciera :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   105
            TabIndex        =   20
            Top             =   225
            Width           =   2370
         End
      End
      Begin VB.Frame FraDestino 
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
         Height          =   1680
         Left            =   -74880
         TabIndex        =   13
         Top             =   420
         Width           =   6435
         Begin Sicmact.TxtBuscar txtBuscarDest 
            Height          =   330
            Left            =   855
            TabIndex        =   6
            Top             =   210
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   582
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            sTitulo         =   ""
         End
         Begin Sicmact.FlexEdit fgObjDest 
            Height          =   975
            Left            =   870
            TabIndex        =   8
            Top             =   600
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   1720
            Cols0           =   6
            HighLight       =   2
            AllowUserResizing=   3
            EncabezadosNombres=   "-Objeto-Descripción-SubCta-cObjetoCod-nCtaObjNiv"
            EncabezadosAnchos=   "350-1200-2700-800-0-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-0"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Objetos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   105
            TabIndex        =   21
            Top             =   570
            Width           =   645
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   105
            TabIndex        =   14
            Top             =   240
            Width           =   675
         End
         Begin VB.Label lblDescCtaDest 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2280
            TabIndex        =   7
            Top             =   225
            Width           =   4020
         End
      End
   End
   Begin MSMask.MaskEdBox txtFechamov 
      Height          =   330
      Left            =   900
      TabIndex        =   0
      Top             =   60
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4230
      TabIndex        =   10
      Top             =   5085
      Width           =   1275
   End
   Begin VB.Frame FraMotivo 
      Height          =   1320
      Left            =   120
      TabIndex        =   29
      Top             =   2700
      Width           =   6720
      Begin VB.CommandButton cmdNegocio 
         Caption         =   "&Operación de Negocio"
         Height          =   345
         Left            =   150
         TabIndex        =   32
         Top             =   240
         Width           =   2325
      End
      Begin VB.TextBox txtMovDesc 
         Height          =   510
         Left            =   855
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   675
         Width           =   5700
      End
      Begin VB.Label lblPersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2580
         TabIndex        =   33
         Top             =   270
         Width           =   3945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Glosa  :"
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
         Height          =   195
         Left            =   150
         TabIndex        =   31
         Top             =   690
         Width           =   675
      End
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monto "
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
      Left            =   285
      TabIndex        =   22
      Top             =   5160
      Width           =   570
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   210
      TabIndex        =   16
      Top             =   105
      Width           =   600
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   345
      Left            =   135
      Top             =   5100
      Width           =   2955
   End
End
Attribute VB_Name = "frmCajaGenOpeDivBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oOpe              As DOperacion
Dim oCtaIf            As NCajaCtaIF
Dim lnTipoObj         As TpoObjetos
Dim oCaja             As nCajaGeneral
Dim rsNeg             As ADODB.Recordset
Dim lsCtaContBanco    As String
Dim lsAjusteCta       As String
Dim lsClaseCta        As String
Dim lbDeposito        As Boolean
Dim lbGetDataNeg      As Boolean
Dim lbRegulaPendiente As Boolean
Dim lnMontoDif        As Currency
Dim objPista As COMManejador.Pista 'ARLO20170217

Public Sub Inicio(Optional pbDeposito As Boolean = True, Optional pbGetDataNeg As Boolean = True)
lbDeposito = pbDeposito
lbGetDataNeg = pbGetDataNeg
Me.Show 1
End Sub

Private Sub cboComision_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAceptar.SetFocus
End If
End Sub

Private Sub MuestraControlDoc(pbActiva As Boolean)
txtNroDoc.Enabled = pbActiva
txtNroDoc.Visible = pbActiva
txtFechaDoc.Visible = pbActiva
lblDocNro.Visible = pbActiva
lblDocFec.Visible = pbActiva
lblComi.Visible = Not pbActiva
cboComision.Visible = Not pbActiva
End Sub

Private Sub cboDocumento_Click()
MuestraControlDoc True
If cboDocumento <> "" Then
    Select Case Val(Right(cboDocumento, 2))
        Case TpoDocCarta, TpoDocCheque, TpoDocOrdenPago, TpoDocNotaAbono, TpoDocNotaCargo
            MuestraControlDoc False
    End Select
End If
If Left(cboDocumento, 7) = "NINGUNO" Then
    txtNroDoc.Enabled = False
End If
End Sub

Private Sub cboDocumento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtNroDoc.Enabled And txtNroDoc.Visible Then
        txtNroDoc.SetFocus
    ElseIf txtFechaDoc.Visible Then
        txtFechaDoc.SetFocus
    ElseIf cboComision.Visible Then
        cboComision.SetFocus
    Else
        cmdAceptar.SetFocus
    End If
End If
End Sub

Private Sub cboMotivo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMovDesc.SetFocus
End If
End Sub
Function Valida() As Boolean
'***Agregado por ELRO el 20120109, según Acta N° 003-2012/TI-D
Dim dFechaCierreMensualContabilidad, dFechaHabil As Date
Dim i, nDias As Integer
Dim oNConstSistemas As NConstSistemas
Set oNConstSistemas = New NConstSistemas
Dim oNContFunciones As New NContFunciones
Set oNContFunciones = New NContFunciones
   
dFechaCierreMensualContabilidad = CDate(oNConstSistemas.LeeConstSistema(gConstSistCierreMensualCont))
'***Fin Agregado por ELRO*************************************

Valida = True
If lsCtaContBanco = "" Then
    MsgBox "Cuenta " & StabCuentas.TabCaption(0) & " no ha sido seleccionada", vbInformation, "Aviso"
    Valida = False
    StabCuentas.Tab = 0
    txtObjOrig.SetFocus
    Exit Function
End If
If Trim(txtObjOrig) = "" And txtObjOrig.rs.RecordCount > 0 Then
    MsgBox "Objeto " & StabCuentas.TabCaption(0) & " no ha sido seleccionada", vbInformation, "Aviso"
    Valida = False
    StabCuentas.Tab = 0
    txtObjOrig.SetFocus
    Exit Function
End If
If Trim(txtBuscarDest) = "" Then
    MsgBox "Cuenta " & StabCuentas.TabCaption(1) & " no ha sido seleccionada", vbInformation, "Aviso"
    Valida = False
    StabCuentas.Tab = 1
    txtBuscarDest.SetFocus
    Exit Function
End If
If Len(Trim(txtMovDesc)) = 0 Then
    MsgBox "Descripción de Operación no ingresada", vbInformation, "aviso"
    Valida = False
    txtMovDesc.SetFocus
    Exit Function
End If
If Len(Trim(cboDocumento)) = 0 Then
    If MsgBox("Documento no ha sido Ingresado. Desea Continuar??", vbQuestion + vbYesNo, "aviso") = vbNo Then
        Valida = False
        cboDocumento.SetFocus
        Exit Function
    End If
End If
If cboDocumento <> "" And Left(cboDocumento, 7) <> "NINGUNO" Then
Select Case Val(Right(cboDocumento, 2))
    Case TpoDocCarta, TpoDocCheque, TpoDocNotaAbono, TpoDocNotaCargo
    Case Else
        If Len(Trim(txtNroDoc)) = 0 Then
            MsgBox "Numero de documento no ha sido ingresado", vbInformation, "aviso"
            Valida = False
            If txtNroDoc.Enabled Then txtNroDoc.SetFocus
            Exit Function
        End If
End Select
End If
If Val(txtImporte) = 0 Then
    MsgBox "Importe de Operación no válido", vbInformation, "aviso"
    Valida = False
    txtImporte.SetFocus
    Exit Function
End If
'***Modificado por ELRO el 20120109, según Acta N° 003-2012/TI-D
'    If Not ValidaFechaContab(txtFechamov, gdFecSis, True) Then
'       Valida = False
'       fEnfoque txtFechamov
'       txtFechamov.SetFocus
'       Exit Function
'    End If
If Month(CDate(txtFechamov)) = Month(dFechaCierreMensualContabilidad) And _
   Year(CDate(txtFechamov)) = Year(dFechaCierreMensualContabilidad) Then
    
    If MsgBox("¿Desea realizar la operación en una fecha que pertenece a un Mes Cerrado?", vbYesNo, "Confirmar") = vbYes Then
    nDias = DateDiff("D", dFechaCierreMensualContabilidad, gdFecSis)
        For i = 1 To nDias
        
            If Not oNContFunciones.EsFeriado(DateAdd("D", i, dFechaCierreMensualContabilidad)) Then
                dFechaHabil = DateAdd("D", i, dFechaCierreMensualContabilidad)
                If DateDiff("D", dFechaHabil, gdFecSis) > 0 Then
                    MsgBox "Solo se puede realizar la operación en un Mes Cerrado hasta " & dFechaHabil, vbInformation, "aviso"
                    Valida = False
                    txtFechamov.SetFocus
                    Exit Function
                    
                End If
            
            End If
         
        Next i
    Else
        Valida = False
        Exit Function
    End If
Else
    If Not ValidaFechaContab(txtFechamov, gdFecSis, True) Then
       Valida = False
       fEnfoque txtFechamov
       txtFechamov.SetFocus
       Exit Function
    End If
End If
'***Fin Modificado por ELRO*************************************

If lbRegulaPendiente Then
    If rsNeg Is Nothing Then
       If MsgBox("Concepto seleccionado puede regularizar una Operación de Negocio" & Chr(10) & " ¿ Desea regularizar Operación de Negocio ? ", vbQuestion + vbYesNo, "¡Aviso!") = vbYes Then
            cmdNegocio.SetFocus
            Valida = False
       End If
    End If
End If
End Function
Private Sub cmdAceptar_Click()
On Error GoTo AceptarErr
Dim lsNroDoc As String
Dim lnTpoDoc As TpoDoc
Dim lsDocumento As String
Dim lsDocNotaAC As String
Dim lsNroVoucher As String
Dim lsNroNotaAC As String
Dim lnMotivoNAC As MotivoNotaAbonoCargo
Dim lsObjetoPadre As String
Dim lsObjetoCod As String

Dim lsCadBol As String
Dim lnMotivoNACAux As MotivoNotaAbonoCargo
Dim lsObjetoPadreAux As String
Dim lsObjetoCodAux As String

Dim oDoc As clsDocPago
Dim oCont As NContFunciones
Dim oDocRec As NDocRec
Dim oContImp As NContImprimir
Dim lsCuentaAho As String
Dim lnTpoDocAux As TpoDoc
                    
Dim lsMovNro As String
Dim lsPersNombre As String
Dim lsPersDireccion As String
Dim lsUbigeo As String
Dim lsCuentaAhoAux As String
Dim lnMontoAux As Currency
'Dim lbDeposito As Boolean
Dim lsEntiOrig As String
Dim lsCtaEntOrig As String
Dim lsEntiDest As String
Dim lsCtaEntDest As String
Dim lsSubCtaIF   As String

Set oDocRec = New NDocRec
Set oContImp = New NContImprimir
Set oDoc = New clsDocPago
Dim lsGlosa As String

Dim lbGrabaNegocio As Boolean


lbGrabaNegocio = False

If Valida = False Then Exit Sub
lsNroDoc = ""
lsDocNotaAC = ""
lsDocumento = ""
lsNroNotaAC = ""
lsNroVoucher = ""
lnTpoDoc = -1
If cboDocumento <> "" Then
    lnTpoDoc = Val(Right(cboDocumento, 2))
    If lbDeposito Then
            lsEntiOrig = gsNomCmac
            lsCtaEntOrig = ""
            lsEntiDest = lblDescbanco
            lsCtaEntDest = Trim(lblDescCtabanco)
    Else
            lsEntiOrig = lblDescbanco
            lsCtaEntOrig = Trim(lblDescCtabanco)
            lsEntiDest = gsNomCmac
            lsCtaEntDest = ""
    End If
    Select Case lnTpoDoc
        Case TpoDocCarta
            oDoc.InicioCarta "", "", gsOpeCod, gsOpeDesc, txtMovDesc, "", txtImporte, gdFecSis, lsEntiOrig, lsCtaEntOrig, lsEntiDest, lsCtaEntDest, ""
            If oDoc.vbOk Then
                lsNroDoc = oDoc.vsNroDoc
                lsDocumento = oDoc.vsDocumento
            Else
                Exit Sub
            End If
        Case TpoDocCheque
        
            lsSubCtaIF = oCtaIf.SubCuentaIF(Mid(txtObjOrig, 4, 13))
            lsEntiDest = gsNomCmac
            'EJVG20121130 ***
            'oDoc.InicioCheque "", True, Mid(Me.txtObjOrig, 4, 13), gsOpeCod, lsEntiDest, gsOpeDesc, txtMovDesc, txtImporte, gdFecSis, gsNomCmacRUC, lsSubCtaIF, lsEntiOrig, lsCtaEntOrig, "", True, gsCodAge, Mid(Me.txtObjOrig, 18, 10)
            If DateDiff("d", CDate(Me.txtFechamov.Text), gdFecSis) <> 0 Then
                MsgBox "Para operaciones con cheque la fecha debe ser la fecha actual del Sistema", vbInformation, "Aviso"
                Exit Sub
            End If
            oDoc.InicioCheque "", True, Mid(Me.txtObjOrig, 4, 13), gsOpeCod, lsEntiDest, gsOpeDesc, txtMovDesc, txtImporte, gdFecSis, gsNomCmacRUC, lsSubCtaIF, lsEntiOrig, lsCtaEntOrig, "", True, gsCodAge, Mid(Me.txtObjOrig, 18, 10), , Mid(Me.txtObjOrig, 1, 2), Mid(Me.txtObjOrig, 4, 13), Mid(Me.txtObjOrig, 18, Len(txtObjOrig))
            'END EJVG *******
            If oDoc.vbOk Then
                lsNroDoc = oDoc.vsNroDoc
                lsDocumento = oDoc.vsDocumento
                lsNroVoucher = oDoc.vsNroVoucher
            Else
                Exit Sub
            End If
        Case TpoDocOrdenPago
            oDoc.InicioOrdenPago "", True, "", gsOpeCod, gsNomCmac, gsOpeDesc, txtMovDesc, CCur(txtImporte), gdFecSis, "", True, gsCodAge
            If oDoc.vbOk Then
                lsNroDoc = oDoc.vsNroDoc
                lsDocumento = oDoc.vsDocumento
                lsNroVoucher = oDoc.vsNroVoucher
            Else
                Exit Sub
            End If
'        Case TpoDocNotaAbono, TpoDocNotaCargo
'              If cboComision.ListIndex = -1 Or Trim(Right(cboComision, 6)) = "" Then
'                 MsgBox "Debe seleccionar Motivo del Cargo", vbInformation, "¡AViso!"
'                 cboComision.SetFocus
'                 Exit Sub
'              End If

'            frmNotaCargoAbono.Inicio lnTpoDoc, CCur(txtImporte), gdFecSis, txtMovDesc, gsOpeCod, False
'            If frmNotaCargoAbono.vbOk Then
'                lsNroDoc = frmNotaCargoAbono.NroNotaCA
'                txtMovDesc = frmNotaCargoAbono.Glosa
'                lsDocumento = frmNotaCargoAbono.NotaCargoAbono
'                lsPersNombre = frmNotaCargoAbono.PersNombre
'                lsPersDireccion = frmNotaCargoAbono.PersDireccion
'                lsUbigeo = frmNotaCargoAbono.PersUbigeo
'                lsCuentaAho = frmNotaCargoAbono.CuentaAhoNro
'                lnMotivoNAC = frmNotaCargoAbono.Motivo
'                lsObjetoPadre = frmNotaCargoAbono.ObjetoMotivoPadre
'                lsObjetoCod = frmNotaCargoAbono.ObjetoMotivo
'
'                'lsNroDoc = oDocRec.GetNroNotaCargoAbono(lnTpoDoc)
'                'lsDocumento = oContImp.ImprimeNotaCargoAbono(lsNroNotaAC, txtMovDesc, CCur(frmNotaCargoAbono.Monto), _
'                                    lsPersNombre, lsPersDireccion, lsUbigeo, gdFecSis, Mid(gsOpeCod, 3, 1), lsCuentaAho, lnTpoDoc, gsNomAge, gsCodUser)
'
'                lsDocumento = oContImp.ImprimeNotaAbono(Format(gdFecSis, gsFormatoFechaView), CCur(frmNotaCargoAbono.Monto), txtMovDesc, lsCuentaAho, lsPersNombre)
'                Unload frmNotaCargoAbono
'                Set frmNotaCargoAbono = Nothing
'            Else
'                Unload frmNotaCargoAbono
'                Set frmNotaCargoAbono = Nothing
'                Exit Sub
'            End If
        Case Else
            lsNroDoc = txtNroDoc
    End Select
End If

If cboDocumento <> "" Then
    Select Case Val(Right(cboDocumento, 2))
        Case TpoDocCarta, TpoDocCheque
            'cargar el formulario de nota de cargo y abono para cargos por depositos o retiros a bancos
            If MsgBox("Desea Generar Nota de Cargo/Abono por Comisión Adicional??", vbYesNo + vbQuestion + vbDefaultButton2, "Aviso") = vbYes Then
                If cboComision.ListIndex = -1 Or Trim(Right(cboComision, 6)) = "" Then
                    MsgBox "Debe seleccionar Motivo del Cargo", vbInformation, "¡AViso!"
                    cboComision.SetFocus
                    Exit Sub
                End If
            
                lnTpoDocAux = TpoDocNotaCargo
                lbGrabaNegocio = True
                frmNotaCargoAbono.Inicio lnTpoDocAux, 0, gdFecSis, txtMovDesc, gsOpeCod, False
                If frmNotaCargoAbono.vbOk Then
                    If frmNotaCargoAbono.Monto > Me.txtImporte Then
                        MsgBox "Importe Cargado no debe superar el total de Operación", vbInformation, "Aviso"
                        Unload frmNotaCargoAbono
                        Set frmNotaCargoAbono = Nothing
                        Exit Sub
                    End If
                    lnMontoAux = frmNotaCargoAbono.Monto
                    lsNroNotaAC = frmNotaCargoAbono.NroNotaCA
                    lsDocNotaAC = frmNotaCargoAbono.Glosa
                    lsCuentaAhoAux = frmNotaCargoAbono.CuentaAhoNro
                    lsPersNombre = frmNotaCargoAbono.PersNombre
                    lsPersDireccion = frmNotaCargoAbono.PersDireccion
                    lsUbigeo = frmNotaCargoAbono.PersUbigeo
                    
                    lnMotivoNACAux = frmNotaCargoAbono.Motivo
                    lsObjetoPadreAux = frmNotaCargoAbono.ObjetoMotivoPadre
                    lsObjetoCodAux = frmNotaCargoAbono.ObjetoMotivo
                    
                    'lsNroNotaAC = oDocRec.GetNroNotaCargoAbono(TpoDocNotaCargo)
                    'lsDocNotaAC = oContImp.ImprimeNotaCargoAbono(lsNroNotaAC, txtMovDesc, CCur(frmNotaCargoAbono.Monto), _
                                    lsPersNombre, lsPersDireccion, lsUbigeo, gdFecSis, Mid(gsOpeCod, 3, 1), lsCuentaAhoAux, TpoDocNotaCargo, gsNomAge, gsCodUser)
                    lsDocNotaAC = oContImp.ImprimeNotaAbono(Format(gdFecSis, gsFormatoFechaView), CCur(frmNotaCargoAbono.Monto), txtMovDesc, lsCuentaAhoAux, lsPersNombre, Trim(Left(cboDocumento, Len(cboDocumento) - 3)) & " " & lsNroDoc)
                    
                    Dim oDis As New NRHProcesosCierre
                    lsCadBol = oDis.ImprimeBoletaCad(CDate(gdFecSis), "CARGO CAJA GENERAL", "CARGO CAJA GENERAL*Nro." & lsNroNotaAC, "", CCur(frmNotaCargoAbono.Monto), lsPersNombre, lsCuentaAho, "", 0, 0, "Nota Cargo", 0, 0, False, False, , , , True, , , , False, gsNomAge) & oImpresora.gPrnSaltoPagina
                    
                    Unload frmNotaCargoAbono
                    Set frmNotaCargoAbono = Nothing
                Else
                    Unload frmNotaCargoAbono
                    Set frmNotaCargoAbono = Nothing
                    Exit Sub
                End If
            End If
    End Select
End If
Set oCont = New NContFunciones

If MsgBox("Desea Grabar operación seleccionada??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsGlosa = Me.txtMovDesc
    lsMovNro = oCont.GeneraMovNro(txtFechamov, gsCodAge, gsCodUser)
    lsCtaContBanco = oOpe.EmiteOpeCta(gsOpeCod, IIf(lbDeposito, "D", "H"), , txtObjOrig, CtaOBjFiltroIF)
    If lsCtaContBanco = "" Then
       MsgBox "Institución Financiera no relacionado con Cuenta de Capital", vbInformation, "¡Aviso!"
       Exit Sub
    End If
    lsAjusteCta = ""
    If lnMontoDif <> 0 Then
      lsAjusteCta = oOpe.EmiteOpeCta(gsOpeCod, IIf((lbDeposito And lnMontoDif > 0) Or (Not lbDeposito And lnMontoDif < 0), "H", "D"), "3")
      If lsAjusteCta = "" Then
         MsgBox "No se definio Cuenta para realizar Operacion con Ajuste en el Orden 3", vbInformation, "¡Aviso!"
         Exit Sub
      End If
    End If
    
    cmdAceptar.Enabled = False
    cmdCancelar.Enabled = False
    oCaja.GrabaDepRetBancosVarios lsMovNro, gsOpeCod, lsGlosa, lsCtaContBanco, _
            txtBuscarDest, CCur(txtImporte), lnTpoDoc, lsNroDoc, txtFechaDoc, lsNroVoucher, _
            fgObjDest.GetRsNew, ObjEntidadesFinancieras, txtObjOrig, lbDeposito, lnMotivoNAC, lsObjetoPadre, lsObjetoCod, lsCuentaAho, _
            lnTpoDocAux, lsNroNotaAC, lnMontoAux, lnMotivoNACAux, lsObjetoPadreAux, lsObjetoCodAux, _
            lsCuentaAhoAux, lbGrabaNegocio, gbBitCentral, Right(cboComision, 6), rsNeg, lsAjusteCta, lnMontoDif, nVal(txtTCBanco)
    
    ImprimeAsientoContable lsMovNro, , lnTpoDoc, lsDocumento, , , , , , , , , , , , lsDocNotaAC
    cmdAceptar.Enabled = True
    cmdCancelar.Enabled = True
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo la Operación "
        Set objPista = Nothing
        '****
    If MsgBox("Desea ingresar otra operación?", vbQuestion + vbYesNo, "¡Confirmacion!") = vbNo Then
        Unload Me
    Else
        txtMovDesc = ""
        txtImporte = "0.00"
    End If
End If
Exit Sub
AceptarErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
    cmdAceptar.Enabled = True
    cmdCancelar.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdNegocio_Click()
Dim i As Integer
Dim lsAgencia As String
On Error GoTo ErrNegocio
If nVal(txtImporte) = 0 Then
    MsgBox "Falta indicar Importe de Operación", vbInformation, "¡Aviso!"
    txtImporte.SetFocus
    Exit Sub
End If
If txtBuscarDest = "" Then
    MsgBox "Falta indicar Cuenta Contable", vbInformation, "¡Aviso!"
    txtBuscarDest.SetFocus
    Exit Sub
End If
  
If fgObjDest.TextMatrix(1, 1) <> "" Then
    For i = 1 To fgObjDest.Rows - 1
        If Val(fgObjDest.TextMatrix(i, 4)) = TpoObjetos.ObjCMACAgenciaArea Then
            lsAgencia = Right(fgObjDest.TextMatrix(i, 1), 2)
        End If
    Next
End If
cmdNegocio.Enabled = False
frmOpeNegVentanilla.Inicio "", Mid(gsOpeCod, 3, 1), nVal(txtImporte), , , , lsAgencia, txtBuscarDest, IIf(lbDeposito, "D", "H")
If Not frmOpeNegVentanilla.lbOk Then
    lnMontoDif = 0
    RSClose rsNeg
    cmdNegocio.Enabled = True
    Exit Sub
End If
Set rsNeg = frmOpeNegVentanilla.rsPago
If frmOpeNegVentanilla.vnDiferencia <> 0 Then
    lnMontoDif = frmOpeNegVentanilla.vnDiferencia
End If
Do While Not rsNeg.EOF
    If rsNeg!OK = "1" Then
        lblPersNombre = rsNeg!Persona
        lblPersNombre.Tag = rsNeg!cPersCod
    End If
    rsNeg.MoveNext
Loop
rsNeg.MoveFirst
txtMovDesc = frmOpeNegVentanilla.vsMotivo
cmdNegocio.Enabled = True
Exit Sub
ErrNegocio:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
   If cmdNegocio.Visible Then
      cmdNegocio.Enabled = True
   End If
End Sub

Private Sub fgObjDest_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMovDesc.SetFocus
End If
End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Set oCtaIf = New NCajaCtaIF
Set oOpe = New DOperacion
Set oCaja = New nCajaGeneral

CentraForm Me
StabCuentas.Tab = 0
txtFechaDoc = gdFecSis
txtFechamov = gdFecSis
'Me.Caption = gsOpeDesc
Me.Caption = Mid(gsOpeDesc, 20, 28)

FraTipoCambio.Visible = False
If Mid(gsOpeCod, 3, 1) = gMonedaExtranjera Then
    'lblTCFijo = Format(gnTipCambio, gsFormatoNumeroView)
    lblTCFijo = Format(gnTipCambio, "##,###,##0.000")
      FraTipoCambio.Visible = True
    txtImporte.BackColor = &HC0FFC0
End If
Me.Label13 = Label13 & gsSimbolo

Set rs = oOpe.CargaOpeDoc(gsOpeCod)
Do While Not rs.EOF
   cboDocumento.AddItem Mid(rs!cDocDesc & Space(100), 1, 100) & Trim(rs!nDocTpo)
   rs.MoveNext
Loop
cboDocumento.AddItem "NINGUNO" & Space(105)
RSClose rs
CambiaTamañoCombo cboDocumento
txtObjOrig.rs = oOpe.GetRsOpeObj(gsOpeCod, "0")

If lbDeposito Then
    txtBuscarDest.psRaiz = "Cuentas Contables"
    txtBuscarDest.rs = oOpe.CargaOpeCta(gsOpeCod, "H", "0")

    StabCuentas.TabCaption(0) = "Destino"
    StabCuentas.TabCaption(1) = "Origen"
    Select Case gsOpeCod
        Case gOpeCGOpeBancosDepDivBancosMN, gOpeCGOpeBancosDepDivBancosME, _
            gOpeCGOpeCMACDepDivMN, gOpeCGOpeCMACDepDivME
    End Select
Else
    StabCuentas.TabCaption(0) = "Origen"
    StabCuentas.TabCaption(1) = "Destino"
    
    txtBuscarDest.psRaiz = "Cuentas Contables"
    txtBuscarDest.rs = oOpe.CargaOpeCta(gsOpeCod, "D", "0")
    
    Select Case gsOpeCod
        Case gOpeCGOpeBancosRetDivBancosMN, gOpeCGOpeBancosRetDivBancosME, _
            gOpeCGOpeCMACRetDivMN, gOpeCGOpeCMACRetDivME
        
        Case gOpeCGOpeGastComBancosMN, gOpeCGOpeGastComBancosME, gOpeCGOpeCMACGastosComMN, gOpeCGOpeCMACGastosComME
    End Select
End If
'Operaciones Comisión en Negocio
Set rs = oOpe.GetOperacionRSRefer(gsOpeCod, 5, gbBitCentral)
If Not rs.EOF Then
    Do While Not rs.EOF
        cboComision.AddItem Mid(rs!cOpeDesc & Space(100), 1, 100) & Trim(rs!cOpeCodRef)
        rs.MoveNext
    Loop
    cboComision.AddItem "NINGUNO" & Space(105)
End If
RSClose rs
If cboComision.ListCount = 0 Then 'And gbBitCentral
   cboComision.Visible = False
End If
cmdNegocio.Visible = True
End Sub

Private Sub txtBuscarDest_EmiteDatos()
Dim oPend As New DMov
If txtObjOrig = "" Then Exit Sub
lblDescCtaDest = txtBuscarDest.psDescripcion
If txtBuscarDest.Text <> "" Then
    AsignaCtaObjFlex txtBuscarDest.Text
    If fgObjDest.Visible Then fgObjDest.SetFocus
    If txtBuscarDest.Text <> "" Then
       If oPend.CuentaEsPendiente(txtBuscarDest, lsClaseCta) Then
          If (lsClaseCta = "D" And lbDeposito) Or (lsClaseCta = "A" And Not lbDeposito) Then
             lbRegulaPendiente = True
          Else
             lbRegulaPendiente = False
          End If
       End If
    End If
End If
Set oPend = Nothing
End Sub

Private Sub txtFechaDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAceptar.SetFocus
End If
End Sub

Private Sub txtFechaDoc_Validate(Cancel As Boolean)
If ValFecha(txtFechaDoc) = False Then
    Cancel = True
End If
End Sub

Private Sub txtFechamov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If FraTipoCambio.Visible Then
        txtTCBanco.SetFocus
    Else
        StabCuentas.Tab = 0
        If txtObjOrig.Enabled Then txtObjOrig.SetFocus
    End If
End If
End Sub

Private Sub txtFechamov_Validate(Cancel As Boolean)
If ValFecha(txtFechamov) = False Then
    Cancel = True
End If
End Sub

Private Sub txtImporte_GotFocus()
fEnfoque txtImporte
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtImporte, KeyAscii, 15, 2)
If KeyAscii = 13 Then
    cboDocumento.SetFocus
End If
End Sub

Private Sub txtImporte_LostFocus()
If Val(txtImporte) = 0 Then txtImporte = 0
txtImporte = Format(txtImporte, "#,#0.00")
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtImporte.SetFocus
End If
End Sub

Private Sub txtNroDoc_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    txtFechaDoc.SetFocus
End If
End Sub

Private Sub txtObjOrig_EmiteDatos()
Dim oContFunct As NContFunciones
Dim nCont As Integer
Set oContFunct = New NContFunciones
lblDescbanco = ""
lblDescCtabanco = ""
If Len(txtObjOrig) > 15 Then
    lblDescbanco = oCtaIf.NombreIF(Mid(txtObjOrig, 4, 13))
    lblDescCtabanco = oCtaIf.EmiteTipoCuentaIF(Mid(txtObjOrig, 18, 10)) + " " + txtObjOrig.psDescripcion
    lsCtaContBanco = oOpe.EmiteOpeCta(gsOpeCod, IIf(lbDeposito, "D", "H"), , txtObjOrig, CtaOBjFiltroIF)
    If lsCtaContBanco = "" Then
        MsgBox "Institución Financiera no tiene definida Cuenta Contable", vbInformation, "Aviso"
    End If
    
    If txtBuscarDest <> "" Then
        txtBuscarDest_EmiteDatos
        'For nCont = 1 To Me.fgObjDest.Rows - 1
        '    If nVal(fgObjDest.TextMatrix(nCont, 4)) = ObjEntidadesFinancieras And fgObjDest.TextMatrix(nCont, 1) = "" Then
        '            fgObjDest.TextMatrix(fgObjDest.Row, 1) = txtObjOrig
        '            fgObjDest.TextMatrix(fgObjDest.Row, 2) = lblDescbanco
        '            If fgObjDest.TextMatrix(fgObjDest.Row, 5) = 1 Then
        '                fgObjDest.TextMatrix(fgObjDest.Row, 3) = oContFunct.GetFiltroObjetos(ObjEntidadesFinancieras, txtBuscarDest, Mid(txtObjOrig, 4, 13), False)
        '            Else
        '                fgObjDest.TextMatrix(fgObjDest.Row, 3) = oContFunct.GetFiltroObjetos(ObjEntidadesFinancieras, txtBuscarDest, txtObjOrig, False)
        '            End If
        '    End If
        'Next
    End If
    StabCuentas.Tab = 1
    If txtBuscarDest.Enabled Then txtBuscarDest.SetFocus
End If
Set oContFunct = Nothing
End Sub

Private Sub txtTCBanco_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTCBanco, KeyAscii, 8, 4)
If KeyAscii = 13 Then
    StabCuentas.Tab = 0
    If txtObjOrig.Enabled Then txtObjOrig.SetFocus
End If
End Sub
Private Sub AsignaCtaObjFlex(ByVal psCtaContCod As String)
Dim sql As String
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim lsRaiz As String
Dim oDescObj As ClassDescObjeto
Dim UP As UPersona
Dim lsFiltro As String
Dim oRHAreas As DActualizaDatosArea
Dim oCtaCont As DCtaCont
Dim oCtaIf As NCajaCtaIF
Dim oEfect As Defectivo
Dim oContFunct As NContFunciones

Set oContFunct = New NContFunciones
Set oEfect = New Defectivo
Set oCtaIf = New NCajaCtaIF
Set oRHAreas = New DActualizaDatosArea
Set oDescObj = New ClassDescObjeto
Set oCtaCont = New DCtaCont
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset

fgObjDest.Clear
fgObjDest.FormaCabecera
fgObjDest.Rows = 2

Set rs1 = oCtaCont.CargaCtaObj(psCtaContCod, , True)
If Not rs1.EOF And Not rs1.BOF Then
    Do While Not rs1.EOF
        lsRaiz = ""
        lsFiltro = ""
        Select Case Val(rs1!cObjetoCod)
            Case ObjCMACAgencias
                Set rs = oRHAreas.GetAgencias(rs1!cCtaObjFiltro)
            Case ObjCMACAgenciaArea
                lsRaiz = "Unidades Organizacionales"
                Set rs = oRHAreas.GetAgenciasAreas(rs1!cCtaObjFiltro)
            Case ObjCMACArea
                Set rs = oRHAreas.GetAreas(rs1!cCtaObjFiltro)
            Case ObjEntidadesFinancieras
                If txtBuscarDest = Left(lsCtaContBanco, Len(txtBuscarDest)) Then
                    lsRaiz = "Cuentas de Entidades Financieras"
                    If rs1!nCtaObjNiv = 1 Then
                       Set rs = oCtaIf.GetCtasInstFinancieras(rs1!cCtaObjFiltro, psCtaContCod)
                    Else
                       Set rs = oCtaIf.CargaCtasIF(Mid(psCtaContCod, 3, 1), rs1!cCtaObjFiltro)
                    End If
                Else
                    lsFiltro = oContFunct.GetFiltroObjetos(Val(rs1!cObjetoCod), psCtaContCod, txtObjOrig, False)
                    'lsFiltro = oContFunct.GetFiltroObjetos(Val(rs1!cObjetoCod), psCtaContCod, Mid(txtObjOrig, 4, 13), False)
                    fgObjDest.AdicionaFila
                    fgObjDest.TextMatrix(fgObjDest.row, 1) = txtObjOrig
                    fgObjDest.TextMatrix(fgObjDest.row, 2) = lblDescbanco
                    fgObjDest.TextMatrix(fgObjDest.row, 3) = lsFiltro
                    fgObjDest.TextMatrix(fgObjDest.row, 4) = rs1!cObjetoCod
                    fgObjDest.TextMatrix(fgObjDest.row, 5) = rs1!nCtaObjNiv
                    Set rs = Nothing
                End If
            Case ObjDescomEfectivo
                lsRaiz = "Denominación"
                Set rs = oEfect.GetBilletajes(rs1!cCtaObjFiltro)
            Case ObjPersona
                Set rs = Nothing
            Case Else
                Set rs = GetObjetos(Val(rs1!cObjetoCod), False)
        End Select
        If Not rs Is Nothing Then
            If rs.State = adStateOpen Then
                If Not rs.EOF And Not rs.BOF Then
                    If rs.RecordCount > 1 Then
                        oDescObj.Show rs, "", lsRaiz
                        If oDescObj.lbOk Then
                            lsFiltro = oContFunct.GetFiltroObjetos(Val(rs1!cObjetoCod), psCtaContCod, oDescObj.gsSelecCod, False)
                            
                            fgObjDest.AdicionaFila
                            fgObjDest.TextMatrix(fgObjDest.row, 1) = oDescObj.gsSelecCod
                            fgObjDest.TextMatrix(fgObjDest.row, 2) = oDescObj.gsSelecDesc
                            fgObjDest.TextMatrix(fgObjDest.row, 3) = lsFiltro
                            fgObjDest.TextMatrix(fgObjDest.row, 4) = rs1!cObjetoCod
                        Else
                            txtBuscarDest = ""
                            lblDescCtaDest = ""
                            Exit Do
                        End If
                    Else
                        fgObjDest.AdicionaFila
                        fgObjDest.TextMatrix(fgObjDest.row, 1) = rs1!cObjetoCod
                        fgObjDest.TextMatrix(fgObjDest.row, 2) = rs1!cObjetoDesc
                        fgObjDest.TextMatrix(fgObjDest.row, 3) = lsFiltro
                        fgObjDest.TextMatrix(fgObjDest.row, 4) = rs1!cObjetoCod
                    End If
                End If
            End If
        Else
            If Val(rs1!cObjetoCod) = ObjPersona Then
                Set UP = frmBuscaPersona.Inicio
                If Not UP Is Nothing Then
                    fgObjDest.AdicionaFila
                    fgObjDest.TextMatrix(fgObjDest.row, 1) = UP.sPersCod
                    fgObjDest.TextMatrix(fgObjDest.row, 2) = UP.sPersNombre
                    fgObjDest.TextMatrix(fgObjDest.row, 3) = ""
                    fgObjDest.TextMatrix(fgObjDest.row, 4) = rs1!cObjetoCod
                End If
            End If
        End If
        rs1.MoveNext
    Loop
End If
rs1.Close
Set rs1 = Nothing
Set oDescObj = Nothing
Set UP = Nothing
Set oCtaCont = Nothing
Set oCtaIf = Nothing
Set oEfect = Nothing
End Sub
