VERSION 5.00
Begin VB.Form FrmMantCargoAutomatico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento Cargo Automatico"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraBotones 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   5880
      Width           =   7455
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3240
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   4560
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame FraDatos 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   7455
      Begin VB.ListBox lstCuenta 
         Height          =   1425
         Left            =   5160
         TabIndex        =   12
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label LblMoneda 
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label LblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   3930
      End
      Begin VB.Label LblSaldo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   960
         TabIndex        =   6
         Top             =   630
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Saldo "
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
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente "
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
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame FraCuenta 
      Height          =   840
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7440
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
         Height          =   405
         Left            =   5610
         TabIndex        =   1
         Top             =   270
         Width           =   1485
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   405
         Left            =   150
         TabIndex        =   2
         Top             =   255
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   714
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
   Begin SICMACT.FlexEdit Flex 
      Height          =   2655
      Left            =   600
      TabIndex        =   14
      Top             =   3120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4683
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Cta Credito-Cta Cargo-Orden-chk"
      EncabezadosAnchos=   "300-2000-2000-1200-500"
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-4"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-4"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-R-C-C"
      FormatosEdit    =   "0-0-0-0-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Cargo Automatico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   0
      Left            =   2880
      TabIndex        =   15
      Top             =   2760
      Width           =   1875
   End
End
Attribute VB_Name = "FrmMantCargoAutomatico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CodPersona As String

Private Sub ActxCta_KeyPress(KeyAscii As Integer)

'Dim rs As ADODB.Recordset
'Dim Col As COMDColocPig.DCOMColPFunciones
Dim oCred As COMDCredito.DCOMCredDoc
Set oCred = New COMDCredito.DCOMCredDoc
'Set Col = New COMDColocPig.DCOMColPFunciones
Dim rsCred As ADODB.Recordset
Dim rsCol As ADODB.Recordset
Dim rsCue As ADODB.Recordset

If KeyAscii = 13 Then
    If Len(ActxCta.NroCuenta) = 18 Then
        Call oCred.CargaDatosCuentaCAutomatico(ActxCta.NroCuenta, rsCred, rsCol, rsCue)
        Set oCred = Nothing
        'Set rs = Cred.RecuperaDatosMantCargoAutomatico(ActxCta.NroCuenta)
        If rsCred.EOF And rsCred.BOF Then
            MsgBox "La Cuenta no existe", vbInformation, "Mensaje"
            Exit Sub
        End If
        CodPersona = rsCred!cPersCod
        Me.LblCliente = rsCred!nombre
        Me.LblSaldo = Format(rsCred!Saldo, "##0.00")
        LblMoneda = IIf(Mid(ActxCta.NroCuenta, 9, 1) = "1", "SOLES", "DOLARES")
        'Set rs = Cred.RecuperaDatosColocCargoAutoma(ActxCta.NroCuenta)
        
        'If rs.EOF And rs.BOF Then
             'MsgBox "No existen Cuentas para Cargo Automatico", vbInformation, "AVISO"
        'End If
        
        While Not rsCol.EOF
            Flex.AdicionaFila
            Flex.TextMatrix(Flex.Rows - 1, 1) = rsCol!cCtaCod
            Flex.TextMatrix(Flex.Rows - 1, 2) = rsCol!cCodCtaAho
            Flex.TextMatrix(Flex.Rows - 1, 3) = rsCol!nOrden
            Flex.TextMatrix(Flex.Rows - 1, 4) = rsCol!cEstado
            rsCol.MoveNext
        Wend
        
        lstCuenta.Clear
        'Set rs = Col.dObtieneCuentasPersona(CodPersona, "1000", IIf(Mid(ActxCta.NroCuenta, 9, 1) = "1", "1", "2"))
        'Set rs = Cred.ObtieneCuentasPersona(CodPersona, "1000", IIf(Mid(ActxCta.NroCuenta, 9, 1) = "1", "1", "2"))
        While Not rsCue.EOF
            lstCuenta.AddItem rsCue!cCtaCod & Space(50) & rsCue!nSaldo
            rsCue.MoveNext
        Wend

    Else
        MsgBox "Nro de Cuenta Incompleto", vbInformation, "AVISO"
    End If
End If
'Set rs = Nothing
'Set oCred = Nothing
'Set Col = Nothing
End Sub

'Private Sub CmdAgregar_Click()
'Dim sCodPers As String
'Dim CtaCod As String
'Dim Pers As COMDPersona.UCOMPersona
'Set Pers = New COMDPersona.UCOMPersona
'Set Pers = frmBuscaPersona.Inicio
'
'If Pers Is Nothing Then
'CtaCod = ""
'Else
'    sCodPers = Pers.sPerscod
'    Call frmColPBuscaContrato.CargaListaContratos(sCodPers, "2020,2021,2022,2030,2031,2032")
'    frmColPBuscaContrato.Caption = "Credito Cliente"
'    frmColPBuscaContrato.Show 1
'
'End If
'CtaCod = Mid(frmColPBuscaContrato.lstContratos.Text, 1, 18)
'
'If CtaCod = "" Then
'    Exit Sub
'End If
'Me.ActxCta.NroCuenta = CtaCod
'Me.ActxCta.SetFocusCuenta
'
'End Sub

Private Sub CmdAceptar_Click()
Dim i As Integer
Dim oCred As COMDCredito.DCOMCredDoc
Dim MatCargoAuto As Variant

If Flex.Rows = 2 And Flex.TextMatrix(Flex.Rows - 1, 1) = "" Then
    MsgBox "No existen Cuentas para hacer Cargo Auotmatico", vbInformation, "AVISO"
    Exit Sub
End If

If MsgBox("Esta seguro de registrar la operación", vbQuestion + vbYesNo, "Confirmación") = vbNo Then Exit Sub

ReDim MatCargoAuto(Flex.Rows - 1, 4)

For i = 1 To Flex.Rows - 1
    MatCargoAuto(i, 1) = Flex.TextMatrix(i, 1)
    MatCargoAuto(i, 2) = Flex.TextMatrix(i, 2)
    MatCargoAuto(i, 3) = Flex.TextMatrix(i, 3)
    MatCargoAuto(i, 4) = Flex.TextMatrix(i, 4)
'    If Flex.TextMatrix(I, 1) <> "" Then
'        If oCred.ExisteCtaCargoAutoma(ActxCta.NroCuenta, Flex.TextMatrix(I, 2)) = True Then
'            Call oCred.ActualizaColocCargoAutoma(ActxCta.NroCuenta, Flex.TextMatrix(I, 2), Flex.TextMatrix(I, 3), IIf(Flex.TextMatrix(I, 4) = ".", 1, 0))
'        Else
'            Call oCred.InsertaColocCargoAutoma(ActxCta.NroCuenta, Flex.TextMatrix(I, 2), IIf(Flex.TextMatrix(I, 4) = ".", "1", "0"), Flex.TextMatrix(I, 3))
'       End If
'    End If
Next i
Set oCred = New COMDCredito.DCOMCredDoc
Call oCred.ActualizaColocCargoAutomaLote(ActxCta.NroCuenta, MatCargoAuto)
Set oCred = Nothing
Limpiar
Me.ActxCta.Enabled = True
CmdBuscar.Enabled = True
CmdBuscar.SetFocus
Call cmdCancelar_Click
End Sub

'Private Sub cmdAgregarCargo_Click()
'Dim I As Integer
'Dim sCodPers As String
'Dim CtaCod As String
'Dim sNombre As String
'Dim Pers As COMDPersona.UCOMPersona
'Set Pers = New COMDPersona.UCOMPersona
'Set Pers = frmBuscaPersona.Inicio
'
'If Pers Is Nothing Then
'    Exit Sub
'Else
'    sNombre = ""
'    CtaCod = ""
'    sCodPers = Pers.sPerscod
'    Call frmColPBuscaContrato.CargaListaCtas(sCodPers, "1000")
'    frmColPBuscaContrato.Caption = "Cuentas Cliente"
'    frmColPBuscaContrato.Show 1
'End If
'
'CtaCod = Mid(frmColPBuscaContrato.lstContratos.Text, 1, 18)
'sNombre = Mid(frmColPBuscaContrato.lstContratos.Text, 20, Len(frmColPBuscaContrato.lstContratos.Text))
'If Trim(CtaCod) = "" Then
'    Exit Sub
'End If
'
'For I = 0 To Flex.Rows - 1
'    If Trim(Flex.TextMatrix(I, 1)) = Trim(CtaCod) Then
'        MsgBox "Ya existe ese Numero de Cuenta", vbInformation, "AVISO"
'        Exit Sub
'    End If
'Next I
'
'If Flex.Rows = 2 And Len(Trim(Flex.TextMatrix(1, 1))) = 0 Then
'    Flex.TextMatrix(Flex.Rows - 1, 1) = CtaCod
'    Flex.TextMatrix(Flex.Rows - 1, 2) = PstaNombre(Trim(sNombre))
'    Flex.TextMatrix(Flex.Rows - 1, 3) = "1"
'Else
'    Flex.Rows = Flex.Rows + 1
'    Flex.TextMatrix(Flex.Rows - 1, 1) = CtaCod
'    Flex.TextMatrix(Flex.Rows - 1, 2) = PstaNombre(Trim(sNombre))
'    Flex.TextMatrix(Flex.Rows - 1, 3) = Flex.TextMatrix(Flex.Rows - 2, 3) + 1
'End If
'CmdAceptar.Enabled = True
'End Sub

Private Sub CmdBuscar_Click()
Dim sCodPers As String
Dim CtaCod As String
Dim Pers As COMDPersona.UCOMPersona
Set Pers = New COMDPersona.UCOMPersona
Set Pers = frmBuscaPersona.Inicio
CtaCod = ""
If Pers Is Nothing Then
    Exit Sub
Else
    CodPersona = Pers.sPersCod
    Call frmColPBuscaContrato.CargaListaContratos(CodPersona, "2020,2021,2022,2030,2031,2032", True)
    frmColPBuscaContrato.Caption = "Credito Cliente"
    frmColPBuscaContrato.Show 1
    CtaCod = Mid(frmColPBuscaContrato.lstContratos.Text, 1, 18)
    If CtaCod = "" Then
        Call cmdCancelar_Click
        Exit Sub
    End If
    Me.ActxCta.NroCuenta = CtaCod
    Me.ActxCta.SetFocusCuenta

End If
End Sub

Sub Limpiar()
    LblMoneda = ""
    LblCliente = ""
    LblSaldo = ""
    ActxCta.NroCuenta = ""
    Call Form_Load
End Sub

Private Sub cmdCancelar_Click()
lstCuenta.Clear
Limpiar

Flex.Rows = 2
Flex.TextMatrix(1, 1) = ""
Flex.TextMatrix(1, 2) = ""
Flex.TextMatrix(1, 3) = ""
Flex.TextMatrix(1, 4) = ""
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Flex_KeyPress(KeyAscii As Integer)
If Flex.Col = 3 Then
    'If Asc("0") < KeyAscii And KeyAscii >= Asc("9") Then
        Flex.TextMatrix(Flex.Row, Flex.Col) = Chr(KeyAscii)
    'End If
End If
End Sub

Private Sub Form_Load()
CodPersona = ""
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Me.ActxCta.CMAC = gsCodCMAC
Me.ActxCta.EnabledCMAC = False
Flex.Enabled = True
End Sub

Private Sub lstCuenta_DblClick()
Dim Ord As Integer
Dim i As Integer
If lstCuenta.ListIndex = -1 Then
    MsgBox "El Cliente no posee cuentas en " & IIf(Mid(ActxCta.NroCuenta, 9, 1) = "1", " Soles", " Dolares"), vbInformation, "AVISO"
    Exit Sub
End If
Ord = 0
For i = 1 To Flex.Rows - 1
    If Flex.TextMatrix(i, 4) = "." And CStr(Ord) < Flex.TextMatrix(i, 3) Then
        Ord = Flex.TextMatrix(i, 3)
    End If
Next i

For i = 1 To Flex.Rows - 1
    If Trim(Flex.TextMatrix(i, 2)) = Trim(Mid(lstCuenta.Text, 1, 18)) Then
        MsgBox "Ya se agregó esa cuenta", vbInformation, "AVISO"
        Exit Sub
    End If
Next i
Flex.AdicionaFila
If Flex.Rows = 2 And Flex.TextMatrix(Flex.Row, Flex.Col) = "" Then
    Flex.TextMatrix(Flex.Rows - 1, 1) = Me.ActxCta.NroCuenta
    Flex.TextMatrix(Flex.Rows - 1, 2) = Trim(Mid(lstCuenta.Text, 1, 18))
    Flex.TextMatrix(Flex.Rows - 1, 3) = Ord + 1
    Flex.TextMatrix(Flex.Rows - 1, 4) = 1
Else
    Flex.TextMatrix(Flex.Rows - 1, 1) = Me.ActxCta.NroCuenta
    Flex.TextMatrix(Flex.Rows - 1, 2) = Trim(Mid(lstCuenta.Text, 1, 18))
    Flex.TextMatrix(Flex.Rows - 1, 3) = Ord + 1
    Flex.TextMatrix(Flex.Rows - 1, 4) = 1
End If

End Sub

Private Sub lstCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call lstCuenta_DblClick
End If
End Sub
