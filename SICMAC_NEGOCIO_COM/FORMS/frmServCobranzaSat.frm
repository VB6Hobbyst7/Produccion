VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmServCobranzaSat 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   105
      TabIndex        =   12
      Top             =   6030
      Width           =   1110
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8670
      TabIndex        =   11
      Top             =   6030
      Width           =   1110
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   7485
      TabIndex        =   10
      Top             =   6030
      Width           =   1110
   End
   Begin VB.Frame fraPapeleta 
      Caption         =   "Papeletas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3360
      Left            =   75
      TabIndex        =   5
      Top             =   1560
      Width           =   9810
      Begin MSComctlLib.ListView lstPapeleta 
         Height          =   3135
         Left            =   120
         TabIndex        =   19
         Top             =   255
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Papeleta"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Infracción"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Placa"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Licencia"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Infracción"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Importe"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Costas"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Gastos"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "TOTAL"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Frame fraDatoBuscar 
      Height          =   1470
      Left            =   1860
      TabIndex        =   4
      Top             =   75
      Width           =   3315
      Begin VB.TextBox txtPapeleta 
         Height          =   350
         Left            =   180
         MaxLength       =   7
         TabIndex        =   9
         Top             =   585
         Width           =   1230
      End
      Begin VB.TextBox txtPlaca 
         Height          =   350
         Left            =   180
         MaxLength       =   7
         TabIndex        =   8
         Top             =   585
         Width           =   1230
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   350
         Left            =   2070
         TabIndex        =   7
         Top             =   585
         Width           =   1020
      End
      Begin MSMask.MaskEdBox txtLicencia 
         Height          =   345
         Left            =   180
         TabIndex        =   6
         Top             =   585
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "CC-#######"
         PromptChar      =   "_"
      End
   End
   Begin VB.Frame fraBusqueda 
      Caption         =   "Tipo Búsqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1470
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   1665
      Begin VB.OptionButton optBuscar 
         Caption         =   "Papeleta"
         Height          =   375
         Index           =   2
         Left            =   330
         TabIndex        =   3
         Top             =   885
         Width           =   1080
      End
      Begin VB.OptionButton optBuscar 
         Caption         =   "N° Placa"
         Height          =   375
         Index           =   1
         Left            =   315
         TabIndex        =   2
         Top             =   585
         Width           =   1080
      End
      Begin VB.OptionButton optBuscar 
         Caption         =   "Lincencia"
         Height          =   375
         Index           =   0
         Left            =   330
         TabIndex        =   1
         Top             =   270
         Width           =   1080
      End
   End
   Begin VB.Label lblTotalPagar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   8205
      TabIndex        =   18
      Top             =   5595
      Width           =   1560
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total a Pagar :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   6660
      TabIndex        =   17
      Top             =   5595
      Width           =   1575
   End
   Begin VB.Label lblComision 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   8205
      TabIndex        =   16
      Top             =   5280
      Width           =   1560
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comisión :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   6660
      TabIndex        =   15
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label lblSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   8205
      TabIndex        =   14
      Top             =   4950
      Width           =   1560
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sub - Total :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   6660
      TabIndex        =   13
      Top             =   4950
      Width           =   1575
   End
End
Attribute VB_Name = "frmServCobranzaSat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function GetValorComision() As Double
Dim rsPar As New ADODB.Recordset
Dim oCap As COMNCaptaGenerales.NCOMCaptaDefinicion
Set oCap = New COMNCaptaGenerales.NCOMCaptaDefinicion

Set rsPar = oCap.GetTarifaParametro(gServCobSATTInfraccion, gMonedaNacional, gCostoComServSATInfraccion)

Set oCap = Nothing
If rsPar.EOF And rsPar.BOF Then
    GetValorComision = 0
Else
    GetValorComision = rsPar("nParValor")
End If
rsPar.Close
Set rsPar = Nothing
End Function

Private Sub LimpiaControles()
txtLicencia.Text = "__-_______"
txtPlaca = ""
txtPapeleta = ""
lstPapeleta.ListItems.Clear
lblSubTotal = "0.00"
lblTotalPagar = "0.00"
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
optBuscar(0).value = True
End Sub

Private Sub cmdBuscar_Click()
Dim sDato As String
Dim nTipo As Integer
If optBuscar(0).value Then
    sDato = Trim(txtLicencia.Text)
    sDato = Replace(sDato, "-", "", 1, , vbTextCompare)
    nTipo = 0
ElseIf optBuscar(1).value Then
    sDato = Trim(txtPlaca.Text)
    nTipo = 1
ElseIf optBuscar(2).value Then
    sDato = Trim(txtPapeleta.Text)
    nTipo = 2
End If

If sDato = "" Then
    MsgBox "Debe ingresar el dato a buscar", vbInformation, "Aviso"
    Exit Sub
End If

Dim oServ As COMNCaptaServicios.NCOMCaptaServicios

Dim rsSat As New ADODB.Recordset
Dim L As MSComctlLib.ListItem
Dim sPapeleta As String, dInfraccion As Date
Dim sPlaca As String, sLicencia As String
Dim sInfraccion As String
Dim nMonto As Double, nGastos As Double, nCostas As Double, nReincide As Double

Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    Set rsSat = oServ.GetServSATPapeletas(sDato, nTipo)
Set oServ = Nothing
If rsSat.EOF And rsSat.BOF Then
    MsgBox "Dato NO Encontrado o NO Posee Infracciones", vbInformation, "Aviso"
Else
    Do While Not rsSat.EOF
        sPapeleta = rsSat("sPapeleta")
        dInfraccion = rsSat("dFecInf")
        sPlaca = rsSat("cPlaca")
        sLicencia = rsSat("sLicCond")
        sInfraccion = rsSat("sTipo") & "-" & rsSat("sFalta")
        nMonto = rsSat("nImporte")
        nCostas = rsSat("nValCostas")
        nGastos = rsSat("nValGastAdm")
        Set L = lstPapeleta.ListItems.Add(, , sPapeleta)
        L.SubItems(1) = Format$(dInfraccion, "dd/mm/yyyy")
        L.SubItems(2) = sPlaca
        L.SubItems(3) = sLicencia
        L.SubItems(4) = sInfraccion
        L.SubItems(5) = Format$(nMonto, "#,##0.00")
        L.SubItems(6) = Format$(nCostas, "#,##0.00")
        L.SubItems(7) = Format$(nGastos, "#,##0.00")
        L.SubItems(8) = Format$(nMonto + nCostas + nGastos, "#,##0.00")
        Set L = Nothing
        rsSat.MoveNext
    Loop
    cmdCancelar.Enabled = True
End If
rsSat.Close
Set rsSat = Nothing
End Sub

Private Sub cmdCancelar_Click()
LimpiaControles
End Sub

Private Sub cmdGrabar_Click()
Dim sPapeleta As String, sLicencia As String
Dim nMonto As Double, nMontoComision As Double, nCostas As Double, nGastosA As Double

Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios


If MsgBox("¿Desea grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then

    Dim clsMov As COMNContabilidad.NCOMContFunciones
    Dim sMovNro As String
    Dim L As MSComctlLib.ListItem
    Dim lsBoleta As String
    Dim nFicSal As Integer
    On Error GoTo ErrGraba
    
    For Each L In lstPapeleta.ListItems
        If L.Checked Then

            Set clsMov = New COMNContabilidad.NCOMContFunciones
            
            sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            Set clsMov = Nothing
            sPapeleta = L.Text
            sLicencia = L.SubItems(3)
            nMonto = CDbl(L.SubItems(5))
            nMontoComision = CDbl(lblComision.Caption)
            nCostas = CDbl(L.SubItems(6))
            nGastosA = CDbl(L.SubItems(7))
            clsServ.CapCobranzaServicios sMovNro, COMDConstantes.gServCobSATTInfraccion, sPapeleta, sLicencia, nMonto, gsNomCmac, gsNomAge, sLpt, nMontoComision, , nGastosA, nCostas, 0, gsCodCMAC, lsBoleta
            
            If Trim(lsBoleta) <> "" Then
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                    Print #nFicSal, lsBoleta
                    Print #nFicSal, ""
                Close #nFicSal
            End If
            
        End If
    Next
    LimpiaControles
End If
Set clsServ = Nothing
Exit Sub
ErrGraba:
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
LimpiaControles
lblComision = Format$(GetValorComision(), "#,##0.00")
lblTotalPagar = lblComision
End Sub

Private Sub lstPapeleta_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim nMonto As Double
Dim nSubTotal As Double, ntotal As Double
nMonto = CDbl(Item.SubItems(8))
nSubTotal = CDbl(lblSubTotal)
ntotal = CDbl(lblTotalPagar)

If Item.Checked Then
    lblSubTotal = Format$(nSubTotal + nMonto, "#,##0.00")
    lblTotalPagar = Format$(ntotal + nMonto, "#,##0.00")
Else
    lblSubTotal = Format$(nSubTotal - nMonto, "#,##0.00")
    lblTotalPagar = Format$(ntotal - nMonto, "#,##0.00")
End If
If CDbl(lblSubTotal) > 0 Then
    cmdGrabar.Enabled = True
Else
    cmdGrabar.Enabled = False
End If
End Sub

Private Sub optBuscar_Click(Index As Integer)
Select Case Index
    Case 0
        txtLicencia.Visible = True
        txtPlaca.Visible = False
        txtPapeleta.Visible = False
    Case 1
        txtLicencia.Visible = False
        txtPlaca.Visible = True
        txtPapeleta.Visible = False
    Case 2
        txtLicencia.Visible = False
        txtPlaca.Visible = False
        txtPapeleta.Visible = True
End Select
End Sub

Private Sub optBuscar_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Index
        Case 0
            txtLicencia.SetFocus
        Case 1
            txtPlaca.SetFocus
        Case 2
            txtPapeleta.SetFocus
    End Select
End If
End Sub

Private Sub txtLicencia_GotFocus()
With txtLicencia
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtLicencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdbuscar.SetFocus
End If
End Sub
    
