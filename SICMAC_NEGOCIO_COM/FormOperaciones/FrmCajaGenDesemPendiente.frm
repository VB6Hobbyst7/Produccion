VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmCajaGenDesemPendiente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desembolso  Caja Chica"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   Icon            =   "FrmCajaGenDesemPendiente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   6
      Top             =   750
      Width           =   7370
      Begin MSComctlLib.ListView lvwVaucher 
         Height          =   3645
         Left            =   60
         TabIndex        =   9
         Top             =   120
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   6429
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro Vaucher"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha "
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Monto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Area"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Agenica"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "CodigoPersona"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Cuenta"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "MovNro"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame fraPeriodo1 
      Caption         =   "Periodo"
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
      Height          =   765
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7365
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   6000
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mskFechaF 
         Height          =   330
         Left            =   3840
         TabIndex        =   1
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFechaI 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
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
         Left            =   2760
         TabIndex        =   4
         Top             =   280
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
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
         TabIndex        =   3
         Top             =   280
         Width           =   1110
      End
   End
End
Attribute VB_Name = "FrmCajaGenDesemPendiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdBuscar_Click()
    
    Dim oCaja As COMDCajaGeneral.DCOMDocumento
    Dim rs As New ADODB.Recordset
    Dim lista As ListItem
    
    If ValidarFechas = False Then Exit Sub
    
    Set oCaja = New COMDCajaGeneral.DCOMDocumento
        Set rs = oCaja.CargaDocVaucherCtaPendiente(mskFechaI.Text, mskFechaF.Text)
    Set oCaja = Nothing
    
    lvwVaucher.ListItems.Clear
    If Not (rs.EOF And rs.BOF) Then
        Do Until rs.EOF
            Set lista = lvwVaucher.ListItems.Add(, , rs(0))
            lista.SubItems(1) = PstaNombre(rs(1), False)
            lista.SubItems(2) = rs(2)
            lista.SubItems(3) = rs(3)
            lista.SubItems(4) = rs(4)
            lista.SubItems(5) = rs(5)
            lista.SubItems(6) = rs(6)
            lista.SubItems(7) = rs(7)
            lista.SubItems(8) = rs(8)
            rs.MoveNext
        Loop
    Else
        MsgBox "No existen Datos", vbInformation, "Aviso"
    End If
    
End Sub


Private Sub CmdGrabar_Click()
    Dim i As Integer
    
    Dim lsCadImp As String
    Dim lsVaucher As String
    Dim lsPersNom As String
    Dim lnMonto As Double
    Dim ldfecha As Date
    Dim lsPersCod As String
    Dim lsArea As String
    Dim lsAge As String
    Dim lsnMovNro As String
    Dim lsCta As String
    Dim lsProcNro As String
    Dim lnSaldo As Currency
    Dim lnMovNroRef As Long
    Dim nResp As Integer
    Dim lbOk As Boolean
    
    Dim oCajaN As COMNCajaGeneral.NCajeroImp 'Forma la cadena de impresion
    Dim oCaja As COMDCajaGeneral.DCOMDocumento ' actualizar saldo
    Dim oMov As COMDMov.DCOMMov 'Generar Mov
    Dim oCajaCH As COMNCajaGeneral.NCOMCajaCtaIF
    
    Dim nFicSal As Integer
    
    i = lvwVaucher.SelectedItem.Index
    lsVaucher = lvwVaucher.ListItems.iTem(i).Text
    lsPersNom = lvwVaucher.ListItems.iTem(i).SubItems(1)
    ldfecha = lvwVaucher.ListItems.iTem(i).SubItems(2)
    lnMonto = lvwVaucher.ListItems.iTem(i).SubItems(3)
    lsArea = lvwVaucher.ListItems.iTem(i).SubItems(4)
    lsAge = lvwVaucher.ListItems.iTem(i).SubItems(5)
    lsPersCod = lvwVaucher.ListItems.iTem(i).SubItems(6)
    lsCta = lvwVaucher.ListItems.iTem(i).SubItems(7)
    lnMovNroRef = lvwVaucher.ListItems.iTem(i).SubItems(8)
    

    
    Set oCaja = New COMDCajaGeneral.DCOMDocumento
        lnSaldo = oCaja.GetDatosCajaChica(lsArea, lsAge, gSaldoActual)
    Set oCaja = Nothing
    
    Set oCaja = New COMDCajaGeneral.DCOMDocumento
        lsProcNro = oCaja.GetDatosCajaChica(lsArea, lsAge, gNroCajaChica)
    Set oCaja = Nothing
     
    Set oMov = New COMDMov.DCOMMov
        lsnMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oMov = Nothing
    
    'If lnSaldo >= lnMonto Then
    Set oCajaCH = New COMNCajaGeneral.NCOMCajaCtaIF
        nResp = oCajaCH.GrabaDesembolsoVaucherCH(lsnMovNro, lsPersCod, gsFormatoFecha, gOtrOpeEgresosDesemCajaChica, "Desembolso de Cta Pendiente", lnMonto, lsArea, lsAge, lsProcNro, gdFecSis, lsVaucher, lnMovNroRef, lsCta, lnSaldo)
        
        If nResp = 1 Then
            MsgBox " Voucher Emitido Correctamente", vbInformation, "Aviso"
        Else
            MsgBox " Fallo en Emision de Voucher", vbCritical, "Aviso"
        End If
    
    Set oCajaCH = Nothing
    'Else
'        MsgBox "Saldo es Menor a Monto ", vbInformation, "Aviso"
'        Exit Sub
'    End If
        
    Set oCajaN = New COMNCajaGeneral.NCajeroImp
        lsCadImp = oCajaN.ImprimeVaucherCtaPendiente(lsVaucher, lsPersNom, lnMonto, ldfecha, gsNomCmac, gsNomAge, gdFecSis)
    Set oCajaN = Nothing
    
    lbOk = True
         Do While lbOk
              nFicSal = FreeFile
              Open sLpt For Output As nFicSal
                  Print #nFicSal, lsCadImp
                  Print #nFicSal, ""
              Close #nFicSal
              If MsgBox("Desea Reimprimir Boleta ??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                  lbOk = False
              End If
         Loop

    CmdBuscar_Click
End Sub
Private Sub CmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    mskFechaI.Text = gdFecSis
    mskFechaF.Text = gdFecSis
End Sub

Public Function ValidarFechas() As Boolean
     ValidarFechas = True
     If Not IsDate(mskFechaI) = True Then
        MsgBox "Ingrese Fecha Correcta", vbInformation, "Aviso"
        ValidarFechas = False
        Exit Function
     End If
     
     If Not IsDate(mskFechaF) = True Then
        MsgBox "Ingrese Fecha Correcta", vbInformation, "Aviso"
        ValidarFechas = False
        Exit Function
     End If

     If CDate(mskFechaI) > CDate(mskFechaF) Then
        MsgBox "La Fecha incial debe de ser Menor", vbInformation, "Aviso"
        ValidarFechas = False
        Exit Function
     End If
     
End Function


