VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCajaGenDesemPendienteExtorno 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extorno - Desembolso  Caja Chica"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   Icon            =   "FrmCajaGenDesemPendienteExtorno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Recibos Emitidos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2175
      Left            =   0
      TabIndex        =   2
      Top             =   30
      Width           =   7370
      Begin MSComctlLib.ListView lvwVaucher 
         Height          =   1845
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   3254
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
            Text            =   "Mov"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.CommandButton CmdExtornar 
      Caption         =   "&Extornar"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "FrmCajaGenDesemPendienteExtorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdExtornar_Click()
     Dim i As Integer
    
    Dim lsCadImp As String
    Dim lsVaucher As String
    Dim lsPersNom As String
    Dim lnMonto As Double
    Dim ldfecha As Date
    Dim lsPersCod As String
    Dim lsArea As String
    Dim lsAge As String
    Dim lsnMovNroExtorno As String
    Dim lsCta As String
    Dim lsProcNro As String
    Dim lnSaldo As Currency
    Dim lnMovNro As Long
    Dim nResp As Integer
    Dim lnMovNroAnt As Long
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
    lnMovNroAnt = lvwVaucher.ListItems.iTem(i).SubItems(8)
    
   
    Set oCaja = New COMDCajaGeneral.DCOMDocumento
        lnSaldo = oCaja.GetDatosCajaChica(lsArea, lsAge, gSaldoActual)
    Set oCaja = Nothing
    
    Set oCaja = New COMDCajaGeneral.DCOMDocumento
        lsProcNro = oCaja.GetDatosCajaChica(lsArea, lsAge, gNroCajaChica)
    Set oCaja = Nothing
     
    Set oMov = New COMDMov.DCOMMov
        lsnMovNroExtorno = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oMov = Nothing
    
       
        Set oCajaCH = New COMNCajaGeneral.NCOMCajaCtaIF
           nResp = oCajaCH.GrabaExtornoDesembolsoVaucher(gdFecSis, ldfecha, gsFormatoFecha, lnMovNroAnt, lsnMovNroExtorno, gCajaGenDesemPendienteExtorno, "Extorno de Desembolso de Cta Pendiente", lsArea, lsAge, lsProcNro, lnMonto)
        
            If nResp = 1 Then
               MsgBox "Recibo Extornado con exito", vbInformation, "Aviso"
               CargaEgresosCtaPendiente
            Else
               MsgBox "Datos No Extornados ", vbCritical, "Aviso"
               Exit Sub
            End If
        
     Set oCajaCH = Nothing
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Call CargaEgresosCtaPendiente
End Sub

Private Sub CargaEgresosCtaPendiente()
Dim oCaja As COMDCajaGeneral.DCOMDocumento
Dim rs As New ADODB.Recordset
Dim lista As ListItem
    
Set oCaja = New COMDCajaGeneral.DCOMDocumento
    
    Set rs = oCaja.CargaDocVaucherCtaPendienteExtorno(gdFecSis)
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
