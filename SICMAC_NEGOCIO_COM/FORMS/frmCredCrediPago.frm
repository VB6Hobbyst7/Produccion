VERSION 5.00
Begin VB.Form frmCredCrediPago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CrediPago"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Credito"
      Height          =   1185
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6885
      Begin SICMACT.ActXCodCta ActXCodCta 
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         Texto           =   "Nro Cta"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar..."
         Height          =   360
         Left            =   3840
         TabIndex        =   22
         Top             =   480
         Width           =   900
      End
      Begin VB.Frame FraListaCred 
         Caption         =   "&Lista Creditos"
         Height          =   960
         Left            =   4920
         TabIndex        =   20
         Top             =   180
         Width           =   1875
         Begin VB.ListBox LstCred 
            Height          =   645
            ItemData        =   "frmCredCrediPago.frx":0000
            Left            =   75
            List            =   "frmCredCrediPago.frx":0002
            TabIndex        =   21
            Top             =   225
            Width           =   1725
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   6855
      Begin VB.Label LblMonCred 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1815
         TabIndex        =   18
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label lblmonto 
         AutoSize        =   -1  'True
         Caption         =   "Prestamo "
         Height          =   195
         Left            =   540
         TabIndex        =   17
         Top             =   225
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Deuda"
         Height          =   195
         Left            =   3780
         TabIndex        =   16
         Top             =   840
         Width           =   885
      End
      Begin VB.Label LblTotDeuda 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4950
         TabIndex        =   15
         Top             =   825
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Capital"
         Height          =   195
         Left            =   3780
         TabIndex        =   14
         Top             =   285
         Width           =   930
      End
      Begin VB.Label LblSalCap 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4950
         TabIndex        =   13
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Interes"
         Height          =   195
         Left            =   3780
         TabIndex        =   12
         Top             =   570
         Width           =   480
      End
      Begin VB.Label LblIntCred 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4950
         TabIndex        =   11
         Top             =   540
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Analista  "
         Height          =   195
         Left            =   540
         TabIndex        =   10
         Top             =   840
         Width           =   645
      End
      Begin VB.Label lblAnalista 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1815
         TabIndex        =   9
         Top             =   780
         Width           =   1020
      End
      Begin VB.Label lblProximaCuota 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1815
         TabIndex        =   8
         Top             =   510
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Prox Cuota"
         Height          =   195
         Left            =   540
         TabIndex        =   7
         Top             =   525
         Width           =   780
      End
   End
   Begin VB.Frame fraNuevoMet 
      Height          =   705
      Left            =   30
      TabIndex        =   3
      Top             =   2400
      Width           =   6825
      Begin VB.CheckBox chkUsarCrediPago 
         Height          =   285
         Left            =   3600
         TabIndex        =   4
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Utilizar el Sistema CrediPago ?"
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
         Height          =   405
         Left            =   450
         TabIndex        =   5
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4650
      TabIndex        =   2
      Top             =   3240
      Width           =   1035
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3570
      TabIndex        =   1
      Top             =   3240
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5730
      TabIndex        =   0
      Top             =   3240
      Width           =   1035
   End
End
Attribute VB_Name = "frmCredCrediPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PersCod As String
Dim fbExisteCrediPago As Boolean

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    LstCred.Clear
    If VerificaCuenta Then
        Call CargaDatos
    Else
        'Call LimpiaForm
        'Me.ActXCodCta.Age = gsCodAge
        'Call CtaCred.Enfoque(2)
    End If
End If
End Sub

Private Sub LimpiaForm()
    LblMonCred.Caption = ""
    Me.LblAnalista.Caption = ""
    Me.lblProximaCuota.Caption = ""
    LblSalCap.Caption = ""
    LblIntCred.Caption = ""
    LblTotDeuda.Caption = ""
    Me.ActXCodCta.NroCuenta = ""
    Me.ActXCodCta.CMAC = gsCodCMAC
    cmdAceptar.Enabled = False
    cmdSalir.Enabled = True
    Me.chkUsarCrediPago.value = 0
End Sub

Private Function ValidaDatos() As Boolean
Dim I, K As Integer
ValidaDatos = True
If Len(Me.ActXCodCta.NroCuenta) <> 18 Then
    MsgBox "No se ha ingresado correcta el Nro de Credito ", vbInformation, "Aviso"
    ValidaDatos = False
    Exit Function
End If
End Function

Private Sub CmdAceptar_Click()

Dim ObjCredi As COMNCredito.NCOMCrediPago
Set ObjCredi = New COMNCredito.NCOMCrediPago

If Not ValidaDatos Then Exit Sub
If fbExisteCrediPago Then
    If MsgBox("Desea Grabar Uso del Sistema CrediPago ? ", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Call ObjCredi.ActualizaCrediPago(Me.ActXCodCta.NroCuenta, Me.chkUsarCrediPago.value)
    End If
Else
    If MsgBox("Desea Grabar Uso del Sistema CrediPago ? ", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Call ObjCredi.InsertCrediPago(Me.ActXCodCta.NroCuenta, Me.chkUsarCrediPago.value, gsCodUser, gdFecSis)
    End If
End If
Call LimpiaForm

cmdAceptar.Enabled = False
Set ObjCredi = Nothing
End Sub


Private Sub CmdBuscar_Click()
Dim Pers As COMDPersona.UCOMPersona
Dim ObjCrediPago As COMNCredito.NCOMCrediPago
Dim rs As New ADODB.Recordset

Set Pers = New COMDPersona.UCOMPersona
Set Pers = frmBuscaPersona.Inicio
If Pers Is Nothing Then
Else
   PersCod = Pers.sPersCod
   Set ObjCrediPago = New COMNCredito.NCOMCrediPago
   Set rs = ObjCrediPago.Busca_x_Cliente(PersCod)
   LstCred.Clear
    Do While Not rs.EOF
        LstCred.AddItem rs!cCtaCod
        rs.MoveNext
    Loop
    rs.Close
    'LstCred.SetFocus
End If
Set Pers = Nothing
End Sub

Private Sub CmdCancelar_Click()
Call LimpiaForm
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And ActXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColPYMEEmp, False)
        If sCuenta <> "" Then
            ActXCodCta.NroCuenta = sCuenta
            ActXCodCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
ActXCodCta.CMAC = gsCodCMAC
'Frame1.Enabled = True
'Me.ActXCodCta.SetFocusProd
End Sub
Sub CargaDatos()

'Dim SQL1 As String
Dim RCred As New ADODB.Recordset
Dim nPenalidad As Double
Dim lsExoPenalidad As String
Dim ObjCredi As COMNCredito.NCOMCrediPago
'Dim nCred As COMNCredito.NCOMCredito
Dim nDeuda As Double
Dim RCrediPago As ADODB.Recordset

Set ObjCredi = New COMNCredito.NCOMCrediPago
'Set nCred = New COMNCredito.NCOMCredito
               
       Call ObjCredi.CargarDatosCrediPago(Me.ActXCodCta.NroCuenta, gdFecSis, RCred, RCrediPago, nDeuda)
       
       'Set RCred = ObjCredi.DatosCred(Me.ActXCodCta.NroCuenta)
        
        If RCred.EOF Then
            MsgBox "El Credito tiene que estar Aprobado", vbInformation, "Mensaje"
            Set RCred = Nothing
            Exit Sub
        End If
        
        LblMonCred.Caption = Format(RCred!nMontoApr, "#0.00")
        LblAnalista.Caption = IIf(IsNull(RCred!cCodAnalista), "", RCred!cCodAnalista)
        Me.lblProximaCuota.Caption = Str(RCred!nNroProxCuota) & " /" & Str(RCred!nCuotasApr)
        LblSalCap.Caption = Format(IIf(IsNull(RCred!nSaldoCap), 0, RCred!nSaldoCap), "#0.00")
        RCred.Close
        
        'Set RCred = ObjCredi.GetCrediPago(Me.ActXCodCta.NroCuenta)
    
    If RCrediPago.BOF And RCrediPago.EOF Then
        fbExisteCrediPago = False
        Me.chkUsarCrediPago.value = 0
    ElseIf RCrediPago!cCrediPago = "N" Then
        Me.chkUsarCrediPago.value = 0
        fbExisteCrediPago = True
    ElseIf RCrediPago!cCrediPago = "S" Then
        Me.chkUsarCrediPago.value = 1
        fbExisteCrediPago = True
    End If
    RCrediPago.Close
    
    'SQL1 = "SELECT SUM (nInteres + isnull(nMora+nIntGra,0) - isnull(nIntComPag,0) - isnull(nIntMorPag,0) - isnull(nIntGraPag,0) ) as Inter from Plandespag where cCodCta='" & CtaCred.Text & "'"
    'RCred.Open SQL1, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
    
    'nDeuda = nCred.MatrizInteresTotalesAFecha(Me.ActXCodCta.NroCuenta, nCred.RecuperaMatrizCalendarioPendiente(Me.ActXCodCta.NroCuenta), gdFecSis)
    LblIntCred.Caption = Format(nDeuda, "#0.00")
   ' RCred.Close
    'Set RCred = Nothing
    
    LblTotDeuda.Caption = CDbl(LblIntCred.Caption) + CDbl(LblSalCap.Caption)
    
    Me.cmdAceptar.Enabled = True
    Me.chkUsarCrediPago.Enabled = True
    Me.chkUsarCrediPago.SetFocus
End Sub

Private Sub LstCred_Click()
  If LstCred.Text <> "" Then
    ActXCodCta.NroCuenta = LstCred.Text
  Else
  
  End If
End Sub

Function VerificaCuenta() As Boolean
Dim RegCred As New ADODB.Recordset
Dim ObjCredi As COMNCredito.NCOMCrediPago
Set ObjCredi = New COMNCredito.NCOMCrediPago
Set RegCred = ObjCredi.EstadoCred(Me.ActXCodCta.NroCuenta)
    If Not RegCred.BOF And Not RegCred.EOF Then
        If RegCred!nPrdEstado = gColocEstCancelado Then
            MsgBox "Credito ya Esta Pagado", vbInformation, "Aviso"
            VerificaCuenta = False
            RegCred.Close
            Set RegCred = Nothing
            Exit Function
        Else
            If RegCred!nPrdEstado = gColocEstVigNorm Or _
               RegCred!nPrdEstado = gColocEstVigVenc Or _
               RegCred!nPrdEstado = gColocEstVigMor Or _
               RegCred!nPrdEstado = gColocEstRefVenc Or _
               RegCred!nPrdEstado = gColocEstRefNorm Or _
               RegCred!nPrdEstado = gColocEstRefMor Or _
               RegCred!nPrdEstado = gColocEstSug Then
                VerificaCuenta = True
                RegCred.Close
                Set RegCred = Nothing
                Exit Function
            Else
            
                MsgBox "Credito No Esta En Estado Pendiente", vbInformation, "Aviso"
                VerificaCuenta = False
                RegCred.Close
                Set RegCred = Nothing
                Exit Function
            End If
        End If
    Else
        MsgBox "Credito No Se Encuentra", vbInformation, "Aviso"
        VerificaCuenta = False
        RegCred.Close
        Set RegCred = Nothing
        Exit Function
    End If
    VerificaCuenta = True
End Function

Private Sub LstCred_GotFocus()
 Me.ActXCodCta.NroCuenta = LstCred.Text
End Sub

Private Sub LstCred_KeyPress(KeyAscii As Integer)
  If LstCred.Text <> "" And KeyAscii = 13 Then
    ActXCodCta.NroCuenta = LstCred.Text
  Else
  
  End If
End Sub
