VERSION 5.00
Begin VB.Form frmPreDesemyExtornoCompraDeuda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pre Desembolso - Compra de Deuda"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8550
   Icon            =   "frmPreDesemCompraDeuda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraCuenta 
      Height          =   840
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   8160
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   405
         Left            =   150
         TabIndex        =   16
         Top             =   240
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
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7080
      TabIndex        =   11
      Top             =   3480
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5640
      TabIndex        =   10
      Top             =   3480
      Width           =   1140
   End
   Begin VB.CommandButton CmdAprobar 
      Caption         =   "&Aprobar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   1140
   End
   Begin VB.Frame fraCredito 
      Caption         =   "Datos del Crédito"
      Enabled         =   0   'False
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
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8160
      Begin VB.TextBox txtDesem 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         TabIndex        =   17
         Top             =   1800
         Width           =   1815
      End
      Begin VB.ListBox lstCuentas 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   4680
         TabIndex        =   12
         Top             =   360
         Width           =   3300
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblFecha 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Desembolso a Comprar :"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblMontoApro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Monto Aprobado :"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblTpoCredito 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Crédito :"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "frmPreDesemyExtornoCompraDeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'***Nombre      : frmPreDesemCompraDeuda
'***Descripción : Formulario para Aprobar y Extornar el Pre - Desembolso de la creditos con destino compra de deuda
'***Creación    : ARLO 20180315, según TI-ERS 070-2017
'************************************************************************************************
Option Explicit
Private fnTipoAsig As Double

Public Sub Inicio(ByVal pnTipo As Integer)
    fnTipoAsig = pnTipo
    Select Case fnTipoAsig
        Case 1: Me.Caption = "Pre Desembolso - Compra de Deuda"
        Case 2: Me.Caption = "Extorno Pre Desembolso - Compra de Deuda"
    End Select
    
    Me.Show 1
End Sub
Private Sub CmdAprobar_Click()
    
        Dim oCred As New COMDCredito.DCOMCreditos
        Dim sMensaje As String
        Dim lcMovNro As String
        
        lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        
        If (fnTipoAsig = 1) Then
            sMensaje = "aprobado"
        Else
            sMensaje = "extornado"
        End If
        
        If Not ValidarDatos Then Exit Sub
        
            If (CDbl(lblMontoApro.Caption) < CDbl(txtDesem.Text)) Then
                MsgBox "Monto Solicitado debe ser mayor o igual al Saldo a Comprar", vbInformation, "Aviso"
                Exit Sub
            End If
         
        If MsgBox("Se va a Grabar la Operación, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
            
        oCred.InsertarPreDesembCompraDeuda ActxCta.NroCuenta, CDbl(lblMontoApro.Caption), CDbl(txtDesem.Text), gdFecSis, fnTipoAsig, lcMovNro
        
        MsgBox "El Nro. de crédito fue " & sMensaje & " con éxito", vbInformation, "Aviso"
        
        Call LimpiaPantalla
        
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub CargarDatos(ByVal psCtaCod As String)
    Dim rsCred As ADODB.Recordset
    Dim rsIfis As ADODB.Recordset
    Dim oCred As New COMDCredito.DCOMCreditos
    
    Set rsCred = New ADODB.Recordset
    Set rsCred = oCred.ListarCreditosPreDesemCompraDeuda(psCtaCod, fnTipoAsig)
    
    Set rsIfis = rsCred.Clone
        
    If Not (rsCred.EOF And rsCred.BOF) Then
        
        lblCliente.Caption = rsCred!cNombre
        lblTpoCredito.Caption = rsCred!cTpoCredito
        lblMontoApro.Caption = Format(rsCred!nMontoApro, "#,##0.00")
        txtDesem.Text = Format(rsCred!nMontoComprar, "#,##0.00")
        EnfocaControl txtDesem
        fEnfoque txtDesem
        
        lstCuentas.Clear
        
        Do While Not rsIfis.EOF
            lstCuentas.AddItem rsIfis("cIFIS")
            rsIfis.MoveNext
        Loop
        lstCuentas.Selected(0) = True
        Me.txtDesem.Enabled = True
        FraCredito.Enabled = True
        If (fnTipoAsig = 2) Then
            Me.txtDesem.Enabled = False
        End If
    Else
        MsgBox "El crédito no Existe", vbInformation, "Aviso"
        cmdCancelar_Click
        FraCredito.Enabled = False
    End If
      
End Sub
Public Function ValidarDatos() As Boolean
    
    If (txtDesem.Text = "" Or CDbl(IIf(txtDesem.Text = "", 0, txtDesem.Text)) = 0) Then
        MsgBox "El campo desembolso a comprar no puede estar vacio, ni puede ser cero (0).", vbInformation, "Aviso"
        ValidarDatos = False
        Exit Function
    End If
    ValidarDatos = True
End Function
Private Sub ActxCta_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        Call CargarDatos(ActxCta.NroCuenta)
     End If
End Sub

Private Sub Form_Load()
    CentraForm Me
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    ActxCta.EnabledCMAC = False
    lblFecha.Caption = gdFecSis
    If (fnTipoAsig = 1) Then
        CmdAprobar.Caption = "Aprobar"
    Else
        CmdAprobar.Caption = "Extornar"
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiaPantalla
End Sub
Private Sub LimpiaPantalla()
    
    lblCliente.Caption = ""
    lblTpoCredito.Caption = ""
    lblMontoApro.Caption = ""
    txtDesem.Text = ""
    ActxCta.NroCuenta = ""
    lstCuentas.Clear
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    FraCredito.Enabled = False
    ActxCta.EnabledCMAC = False
End Sub
Private Sub txtDesem_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosDecimales(txtDesem, KeyAscii, 15)
     If KeyAscii = 13 Then
        Me.txtDesem.Text = Format(Me.txtDesem.Text, "#,##0.00")
        Me.CmdAprobar.SetFocus
     End If
End Sub

