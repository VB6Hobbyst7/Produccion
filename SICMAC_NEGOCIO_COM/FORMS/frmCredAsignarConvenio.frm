VERSION 5.00
Begin VB.Form frmCredAsignarConvenio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Datos de Solicitud por Convenio"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmCredAsignarConvenio.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   3960
      Width           =   1380
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   3960
      Width           =   1380
   End
   Begin VB.Frame fraConvenio 
      Height          =   1965
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   6915
      Begin VB.ComboBox cmbInstitucion 
         Height          =   315
         Left            =   1170
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   225
         Width           =   5565
      End
      Begin VB.ComboBox cmbModular 
         Height          =   315
         Left            =   1170
         TabIndex        =   8
         Top             =   690
         Width           =   1815
      End
      Begin VB.TextBox txtCargo 
         Height          =   315
         Left            =   3960
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "000000"
         Top             =   675
         Width           =   765
      End
      Begin VB.TextBox txtCARBEN 
         Height          =   315
         Left            =   6135
         MaxLength       =   4
         TabIndex        =   6
         Text            =   "0000"
         Top             =   675
         Width           =   615
      End
      Begin VB.Frame fraTipoPlanilla 
         Caption         =   "Tipo Planilla"
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
         Height          =   690
         Left            =   150
         TabIndex        =   3
         Top             =   1125
         Width           =   6600
         Begin VB.OptionButton optT_Plani 
            Caption         =   "CAS"
            Height          =   390
            Index           =   2
            Left            =   360
            TabIndex        =   27
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optT_Plani 
            Caption         =   "Descuento para Cesantes"
            Height          =   390
            Index           =   1
            Left            =   4200
            TabIndex        =   5
            Top             =   225
            Width           =   2190
         End
         Begin VB.OptionButton optT_Plani 
            Caption         =   "Descuento para Activos"
            Height          =   390
            Index           =   0
            Left            =   1680
            TabIndex        =   4
            Top             =   240
            Width           =   2190
         End
      End
      Begin VB.Label lblinstitucion 
         AutoSize        =   -1  'True
         Caption         =   "Institución :"
         Height          =   195
         Left            =   135
         TabIndex        =   14
         Top             =   285
         Width           =   810
      End
      Begin VB.Label lblModular 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Modular :"
         Height          =   195
         Left            =   75
         TabIndex        =   13
         Top             =   675
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Cargo:"
         Height          =   240
         Left            =   3285
         TabIndex        =   12
         Top             =   750
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Correlativo Sobreviviente:"
         Height          =   390
         Left            =   4935
         TabIndex        =   11
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "o DNI :"
         Height          =   195
         Left            =   75
         TabIndex        =   10
         Top             =   900
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Crédito"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6975
      Begin VB.Label Label13 
         Caption         =   "Estado :"
         Height          =   255
         Left            =   4320
         TabIndex        =   26
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblEstado 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4920
         TabIndex        =   25
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblTipoProd 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   24
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label lblTipoCred 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   23
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label lblClienteDOI 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5520
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblClienteNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label7 
         Caption         =   "DOI :"
         Height          =   255
         Left            =   5040
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo Producto :"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo Crédito:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
   End
   Begin SICMACT.ActXCodCta ActXCodCta 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Texto           =   "Credito :"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
End
Attribute VB_Name = "frmCredAsignarConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public fsCodInstitucion As String
Public fsCodModular As String
Public fsCargo As String
Public fsCARBEN As String
Public fsT_Plani As String
Dim nTpoAcc As Integer
Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    Dim loCred As New COMDCredito.DCOMCreditos
    Dim loDR  As New ADODB.Recordset
     If KeyAscii = 13 Then
        Set loCred = New COMDCredito.DCOMCreditos
        ActXCodCta.Enabled = False
        Set loDR = loCred.ObtieneDatosCredPersona(ActXCodCta.NroCuenta)
        
        If Not (loDR.EOF And loDR.BOF) Then
            If (ValidaExisteConvenio(ActXCodCta.NroCuenta) = True) Then
                MsgBox "Cuenta ya tiene convenio", vbCritical, "Mensaje"
                Call Limpiar
                Exit Sub
            Else
                lblClienteNombre.Caption = loDR!cPersNombre
                lblClienteDOI.Caption = loDR!cPersIDnro
                lblTipoCred.Caption = loDR!cTpoCred
                lblTipoProd.Caption = loDR!CTpoProd
                lblEstado.Caption = loDR!cEstado
                Call CargarInst
                nTpoAcc = 1
                cmdCancelar.Caption = "Cancelar"
            End If
        Else
            MsgBox "No se encontro datos de crédito, verifique el número", vbExclamation, "Aviso"
            Call Limpiar
            nTpoAcc = 0
        End If
     End If
End Sub

Private Sub cmbInstitucion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbModular.SetFocus
    End If
End Sub

Private Sub cmbModular_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCargo.SetFocus
    End If
End Sub

Private Sub CmdAceptar_Click()
    Dim loCred As New COMDCredito.DCOMCreditos
    Set loCred = New COMDCredito.DCOMCreditos
    
       'JAME ***********'
    Dim loCredHis As New COMDCredito.DCOMCredito
    Set loCredHis = New COMDCredito.DCOMCredito
    'JAME FIN'
    
    If cmbInstitucion.ListIndex = -1 Then
        MsgBox "Debe seleccionar una institucion", vbInformation, "Mensaje"
        Exit Sub
    End If
    If cmbModular.Text = "" Then
        MsgBox "Debe indicar un Codigo Modular", vbInformation, "Mensaje"
        Exit Sub
    End If
    If txtCargo.Text = "" Then
        MsgBox "Debe indicar un Cargo", vbInformation, "Mensaje"
        Exit Sub
    End If
    If txtCARBEN.Text = "" Then
        MsgBox "Debe indicar un Correlativo de Sobreviviente", vbInformation, "Mensaje"
        Exit Sub
    End If
    If txtCargo.Text = "" Then
        MsgBox "Debe indicar un Cargo", vbInformation, "Mensaje"
        Exit Sub
    End If
    If optT_Plani(0).Value = False And optT_Plani(1).Value = False And optT_Plani(2).Value = False Then 'WIOR 20150217 AGREGO optT_Plani(2).Value = False
        MsgBox "Debe indicar tipo de Planilla", vbInformation, "Mensaje"
        Exit Sub
    End If
    'ARCV 30-01-2007
    'If Len(cmbModular.Text) <> 10 Then
    '    MsgBox "El Codigo Modular debe tener 10 caracteres", vbInformation, "Mensaje"
    '    Exit Sub
    'End If
    If Len(txtCargo.Text) <> 6 Then
        MsgBox "El Cargo debe tener 6 caracteres", vbInformation, "Mensaje"
        Exit Sub
    End If
    If Len(txtCARBEN.Text) <> 4 Then
        MsgBox "El Correlativo de Sobreviviente debe tener 4 caracteres", vbInformation, "Mensaje"
        Exit Sub
    End If
    If txtCargo.Text = "" Then
        MsgBox "Debe indicar un Cargo", vbInformation, "Mensaje"
        Exit Sub
    End If
    
    
    fsCodInstitucion = Trim(Right(cmbInstitucion.Text, 15))
    fsCodModular = cmbModular.Text
    fsCargo = txtCargo.Text
    fsCARBEN = txtCARBEN.Text
    fsT_Plani = IIf(optT_Plani(0).Value, "A", IIf(optT_Plani(2).Value, "CA", "C")) 'WIOR 20150217
    
    
    loCred.RegistraConvenioCredPosDesembolso ActXCodCta.NroCuenta, fsCodInstitucion, fsCodModular, fsCargo, fsCARBEN, fsT_Plani, Format(gdFecSis, "yyyy/MM/dd")
     'JAME *****'
    
    loCredHis.RegistroHistorialConvenio ActXCodCta.NroCuenta, gsCodUser, Format(gdFecSis & " " & GetHoraServer, "yyyy/MM/dd hh:mm:ss"), fsCodInstitucion, " ", 1
    'loCredHis.RegistroHistorialConvenio ActXCodCta.NroCuenta, gsCodUser, gdFecSis & GetHoraServer, fsCodInstitucion, " ", 1
    'JAME FIN'
    MsgBox "La asignación se completo de forma satisfactoria", vbInformation, "Mensaje"
    Call Limpiar
    'Unload Me
End Sub

Public Sub Limpiar()
    fsCodInstitucion = ""
    fsCodModular = ""
    fsCargo = ""
    fsCARBEN = ""
    fsT_Plani = ""
    cmbInstitucion.Clear
    cmbModular.Clear
    txtCargo.Text = "000000"
    txtCARBEN.Text = "0000"
    optT_Plani(0).Value = False
    optT_Plani(1).Value = False
    ActXCodCta.NroCuenta = ""
    lblClienteNombre.Caption = ""
    lblClienteDOI.Caption = ""
    lblTipoCred.Caption = ""
    lblTipoProd.Caption = ""
    lblEstado.Caption = ""
    nTpoAcc = 0
    cmdCancelar.Caption = "Salir"
    ActXCodCta.Enabled = True
End Sub

Public Sub CargarInst()
    Dim oPersonas As New COMDPersona.DCOMPersonas
    Dim rsInstituc As ADODB.Recordset
    Set rsInstituc = New ADODB.Recordset
    'Set oPersonas = New comdpersona.dcompersonas
    
    Set rsInstituc = oPersonas.RecuperaPersonasTipo(gPersTipoConvenio)
    Set oPersonas = Nothing
    
     cmbInstitucion.Clear
    Do While Not rsInstituc.EOF
        cmbInstitucion.AddItem PstaNombre(rsInstituc!cPersNombre) & Space(250) & rsInstituc!cPersCod
        rsInstituc.MoveNext
    Loop
End Sub

Private Sub cmdCancelar_Click()
    If nTpoAcc = 0 Then
        Unload Me
    Else
        Call Limpiar
    End If
End Sub

Private Sub optT_Plani_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub

Private Sub txtCARBEN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        optT_Plani(0).SetFocus
    End If
End Sub

Private Sub txtCargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCARBEN.SetFocus
    End If
End Sub

Public Function ValidaExisteConvenio(ByVal psCtaCod As String) As Boolean
    Dim loCred As New COMDCredito.DCOMCreditos
    Dim loDR As New ADODB.Recordset
    Set loCred = New COMDCredito.DCOMCreditos
    
    Set loDR = loCred.ObtieneDatosCreditoConvenio(psCtaCod)
    
    If Not (loDR.EOF And loDR.BOF) Then
        ValidaExisteConvenio = True
    Else
        ValidaExisteConvenio = False
    End If
End Function
