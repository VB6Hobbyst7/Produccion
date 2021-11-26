VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCredSolicAfpClie 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuevo Cliente que Solicitó su 25% de AFP"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6240
   Icon            =   "frmCredSolicAfpClie.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmbBuscar 
      Caption         =   "&Mostrar"
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
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1150
      Width           =   1215
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ComboBox cmbDestino 
      Height          =   315
      ItemData        =   "frmCredSolicAfpClie.frx":030A
      Left            =   960
      List            =   "frmCredSolicAfpClie.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin SICMACT.ActXCodCta txtCuenta 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   661
      Texto           =   "Cuenta N°:"
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton cmdBcliente 
         Caption         =   "&Buscar Créditos"
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
         Left            =   4130
         TabIndex        =   22
         Top             =   1000
         Width           =   1575
      End
      Begin VB.Frame FraListaCred 
         Caption         =   "&Lista Créditos"
         Height          =   1320
         Left            =   3840
         TabIndex        =   20
         Top             =   120
         Width           =   2115
         Begin VB.ListBox LstCred 
            Height          =   645
            ItemData        =   "frmCredSolicAfpClie.frx":0335
            Left            =   75
            List            =   "frmCredSolicAfpClie.frx":0337
            TabIndex        =   21
            Top             =   225
            Width           =   1980
         End
      End
      Begin VB.CheckBox chkAsignacion 
         Caption         =   "Asignación por Destino"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lblvalor 
         Height          =   255
         Left            =   3000
         TabIndex        =   18
         Top             =   1080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Destino"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame frmAgrupa 
      Enabled         =   0   'False
      Height          =   2165
      Left            =   1320
      TabIndex        =   13
      Top             =   2280
      Width           =   3800
      Begin SICMACT.EditMoney edtDisponible 
         Height          =   300
         Left            =   1560
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox cmbAfp 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtFechaAbono 
         Height          =   300
         Left            =   1560
         TabIndex        =   7
         Top             =   1680
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaCarta 
         Height          =   300
         Left            =   1560
         TabIndex        =   5
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "AFP :"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Carta :"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Abono :"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Imp. Disponible :"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1170
      End
   End
   Begin VB.Label LblCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000006&
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   3780
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Cliente :"
      Height          =   195
      Left            =   1320
      TabIndex        =   10
      Top             =   1680
      Width           =   570
   End
End
Attribute VB_Name = "frmCredSolicAfpClie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub chkAsignacion_Click()
 If chkAsignacion.value = 0 Then
   cmbDestino.Enabled = True
 Else
   cmbDestino.ListIndex = 1
   cmbDestino.Enabled = False
 End If
End Sub

Private Sub cmbAfp_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 And cmbAfp.Text <> "" Then SendKeys "{TAB}", True
End Sub

Private Sub cmbBuscar_Click()

Dim oCred As COMDCredito.DCOMCredito
Dim bCargado As New ADODB.Recordset

If chkAsignacion.value = 0 Then

        If txtCuenta.NroCuenta <> "" Then
            Set oCred = New COMDCredito.DCOMCredito
            Set bCargado = oCred.RecuperaCredito(txtCuenta.NroCuenta, cmbDestino.Text, chkAsignacion.value)
            Set oCred = Nothing
            
            If Not (bCargado.EOF And bCargado.BOF) Then
              LblCliente.Caption = bCargado!Nombre_Cliente
              
              frmAgrupa.Enabled = True
              cmbAfp.SetFocus
                    
            Else
              MsgBox "El Destino de la Solicitud del 25% del AFP: No puede ser " & cmbDestino.Text & " ", vbInformation, "AVISO"
                frmAgrupa.Enabled = False
               ' txtCuenta.NroCuenta = ""
                txtCuenta.SetFocusProd
                cmbBuscar.Enabled = False
            End If
        Else
              MsgBox "Debe ingresar un Nro. de Cuenta Válida", vbInformation, "AVISO"
              frmAgrupa.Enabled = False
              txtCuenta.SetFocusProd
              cmbBuscar.Enabled = False
        End If

Else
   If chkAsignacion.value = 1 Then

        If txtCuenta.NroCuenta <> "" Then
            Set oCred = New COMDCredito.DCOMCredito
            Set bCargado = oCred.RecuperaCredito(txtCuenta.NroCuenta, cmbDestino.Text, chkAsignacion.value)
            Set oCred = Nothing
            
            If Not (bCargado.EOF And bCargado.BOF) Then
              LblCliente.Caption = bCargado!Nombre_Cliente
              
              frmAgrupa.Enabled = True
              cmbAfp.SetFocus
                    
            Else
              MsgBox "El Destino de la Solicitud del 25% del AFP: No puede ser " & cmbDestino.Text & " ", vbInformation, "AVISO"
               frmAgrupa.Enabled = False
              'txtCuenta.NroCuenta = ""
              txtCuenta.SetFocusProd
              cmbBuscar.Enabled = False
            End If
        Else
              MsgBox "Debe ingresar un Nro. de Cuenta Válida", vbInformation, "AVISO"
               frmAgrupa.Enabled = False
              txtCuenta.SetFocusProd
              cmbBuscar.Enabled = False
        End If
   
   End If
End If
End Sub

Private Sub cmbDestino_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}", True: Exit Sub
End Sub

Private Sub cmdBcliente_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim oPers As COMDPersona.UCOMPersona
    
    LstCred.Clear
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Set oCredito = New COMDCredito.DCOMCredito
        Set R = oCredito.RecuperaCreditosVigentes(oPers.sPersCod, , Array(2021, gColocEstVigMor, gColocEstVigNorm, gColocEstVigNorm, gColocEstRefMor, gColocEstRefNorm, 2001, 2002, 2000, 2004))
        Do While Not R.EOF
            LstCred.AddItem R!cCtaCod
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
        Set oCredito = Nothing
    End If
    If LstCred.ListCount = 0 Then
        MsgBox "El Cliente No Tiene Creditos Vigentes", vbInformation, "AVISO"
    End If
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdFechExp_Click()

End Sub

Private Sub Cmdguardar_Click()
Dim oBuscarafp As COMDCredito.DCOMCredito
Dim lrsBusca As New ADODB.Recordset

On Error GoTo ErrDatos

Set oBuscarafp = New COMDCredito.DCOMCredito

If lblvalor.Caption = 1 Then
'grabara en caso de nuevo
'verificar si existe o no
    If txtCuenta.NroCuenta <> "" And cmbAfp.ListIndex <> -1 And cmbDestino.ListIndex <> -1 Then
        Set lrsBusca = oBuscarafp.BuscarCreditoAfp(txtCuenta.NroCuenta)
        
        If lrsBusca.RecordCount > 0 Then
        
            MsgBox "El Crédito se encuentra en la lista de Solicitud de Aporte del 25% del AFP correspondiente, No se puede volver a registralo", vbInformation, "AVISO"
            Call Limpiar
'            txtCuenta.SetFocus
         
        Else

            If IsDate(txtFechaCarta) = False Then MsgBox "Fecha de la Carta no es Válida", vbInformation, "AVISO": txtFechaCarta.SelStart = 0: txtFechaCarta.SelLength = Len(txtFechaCarta.Text): txtFechaCarta.SetFocus: Exit Sub
            If IsDate(txtFechaAbono) = False Then MsgBox "Fecha del Abono no es Válida", vbInformation, "AVISO": txtFechaAbono.SelStart = 0: txtFechaAbono.SelLength = Len(txtFechaAbono.Text): txtFechaAbono.SetFocus: Exit Sub
            If CDate(txtFechaAbono.Text) < CDate(Me.txtFechaCarta.Text) Then MsgBox "La fecha de Abono no puede ser menor a la fecha de la Carta", vbInformation, "AVISO": txtFechaAbono.SetFocus: Exit Sub

           Set lrsBusca = oBuscarafp.OperacionesVarios(txtCuenta.NroCuenta, UCase(cmbAfp.Text), txtFechaCarta.Text, cmbDestino.Text, edtDisponible.Text, txtFechaAbono.Text, lblvalor.Caption)
           Set oBuscarafp = Nothing
            Unload Me
            frmCredSolicAfp.LlenarGrilla
            
        End If
    Else
      MsgBox "Debe ingresar todos los campos. Por favor verificar", vbInformation, "AVISO"
      txtCuenta.SetFocus
    End If
    

ElseIf lblvalor.Caption = 2 Then
'grabara las modificaciones

        If IsDate(txtFechaCarta) = False Then MsgBox "Fecha de la Carta no es Válida", vbInformation, "AVISO": txtFechaCarta.SelStart = 0: txtFechaCarta.SelLength = Len(txtFechaCarta.Text): txtFechaCarta.SetFocus: Exit Sub
        If IsDate(txtFechaAbono) = False Then MsgBox "Fecha del Abono no es Válida", vbInformation, "AVISO": txtFechaAbono.SelStart = 0: txtFechaAbono.SelLength = Len(txtFechaAbono.Text): txtFechaAbono.SetFocus: Exit Sub
        If CDate(txtFechaAbono.Text) < CDate(Me.txtFechaCarta.Text) Then MsgBox "La fecha de Abono no puede ser menor a la fecha de la Carta", vbInformation, "AVISO": txtFechaAbono.SetFocus: Exit Sub

        Set lrsBusca = oBuscarafp.OperacionesVarios(txtCuenta.NroCuenta, UCase(cmbAfp.Text), txtFechaCarta.Text, cmbDestino.Text, edtDisponible.Text, txtFechaAbono.Text, lblvalor.Caption)
        Set oBuscarafp = Nothing
        Unload Me
        frmCredSolicAfp.LlenarGrilla

End If

ErrDatos:
If Err.Number <> 0 Then
'Dim valor As Integer
'valor = Err.Number

MsgBox Err.Number & " : " & Err.Description & " - Verifique sus Datos", vbInformation, "AVISO"


End If

End Sub
Sub Limpiar()
txtFechaAbono.Text = Format(Date, "dd/mm/yyyy")
txtFechaCarta.Text = Format(Date, "dd/mm/yyyy")
txtCuenta.NroCuenta = ""
cmbDestino.ListIndex = -1
cmbAfp.ListIndex = -1
Me.chkAsignacion.value = 0
edtDisponible.Text = ""
LblCliente.Caption = ""
Me.txtCuenta.CMAC = "109"
Me.txtCuenta.Age = "01"
Me.txtCuenta.SetFocusProd
End Sub



Private Sub edtDisponible_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{TAB}", True
End Sub

Private Sub Form_Activate()
If CInt(lblvalor.Caption) = 2 Then
  cmbDestino.Enabled = False
  frmAgrupa.Enabled = True
  txtCuenta.Enabled = False
  cmbBuscar.Enabled = False
  cmdBcliente.Enabled = False
  chkAsignacion.Enabled = False

Else
  txtCuenta.Enabled = True
  cmbDestino.Enabled = False
  txtCuenta.CMAC = "109"
  txtCuenta.Age = "01"
  txtCuenta.Prod = ""
  txtCuenta.SetFocusProd
End If
End Sub

Private Sub Form_Load()
Call CargaComboAfp
txtFechaAbono.Text = Format(Date, "dd/mm/yyyy")
txtFechaCarta.Text = Format(Date, "dd/mm/yyyy")

End Sub

Sub CargaComboAfp()

Dim oAfp As COMDCredito.DCOMCredito
Dim lrsafp As New ADODB.Recordset  'ferimoro 27AGO2018
Set oAfp = New COMDCredito.DCOMCredito

Set lrsafp = oAfp.MostrarAfps
Set oAfp = Nothing

Do While Not lrsafp.EOF

    cmbAfp.AddItem UCase(lrsafp!cNombAfp)

    lrsafp.MoveNext
Loop

End Sub

Private Sub MaskEdBox1_Change()

End Sub

Private Sub LstCred_Click()
        If LstCred.ListCount > 0 And LstCred.ListIndex <> -1 Then
            txtCuenta.NroCuenta = LstCred.Text
            txtCuenta.SetFocusCuenta
        End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtCuenta.NroCuenta Then
          If chkAsignacion.value = 0 Then
          cmbDestino.ListIndex = 0
          cmbDestino.Enabled = True
          cmbDestino.SetFocus
          cmbBuscar.Enabled = True
          Else
          cmbBuscar.Enabled = True
          End If
        End If
    End If
End Sub
Function ValidaFecha(ByVal valor As String) As Integer
Dim aFecha() As String
    aFecha = Split(valor, "/")
    
    If UBound(aFecha) <> 2 Then ValidaFecha = 4: Exit Function
    If aFecha(0) = "__" Then ValidaFecha = 5: Exit Function
    If aFecha(1) = "__" Then ValidaFecha = 5: Exit Function
    If aFecha(2) = "____" Then ValidaFecha = 5: Exit Function
        
    If aFecha(0) <= 0 Or aFecha(0) > 31 Then
        ValidaFecha = 1
        Exit Function
    ElseIf aFecha(1) <= 0 Or aFecha(1) > 12 Then
        ValidaFecha = 2
        Exit Function
    ElseIf aFecha(2) > Format(Date, "yyyy") Or aFecha(2) < Format(Date, "yyyy") Then
        ValidaFecha = 3
        Exit Function
    Else
       ValidaFecha = 0
    End If
End Function


Private Sub txtFechaAbono_KeyPress(KeyAscii As Integer)
Dim nError As Integer
 
    If KeyAscii <> 13 Then Exit Sub
    nError = ValidaFecha(txtFechaAbono)

    If nError = 0 Then
        If CDate(txtFechaAbono.Text) < CDate(Me.txtFechaCarta.Text) Then MsgBox "La fecha de Abono no puede ser menor a la fecha de la Carta", vbInformation, "AVISO": txtFechaAbono.SetFocus: Exit Sub
        SendKeys "{TAB}", True
        Exit Sub
    End If
    If nError = 1 Or nError = 2 Or nError = 3 Or nError = 5 Then MsgBox "Ingrese una fecha válida", vbInformation, "AVISO": txtFechaAbono.SelStart = 0: txtFechaAbono.SelLength = Len(txtFechaAbono.Text): txtFechaAbono.SetFocus: Exit Sub

End Sub

Private Sub txtFechaCarta_KeyPress(KeyAscii As Integer)
Dim nError As Integer
 
    If KeyAscii <> 13 Then Exit Sub

    nError = ValidaFecha(txtFechaCarta)

    If nError = 0 Then SendKeys "{TAB}", True: Exit Sub
    If nError = 1 Or nError = 2 Or nError = 3 Or nError = 5 Then MsgBox "Ingrese una fecha válida", vbInformation, "AVISO": txtFechaCarta.SelStart = 0: txtFechaCarta.SelLength = Len(txtFechaCarta.Text): txtFechaCarta.SetFocus: Exit Sub

End Sub
