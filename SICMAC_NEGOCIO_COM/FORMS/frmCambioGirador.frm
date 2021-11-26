VERSION 5.00
Begin VB.Form frmCambioGirador 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambiar Girador"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdCambiar 
      Caption         =   "Cambiar"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Frame frmaNuevoGir 
      Caption         =   "Nuevo Girador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   6735
      Begin SICMACT.TxtBuscar txtInst 
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   609
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
         EnabledText     =   0   'False
      End
      Begin VB.Label lblDOI 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblInstDesc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   6465
      End
   End
   Begin VB.Frame fraGirActual 
      Caption         =   "Girador Actual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   6735
      Begin VB.Label lblGirActual 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   6495
      End
   End
   Begin VB.Frame fraDatCheque 
      Caption         =   "Datos del Cheque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtCheque 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   800
         Width           =   1695
      End
      Begin VB.ComboBox CmbBancos 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   4380
      End
      Begin VB.Label lblNroCheque 
         Caption         =   "Nº Cheque:"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Banco : "
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   390
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmCambioGirador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnTipoDoc As Integer
Dim lnMonto As Currency
Dim lsNumeDoc As String
Dim lsCodIns As String
Dim lsCodGirAct As String
Dim oFinan As comdpersona.DCOMInstFinac

Private Sub LimpiarCampos(ByVal nTpo As Integer)
If nTpo = 1 Then
    txtCheque.Text = ""
    lblGirActual.Caption = ""
    txtInst.Text = ""
    lblInstDesc.Caption = ""
    lblDOI.Caption = ""
ElseIf nTpo = 2 Then
    txtInst.Text = ""
    lblInstDesc.Caption = ""
    lblDOI.Caption = ""
End If
End Sub

Private Sub CmbBancos_Click()
    If CmbBancos.ListIndex <> -1 Then
         LimpiarCampos 1
    End If
End Sub

Private Sub cmdCambiar_Click()
Dim oMov As COMDMov.DCOMMov
Set oMov = New COMDMov.DCOMMov
Dim oCon As COMNContabilidad.NCOMContFunciones
Set oCon = New COMNContabilidad.NCOMContFunciones
Dim lsMovNro  As String
Dim lnMovNro  As Long
Dim nSaldo As Currency
Dim lbTrans As Boolean
Dim rsInst As ADODB.Recordset

'gsOpeCod = "900035" 'solo por ahora
    If Me.txtCheque = "" Or Me.lblGirActual.Caption = "" Or Me.txtInst = "" Then
        MsgBox "Verifique los datos!; Falta completar"
        Exit Sub
    End If

    Set oFinan = New comdpersona.DCOMInstFinac
    Set rsInst = oFinan.CargaCodGirador(lsCodIns, txtCheque.Text)
    If rsInst.RecordCount > 0 Then
        nSaldo = rsInst!nMonto - oFinan.CargaChequeMontoUsado(rsInst!nTpoDoc, txtCheque.Text)
        If nSaldo > 0 Then
            If MsgBox("¿Desea realizar cambio de Girador seleccionado?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                oMov.BeginTrans
                lsMovNro = oCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                oMov.InsertaMov lsMovNro, gsOpeCod, "Cambio Girador", gMovEstContabNoContable
                lnMovNro = oMov.GetnMovNro(lsMovNro)
                
                oFinan.ActualizaGiradorCheque txtInst.Text, txtCheque.Text
                oMov.InsertaMovCambioGirador lnMovNro, lsMovNro, txtCheque.Text, rsInst!cPersCodGirador, Me.txtInst
                oMov.CommitTrans
                lbTrans = False
                Set oMov = Nothing
                If MsgBox("El cambio de girador se ralizo con exito!, ¿Desea realizar otro cambio?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                    LimpiarCampos 1
                    CargaBancos
                Else
                    Unload Me
                End If
            Else
                LimpiarCampos 2
            End If
        Else
            MsgBox "El cheque seleccionado no tiene saldo!!!"
        End If
    End If
    Set rsInst = Nothing
    Set oFinan = Nothing
    Set oCon = Nothing
   
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CargaBancos
End Sub

Private Sub CargaBancos()
Dim R As ADODB.Recordset
    
    CmbBancos.Clear
    Set oFinan = New comdpersona.DCOMInstFinac
    Set R = oFinan.RecuperaBancos(False, "")
    Do While Not R.EOF
        CmbBancos.AddItem PstaNombre(R!cPersNombre) & Space(80) & R!cPersCod
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set oFinan = Nothing
End Sub

Private Sub cargarDatosCheque()
Dim rsCodInst As ADODB.Recordset
Dim rsInst As ADODB.Recordset
    Set oFinan = New comdpersona.DCOMInstFinac
    lsCodIns = Trim(Right(CmbBancos.Text, 30))
    Set rsCodInst = oFinan.CargaCodGirador(lsCodIns, txtCheque.Text)
    If rsCodInst.RecordCount > 0 Then
        Set rsInst = oFinan.CargaNombreGirador(rsCodInst!cPersCodGirador)
            If rsInst.RecordCount > 0 Then
                Me.lblGirActual.Caption = rsInst!cPersNombre
            Else
                MsgBox "No existe la Institución Girador buscada!!!"
            End If
            rsInst.Close
            Set rsInst = Nothing
    Else
        MsgBox "No existe el Nº de cheque relacionado al Banco selecccionado!!!"
    End If
    CargaGiradores
    rsCodInst.Close
    Set rsCodInst = Nothing
    Set oFinan = Nothing
End Sub

Private Sub txtCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cargarDatosCheque
End Sub

Private Sub CargaGiradores()
Dim oPersonas As comdpersona.DCOMPersonas

    Set oPersonas = New comdpersona.DCOMPersonas
    txtInst.rs = oPersonas.RecuperaPersonasTipo_Arbol_AgenciaTodos(gPersTipoConvenio)

    Set oPersonas = Nothing
End Sub


Private Sub txtInst_EmiteDatos()
Dim rsDOIInst As ADODB.Recordset
Dim psCodInst As String
Dim oPersonas As comdpersona.DCOMPersonas
    Me.lblInstDesc = Trim(txtInst.psDescripcion)
    psCodInst = Trim(txtInst)
    Set oPersonas = New comdpersona.DCOMPersonas
    Set rsDOIInst = oPersonas.RecuperaNroDOI_Inst(psCodInst)
        If rsDOIInst.RecordCount > 0 Then
           lblDOI.Caption = rsDOIInst!NroDOI
        End If
    Set oPersonas = Nothing
End Sub
