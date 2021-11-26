VERSION 5.00
Begin VB.Form frmCajeroCorte 
   Caption         =   "Pre-Cuadre Operaciones "
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6990
   Icon            =   "frmCajeroCorte.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txttotdol 
         Height          =   340
         Left            =   5040
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txttotsol 
         Height          =   340
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   1690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL DOLARES :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   165
         Index           =   1
         Left            =   3600
         TabIndex        =   4
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "TOTAL SOLES :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   165
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1155
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Index           =   1
         Left            =   3480
         Top             =   240
         Width           =   3240
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   345
         Index           =   0
         Left            =   120
         Top             =   240
         Width           =   3240
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   840
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   840
      Width           =   1200
   End
End
Attribute VB_Name = "frmCajeroCorte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbSalir As Boolean
Dim lbModifica As Boolean
Dim lnParamDif As Double
Dim sUsuario As String
Dim lsOpeCod As Long
Dim fnNumMaxRegEfecSinSolicExt As Integer
Dim fnVecesTotal As Integer
Dim fnVecesVigente As Integer
Dim fnMovNroRegEfec As Long, fnMovNroRegEfecUlt As Long
Dim fnVecesExtorno As Integer
'*****************************************************
'**MADM 20101006 *************************************
Dim loVistoElectronico As frmVistoElectronico
Dim lbVistoVal As Boolean
Dim lbIndPre As Boolean
'*****************************************************
Dim nSaldoEfectivoAyer As Double, nSaldoEfectivoHoy As Double

Private Sub cmdAceptar_Click()
Dim oCont As COMNContabilidad.NCOMContFunciones
Dim oCajero As COMNCajaGeneral.NCOMCajero
Dim lbNuevo As Boolean
Dim lsMovNro As String
 If Me.txttotsol.Text = "" Or Me.txttotdol.Text = "" Then
    MsgBox "Debe Completar los Montos Totales", vbInformation, "Aviso"
    Exit Sub
 End If
 
 '''If MsgBox("Desea Registrar Pre-Cuadre de Operaciones  por S/." & txttotsol.Text & " y $." & txttotdol.Text, vbYesNo + vbQuestion, "Aviso") = vbYes Then 'marg ers044-2016
 If MsgBox("Desea Registrar Pre-Cuadre de Operaciones  por " & gcPEN_SIMBOLO & txttotsol.Text & " y $." & txttotdol.Text, vbYesNo + vbQuestion, "Aviso") = vbYes Then 'marg ers044-2016
    Set oCont = New COMNContabilidad.NCOMContFunciones
    Set oCajero = New COMNCajaGeneral.NCOMCajero
    lbNuevo = True
    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
   
    Set oCajero = New COMNCajaGeneral.NCOMCajero
    
    If oCajero.GrabaRegistroEfectivoCorte(gsFormatoFecha, lsMovNro, _
          lsOpeCod, "Pre-Cuadre Operaciones", CDbl(txttotsol.Text), CDbl(txttotdol.Text), sUsuario, lbNuevo, fnMovNroRegEfec) = 0 Then
    
        frmCajeroIngEgre.Inicia False, False, , lsMovNro, fnMovNroRegEfec, fnVecesTotal + 1, lsOpeCod
        Me.txttotsol.Text = 0#
        Me.txttotdol.Text = 0#
        Me.txttotsol.SetFocus
        Me.cmdAceptar.Enabled = False
    End If
 Else
       Exit Sub
 End If


End Sub

Function valida_controles() As Boolean
    
    If Len(Me.txttotdol.Text) = 0 Or Len(Me.txttotsol.Text) Then
         valida_controles = False
    End If
    
End Function

Private Sub cmdCancelar_Click()
    Unload Me
    'MIOL 20120727, POR CASO HUANUCO **************************
    frmCajaGenEfectivo.cmdCancelar = True
    'END MIOL *************************************************
End Sub

Private Sub Form_Load()
Dim oCajero As New COMNCajaGeneral.NCOMCajero
Dim oGen As COMDConstSistema.DCOMGeneral
Dim lrs As ADODB.Recordset

CentraForm Me
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
Me.Height = 1860

Set oGen = New COMDConstSistema.DCOMGeneral
Set oCajero = New COMNCajaGeneral.NCOMCajero

'Tiene para 3 precuadres
fnNumMaxRegEfecSinSolicExt = oGen.GetParametro(4000, 1003)
Set oGen = Nothing
lbSalir = False
lbModifica = False
lsOpeCod = 901040
'gsOpeCod = lsOpeCod
sUsuario = gsCodUser
lbModifica = True

        If oCajero.YaRealizoDevBilletaje(gsCodUser, gdFecSis, gsCodAge) Then
            MsgBox "Ud. ha realizado la operación de registro de efectivo, esta operación no esta disponible después del registro de efectivo", vbInformation, "Aviso"
            cmdAceptar.Enabled = False
            lbModifica = False
        End If

        If lbModifica = False Then
            Me.cmdAceptar.Enabled = False
            Exit Sub
        End If

        Set lrs = oCajero.GetValidacionRegistroPreCuadre(sUsuario, gsCodAge, Format(gdFecSis, "yyyymmdd"), "901040")
        If Not lrs.EOF And Not lrs.BOF Then
                fnMovNroRegEfecUlt = lrs!nMovNro
                fnVecesTotal = lrs!nVecesTotal
                fnVecesVigente = lrs!nVecesVigente
                fnVecesExtorno = lrs!nVecesExtorno
                
                If fnVecesTotal >= fnNumMaxRegEfecSinSolicExt And fnVecesVigente >= 3 Then
                    MsgBox "Ud. ya ha realizado 1 registro de efectivo, para realizar esta operación es necesario que realice un : Extorno de Devolución por Billetaje o Extorno de Registro de Efectivo ", vbInformation, "Aviso"
                    cmdAceptar.Enabled = False
                    lbModifica = False
                End If
                'MADM 20111811
                 
                 'Comentado Por MIOL 20120705 ***
                 'If fnVecesExtorno >= 1 Then
                 '   Set loVistoElectronico = New frmVistoElectronico
                 '   lbVistoVal = loVistoElectronico.Inicio(3, gsOpeCod)
                 '   If Not lbVistoVal Then
                 '       cmdAceptar.Enabled = False
                 '       lbModifica = False
                 '   End If
                 'End If
                 '***
                 
                'END MADM
                   
            Set oCajero = Nothing
        End If
   
        If lbModifica = False Then
            Me.cmdAceptar.Enabled = False
            Exit Sub
        End If
    'MIOL 20120601, SEGUN RQ1209 *************************************************************
        Me.txttotsol.Text = frmCajaGenEfectivo.lblTotal(0)
        Me.txttotdol.Text = frmCajaGenEfectivo.lblTotal(1)
    'END MIOL ********************************************************************************
End Sub

Private Sub txttotdol_GotFocus()
fEnfoque txttotdol
End Sub

Private Sub txttotdol_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txttotdol, KeyAscii, 15)
    If KeyAscii <> 13 Then Exit Sub
    
    If lbModifica Then
        If Not IsNumeric(txttotdol.Text) Then
            MsgBox "Ingrese un monto válido", vbInformation, "Mensaje"
            Exit Sub
        End If
        
        Me.cmdAceptar.Enabled = True
        Me.cmdAceptar.SetFocus
    End If
End Sub

Private Sub txttotsol_GotFocus()
fEnfoque txttotsol
End Sub

Private Sub txttotsol_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txttotsol, KeyAscii, 15)
    If KeyAscii <> 13 Then Exit Sub
    
    If Not IsNumeric(txttotsol.Text) Then
        MsgBox "Ingrese un monto válido", vbInformation, "Mensaje"
        Exit Sub
    End If
    
    txttotdol.SetFocus
End Sub


