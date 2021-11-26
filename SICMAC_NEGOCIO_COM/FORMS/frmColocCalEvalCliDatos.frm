VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmColocCalEvalCliDatos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colocaciones - Datos de Evaluación de Credito"
   ClientHeight    =   4425
   ClientLeft      =   1815
   ClientTop       =   2310
   ClientWidth     =   7485
   Icon            =   "frmColocCalEvalCliDatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      Top             =   3945
      Width           =   1350
   End
   Begin VB.Frame fraDatos 
      Height          =   3405
      Left            =   120
      TabIndex        =   13
      Top             =   450
      Width           =   7215
      Begin VB.CheckBox chkVigente 
         Caption         =   "Vigente"
         Height          =   225
         Left            =   4200
         TabIndex        =   29
         Top             =   1800
         Width           =   885
      End
      Begin VB.OptionButton optSeleccion 
         Caption         =   "Datos a &Fecha :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   28
         Top             =   165
         Width           =   2685
      End
      Begin VB.OptionButton optSeleccion 
         Caption         =   "Datos Ultimo Fin de &Mes"
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
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   27
         Top             =   885
         Value           =   -1  'True
         Width           =   2790
      End
      Begin VB.Frame FraDatosActual 
         Enabled         =   0   'False
         Height          =   690
         Left            =   90
         TabIndex        =   22
         Top             =   135
         Width           =   7005
         Begin VB.TextBox txtDiasAtraso 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6240
            MaxLength       =   4
            TabIndex        =   4
            Top             =   300
            Width           =   570
         End
         Begin VB.TextBox txtNota 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   4515
            MaxLength       =   2
            TabIndex        =   3
            Top             =   255
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.TextBox txtSaldoCap 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2880
            TabIndex        =   2
            Top             =   255
            Width           =   1155
         End
         Begin MSMask.MaskEdBox txtFecha 
            Height          =   345
            Left            =   840
            TabIndex        =   1
            Top             =   255
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Dias Atraso"
            Height          =   195
            Left            =   5280
            TabIndex        =   26
            Top             =   315
            Width           =   810
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Nota"
            Height          =   195
            Left            =   4200
            TabIndex        =   25
            Top             =   330
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Left            =   180
            TabIndex        =   24
            Top             =   300
            Width           =   450
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Saldo"
            Height          =   195
            Left            =   2220
            TabIndex        =   23
            Top             =   330
            Width           =   405
         End
      End
      Begin VB.Frame fraDatosFM 
         Enabled         =   0   'False
         Height          =   705
         Left            =   90
         TabIndex        =   17
         Top             =   885
         Width           =   7005
         Begin VB.TextBox txtSaldoCapFM 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2880
            TabIndex        =   6
            Top             =   240
            Width           =   1155
         End
         Begin VB.TextBox txtNotaFM 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   4530
            MaxLength       =   2
            TabIndex        =   7
            Top             =   240
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.TextBox txtDiasAtrasoFM 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   6240
            MaxLength       =   4
            TabIndex        =   8
            Top             =   255
            Width           =   570
         End
         Begin MSMask.MaskEdBox txtFechaFM 
            Height          =   345
            Left            =   825
            TabIndex        =   5
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Saldo"
            Height          =   195
            Left            =   2220
            TabIndex        =   21
            Top             =   315
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   285
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nota"
            Height          =   195
            Left            =   4200
            TabIndex        =   19
            Top             =   315
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Dias Atraso"
            Height          =   195
            Left            =   5280
            TabIndex        =   18
            Top             =   315
            Width           =   810
         End
      End
      Begin VB.TextBox txtObs 
         Height          =   1140
         Left            =   120
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2175
         Width           =   6945
      End
      Begin VB.TextBox txtCalificacion 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   6315
         MaxLength       =   2
         TabIndex        =   9
         Top             =   1770
         Width           =   510
      End
      Begin VB.Label lblCalxDiasAtraso 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1680
         Width           =   3795
      End
      Begin VB.Label Label4 
         Caption         =   "Observaciones :"
         Height          =   240
         Left            =   150
         TabIndex        =   15
         Top             =   1950
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Calificación :"
         Height          =   195
         Left            =   5370
         TabIndex        =   14
         Top             =   1815
         Width           =   915
      End
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   3780
      TabIndex        =   30
      Top             =   120
      Width           =   3315
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Crédito :"
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
      Left            =   180
      TabIndex        =   16
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblCodCredito 
      Alignment       =   2  'Center
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
      Height          =   270
      Left            =   1020
      TabIndex        =   0
      Top             =   75
      Width           =   2235
   End
End
Attribute VB_Name = "frmColocCalEvalCliDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* COLOCACIONES - CALIFICACION - EVALUACION DE CLIENTES
'Archivo:  frmColocCalEvalDetalle.frm
'LAYG   :  01/10/2002.
'Resumen:  Registra el Detalle de la evaluacion de la Calificacion

'** Registro de Evaluacion Detalle
Option Explicit

Public fsCodCta As String
Public fsCodPers As String
Public fbOk As Boolean
Public fbNuevo As Boolean
Public lbNuevaPers As Boolean
Public fdFechaEval As Date
Public fnTipoEval As Integer  ' 0=Evaluacion / 1=Revision
Public fnSaldoCap As Double
Public fnDiasAtraso As Double

Dim lbFinMes As Boolean

Private Sub CmdAceptar_Click()
Dim loNCal As COMNCredito.NCOMColocEval ' nColocEvalCal
Dim loContFunct As COMNContabilidad.NCOMContFunciones  'NContFunciones
Dim lsMovNro As String

If Valida = False Then Exit Sub
If MsgBox("Desea Guardar los Datos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    'Genera el Mov Nro
    Set loContFunct = New COMNContabilidad.NCOMContFunciones   'NContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    If fbNuevo Then
        Set loNCal = New COMNCredito.NCOMColocEval  'nColocEvalCal
            Call loNCal.nCalifDetalleNuevo(Me.fsCodPers, fnTipoEval, fsCodCta, fdFechaEval, "", _
                Trim(txtCalificacion), fnSaldoCap, fnDiasAtraso, lsMovNro, Trim(txtObs), False)
        Set loNCal = Nothing
    Else
        Set loNCal = New COMNCredito.NCOMColocEval
            Call loNCal.nCalifDetalleModifica(Me.fsCodPers, fnTipoEval, fsCodCta, fdFechaEval, "", _
            Trim(txtCalificacion), fnSaldoCap, fnDiasAtraso, lsMovNro, Trim(txtObs), False)
        Set loNCal = Nothing
    End If
End If
fbOk = True
Unload Me

End Sub
Private Sub cmdSalir_Click()
Me.fbOk = False
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Me.lblCodCredito = Trim(fsCodCta)
Me.txtFecha = gdFecSis
Me.fbOk = False
lbFinMes = True
If fbNuevo = False And lbNuevaPers = False Then
    txtFecha.Enabled = True
    Me.chkVigente.Enabled = True
Else
    Me.chkVigente.Enabled = False
    txtFecha.Enabled = True
End If
End Sub

Private Function Valida() As Boolean
Valida = True
'If optSeleccion(0).Value = 0 Then
'    If ValFecha(txtFecha) = False Then
'        Valida = False
'        Exit Function
'    End If
'    If Val(txtSaldoCap) = 0 Then
'        MsgBox "Saldo de Capital no válido", vbInformation, "aviso"
'        txtSaldoCap.SetFocus
'        Valida = False
'        Exit Function
'    End If
'    If Len(Trim(txtNota)) = 0 Then
'        MsgBox "Nota Ingresada no válida", vbInformation, "aviso"
'        txtNota.SetFocus
'        Valida = False
'        Exit Function
'    End If
'    If txtDiasAtraso = "" Then
'        MsgBox "Nro de Dias de Atraso no válido", vbInformation, "aviso"
'        txtDiasAtraso.SetFocus
'        Valida = False
'        Exit Function
'    End If
'    If Me.lbNuevo = True And Me.lbNuevaPers = False And VerAuditDet(fsCodPers, fsCodCta, txtFecha) Then
'        MsgBox "Credito se encuentra ingresado con la fecha indicada", vbInformation, "Aviso"
'        Me.txtFecha.SetFocus
'        Valida = False
'        Exit Function
'    End If
'Else
'    If Me.fraDatosFM.Visible = True Then
'        If ValFecha(txtFechaFM) = False Then
'            Valida = False
'            Exit Function
'        End If
'        If Val(txtSaldoCapFM) = 0 Then
'            MsgBox "Saldo de Capital no válido", vbInformation, "aviso"
'            txtSaldoCap.SetFocus
'            Valida = False
'            Exit Function
'        End If
'        If Len(Trim(txtNotaFM)) = 0 Then
'            MsgBox "Nota Ingresada no válida", vbInformation, "aviso"
'            txtNota.SetFocus
'            Valida = False
'            Exit Function
'        End If
'        If txtDiasAtrasoFM = "" Then
'            MsgBox "Nro de Dias de Atraso no válido", vbInformation, "aviso"
'            txtDiasAtraso.SetFocus
'            Valida = False
'            Exit Function
'        End If
'        If lbNuevo = True And Me.lbNuevaPers = False And VerAuditDet(fsCodPers, fsCodCta, txtFechaFM) Then
'            MsgBox "Credito se encuentra ingresado con la fecha indicada", vbInformation, "Aviso"
'            Me.txtFechaFM.SetFocus
'            Valida = False
'            Exit Function
'        End If
'    End If
'End If
If Len(Trim(txtCalificacion)) = 0 Then
    MsgBox "Calificación no válida", vbInformation, "aviso"
    txtCalificacion.SetFocus
    Valida = False
    Exit Function
Else
    If Val(Me.txtCalificacion) >= 5 Then
        MsgBox "Calificación no puede ser mayor que 5", vbInformation, "aviso"
        txtCalificacion.SetFocus
        Valida = False
        Exit Function
    End If
End If
End Function

Private Sub optSeleccion_Click(Index As Integer)
'Select Case Index
'    Case 0
'        lbFinMes = True
'        fraDatosFM.Enabled = True
'        Me.FraDatosActual.Enabled = False
'    Case 1
'        lbFinMes = False
'        fraDatosFM.Enabled = False
'        Me.FraDatosActual.Enabled = True
'
'End Select
End Sub

Private Sub txtCalificacion_GotFocus()
    fEnfoque txtCalificacion
End Sub
Private Sub TxtCalificacion_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    txtObs.SetFocus
End If
End Sub

Private Sub txtDiasAtraso_GotFocus()
fEnfoque Me.txtDiasAtraso
End Sub

Private Sub txtDiasAtraso_KeyPress(KeyAscii As Integer)
If KeyAscii <> 45 Then
    KeyAscii = NumerosEnteros(KeyAscii)
End If
If KeyAscii = 13 Then
    Me.txtCalificacion.SetFocus
End If
End Sub

Private Sub txtDiasAtrasoFM_GotFocus()
fEnfoque Me.txtDiasAtrasoFM
End Sub

Private Sub txtDiasAtrasoFM_KeyPress(KeyAscii As Integer)
If KeyAscii <> 45 Then
    KeyAscii = NumerosEnteros(KeyAscii)
End If
If KeyAscii = 13 Then
    txtCalificacion.SetFocus
End If
End Sub


Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtSaldoCap.SetFocus
End If
End Sub

Private Sub txtFechaFM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtSaldoCapFM.SetFocus
End If
End Sub

Private Sub txtNota_GotFocus()
fEnfoque txtNota
End Sub

Private Sub TxtNota_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    Me.txtDiasAtraso.SetFocus
End If
End Sub

Private Sub txtNotaFM_GotFocus()
fEnfoque txtNota
End Sub
Private Sub TxtNotaFM_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    txtDiasAtrasoFM.SetFocus
End If
End Sub
Private Sub txtObs_GotFocus()
fEnfoque txtObs
End Sub

Private Sub txtObs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdAceptar.SetFocus
End If
End Sub
Private Sub txtSaldoCap_GotFocus()
fEnfoque txtSaldoCap
End Sub

Private Sub txtSaldoCap_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtSaldoCap, KeyAscii)
If KeyAscii = 13 Then
    Me.txtNota.SetFocus
End If
End Sub

Private Sub txtSaldoCap_LostFocus()
If Val(txtSaldoCap) = 0 Then txtSaldoCap = 0
txtSaldoCap = Format(txtSaldoCap, "#0.00")
End Sub
Private Sub txtSaldoCapFM_LostFocus()
If Val(txtSaldoCapFM) = 0 Then txtSaldoCapFM = 0
txtSaldoCapFM = Format(txtSaldoCapFM, "#0.00")
End Sub
Private Sub txtSaldoCapFM_GotFocus()
fEnfoque txtSaldoCapFM
End Sub

Private Sub txtSaldoCapFM_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtSaldoCapFM, KeyAscii)
If KeyAscii = 13 Then
    Me.txtNotaFM.SetFocus
End If
End Sub

