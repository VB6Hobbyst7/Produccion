VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAsientosFractal 
   Caption         =   "Asiento Fractal"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5280
   Icon            =   "frmAsientosFractal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAsiento 
      Caption         =   "Asiento"
      Height          =   375
      Left            =   720
      TabIndex        =   14
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Asiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   5055
      Begin VB.OptionButton optProvisiones 
         Caption         =   "Asientos de Provisiones"
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   720
         Width           =   2775
      End
      Begin VB.OptionButton optGastos 
         Caption         =   "Asiento de Gastos"
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fecha y Tipo cambio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txtTipCambio 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2235
         MaxLength       =   16
         TabIndex        =   4
         Top             =   840
         Width           =   1125
      End
      Begin VB.TextBox txtTipCambioVenta 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2235
         MaxLength       =   16
         TabIndex        =   3
         Top             =   1440
         Width           =   1125
      End
      Begin VB.TextBox txtTipCambioCompra 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2235
         MaxLength       =   16
         TabIndex        =   2
         Top             =   2040
         Width           =   1125
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   2235
         TabIndex        =   1
         Top             =   360
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
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
      Begin VB.Label Label1 
         Caption         =   "Fecha de Asiento"
         Height          =   315
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Cambio Venta"
         Height          =   315
         Left            =   360
         TabIndex        =   7
         Top             =   1560
         Width           =   1680
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Cambio Fijo"
         Height          =   315
         Left            =   360
         TabIndex        =   6
         Top             =   900
         Width           =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Cambio Compra"
         Height          =   315
         Left            =   375
         TabIndex        =   5
         Top             =   2160
         Width           =   1920
      End
   End
End
Attribute VB_Name = "frmAsientosFractal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAsiento_Click()
Dim oAsiento As DAjusteCont
Dim lsMensaje As String
Dim lsMovNro As String
Set oAsiento = New DAjusteCont

    If optGastos.value = True Then
        lsMensaje = oAsiento.GenerarAsientoFractal(txtFecha.Text, "691098", gsCodUser, lsMovNro)
    End If
    If optProvisiones.value = True Then
        lsMensaje = oAsiento.GenerarAsientoFractal(txtFecha.Text, "691099", gsCodUser, lsMovNro)
    End If
        
    If Len(lsMovNro) = 0 Then
        MsgBox lsMensaje, vbApplicationModal, "Asiento Fractal"
    Else
        ImprimeAsientoContable lsMovNro, , , , , , , , , , , , 1
    End If
Set oAsiento = Nothing
End Sub

Private Sub cmdImprimir_Click()
Dim oAsiento As DAjusteCont
Dim lsMensaje As String
Dim lsMovNro As String
Set oAsiento = New DAjusteCont

    If optGastos.value = True Then
        lsMensaje = oAsiento.ObtieneAsientoFractal(txtFecha.Text, "691098", lsMovNro)
    End If
    If optProvisiones.value = True Then
        lsMensaje = oAsiento.ObtieneAsientoFractal(txtFecha.Text, "691099", lsMovNro)
    End If
        
    If Len(lsMovNro) = 0 Then
        MsgBox lsMensaje, vbApplicationModal, "Asiento Fractal"
    Else
        ImprimeAsientoContable lsMovNro, , , , , , , , , , , , 1
    End If
Set oAsiento = Nothing


End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  txtFecha.Text = gdFecSis
End Sub

Private Sub optGastos_Click()
    Call ActivarAsiento
End Sub

Sub ActivarAsiento()
    If optGastos.value = True Then
        optProvisiones.value = False
        optGastos.value = True
    End If
    If optProvisiones.value = True Then
        optProvisiones.value = True
        optGastos.value = False
    End If
End Sub

Private Sub optProvisiones_Click()
    Call ActivarAsiento
End Sub
Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub
Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtTipCambio = TiposCambiosCierreMensual(Format(txtFecha, "YYYY"), CInt(Format(txtFecha, "MM")), False, 1)
    txtTipCambioVenta = TiposCambiosCierreMensual(Format(txtFecha, "YYYY"), CInt(Format(txtFecha, "MM")), False, 2)
    txtTipCambioCompra = TiposCambiosCierreMensual(Format(txtFecha, "YYYY"), CInt(Format(txtFecha, "MM")), False, 3)
    txtFecha.SetFocus
End If
End Sub
