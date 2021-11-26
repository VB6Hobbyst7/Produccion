VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredPolizaCobrar 
   Caption         =   "Póliza Contra Incendio"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3795
   Icon            =   "frmCredPolizaCobrar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3795
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin Spinner.uSpinner SpnCuota 
      Height          =   330
      Left            =   2760
      TabIndex        =   3
      Top             =   240
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   582
      Max             =   500
      Min             =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
   End
   Begin VB.Label Label1 
      Caption         =   "¿Desde qué cuota se cobrará el seguro contra incendio?"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmCredPolizaCobrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnCantidadCuotasSinSeguro As Integer
Dim lnCantidadCuotas As Integer
Public Function MostrarCuotaCobrar(Optional ByVal pnCantidadCuotasSinSeguro As Integer = 0, Optional ByVal pnCantidadCuotas As Integer = 110)
    lnCantidadCuotasSinSeguro = pnCantidadCuotasSinSeguro
    SpnCuota.valor = IIf(lnCantidadCuotasSinSeguro = -1, 1, lnCantidadCuotasSinSeguro)
    'SpnCuota.Max = pnCantidadCuotas
    lnCantidadCuotas = pnCantidadCuotas
    SpnCuota.Min = 1
    Me.Show 1
    MostrarCuotaCobrar = lnCantidadCuotasSinSeguro
End Function

Private Sub CmdAceptar_Click()
lnCantidadCuotasSinSeguro = CInt(SpnCuota.valor)
If lnCantidadCuotasSinSeguro > lnCantidadCuotas Then
    MsgBox "No puede seleccionar ésta la cantidad de cuotas, es mayor que las cuotas totales", vbCritical, "Aviso!"
Else
    Unload Me
End If
End Sub

Private Sub cmdCancelar_Click()
    lnCantidadCuotasSinSeguro = 1
    Unload Me
End Sub

