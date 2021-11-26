VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmExtornoCierre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extorno de Cierre"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "frmExtornoCierre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Extorno de Cierre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdExtornar 
         Caption         =   "&Extornar"
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtGlosa 
         Height          =   735
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   720
         Width           =   3015
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   1560
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mskFechaExt 
         Height          =   255
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Glosa"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmExtornoCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Private Sub cmdExtornar_Click()
If txtGlosa.Text = "" Then
   MsgBox "Debe Ingresar una Glosa", vbInformation, "Aviso"
   Exit Sub
End If
ModificarFechaCierre
End Sub

Public Function ModificarFechaCierre() As Boolean
Dim oConsist As DConstSistemas
Dim oMov As DMov
Dim RespIns As Integer
Dim resp As Integer
Dim lsMovNro As String
Dim lnDifMes As Integer ' Diferencia entre meses
Dim lnDias As Integer ' nro dias del mes
Dim ldFechaNueva As Date ' Fecha a grabar

On Error GoTo errorx
'oMov.BeginTrans

lnDifMes = DateDiff("M", CDate(mskFechaExt.Text), gdFecSis)

If lnDifMes > 1 Then
    MsgBox "No se puede realizar el Extorno", vbInformation, "Aviso"
    txtGlosa.Text = ""
    Exit Function
Else
    Set oMov = New DMov
        lsMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        oMov.InsertaMov lsMovNro, gContExtCierre, txtGlosa.Text, gMovEstContabNoContable, gMovFlagVigente
    Set oMov = Nothing
        
    lnDias = Day(mskFechaExt.Text)
    ldFechaNueva = CDate(mskFechaExt.Text) - lnDias
    Set oConsist = New DConstSistemas
        oConsist.ActualizaConsSistemas gConstSistCierreMensualCont, lsMovNro, ldFechaNueva
    Set oConsist = Nothing
    MsgBox "Se Realizo Extorno", vbInformation, "Aviso"
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaExtornoCierreMensual
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Se Realizo Extorno Mensual con Fecha de Cierre " & mskFechaExt.Text & " Glosa :" & txtGlosa.Text
            Set objPista = Nothing
            '*******
    Limpiar
    
End If
'oMov.CommitTrans
    Exit Function
errorx:
'    oMov.RollbackTrans
    Set oMov = Nothing
    Set oConsist = Nothing
    MsgBox Err.Description, vbInformation
End Function

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub RecuperaUltimaFecha()
Dim oConsist As DConstSistemas

Set oConsist = New DConstSistemas
mskFechaExt.Text = oConsist.RecuperaUltimaFecha(gConstSistCierreMensualCont)

Set oConsist = Nothing
End Sub

Private Sub Form_Load()
RecuperaUltimaFecha
CentraForm Me
End Sub

Private Sub Limpiar()
Me.txtGlosa.Text = ""
RecuperaUltimaFecha
End Sub
