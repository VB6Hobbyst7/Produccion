VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCierreContDia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre Diario de Operaciones"
   ClientHeight    =   4245
   ClientLeft      =   4440
   ClientTop       =   3390
   ClientWidth     =   5220
   Icon            =   "frmCierreContDia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   4245
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   3600
      Left            =   75
      TabIndex        =   4
      Top             =   90
      Width           =   5085
      Begin VB.PictureBox Animation1 
         Enabled         =   0   'False
         Height          =   1185
         Left            =   4980
         ScaleHeight     =   1125
         ScaleWidth      =   1215
         TabIndex        =   8
         Top             =   3330
         Width           =   1275
      End
      Begin VB.Frame fraFecha 
         Caption         =   "Cierre Contable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   810
         Left            =   240
         TabIndex        =   5
         Top             =   2130
         Width           =   4725
         Begin MSMask.MaskEdBox txtFechaIni 
            Height          =   345
            Left            =   720
            TabIndex        =   0
            Top             =   255
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   609
            _Version        =   393216
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtFechaFin 
            Height          =   345
            Left            =   2655
            TabIndex        =   1
            Top             =   255
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   609
            _Version        =   393216
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "AL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2250
            TabIndex        =   7
            Top             =   345
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "DEL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   270
            TabIndex        =   6
            Top             =   345
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
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
         Height          =   375
         Left            =   2715
         TabIndex        =   3
         Top             =   3045
         Width           =   1245
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Default         =   -1  'True
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
         Left            =   1305
         TabIndex        =   2
         Top             =   3045
         Width           =   1245
      End
      Begin VB.Image imgAlerta 
         Height          =   480
         Left            =   1035
         Picture         =   "frmCierreContDia.frx":030A
         Top             =   210
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ADVERTENCIA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   1725
         TabIndex        =   10
         Top             =   315
         Width           =   2100
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   $"frmCierreContDia.frx":045B
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   960
         Left            =   615
         TabIndex        =   9
         Top             =   870
         Width           =   4020
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   1065
         X2              =   4935
         Y1              =   750
         Y2              =   750
      End
   End
   Begin MSComctlLib.ProgressBar PROGRESO 
      Height          =   210
      Left            =   60
      TabIndex        =   12
      Top             =   3990
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   370
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblProceso 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   75
      TabIndex        =   11
      Top             =   3735
      Width           =   75
   End
End
Attribute VB_Name = "frmCierreContDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String
Dim rs   As ADODB.Recordset
Dim lSalir As Boolean
Dim lFecha As Boolean
Dim bEdit As Boolean

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub Inicio(plFecha As Boolean, Optional pbEdit As Boolean = True)
lFecha = plFecha
bEdit = pbEdit
Me.Show 1
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If ValidaFecha(txtFechaIni) <> "" Then
   MsgBox "Fecha Inicial no Válida", vbInformation, "Aviso"
   txtFechaIni.SetFocus
   Exit Function
End If
If ValidaFecha(txtFechaFin) <> "" Then
   MsgBox "Fecha Final no Válida", vbInformation, "Aviso"
   txtFechaFin.SetFocus
   Exit Function
End If
If CDate(txtFechaFin) < CDate(txtFechaIni) Then
   MsgBox "Fecha inicial debe ser menor a Fecha Final", vbInformation, "Aviso"
   txtFechaIni.SetFocus
   Exit Function
End If

If lFecha And CDate(txtFechaFin) > gdFecSis Then
   MsgBox "Sólo se pueden actualizar los saldos hasta la fecha de Hoy", vbInformation, "Aviso"
   txtFechaFin.SetFocus
   Exit Function
End If
ValidaDatos = True
End Function
Private Sub cmdProcesar_Click()
Dim n As Integer
Dim ldFecha As Date, ldFechaAct As Date
Dim oCtaSaldo As New DCtaSaldo
On Error GoTo ErrProc
If Not ValidaDatos Then
   Exit Sub
End If
Dim oCont As New NContFunciones
If Not oCont.PermiteModificarAsiento(Format(txtFechaIni, gsFormatoMovFecha), False) Then
   MsgBox "No se puede Actualizar Saldos de Mes Contable Cerrado!", vbInformation, "¡Aviso!"
   Exit Sub
End If
Set oCont = Nothing

If MsgBox(" ¿ Seguro de Iniciar Proceso ? ", vbQuestion + vbOKCancel, "Confirmación") = vbCancel Then
   Exit Sub
End If

cmdProcesar.Enabled = False
cmdCancelar.Enabled = False
Me.Enabled = False

If Not oCtaSaldo.PermiteActualizarSaldos(gsCodUser) Then
   Exit Sub
End If
ldFechaAct = GetFechaHoraServer()
oCtaSaldo.InsertaCtaSaldoEstad Format(ldFechaAct, gsFormatoFechaHora), gsCodUser, Format(txtFechaIni, gsFormatoFecha), Format(txtFechaFin, gsFormatoFecha), 0
ldFecha = CDate(txtFechaIni)
PROGRESO.Max = (CDate(txtFechaFin) - ldFecha + 1) * 2 + 1
n = 1
Do While ldFecha <= CDate(txtFechaFin)
   lblProceso.Caption = "Procesando [" & ldFecha & "] : " & Format(Int(n * 100# / PROGRESO.Max), "##0") & "%"
   DoEvents
   PROGRESO.value = n
   'oCtaSaldo.GeneraSaldosContables Format(ldFecha, gsFormatoFecha)
   oCtaSaldo.GeneraSaldosContablesNew Format(ldFecha, gsFormatoFecha)
   n = n + 1
   lblProceso.Caption = "Procesando [" & ldFecha & "] : " & Format(Int(n * 100# / PROGRESO.Max), "##0") & "%"
   PROGRESO.value = n
   ldFecha = ldFecha + 1
   n = n + 1
Loop
oCtaSaldo.ActualizaCtaSaldoEstad Format(ldFechaAct, gsFormatoFechaHora), gsCodUser, 1
Dim oConst As New NConstSistemas
   gsMovNro = GeneraMovNroActualiza(CDate(Format(ldFechaAct, gsFormatoFechaView)), gsCodUser, gsCodCMAC, gsCodAge)
   oConst.ActualizaConstSistemas gConstSistUltActSaldos, gsMovNro, txtFechaFin
   oConst.ActualizaConstSistemas gConstSistCierreDiaCont, gsMovNro, txtFechaFin
   If Not lFecha Then
      oConst.ActualizaConstSistemas gConstSistFechaInicioDia, gsMovNro, gdFecSis + 1
   End If
Set oCont = Nothing
lblProceso.Caption = "Procesando [" & txtFechaFin & "] : " & Format(Int(n * 100# / PROGRESO.Max), "##0") & "%"
PROGRESO.value = n
MsgBox " ¡ Proceso terminado ! ", vbInformation, ""
If CDate(txtFechaFin) >= gdFecSis Then
   glDiaCerrado = True
End If
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            If (Me.Caption = "  ACTUALIZACION DE SALDOS") Then
            gsOpeCod = LogPistaActSaldos
            Else: gsOpeCod = LogPistaCierreDiarioCont
            End If
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " del  " & txtFechaIni.Text & " al  " & txtFechaFin.Text
            Set objPista = Nothing
            '*******

Me.Enabled = True
cmdProcesar.Enabled = True
cmdCancelar.Enabled = True
Unload Me
Exit Sub
ErrProc:
   MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
   cmdProcesar.Enabled = True
   cmdCancelar.Enabled = True
   Me.Enabled = True
End Sub

Private Sub Form_Activate()
If lSalir Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
Dim sFecha As String
Dim sFechaSdo As String
Dim oConec As DConecta
Set oConec = New DConecta

lSalir = False
oConec.AbreConexion
'If Dir(App.path & "\videos\findfile.avi") <> "" Then
'    Animation1.Open (App.path & "\videos\FindFile.avi")
'End If
CentraForm frmCierreContDia
If lFecha Then
   Me.Caption = "  ACTUALIZACION DE SALDOS"
   fraFecha.Caption = "  ACTUALIZACION DE SALDOS"
    
   txtFechaIni.Enabled = bEdit
   
   sFecha = Trim(LeeConstanteSist(gConstSistCierreDiaCont))
   If ValidaFecha(sFecha) <> "" Then
      txtFechaIni = gdFecSis
      txtFechaFin = gdFecSis
   Else
      If CDate(sFecha) >= gdFecSis And Not lFecha Then
         MsgBox "Cierre de Operaciones del Día ya se realizaron. " & oImpresora.gPrnSaltoLinea & _
                "Por Favor verificar...", vbInformation, "Advertencia"
         lSalir = True
         Exit Sub
      End If
      
      If bEdit = True Then
        txtFechaIni = Trim(sFecha)
      Else
        txtFechaIni = "01" & Mid(sFecha, 3, 8)
      End If
      txtFechaFin = gdFecSis
      sFechaSdo = Trim(LeeConstanteSist(gConstSistUltActSaldos))
      If ValidaFecha(sFechaSdo) = "" Then
         If CDate(sFechaSdo) < CDate(txtFechaIni) Then
            txtFechaIni = CDate(Trim(sFechaSdo)) + 1
         End If
      End If
   End If
Else
   txtFechaIni = gdFecSis
   txtFechaFin = gdFecSis
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim oConec As DConecta
Set oConec = New DConecta

oConec.CierraConexion
End Sub


