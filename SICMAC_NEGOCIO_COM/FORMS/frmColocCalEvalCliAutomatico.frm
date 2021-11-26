VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColocCalEvalCliAutomatico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colocaciones - Evaluación  Automatica"
   ClientHeight    =   3540
   ClientLeft      =   2400
   ClientTop       =   2940
   ClientWidth     =   6000
   Icon            =   "frmColocCalEvalCliAutomatico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Evaluacion Automatica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      Begin VB.CommandButton cmdEvaluaAutomatico 
         Caption         =   "Evaluacion Automatica"
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
         Left            =   240
         TabIndex        =   1
         Top             =   2160
         Width           =   4095
      End
      Begin MSComctlLib.ProgressBar Barra 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
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
         Left            =   4440
         TabIndex        =   0
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Data"
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
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label lblFechaData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2460
         TabIndex        =   7
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ultima Calif"
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
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   780
         Width           =   1560
      End
      Begin VB.Label lblFechaCalif 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2460
         TabIndex        =   5
         Top             =   780
         Width           =   1515
      End
      Begin VB.Label lblDato 
         AutoSize        =   -1  'True
         Caption         =   "lblDato"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmColocCalEvalCliAutomatico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* COLOCACIONES - CALIFICACION - EVALUACION AUTOMATICA
'Archivo:  frmColocCalEvalCli.frm
'LAYG   :  22/06/2004.
'Resumen:  Evaluacion Automatica para la Calificacion

Option Explicit
Dim fsServConsol As String

Private Sub cmdEvaluaAutomatico_Click()
Dim loDatos As COMNCredito.NCOMColocEval
Dim lrDatos As ADODB.Recordset

Dim lsCalificacion As String
Dim lnTotal As Integer, j As Integer
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim lsMovNro As String


If DateDiff("d", lblFechaCalif, lblFechaData) <= 0 Then
    MsgBox "Ya se Realizo la Evaluacion Automatica...", vbInformation, "AVISO"
    Exit Sub
End If
If MsgBox("Esta seguro de realizar Evaluacion Automatica de la Cartera ", vbInformation + vbYesNo, "AVISO") = vbNo Then
    Exit Sub
End If

'Genera el Mov Nro
Set loContFunct = New COMNContabilidad.NCOMContFunciones
    lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
Set loContFunct = Nothing

        
Set loDatos = New COMNCredito.NCOMColocEval
    Set lrDatos = loDatos.nObtieneDatosEvaluacionAutomatica(lblFechaCalif, lblFechaData)

    If lrDatos Is Nothing Then
        MsgBox "No existen Datos para Evaluacion Automatica", vbInformation, "Aviso"
        Exit Sub
    End If
    If lrDatos.BOF And lrDatos.EOF Then
        MsgBox "No existe datos La Data ya fue Transferida", vbInformation, "AVISO"
        Exit Sub
    End If

    lnTotal = lrDatos.RecordCount
    Do While Not lrDatos.EOF
        'Obtengo la Calificacion
        
        'lsCalificacion = loDatos.nGeneraCalificacionAutomatica(lrDatos!cCtaCod, fsServConsol)
        lsCalificacion = lrDatos!cEvalCalifDet
        
        If lsCalificacion <> "" Then
            Call loDatos.nCalifDetalleNuevo(lrDatos!cPersCod, 0, lrDatos!cCtaCod, lblFechaData, "", lsCalificacion, _
                lrDatos!nSaldoCap, lrDatos!nDiasAtraso, lsMovNro, "Eval.Automat", False)
        End If

        barra.value = Int(j / lnTotal * 100)
        'Me.lblDato.Caption = Trim(lrDatos!cPersCod) & "  - " & lnNuevos
        lrDatos.MoveNext
    Loop
Set loDatos = Nothing
Set lrDatos = Nothing
MsgBox "Transferencia satisfactoria", vbInformation, "AVISO"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim lsFecha As Date
Dim loConstSist As COMDConstSistema.NCOMConstSistema
    Set loConstSist = New COMDConstSistema.NCOMConstSistema
        fsServConsol = loConstSist.LeeConstSistema(gConstSistServCentralRiesgos)
        lblFechaData = loConstSist.LeeConstSistema(gConstSistCierreMesNegocio)
        lblFechaCalif = loConstSist.LeeConstSistema(141) ' Fecha Ultima Calificacion Automatica
    Set loConstSist = Nothing
   
    Me.lblDato = ""
    
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

