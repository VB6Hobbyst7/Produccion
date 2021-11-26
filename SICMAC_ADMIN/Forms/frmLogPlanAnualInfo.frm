VERSION 5.00
Begin VB.Form frmLogPlanAnualInfo 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5940
   ClientLeft      =   2775
   ClientTop       =   1980
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraNivel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   5
      Left            =   420
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Image imgNivel 
         Height          =   315
         Index           =   5
         Left            =   5100
         Top             =   120
         Width           =   315
      End
      Begin VB.Label lblNivel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel 5"
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
         Index           =   5
         Left            =   240
         TabIndex        =   13
         Top             =   180
         Width           =   4515
      End
      Begin VB.Shape shpNivel 
         BackColor       =   &H00ECFFEF&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   5
         Left            =   0
         Top             =   0
         Width           =   4995
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   4980
         Top             =   0
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4740
      TabIndex        =   10
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Frame fraNivel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   420
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Image imgNivel 
         Height          =   315
         Index           =   4
         Left            =   5100
         Top             =   120
         Width           =   315
      End
      Begin VB.Label lblNivel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel 4"
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
         Index           =   4
         Left            =   240
         TabIndex        =   9
         Top             =   180
         Width           =   4515
      End
      Begin VB.Shape shpNivel 
         BackColor       =   &H00ECFFEF&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   4
         Left            =   0
         Top             =   0
         Width           =   4995
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   4980
         Top             =   0
         Width           =   555
      End
   End
   Begin VB.Frame fraNivel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   3
      Left            =   420
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Label lblCargo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   4650
      End
      Begin VB.Image imgNivel 
         Height          =   315
         Index           =   3
         Left            =   5100
         Top             =   120
         Width           =   315
      End
      Begin VB.Label lblNivel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel 3"
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
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   420
         Width           =   4515
      End
      Begin VB.Shape shpNivel 
         BackColor       =   &H00ECFFEF&
         BackStyle       =   1  'Opaque
         Height          =   735
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   4875
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   4980
         Top             =   0
         Width           =   555
      End
   End
   Begin VB.Frame fraNivel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   2
      Left            =   420
      TabIndex        =   4
      Top             =   1980
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Label lblCargo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   15
         Top             =   180
         Width           =   4650
      End
      Begin VB.Image imgNivel 
         Height          =   315
         Index           =   2
         Left            =   5100
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblNivel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   420
         Width           =   4635
      End
      Begin VB.Shape shpNivel 
         BackColor       =   &H00EAFFFF&
         BackStyle       =   1  'Opaque
         Height          =   735
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   4875
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         Height          =   735
         Left            =   4980
         Top             =   0
         Width           =   555
      End
   End
   Begin VB.Frame fraNivel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   1
      Left            =   420
      TabIndex        =   2
      Top             =   1020
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Label lblCargo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   180
         Width           =   4710
      End
      Begin VB.Image imgNivel 
         Height          =   315
         Index           =   1
         Left            =   5100
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblNivel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   420
         Width           =   4695
      End
      Begin VB.Shape shpNivel 
         BackColor       =   &H00EAFFFF&
         BackStyle       =   1  'Opaque
         Height          =   735
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   4875
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   735
         Left            =   4980
         Top             =   0
         Width           =   555
      End
   End
   Begin VB.Frame fraNivel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   60
      Width           =   5535
      Begin VB.Label lblCargo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   180
         Width           =   4650
      End
      Begin VB.Label lblNivel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   420
         Width           =   4635
      End
      Begin VB.Shape shpNivel 
         BackColor       =   &H00EAFFFF&
         BackStyle       =   1  'Opaque
         Height          =   735
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   4875
      End
   End
   Begin VB.Image imgFlecha 
      Height          =   240
      Index           =   3
      Left            =   2820
      Picture         =   "frmLogPlanAnualInfo.frx":0000
      Top             =   2700
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgFlecha 
      Height          =   240
      Index           =   2
      Left            =   2820
      Picture         =   "frmLogPlanAnualInfo.frx":0342
      Top             =   1740
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgNivel 
      Height          =   315
      Index           =   0
      Left            =   6720
      Top             =   2820
      Width           =   315
   End
   Begin VB.Image imgFlecha 
      Height          =   360
      Index           =   0
      Left            =   6660
      Picture         =   "frmLogPlanAnualInfo.frx":0684
      Top             =   2400
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgXX 
      Height          =   240
      Left            =   6660
      Picture         =   "frmLogPlanAnualInfo.frx":0D86
      Top             =   2040
      Width           =   240
   End
   Begin VB.Image imgOK 
      Height          =   240
      Left            =   6660
      Picture         =   "frmLogPlanAnualInfo.frx":10C8
      Top             =   1680
      Width           =   240
   End
   Begin VB.Image imgFlecha 
      Height          =   240
      Index           =   1
      Left            =   2820
      Picture         =   "frmLogPlanAnualInfo.frx":140A
      Top             =   780
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmLogPlanAnualInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nPlanNro As Integer, nNroReq As Integer, bReqNroPro As Boolean

Public Sub PlanAnual(vPlanNro As Long, Optional ByVal pbReqNroPro As Boolean = False, Optional ByVal pnNroReq As Integer = 0)
bReqNroPro = pbReqNroPro
If bReqNroPro Then
    nNroReq = pnNroReq
Else
    nPlanNro = vPlanNro
End If
Me.Show 1
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
GeneraVisorTramite
End Sub

Sub GeneraVisorTramite()
Dim oConn As New DConecta
Dim rs As New ADODB.Recordset
Dim sSQL As String

If oConn.AbreConexion Then

    If bReqNroPro Then
        sSQL = "select r.cPersCod, cPersona = replace(p.cPersNombre,'/',' '), cCargo=t.cRHCargoDescripcion " & _
              "  from LogProSelReq r inner join Persona p on r.cPersCod = p.cPersCod " & _
              "       inner join RHCargosTabla t on r.cRHCargoCod = t.cRHCargoCod " & _
              " where r.nProSelReqNro=" & nNroReq & ""
    Else
        sSQL = "select r.cPersCod, cPersona = replace(p.cPersNombre,'/',' '), cCargo=t.cRHCargoDescripcion " & _
              "  from LogPlanAnualReq r inner join Persona p on r.cPersCod = p.cPersCod " & _
              "       inner join RHCargosTabla t on r.cRHCargoCod = t.cRHCargoCod " & _
              " where r.nPlanReqNro=" & nPlanNro & ""
    End If
          
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      lblNivel(0).Caption = UCase(rs!cPersona)
      lblCargo(0).Caption = rs!cCargo
   End If
   
   If bReqNroPro Then
        sSQL = "select nNivel=a.nNivelAprobacion,cCargo=t.cRHCargoDescripcion, nEstado=a.nEstadoAprobacion " & _
               "  from LogProSelAprobacion a inner join RHCargosTabla t on a.cRHCargoCodAprobacion = t.cRHCargoCod " & _
               " where a.nProSelReqNro = " & nNroReq & ""
   Else
        sSQL = "select nNivel=a.nNivelAprobacion,cCargo=t.cRHCargoDescripcion, nEstado=a.nEstadoAprobacion " & _
               "  from LogPlanAnualAprobacion a inner join RHCargosTabla t on a.cRHCargoCodAprobacion = t.cRHCargoCod " & _
               " where a.nPlanReqNro = " & nPlanNro & ""
   End If
          
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      Do While Not rs.EOF
         fraNivel(rs!nNivel).Visible = True
         lblCargo(rs!nNivel).Caption = rs!cCargo
         'lblNivel(rs!nNivel).Caption = rs!cCargo
         imgFlecha(rs!nNivel).Visible = True
         If rs!nEstado = 1 Then
            Set imgNivel(rs!nNivel) = imgOK
         Else
            Set imgNivel(rs!nNivel) = imgXX
         End If
         rs.MoveNext
      Loop
   End If
   
   
End If

End Sub

