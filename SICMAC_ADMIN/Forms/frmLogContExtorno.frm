VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLogContExtorno 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contratación: Extorno de Contrato"
   ClientHeight    =   4155
   ClientLeft      =   2220
   ClientTop       =   1200
   ClientWidth     =   8070
   Icon            =   "frmLogContExtorno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTContratos 
      Height          =   4020
      Left            =   80
      TabIndex        =   0
      Top             =   80
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7091
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Extorno"
      TabPicture(0)   =   "frmLogContExtorno.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancelar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdExtornar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraProv"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame fraProv 
         Caption         =   "Contrato"
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
         Height          =   2805
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   7680
         Begin VB.TextBox txtGlosa 
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
            Height          =   1170
            Left            =   1080
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   1200
            Width           =   6180
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Glosa:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   480
            TabIndex        =   10
            Top             =   1200
            Width           =   450
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Nº Contrato:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   270
            Width           =   870
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   750
            Width           =   780
         End
         Begin VB.Label lblNContrato 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   7
            Tag             =   "txtnombre"
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lblProveedor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   6
            Tag             =   "txtnombre"
            Top             =   720
            Width           =   6135
         End
      End
      Begin VB.CommandButton cmdExtornar 
         Caption         =   "&Extornar"
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
         Left            =   5735
         TabIndex        =   4
         Top             =   3360
         Width           =   1005
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
         Height          =   360
         Left            =   6750
         TabIndex        =   3
         Top             =   3360
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "_"
         ForeColor       =   &H80000008&
         Height          =   75
         Left            =   3240
         TabIndex        =   1
         Top             =   2760
         Width           =   90
      End
   End
   Begin VB.PictureBox CdlgFile 
      Height          =   615
      Left            =   7440
      ScaleHeight     =   555
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   800
   End
End
Attribute VB_Name = "frmLogContExtorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim fsNContrato As String
Dim fnContRef As Integer 'pasi20140823 ti-ers077-2014
Dim fnEstado As Integer
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdExtornar_Click()
On Error GoTo ErrorExtorno
If ValidaExtorno Then
    If MsgBox("Esta seguro de grabar el Extorno?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    Dim oLog As DLogGeneral
    Set oLog = New DLogGeneral
    
        If oLog.RegistrarExtornoAdenda(Trim(fsNContrato), fnContRef, , , Trim(Me.txtGlosa.Text), GeneraMov(gdFecSis, "109", gsCodAge, gsCodUser), 1) = 0 Then
            If oLog.ExtornarContrato(Trim(fsNContrato), fnContRef) = 0 Then
                MsgBox "Extono de Contrato registrada Satisfactoriamente", vbInformation, "Aviso"
                'ARLO 20160126 ***
                gsOpeCod = LogPistaExtornoContrato
                Set objPista = New COMManejador.Pista
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Extorno Contrato N° : " & fsNContrato & " | Por Motivo : " & Trim(Me.txtGlosa.Text)
                Set objPista = Nothing
                '***
                LimpiarDatos
                frmLogContSeguimiento.CargarGrid
                Unload Me
            Else
                MsgBox "No se registro el extorno del Contrato", vbInformation, "Aviso"
            End If
        Else
            MsgBox "No se registro el extorno del Contrato", vbInformation, "Aviso"
        End If
    
    End If
End If
Exit Sub
ErrorExtorno:
    MsgBox Err.Number & " - " & Err.Description, vbInformation, "Error"


End Sub
Private Function ValidaExtorno() As Boolean
If Trim(Me.txtGlosa.Text) = "" Then
    MsgBox "Ingrese la Glosa del Extorno", vbInformation, "Aviso"
    ValidaExtorno = False
    Exit Function
End If

ValidaExtorno = True
End Function

Public Sub Inicio(ByVal psNContrato As String, ByVal pnContRef As Integer) 'pnContRef agregado PASI20140823 ti-ers077-2014
    fsNContrato = psNContrato
    fnContRef = pnContRef
    If Not ExtornoOK Then Exit Sub
    Call CargaDatos
    Show 1
End Sub


Private Sub CargaDatos()
Dim oLog As DLogGeneral
Dim rsLog As ADODB.Recordset

Set oLog = New DLogGeneral
Set rsLog = oLog.ListarDatosContratos(fsNContrato, fnContRef)

If rsLog.RecordCount > 0 Then
    Me.lblNContrato.Caption = Space(1) & rsLog!NContrato
    Me.lblProveedor.Caption = Space(1) & rsLog!Proveedor
End If

End Sub
Sub LimpiarDatos()
Me.txtGlosa.Text = ""

End Sub
Private Function ExtornoOK() As Boolean
    Dim oLog As New DLogGeneral
    ExtornoOK = True
    If oLog.RealizoPagoContrato(fsNContrato, fnContRef) Then
        ExtornoOK = False
        MsgBox "No se puede realizar el extorno del Contrato porque ya se realizaron movimiento para pagos de sus cuotas, verifique..", vbInformation, "Aviso"
        Exit Function
    End If
End Function

Private Sub SSTContratos_DblClick()

End Sub
