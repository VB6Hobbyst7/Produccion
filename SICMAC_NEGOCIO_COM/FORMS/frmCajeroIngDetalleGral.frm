VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCajeroIngDetalleGral 
   Caption         =   "Detalle de Operaciones General"
   ClientHeight    =   780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8070
   Icon            =   "frmCajeroIngDetalleGral.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   780
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frageneral 
      Height          =   720
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7905
      Begin VB.ComboBox cbocajero 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   3495
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6600
         TabIndex        =   1
         Top             =   240
         Width           =   1230
      End
      Begin MSComCtl2.DTPicker txtFecha 
         Height          =   330
         Left            =   5160
         TabIndex        =   2
         Top             =   240
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   582
         _Version        =   393216
         Format          =   66191361
         CurrentDate     =   37090
      End
      Begin VB.Label lblCaptionUser 
         AutoSize        =   -1  'True
         Caption         =   "Cajero :"
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
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   4560
         TabIndex        =   3
         Top             =   360
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmCajeroIngDetalleGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oGen As COMDConstSistema.DCOMGeneral  'DGeneral
Dim oPrevio As previo.clsprevio
'MADM 20101020

Private Sub CargarCajero_Combo()
Dim rs As ADODB.Recordset
Dim orep As New COMDCaptaGenerales.COMDCaptaReportes

    Set orep = New COMDCaptaGenerales.COMDCaptaReportes
        Set rs = orep.RecuperaRfxAgencia(gsCodAge, True)
    Set orep = Nothing
    
    Call llenar_cbo_Cajero(rs, cbocajero)
    
    Exit Sub
End Sub

Sub llenar_cbo_Cajero(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
pcboObjeto.Clear
Do While Not pRs.EOF
    pcboObjeto.AddItem Trim(pRs!cPersNombre) & Space(100) & Trim(pRs!cPersCod)
    pRs.MoveNext
Loop
    pcboObjeto.ListIndex = CInt(0)
pRs.Close
End Sub

Sub presentar_reporte(ByVal X As Boolean)
        Dim orep1 As COMNCaptaGenerales.NCOMCaptaReportes
        Dim sCad As String
        
        Set orep1 = New COMNCaptaGenerales.NCOMCaptaReportes
        
        If X Then
            sCad = orep1.ReporteTrasTotSM("DETALLE DE OPERACIONES GENERAL", False, gsCodUser, Format$(gdFecSis, "yyyymmdd"), gsCodAge)
        Else
            
            sCad = orep1.ReporteTrasTotSM("DETALLE DE OPERACIONES GENERAL", False, Right(cbocajero.Text, 4), Format$(Me.txtFecha.value, "yyyymmdd"))
        End If
        
        Set orep1 = Nothing

        Set oPrevio = New previo.clsprevio
        oPrevio.Show sCad, "DETALLE DE OPERACIONES GENERAL", True
        Set oPrevio = Nothing
End Sub

Private Sub cmdProcesar_Click()

      If cbocajero.ListIndex = 0 Then
        presentar_reporte (True)
      Else
        presentar_reporte (False)
      End If

    
End Sub

Private Sub Form_Load()
    txtFecha.value = gdFecSis
    Me.Caption = "Detalle de Operaciones General"
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    cbocajero.Enabled = True
    CargarCajero_Combo
End Sub

