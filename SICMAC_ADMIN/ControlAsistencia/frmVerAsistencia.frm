VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmVerAsistencia 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "frmVerAsistencia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grAsistencia 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   105
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6165
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   400
      Left            =   2760
      TabIndex        =   0
      Top             =   3720
      Width           =   1455
   End
End
Attribute VB_Name = "frmVerAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FormatGrid()
    grAsistencia.ColWidth(0) = 0
    grAsistencia.ColWidth(1) = 3500
    grAsistencia.ColWidth(2) = 2000
    grAsistencia.ColWidth(3) = 2000
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsVer As Recordset
    Set rsVer = New Recordset
    
    VSQL = " Select PE.cPersNombre,dRHAsistenciaIngreso,dRHAsistenciaSalida" _
         & " From RHAsistenciaDet TA" _
         & " Inner Join Persona PE ON TA.cPersCod = PE.cPersCod" _
         & " where datediff(day,dRHAsistenciaFechaRef,'" & Format(CDate(frmAsistencia.mskFecha), "mm/dd/yyyy") & "') = 0 Order by cPersNombre"
    Set rsVer = oCon.CargaRecordSet(VSQL)
    Set grAsistencia.Recordset = rsVer
    rsVer.Close
    Set rsVer = Nothing
    FormatGrid
End Sub
