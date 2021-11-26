VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMuestraPersRCC 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mostrar Relacion de Personal"
   ClientHeight    =   3660
   ClientLeft      =   1770
   ClientTop       =   3930
   ClientWidth     =   7575
   Icon            =   "frmMuestraPersRCC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdcancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   6060
      TabIndex        =   2
      Top             =   3165
      Width           =   1275
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   4770
      TabIndex        =   1
      Top             =   3180
      Width           =   1275
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2880
      Left            =   105
      TabIndex        =   0
      Top             =   195
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   5080
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "cnomcli"
         Caption         =   "Nombre de Persona"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "cnudoci"
         Caption         =   "N° Doc. Natural"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "cnudotr"
         Caption         =   "N° Doc.Trib."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   3525.166
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMuestraPersRCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sRutFil As String
Dim lsNombre As String
Dim psCodSbs As String
Dim psNomCli As String
Dim lsAnio As String
Dim lsMes As String
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Public Sub Inicio(ByVal lsRutFil As String, ByVal psNombre As String, ByVal psMes As String, ByVal psAnio As String)
sRutFil = lsRutFil
lsNombre = psNombre
lsMes = psMes
lsAnio = psAnio
Me.Show 1
End Sub

Private Sub CmdAceptar_Click()
Unload Me
End Sub

Private Sub cmdCancelar_Click()
psCodSbs = ""
psNomCli = ""
Unload Me
End Sub
Private Sub DataGrid1_GotFocus()
DataGrid1.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub DataGrid1_LostFocus()
DataGrid1.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub Form_Load()
CentraForm Me
CargaLista sRutFil, lsNombre
End Sub
Sub CargaLista(ByVal lsRuta As String, ByVal lsNombre As String)
Dim sCadSQl As String
Dim oCon As DConecta
Set oCon = New DConecta

oCon.AbreConexion

psCodSbs = ""

sCadSQl = "SELECT  Nom_Deu AS cNomCli,Tip_Pers AS cTipPers,Cod_Sbs AS CCODSBS, Cod_Doc_Id AS cNuDocI , Cod_Doc_Trib AS cNudOtr, Can_Ents AS nCanEmp, " _
        & "         Calif_0 AS nCalif0,Calif_1 AS nCalif1,Calif_2 AS nCalif2,Calif_3 AS nCalif3,Calif_4 AS nCalif4  " _
        & "FROM   " & lsRuta & "RCCTOTAL " _
        & "WHERE Nom_Deu LIKE '" & lsNombre & "%' and month(Fec_Rep)=" & Val(lsMes) & " and year(Fec_Rep)=" & Val(lsAnio) & " ORDER BY Nom_Deu "

'AbreConexion
Set rs = oCon.CargaRecordSet(sCadSQl)
'CierraConexion

Set DataGrid1.DataSource = rs
DataGrid1.Refresh

oCon.CierraConexion

Set oCon = Nothing
End Sub
Public Property Get lsCodSBS() As Variant
    lsCodSBS = psCodSbs
End Property
Public Property Let lsCodSBS(ByVal vNewValue As Variant)
    lsCodSBS = vNewValue
End Property
Private Sub rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If Not pRecordset.EOF And Not pRecordset.BOF Then
    psCodSbs = pRecordset!cCodSBS
    psNomCli = pRecordset!cNomCli
End If
End Sub
Public Property Get lsNomCli() As Variant
lsNomCli = psNomCli
End Property
Public Property Let lsNomCli(ByVal vNewValue As Variant)
End Property
