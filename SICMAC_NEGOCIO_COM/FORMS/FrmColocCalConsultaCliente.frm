VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmColocCalConsultaCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Calificacion del Cliente en Sistema"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   Icon            =   "FrmColocCalConsultaCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   6615
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   3000
         TabIndex        =   13
         Top             =   170
         Width           =   1095
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5400
         TabIndex        =   12
         Top             =   170
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4200
         TabIndex        =   11
         Top             =   170
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   3855
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label lblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label lblPersCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   495
      End
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"FrmColocCalConsultaCliente.frx":030A
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   7
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin VB.ListBox LstMeses 
      Height          =   1860
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   2520
      X2              =   6360
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Unidad de Riesgos"
      Height          =   195
      Left            =   2520
      TabIndex        =   14
      Top             =   400
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Caja Municipal de Cusco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2520
      TabIndex        =   9
      Top             =   105
      Width           =   2220
   End
End
Attribute VB_Name = "FrmColocCalConsultaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sServidor As String

Private Sub CargaPersona(ByVal pPersona As COMDPersona.UCOMPersona)
    If Not pPersona Is Nothing Then
        lblPersCodigo.Caption = Trim(pPersona.sPersCod)
        lblCliente.Caption = Trim(pPersona.sPersNombre)
    End If
End Sub

Private Sub cmdBuscar_Click()
Dim sql As String
Dim I As Integer
Dim rs As New ADODB.Recordset

'Dim oPer As DPersona
Dim s As COMDPersona.UCOMPersona

Dim Riesgo As COMDCredito.DCOMColocEval


MSH.Clear
MSH.Rows = 2
lblPersCodigo = ""
lblCliente = ""
Marco
Dim dFecha As String
Set rs = New ADODB.Recordset
Call CargaPersona(frmBuscaPersona.inicio)

If Len(lblPersCodigo) = 0 Then Exit Sub

Set Riesgo = New COMDCredito.DCOMColocEval

For I = 0 To LstMeses.ListCount - 1
 If LstMeses.Selected(I) = True Then
    If Len(Me.lblPersCodigo) > 0 Then
        Set rs = Riesgo.FechaCalificacion(sServidor, Format(LstMeses.List(I), "YYYY/MM/DD"), Me.lblPersCodigo)
        While Not rs.EOF
            If MSH.Rows >= 2 And MSH.TextMatrix(MSH.Rows - 1, 1) = "" Then
            Else
                MSH.Rows = MSH.Rows + 1
            End If
            MSH.TextMatrix(MSH.Rows - 1, 1) = Format(rs!dFecha, "DD/MM/YYYY")
            MSH.TextMatrix(MSH.Rows - 1, 2) = rs!cCtaCod
            MSH.TextMatrix(MSH.Rows - 1, 3) = rs!Calgen
            MSH.TextMatrix(MSH.Rows - 1, 4) = rs!CalSis
            MSH.TextMatrix(MSH.Rows - 1, 5) = rs!CMAC
            MSH.TextMatrix(MSH.Rows - 1, 6) = rs!CalHist
            rs.MoveNext
        Wend
        rs.Close
        'If MSH.Rows > 2 Then MSH.Rows = MSH.Rows - 1
    End If
 End If
Next I

End Sub

Private Sub cmdCancelar_Click()
Dim I As Integer
MSH.Clear
MSH.Rows = 2
Marco
lblCliente = ""
lblPersCodigo = ""
rtf = ""
For I = 0 To LstMeses.ListCount - 1
  LstMeses.Selected(I) = False
Next I
End Sub

Private Sub CmdImprimir_Click()
Dim oPrevio As previo.clsPrevio
Set oPrevio = New previo.clsPrevio

'Dim lscadena As String
'Dim nFicSal As Integer
'nFicSal = FreeFile
''If Len(MSH.TextMatrix(MSH.Rows - 1, 1)) = 0 Then
''   MsgBox "La persona no posee calificacion", vbInformation, "AVISO"
''   Exit Sub
''End If
'
'lscadena = gsNomCmac & "                                                             " & gdFecSis & Chr(10) & Chr(10)
'lscadena = lscadena & "                     CONSULTA DE CALIFICACION DE CLIENTE" & Chr(10)
'lscadena = lscadena & "Nombre     : " & lblCliente & Chr(10)
'lscadena = lscadena & "Direccion  : " & DirGrid & Chr(10)
'lscadena = lscadena & "-----------------------------------------------------------------------------" & Chr(10)
'lscadena = lscadena & " Fecha    Nro Credito       Cal. Final.  Cal. Sist F.  Cal.CMAC   Cal Riesgo." & Chr(10)
'lscadena = lscadena & "-----------------------------------------------------------------------------" & Chr(10)
'For i = 1 To MSH.Rows - 1
'    lscadena = lscadena & MSH.TextMatrix(i, 1) & Space(6) & MSH.TextMatrix(i, 2) & Space(8) & MSH.TextMatrix(i, 3) & Space(11) & MSH.TextMatrix(i, 4) & Space(13) & MSH.TextMatrix(i, 5) & Space(9) & MSH.TextMatrix(i, 6) & Chr(10)
'Next i
'lscadena = lscadena & "-----------------------------------------------------------------------------" & Chr(10)
'lscadena = lscadena & gPrnSaltoPagina   'expulsa la pagina
'rtf = lscadena
'
''frmPrevio.Previo rtf, "Consulta de Calificacion de Cliente", True, 25
'oPrevio.Show lscadena, Me.Caption, True, 66
'
'Set oPrevio = Nothing
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim I As Integer
Dim Fecha As Date
Dim Rcc As COMDCredito.DCOMColocEval

For I = 0 To 11
    Fecha = DateAdd("m", -I, gdFecDataFM)
    LstMeses.AddItem Format(Fecha, "MMM - YYYY")
Next I
Marco
Set Rcc = New COMDCredito.DCOMColocEval
sServidor = Rcc.ServConsol(gConstSistServCentralRiesgos)

Set Rcc = Nothing
End Sub

Sub Marco()
With MSH
    .ColWidth(0) = 100
    .ColWidth(1) = 1200 '1500
    .ColWidth(2) = 2000 '1200
    .ColWidth(3) = 800
    .ColWidth(4) = 800
    .ColWidth(5) = 800
    .ColWidth(6) = 900
    .ColAlignment(3) = flexAlignCenterCenter
    .ColAlignment(4) = flexAlignCenterCenter
    .ColAlignment(5) = flexAlignCenterCenter
    .ColAlignment(6) = flexAlignCenterCenter
    .TextMatrix(0, 1) = "Fecha"
    .TextMatrix(0, 2) = "Credito"
    .TextMatrix(0, 3) = "Cal. Final."
    .TextMatrix(0, 4) = "Cal. Sist. F."
    .TextMatrix(0, 5) = "Cal. CMAC"
    .TextMatrix(0, 6) = "Cal. Riesgo"
End With
End Sub
