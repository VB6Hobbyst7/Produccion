VERSION 5.00
Begin VB.Form frmSeleccionCtaAhorroERViaticos 
   Caption         =   "SELECCIÓN DE CUENTA DE AHORRO"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   7530
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Frame fraCtaAhorroDisp 
      Caption         =   "Cuentas de Ahorro Disponibles"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   240
      TabIndex        =   13
      Top             =   2400
      Width           =   6975
      Begin Sicmact.FlexEdit flxCtaAHDisp 
         Height          =   2655
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4683
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Check-Cuenta-Moneda-Programa"
         EncabezadosAnchos=   "500-750-1800-900-2500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-X-X"
         ListaControles  =   "0-4-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-L"
         FormatosEdit    =   "0-0-0-0-0"
         CantEntero      =   20
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame fraDatosCom 
      Caption         =   "Datos Personales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cargo :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   12
         Top             =   1485
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Area :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Agencia :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   10
         Top             =   1170
         Width           =   750
      End
      Begin VB.Label txtCargoDes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1140
         TabIndex        =   9
         Top             =   1440
         Width           =   5475
      End
      Begin VB.Label txtAreaCod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1140
         TabIndex        =   8
         Top             =   780
         Width           =   585
      End
      Begin VB.Label txtAgecod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1140
         TabIndex        =   7
         Top             =   1110
         Width           =   585
      End
      Begin VB.Label txtAreaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1740
         TabIndex        =   6
         Top             =   780
         Width           =   4875
      End
      Begin VB.Label txtAgeDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1740
         TabIndex        =   5
         Top             =   1110
         Width           =   4875
      End
      Begin VB.Label txtpersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1140
         TabIndex        =   4
         Top             =   450
         Width           =   5475
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   3
         Top             =   495
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmSeleccionCtaAhorroERViaticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'*** Nombre : frmSeleccionCtaAhorroERViaticos
'*** Descripción : Formulario para seleccionar la cuenta de Ahorro, para aprobar la Solicitud de E/R y Viáticos
'*** Creación : NAGL el 20171115
'********************************************************************************
Dim oRend As New NARendir
Dim ValorCel1, ValorCel2, ValorCel3 As String
Public lsCtaAhorro As String
Public Sub Inicio(psPersCod As String, psMoneda As String)
Call CargaDataComisionado(psPersCod)
If CargarCtasAhorroComis(psPersCod, psMoneda) Then
    CentraForm Me
    Me.Show 1
Else
    lsCtaAhorro = ""
End If
End Sub

Public Sub CargaDataComisionado(psPersCod As String)
Dim rsRendCom As New ADODB.Recordset
Set rsRendCom = oRend.ObtenerDatosGeneralComisionado(psPersCod)

If Not (rsRendCom.EOF And rsRendCom.BOF) Then
    txtpersNombre = PstaNombre(rsRendCom!cPersNombre)
    txtAreaCod = rsRendCom!cAreaCod
    txtAreaDesc = rsRendCom!cAreaDescripcion
    txtAgecod = rsRendCom!cAgeCod
    txtAgeDesc = rsRendCom!cAgeDescripcion
    txtCargoDes = rsRendCom!cRHCargoDes
End If
End Sub

Public Function CargarCtasAhorroComis(psPersCod As String, psMoneda As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim X As Integer
        Set rs = oRend.ObtenerCtaAhorroComis(psPersCod, psMoneda)
        flxCtaAHDisp.Clear
        FormateaFlex flxCtaAHDisp
        If Not (rs.EOF And rs.BOF) Then
            For X = 1 To rs.RecordCount
                flxCtaAHDisp.AdicionaFila , , True
                flxCtaAHDisp.TextMatrix(X, 2) = rs!cCtaCod
                flxCtaAHDisp.TextMatrix(X, 3) = rs!Moneda
                flxCtaAHDisp.TextMatrix(X, 4) = rs!DescripProgram
                rs.MoveNext
            Next
        Else
            MsgBox "El Usuario " & txtpersNombre & " no posee ninguna Cuenta de Ahorro Disponible!!", vbInformation, "Aviso"
            Exit Function
        End If
    CargarCtasAhorroComis = True
End Function

Private Sub cmdAceptar_Click()
Dim nFilasCheck, i, nFilasTotal As Integer
nFilasCheck = 0
nFilasTotal = flxCtaAHDisp.Rows - 1
For i = 1 To nFilasTotal
    If flxCtaAHDisp.TextMatrix(i, 1) = "." Then
       nFilasCheck = nFilasCheck + 1
    End If
Next i

If nFilasCheck > 1 Then
    MsgBox "Seleccione solo una cuenta de Ahorro..!!", vbInformation, "Aviso"
    For i = 1 To nFilasTotal
        If flxCtaAHDisp.TextMatrix(i, 1) = "." Then
           flxCtaAHDisp.TextMatrix(i, 1) = ""
        End If
    Next i
    Exit Sub
ElseIf nFilasCheck = 0 Then
    MsgBox "Para Continuar, se debe seleccionar una cuenta Ahorro..!!", vbInformation, "Aviso"
    Exit Sub
Else
    lsCtaAhorro = flxCtaAHDisp.TextMatrix(flxCtaAHDisp.Row, 2)
    Unload Me
End If
End Sub

Private Sub cmdCancelar_Click()
  lsCtaAhorro = ""
  Unload Me
End Sub

'****NAGL 20180225****
Private Sub flxCtaAHDisp_EnterCell()
  If flxCtaAHDisp.col = 2 Then
        ValorCel1 = flxCtaAHDisp.TextMatrix(flxCtaAHDisp.Row, flxCtaAHDisp.col)
  ElseIf flxCtaAHDisp.col = 3 Then
        ValorCel2 = flxCtaAHDisp.TextMatrix(flxCtaAHDisp.Row, flxCtaAHDisp.col)
  ElseIf flxCtaAHDisp.col = 4 Then
        ValorCel3 = flxCtaAHDisp.TextMatrix(flxCtaAHDisp.Row, flxCtaAHDisp.col)
  End If
End Sub

Private Sub flxCtaAHDisp_OnCellChange(pnRow As Long, pnCol As Long)
  If (pnCol = 2) Then
        flxCtaAHDisp.TextMatrix(pnRow, pnCol) = ValorCel1
  ElseIf (pnCol = 3) Then
        flxCtaAHDisp.TextMatrix(pnRow, pnCol) = ValorCel2
  ElseIf (pnCol = 4) Then
        flxCtaAHDisp.TextMatrix(pnRow, pnCol) = ValorCel3
  End If
End Sub
'***END NAGL 20180205***
