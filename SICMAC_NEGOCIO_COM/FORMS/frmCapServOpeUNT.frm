VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCapServOpeUNT 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   Icon            =   "frmCapServOpeUNT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEstudiante 
      Caption         =   "Estudiante"
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
      Height          =   1005
      Left            =   120
      TabIndex        =   27
      Top             =   60
      Width           =   6765
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2300
         TabIndex        =   1
         Top             =   232
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskCodAlumno 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "CC#####-##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblNom 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   100
         TabIndex        =   32
         Top             =   630
         Width           =   645
      End
      Begin VB.Label lblEsc 
         AutoSize        =   -1  'True
         Caption         =   "Escuela :"
         Height          =   195
         Left            =   2880
         TabIndex        =   31
         Top             =   300
         Width           =   660
      End
      Begin VB.Label lblCodEst 
         AutoSize        =   -1  'True
         Caption         =   "Codigo :"
         Height          =   195
         Left            =   105
         TabIndex        =   30
         Top             =   300
         Width           =   585
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         TabIndex        =   29
         Top             =   600
         Width           =   5760
      End
      Begin VB.Label lblEscuela 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   28
         Top             =   240
         Width           =   3000
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   6750
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   6750
      Width           =   1095
   End
   Begin VB.Frame fraConcepto 
      Caption         =   "Orden de Pago"
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
      Height          =   5550
      Left            =   120
      TabIndex        =   11
      Top             =   1125
      Width           =   6765
      Begin TabDlg.SSTab tabConceptos 
         Height          =   5175
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   9128
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabHeight       =   520
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&Matrícula"
         TabPicture(0)   =   "frmCapServOpeUNT.frx":030A
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txtMonto"
         Tab(0).Control(1)=   "Frame1"
         Tab(0).Control(2)=   "grdConcepto"
         Tab(0).Control(3)=   "Label2"
         Tab(0).Control(4)=   "lblTotal"
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "&Otros"
         TabPicture(1)   =   "frmCapServOpeUNT.frx":0326
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label4"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "lblTotOtros"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "grdOtros"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "txtMonOtros"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Frame2"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
         Begin VB.TextBox txtMonto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -73920
            TabIndex        =   7
            Top             =   1470
            Width           =   1500
         End
         Begin VB.Frame Frame1 
            Height          =   975
            Left            =   -74880
            TabIndex        =   18
            Top             =   360
            Width           =   6255
            Begin VB.CheckBox chkSegProf 
               Alignment       =   1  'Right Justify
               Caption         =   "&2da. Profesión"
               Height          =   255
               Left            =   3720
               TabIndex        =   4
               Top             =   240
               Width           =   1395
            End
            Begin VB.OptionButton optMatricula 
               Caption         =   "&Regular"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   2
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton optMatricula 
               Caption         =   "&Extemporánea"
               Height          =   255
               Index           =   1
               Left            =   1200
               TabIndex        =   3
               Top             =   240
               Width           =   1335
            End
            Begin MSMask.MaskEdBox mskNumRecibo 
               Height          =   315
               Left            =   1800
               TabIndex        =   5
               Top             =   585
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Line Line1 
               BorderWidth     =   3
               X1              =   1575
               X2              =   1695
               Y1              =   735
               Y2              =   735
            End
            Begin VB.Label lblRecibo 
               AutoSize        =   -1  'True
               Caption         =   "Orden Pago :"
               Height          =   195
               Left            =   120
               TabIndex        =   20
               Top             =   645
               Width           =   945
            End
            Begin VB.Label lblCodEsc 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1095
               TabIndex        =   19
               Top             =   585
               Width           =   375
            End
         End
         Begin VB.Frame Frame2 
            Height          =   975
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   6135
            Begin MSMask.MaskEdBox mskOrdenOtros 
               Height          =   315
               Left            =   1815
               TabIndex        =   15
               Top             =   360
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label lblEscOtros 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1200
               TabIndex        =   17
               Top             =   360
               Width           =   375
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Orden Pago :"
               Height          =   195
               Left            =   120
               TabIndex        =   16
               Top             =   420
               Width           =   945
            End
            Begin VB.Line Line2 
               BorderWidth     =   3
               X1              =   1575
               X2              =   1695
               Y1              =   510
               Y2              =   510
            End
         End
         Begin VB.TextBox txtMonOtros 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   13
            Top             =   1800
            Width           =   1500
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdOtros 
            Height          =   2415
            Left            =   120
            TabIndex        =   21
            Top             =   1560
            Width           =   6150
            _ExtentX        =   10848
            _ExtentY        =   4260
            _Version        =   393216
            TextStyleFixed  =   1
            ScrollBars      =   0
            Appearance      =   0
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
            _Band(0).BandIndent=   2
            _Band(0).Cols   =   2
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   1
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdConcepto 
            Height          =   3375
            Left            =   -74880
            TabIndex        =   6
            Top             =   1380
            Width           =   6210
            _ExtentX        =   10954
            _ExtentY        =   5953
            _Version        =   393216
            TextStyleFixed  =   1
            ScrollBars      =   0
            Appearance      =   0
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
            _Band(0).BandIndent=   2
            _Band(0).Cols   =   2
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   1
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total : S/."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   -71655
            TabIndex        =   25
            Top             =   4740
            Width           =   1500
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   -70170
            TabIndex        =   24
            Top             =   4740
            Width           =   1500
         End
         Begin VB.Label lblTotOtros 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   4710
            TabIndex        =   23
            Top             =   3960
            Width           =   1500
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total : S/."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   3225
            TabIndex        =   22
            Top             =   3960
            Width           =   1500
         End
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   10
      Top             =   6750
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox R 
      Height          =   135
      Left            =   0
      TabIndex        =   26
      Top             =   60
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   238
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmCapServOpeUNT.frx":0342
   End
End
Attribute VB_Name = "frmCapServOpeUNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCurr As String
Public s2daProf As String
Public sTipoMat As String
Public sIngresante As String
Public sCuenta As String
Public sCuentaDonacion As String

Private Function GetConceptoGrid(ByVal nConcepto As CapServUNTConcepto, grid As MSHFlexGrid) As Double
Dim i As Long
Dim nConcep As CapServUNTConcepto
Dim nValor As Double
nValor = 0
For i = 1 To grid.Rows - 1
    nConcep = CLng(grid.TextMatrix(i, 1))
    If nConcep = nConcepto Then
        nValor = CDbl(grid.TextMatrix(i, 3))
        Exit For
    End If
Next i
GetConceptoGrid = nValor
End Function

Private Function GetRecordsetGrid(grid As MSHFlexGrid) As Recordset
Dim i As Long
Dim J As Long
Dim rsAux As ADODB.Recordset
Dim nFila As Long
Dim nCol As Long
Dim sTipoDato As DataTypeEnum
Dim sTamCampo As Long
If grid.TextMatrix(1, 1) <> "" Then
    nFila = 0
    'formamos generamos del recordset
    Set rsAux = New Recordset
    For i = 1 To grid.Cols - 1
        If Len(Trim(grid.TextMatrix(nFila + 1, i))) >= 16 Then
            sTipoDato = adVarChar
        Else
            If IsNumeric(grid.TextMatrix(nFila + 1, i)) And grid.ColAlignment(i) >= 7 Then
                sTipoDato = adDouble
            Else
                If ValidaFecha(grid.TextMatrix(nFila + 1, i)) = "" Then
                    sTipoDato = adDate
                Else
                    sTipoDato = adVarChar
                End If
            End If
        End If
        If sTipoDato = adVarChar Then
            rsAux.Fields.Append grid.TextMatrix(nFila, i), sTipoDato, Val(adVarChar), adFldMayBeNull
        Else
            rsAux.Fields.Append grid.TextMatrix(nFila, i), sTipoDato, , adFldMayBeNull
        End If
    Next
    rsAux.Open
    For i = 1 To grid.Rows - 1
        rsAux.AddNew
        'columnas
        For J = 1 To grid.Cols - 1
            If rsAux.Fields(grid.TextMatrix(0, J)).Type = adDouble Then
                rsAux.Fields(grid.TextMatrix(0, J)) = CCur(IIf(grid.TextMatrix(i, J) = "", "0", grid.TextMatrix(i, J)))
            Else
                If rsAux.Fields(grid.TextMatrix(0, J)).Type = adDate Then
                    rsAux.Fields(grid.TextMatrix(0, J)) = IIf(grid.TextMatrix(i, J) = "", Null, grid.TextMatrix(i, J))
                Else
                    rsAux.Fields(grid.TextMatrix(0, J)) = grid.TextMatrix(i, J)
                End If
            End If
        Next
        rsAux.Update
    Next
    rsAux.MoveFirst
    Set GetRecordsetGrid = rsAux
End If
End Function

Private Sub SetupGrid(grid As MSHFlexGrid)
grid.Clear
grid.Cols = 4
grid.Rows = 2
grid.TextMatrix(0, 0) = ""
grid.TextMatrix(0, 1) = "Codigo"
grid.TextMatrix(0, 2) = "Concepto"
grid.TextMatrix(0, 3) = "Monto"
grid.ColWidth(0) = 150
grid.ColWidth(1) = 0
grid.ColWidth(2) = 4500
grid.ColWidth(3) = 1500
grid.ColAlignment(0) = 4
grid.ColAlignment(1) = 4
grid.ColAlignment(2) = 1
grid.ColAlignment(3) = 7
grid.ColAlignmentFixed(0) = 4
grid.ColAlignmentFixed(1) = 4
grid.ColAlignmentFixed(2) = 4
grid.ColAlignmentFixed(3) = 4
End Sub

Private Sub ActualizaMatricula()
If grdConcepto.Rows <= 1 Then
    Exit Sub
End If
Dim rsMat As ADODB.Recordset
Dim clsServ As NCapServicios
Set clsServ = New NCapServicios
Set rsMat = New ADODB.Recordset
Set rsMat = clsServ.GetUNTConceptosAlumno(sCurr, sTipoMat, sIngresante, s2daProf)
Set clsServ = Nothing
If Not (rsMat.EOF And rsMat.BOF) Then
    Dim i As Integer, nFila As Integer, nValor As Integer
    nFila = grdConcepto.Row
    For i = 1 To grdConcepto.Rows - 1
        nValor = CInt(grdConcepto.TextMatrix(i, 1))
        If nValor = gUNTMatReg Or nValor = gUNTMatExt Or nValor = gUNTMatReg2daProf Or nValor = gUNTMatExt2daProf Then
            grdConcepto.TextMatrix(i, 1) = rsMat("nConceptoCod")
            grdConcepto.TextMatrix(i, 2) = rsMat("Concepto")
            grdConcepto.TextMatrix(i, 3) = Format$(rsMat("Monto"), "#,##0.00")
            GetTotalPagar grdConcepto, lblTotal
            Exit For
        End If
    Next i
    grdConcepto.Row = nFila
End If
rsMat.Close
Set rsMat = Nothing
End Sub

Private Sub GetTotalPagar(grid As MSHFlexGrid, lbl As Label)
Dim i As Integer, nFila As Integer
Dim nMonto As Double, nImporte As Double
nMonto = 0
grid.Col = 3
nFila = grid.Row
For i = 1 To grid.Rows - 1
    grid.Row = i
    If grid <> "" Then
        nImporte = CDbl(grid)
    Else
        nImporte = 0
    End If
    nMonto = nMonto + nImporte
    grid = Format$(nImporte, "#,##0.00")
Next i
lbl = Format$(nMonto, "#,##0.00")
grid.Row = nFila
End Sub

Private Sub BuscaAlumno(ByVal sMatricula As String)
Dim rsUNT As ADODB.Recordset
Dim clsServ As NCapServicios
Set clsServ = New NCapServicios
Set rsUNT = clsServ.GetUNTAlumno(sMatricula)
Set clsServ = Nothing
If rsUNT.EOF And rsUNT.BOF Then
    MsgBox "Alumno no encontrado", vbInformation, "Aviso"
    mskCodAlumno.SetFocus
Else
    lblNombre = Trim(rsUNT("cNombre"))
    lblEscuela = Trim(rsUNT("cDescripcion"))
    lblCodEsc = Trim(rsUNT("cEscuelaCod"))
    lblEscOtros = lblCodEsc
    sCurr = rsUNT("cCurriculo")
    fraConcepto.Enabled = True
    optMatricula(0).SetFocus
End If
rsUNT.Close
Set rsUNT = Nothing
End Sub

Private Sub EditKeyCode(grid As MSHFlexGrid, Edit As TextBox, KeyCode As Integer, Shift As Integer, lbl As Label)
Select Case KeyCode
    Case 27 'ESC
        Edit.Visible = False
        grid.SetFocus
        GetTotalPagar grid, lbl
    Case 38 'Arriba
        grid.SetFocus
        DoEvents
        If grid.Row > grid.FixedRows Then
            grid.Row = grid.Row - 1
        End If
        GetTotalPagar grid, lbl
    Case 40, 13 'Abajo
        grid.SetFocus
        DoEvents
        If grid.Row < grid.Rows - 1 Then
            grid.Row = grid.Row + 1
        Else
            cmdGrabar.SetFocus
        End If
        GetTotalPagar grid, lbl
End Select

End Sub

Private Sub GridEdit(grid As MSHFlexGrid, Edit As TextBox, KeyAscii As Integer)
Select Case KeyAscii
    Case 0 To 32
        Edit = grid
        Edit.SelStart = Len(Edit.Text)
    Case Else
        Edit = Chr(KeyAscii)
        Edit.SelStart = 1
End Select
Edit.Move grid.Left + grid.CellLeft, grid.Top + grid.CellTop, grid.CellWidth - 8, grid.CellHeight - 8
Edit.Visible = True
Edit.SetFocus
End Sub

Private Sub chkSegProf_Click()
s2daProf = IIf(chkSegProf.value = 1, "S", "N")
If grdConcepto.Rows > 2 Then
    ActualizaMatricula
    GetTotalPagar grdConcepto, lblTotal
End If
End Sub

Private Sub chkSegProf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    mskNumRecibo.SetFocus
End If
End Sub

Private Sub cmdBuscar_Click()
frmCapServBuscaCliUNT.Show 1
If Len(Trim(lblNombre)) > 0 Then
    optMatricula(0).SetFocus
Else
    mskCodAlumno.SetFocus
End If
End Sub

Private Sub cmdCancelar_Click()
ClearScreen
mskCodAlumno.SetFocus
End Sub

Private Sub cmdGrabar_Click()
Dim sOrden As String * 8, sConcepto As String * 2
Dim sAlumno As String * 9
Dim sEstado As String * 1, sMonto As String
Dim sFecha As String, sFecOrden As String
Dim sOrdenOtros As String * 8
Dim sMonCuoFac As String, sMonCarnet As String
Dim clsMov As NContFunciones
Dim clsServ As NCapServicios
Dim sMovNro As String
Dim rsOrden As Recordset
sOrden = Trim(lblCodEsc) & Trim(mskNumRecibo)
sOrdenOtros = Trim(lblEscOtros) & Trim(mskOrdenOtros)
sAlumno = Left(mskCodAlumno, 7) & Right(mskCodAlumno, 2)

'Validaciones basicas
If (Trim(mskOrdenOtros) = "" Or Val(Trim(mskOrdenOtros)) = 0) And (Trim(mskNumRecibo) = "" Or Val(Trim(mskNumRecibo)) = 0) Then
    MsgBox "Número de Orden No Válido", vbInformation, "Aviso"
    mskOrdenOtros.SetFocus
    Exit Sub
End If

If Val(lblTotal) <= 0 And Val(lblTotOtros) <= 0 Then
    MsgBox "Monto a pagar debe ser mayor que cero", vbInformation, "Aviso"
    Exit Sub
End If

If sCuenta <> "" Then
    sFecha = FechaHora(gdFecSis)
    If MsgBox("¿Desea grabar la operación?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Set clsMov = New NContFunciones
        Set clsServ = New NCapServicios
        If CDbl(lblTotal) > 0 Then
            sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            Set rsOrden = GetRecordsetGrid(grdConcepto)
            sMonCuoFac = Trim(GetConceptoGrid(gUNTCuotaFacultad, grdConcepto))
            sMonCarnet = Trim(GetConceptoGrid(gUNTCarnetUniv, grdConcepto))
            clsServ.ActualizaOrdenPagoUNT sOrden, rsOrden, CDbl(lblTotal), sAlumno, sMovNro, sCuenta, sCuentaDonacion, "M", CDbl(sMonCuoFac), gsNomAge, sLpt
        Else
            sOrden = ""
        End If
        If CDbl(lblTotOtros) > 0 Then
            sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            Set rsOrden = GetRecordsetGrid(grdOtros)
            clsServ.ActualizaOrdenPagoUNT sOrdenOtros, rsOrden, CDbl(lblTotOtros), sAlumno, sMovNro, sCuenta, sCuentaDonacion, "E", , gsNomAge, sLpt
        Else
            sOrdenOtros = ""
        End If
        Set clsMov = Nothing
        On Error GoTo ErrImp
        Do
            clsServ.ImpBolUNT sAlumno, sOrden, lblNombre, lblTotal, lblEscuela, sFecha, sFecha, lblTotOtros, sOrdenOtros, sMonCuoFac, sMonCarnet, gsNomAge, gsCodUser, sLpt
            clsServ.ImpBolUNT sAlumno, sOrden, lblNombre, lblTotal, lblEscuela, sFecha, sFecha, lblTotOtros, sOrdenOtros, sMonCuoFac, sMonCarnet, gsNomAge, gsCodUser, sLpt
        Loop Until MsgBox("Desea Reimprimir???", vbQuestion + vbYesNo, "Aviso") = vbNo
        Set clsServ = Nothing
    End If
Else
    MsgBox "No se encontraron cuentas para el abono. Consulte con el Area de Sistemas", vbExclamation, "Error"
    Exit Sub
End If
ClearScreen
mskCodAlumno.SetFocus
Exit Sub
ErrGraba:
    MsgBox Err.Description, vbCritical, "Error"
    Exit Sub
ErrImp:
    MsgBox "Error de Impresion de boletas. Verifique la impresora, extorne la operacion y vuelva a hacer la operacion.", vbInformation, "Error"
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim rsPerm As Recordset
Dim clsServ As NCapServicios
Set clsServ = New NCapServicios
Set rsPerm = New Recordset
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
rsPerm.CursorLocation = adUseClient
Set rsPerm = clsServ.GetServConvCuentas(, gCapConvUNT)
Set clsServ = Nothing
If Not (rsPerm.EOF And rsPerm.BOF) Then
    Do While Not rsPerm.EOF
        If rsPerm("nTpoCuenta") = gCapConvTpoCtaPension Then
            sCuenta = rsPerm("cCtaCod")
        ElseIf rsPerm("nTpoCuenta") = gCapConvTpoCtaMora Then
            sCuentaDonacion = rsPerm("cCtaCod")
        End If
        rsPerm.MoveNext
    Loop
Else
    sCuenta = ""
End If
rsPerm.Close
Set rsPerm = Nothing
If sCuenta = "" Or sCuentaDonacion = "" Then
    fraConcepto.Enabled = False
    fraEstudiante.Enabled = False
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    cmdSalir.Enabled = True
Else
    ClearScreen
End If
txtMonto = ""
txtMonto.Visible = False
txtMonOtros = ""
txtMonOtros.Visible = False
End Sub

Private Sub grdConcepto_DblClick()
If grdConcepto.Col = 3 Then
    GridEdit grdConcepto, txtMonto, 32
End If
End Sub

Private Sub grdConcepto_GotFocus()
If Not txtMonto.Visible Then Exit Sub
If grdConcepto.Col = 3 Then
    grdConcepto = txtMonto
End If
txtMonto.Visible = False
End Sub

Private Sub grdConcepto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If grdConcepto.Row >= grdConcepto.Rows - 1 Then
        cmdGrabar.SetFocus
    Else
        grdConcepto.Row = grdConcepto.Row + 1
    End If
    Exit Sub
Else
    If grdConcepto.Col = 3 Then
        GridEdit grdConcepto, txtMonto, KeyAscii
    End If
End If
End Sub

Private Sub grdConcepto_LeaveCell()
If Not txtMonto.Visible Then Exit Sub
If grdConcepto.Col = 3 Then
    grdConcepto = txtMonto
End If
txtMonto.Visible = False
End Sub
Private Sub grdOtros_DblClick()
If grdOtros.Col = 3 Then
    GridEdit grdOtros, txtMonOtros, 32
End If
End Sub

Private Sub grdOtros_GotFocus()
If Not txtMonOtros.Visible Then Exit Sub
If grdOtros.Col = 3 Then
    grdOtros = txtMonOtros
End If
txtMonOtros.Visible = False
End Sub

Private Sub grdOtros_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If grdOtros.Row >= grdOtros.Rows - 1 Then
        cmdGrabar.SetFocus
    Else
        grdOtros.Row = grdOtros.Row + 1
    End If
    Exit Sub
Else
    If grdOtros.Col = 3 Then
        GridEdit grdOtros, txtMonOtros, KeyAscii
    End If
End If
End Sub

Private Sub grdOtros_LeaveCell()
If Not txtMonOtros.Visible Then Exit Sub
If grdOtros.Col = 3 Then
    grdOtros = txtMonOtros
End If
txtMonOtros.Visible = False
End Sub


Private Sub mskCodAlumno_GotFocus()
mskCodAlumno.SelStart = 0
mskCodAlumno.SelLength = Len(mskCodAlumno.Text)
End Sub

Private Sub mskCodAlumno_KeyPress(KeyAscii As Integer)
Dim sCod As String
If KeyAscii = 13 Then
    sCod = Mid(mskCodAlumno, 1, 7) & Right(mskCodAlumno, 2)
    If sCod = "" Then
        cmdBuscar.SetFocus
    Else
        BuscaAlumno sCod
    End If
End If
End Sub

Private Sub mskNumRecibo_GotFocus()
mskNumRecibo.SelStart = 0
mskNumRecibo.SelLength = Len(mskNumRecibo.Text)
End Sub

Private Sub mskNumRecibo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(mskNumRecibo) = "" Then
        MsgBox "Número de Orden No Válido", vbInformation, "Aviso"
        mskNumRecibo.SetFocus
        Exit Sub
    End If
    mskNumRecibo = FillNum(Trim(mskNumRecibo), 6, "0")
    Dim sAno As String * 2
    Dim rsUNT As ADODB.Recordset
    Dim clsServ As NCapServicios
    
    sAno = Right(mskCodAlumno.Text, 2)
    Set rsUNT = New ADODB.Recordset
    rsUNT.CursorLocation = adUseClient
    sTipoMat = IIf(optMatricula(0).value, "R", "E")
    s2daProf = IIf(chkSegProf.value = 1, "S", "N")
    
    If sAno = "01" Then
        sIngresante = IIf(Left(lblCodEsc, 1) = "J", "IM", "IT")
    Else
        sIngresante = "NT"
    End If
    
    Set clsServ = New NCapServicios
    Set rsUNT = clsServ.GetUNTCurriculoConcepto(sCurr, sTipoMat, sIngresante, s2daProf)
    Set grdConcepto.Recordset = rsUNT
    GetTotalPagar grdConcepto, lblTotal
    fraConcepto.Enabled = True
    fraEstudiante.Enabled = False
    grdConcepto.SetFocus
    rsUNT.Close
    Set rsUNT = Nothing
    'GetTotalPagar grdConcepto, lblTotal
    cmdGrabar.Enabled = True
End If
End Sub

Private Sub ClearScreen()
sCurr = ""
s2daProf = ""
sTipoMat = ""
sIngresante = ""
lblEscuela = ""
lblNombre = ""
mskCodAlumno = "_______-__"
mskNumRecibo = "      "
mskCodAlumno.Mask = "CC#####-##"
mskNumRecibo.Mask = "######"
mskOrdenOtros = "      "
mskOrdenOtros.Mask = "######"
lblCodEsc = ""
lblEscOtros = ""
optMatricula(0).value = True
SetupGrid grdConcepto
SetupGrid grdOtros
fraConcepto.Enabled = False
fraEstudiante.Enabled = True
lblTotal = "0.00"
lblTotOtros = "0.00"
tabConceptos.Tab = 0
cmdGrabar.Enabled = False
End Sub

Private Sub mskOrdenOtros_GotFocus()
mskOrdenOtros.SelStart = 0
mskOrdenOtros.SelLength = Len(mskOrdenOtros.Text)
End Sub

Private Sub mskOrdenOtros_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(mskOrdenOtros) = "" Then
        MsgBox "Número de Orden No Válido", vbInformation, "Aviso"
        mskOrdenOtros.SetFocus
        Exit Sub
    End If
    mskOrdenOtros = FillNum(Trim(mskOrdenOtros), 6, "0")
    Dim rsUNT As ADODB.Recordset
    Dim clsServ As NCapServicios
    
    Set rsUNT = New ADODB.Recordset
    rsUNT.CursorLocation = adUseClient
    Set clsServ = New NCapServicios
    Set rsUNT = clsServ.GetUNTConceptosOtros()
    Set grdOtros.Recordset = rsUNT
    grdOtros.SetFocus
    rsUNT.Close
    Set rsUNT = Nothing
    GetTotalPagar grdOtros, lblTotOtros
    cmdGrabar.Enabled = True
End If
End Sub

Private Sub optMatricula_Click(Index As Integer)
Select Case Index
    Case 0
        sTipoMat = "R"
    Case 1
        sTipoMat = "E"
End Select
ActualizaMatricula
End Sub

Private Sub optMatricula_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    chkSegProf.SetFocus
End If
End Sub

Private Sub tabConceptos_Click(PreviousTab As Integer)
Select Case PreviousTab
    Case 0
        If mskOrdenOtros.Enabled Then mskOrdenOtros.SetFocus
    Case 1
        If mskNumRecibo.Enabled Then mskNumRecibo.SetFocus
End Select
End Sub

Private Sub txtMonOtros_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCode grdOtros, txtMonOtros, KeyCode, Shift, lblTotOtros
End Sub

Private Sub txtMonOtros_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMonOtros, KeyAscii)
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
EditKeyCode grdConcepto, txtMonto, KeyCode, Shift, lblTotal
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMonto, KeyAscii)
End Sub

