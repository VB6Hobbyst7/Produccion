VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRHformatoPDT 
   Caption         =   "Formato PDT"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12690
   Icon            =   "frmRHFormatoPDT.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6930
   ScaleWidth      =   12690
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir "
      Height          =   375
      Left            =   11400
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdexporta 
      Caption         =   "Formato PDT"
      Height          =   375
      Left            =   10200
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFPDT 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   11245
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Sicmact.TxtBuscar txtPlanillaIns 
      Height          =   300
      Left            =   855
      TabIndex        =   1
      Top             =   120
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   529
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      TipoBusqueda    =   2
   End
   Begin VB.Label lblPlanilla 
      Caption         =   "Planilla :"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   135
      Width           =   735
   End
   Begin VB.Label lblPlaniInstRes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2565
      TabIndex        =   2
      Top             =   120
      Width           =   6195
   End
End
Attribute VB_Name = "frmRHformatoPDT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oPla As DActualizaDatosConPlanilla
Dim rs As ADODB.Recordset
Private Sub cmdexporta_Click()
Dim sAno As String
Dim sMes As String
Dim lsDts As String
Dim X As Double
Dim sRuta As String
Dim npos As Integer
On Error GoTo ErrorValidaDato


If txtPlanillaIns.Text = "" Then
    MsgBox "Antes debe seleccionar un periodo de Planilla ", vbInformation, "Seleccione un Periodo"
    Exit Sub
End If

If MsgBox("Desea Generar el Formato PDT600 ?", vbQuestion + vbYesNo, "Generar FormatoPDT") = vbYes Then
    sAno = Left(Trim(txtPlanillaIns), 4)
    sMes = Mid(Trim(txtPlanillaIns), 5, 2)
    CommonDialog1.FileName = "C:\0600" + sAno + sMes + "20104888934.djt"
    CommonDialog1.Filter = " (*.djt;*.txt;*.*)|*djt;*.txt;*.*"
    CommonDialog1.DialogTitle = "Seleccione la Ruta "
    CommonDialog1.ShowSave
    
    npos = InStrRev(Trim(CommonDialog1.FileName), "\")
    sRuta = Left(Trim(CommonDialog1.FileName), npos)
    oPla.ActualizaRHPDT sAno, sMes, sRuta
    'dtsrun /Sserver_name /Uuser_nName /Ppassword /Npackage_name /Mpackage_password
    lsDts = "dtsrun /S01SRVSICMAC01 /Usa /Pcmacica /NDTSRHPDT /M"
    X = Shell(lsDts, vbMaximizedFocus)
    
End If
Exit Sub
ErrorValidaDato:
    If Err.Number = 32755 Then
            MsgBox "Usted Eligio Accion Cancelar", vbInformation, "Accion Cancelada"
    Else
            MsgBox Err.Number & " " & Err.Description, vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdSalir_Click()
Unload Me

End Sub

Private Sub Form_Load()
Set oPla = New DActualizaDatosConPlanilla

Set rs = New ADODB.Recordset
CargaPlanillasTpo "E01"

MSHFPDT.Cols = 12

MSHFPDT.ColAlignment(5) = 4
MSHFPDT.ColAlignment(6) = 7
MSHFPDT.ColAlignment(7) = 7
MSHFPDT.ColAlignment(8) = 7
MSHFPDT.ColAlignment(9) = 7
MSHFPDT.ColAlignment(10) = 7
MSHFPDT.ColAlignment(11) = 7
MSHFPDT.ColWidth(0) = 1300
MSHFPDT.ColWidth(1) = 800
MSHFPDT.ColWidth(2) = 2000
MSHFPDT.ColWidth(3) = 750
MSHFPDT.ColWidth(4) = 900

End Sub
Private Sub CargaPlanillasTpo(psPlaCodigo As String)
    
    'If TxtPlanillas.Text <> "" And Not fraPla.Enabled Then fraPla.Enabled = True
    Set rs = oPla.GetPlanillasTpo(psPlaCodigo)
    Me.txtPlanillaIns.rs = rs
    'IniFlex False
End Sub


Private Sub txtPlanillaIns_EmiteDatos()
Me.lblPlaniInstRes.Caption = txtPlanillaIns.psDescripcion
'Actualiza tabla de control rhpdt
'Muestra Formato
Dim nRemIES As Double
Dim nRemONP  As Double
Dim nRemESALUD As Double
Dim nRemFSA As Double
Dim nRem5taCat As Double
Dim nTrib5taCat As Double

Dim f As Integer
Dim c As Integer

Set rs = oPla.GetRHListaFormatoPDT(Left(Trim(txtPlanillaIns.Text), 6))
If rs.EOF = True Then
        MsgBox "No existen Datos para esta Planilla", vbInformation, "No existen Datos"
        Exit Sub
    Else
        Set MSHFPDT.Recordset = rs
        MSHFPDT.AddItem "Trabajadores:"
        MSHFPDT.TextMatrix(MSHFPDT.Rows - 1, 1) = Str(MSHFPDT.Rows - 1)
        
        For f = 1 To MSHFPDT.Rows - 2
            For c = 0 To MSHFPDT.Cols - 1
                Select Case c
                    Case Is < 6
                        MSHFPDT.CellBackColor = RGB(100, 200, 300)
                    Case 6  'nRemIES
                        nRemIES = nRemIES + IIf(MSHFPDT.TextMatrix(f, c) = "", 0, MSHFPDT.TextMatrix(f, c))
                    Case 7  'nRemONP
                        nRemONP = nRemONP + IIf(MSHFPDT.TextMatrix(f, c) = "", 0, MSHFPDT.TextMatrix(f, c))
                    Case 8  'nRemESALUD
                        nRemESALUD = nRemESALUD + IIf(MSHFPDT.TextMatrix(f, c) = "", 0, MSHFPDT.TextMatrix(f, c))
                    Case 9  'nRemFSA
                        nRemFSA = nRemFSA + IIf(MSHFPDT.TextMatrix(f, c) = "", 0, MSHFPDT.TextMatrix(f, c))
                    Case 10  'nRem5taCat
                        nRem5taCat = nRem5taCat + IIf(MSHFPDT.TextMatrix(f, c) = "", 0, MSHFPDT.TextMatrix(f, c))
                    Case 11 'nTrib5taCat
                        nTrib5taCat = nTrib5taCat + IIf(MSHFPDT.TextMatrix(f, c) = "", 0, MSHFPDT.TextMatrix(f, c))
                    
                End Select
            Next
        
        Next
        MSHFPDT.TextMatrix(MSHFPDT.Rows - 1, 6) = nRemIES
        MSHFPDT.TextMatrix(MSHFPDT.Rows - 1, 7) = nRemONP
        MSHFPDT.TextMatrix(MSHFPDT.Rows - 1, 8) = nRemESALUD
        MSHFPDT.TextMatrix(MSHFPDT.Rows - 1, 9) = nRemFSA
        MSHFPDT.TextMatrix(MSHFPDT.Rows - 1, 10) = nRem5taCat
        MSHFPDT.TextMatrix(MSHFPDT.Rows - 1, 11) = nTrib5taCat
        
        For i = 0 To MSHFPDT.Cols - 1
        MSHFPDT.Col = i
        MSHFPDT.Row = MSHFPDT.Rows - 1
        MSHFPDT.CellBackColor = RGB(100, 200, 300)
        Next
        
        
        
        
End If


End Sub

