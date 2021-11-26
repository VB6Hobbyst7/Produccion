VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCapAbonoPersParam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametros de Deposito por Agencia"
   ClientHeight    =   4935
   ClientLeft      =   3885
   ClientTop       =   3660
   ClientWidth     =   7275
   Icon            =   "frmCapAbonoPersParam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7275
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   7095
      Begin VB.CommandButton cdmSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   5640
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Montos Minimos por Agencia"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAgencia 
         Height          =   3615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6376
         _Version        =   393216
         Cols            =   5
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
End
Attribute VB_Name = "frmCapAbonoPersParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim objCap As COMNCaptaGenerales.NCOMCaptaMovimiento
'Dim matris() As String
'
'Private Sub cdmSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdCancelar_Click()
'    grdAgencia.Enabled = False
'    cmdGrabar.Enabled = False
'    cmdEditar.Enabled = True
'    grdAgencia_LeaveCell
'
'End Sub
'
'Private Sub cmdEditar_Click()
'    grdAgencia.Enabled = True
'    cmdGrabar.Enabled = True
'    cmdEditar.Enabled = False
'    grdAgencia.row = 1
'    grdAgencia.col = 3
'    grdAgencia.SetFocus
'    grdAgencia_EnterCell
'End Sub
'
'Private Sub CmdGrabar_Click()
'    Set objCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
'    Dim ClsMov As COMNContabilidad.NCOMContFunciones
'    Set ClsMov = New COMNContabilidad.NCOMContFunciones
'    Dim sMovNro As String
'
'
'    If MsgBox("Esta Seguro de Registrar los Parametros?", vbInformation + vbYesNo) = vbYes Then
'        sMovNro = ClsMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
'        Dim i As Integer
'        objCap.actualizaCapAbonoAgeParam
'        For i = 1 To grdAgencia.Rows - 1
'            objCap.insertarCapAbonoAgeParam grdAgencia.TextMatrix(i, 1), matris(i, 3), matris(i, 4), sMovNro
'        Next
'        MsgBox "Se registraron los Parametros con Exito!", vbExclamation
'        cmdGrabar.Enabled = False
'        cmdEditar.Enabled = True
'        grdAgencia.Enabled = False
'    End If
'End Sub
'
'Private Sub Form_Load()
'grdAgencia.Enabled = False
'ConfigGridAgencia
'CargarAgenciasParametro
'
'End Sub
'Private Sub CargarAgenciasParametro()
'    Dim rsAgencia As New ADODB.Recordset
'
'    Set objCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
'    Set rsAgencia.DataSource = objCap.getCapAbonoAgeParamListar
'    Dim i As Integer
'    If Not rsAgencia.EOF Or rsAgencia.BOF Then
'            grdAgencia.Rows = grdAgencia.Rows + rsAgencia.RecordCount - 1
'            ReDim matris(grdAgencia.Rows, grdAgencia.Cols)
'            For i = 0 To rsAgencia.RecordCount - 1
'                grdAgencia.TextMatrix(i + 1, 0) = i + 1
'                grdAgencia.TextMatrix(i + 1, 1) = rsAgencia!cAgeCod
'                grdAgencia.TextMatrix(i + 1, 2) = rsAgencia!cAgeDescripcion
'                grdAgencia.TextMatrix(i + 1, 3) = Format(rsAgencia!nMontoMin, "$ ##,##0.00")
'                matris(i + 1, 3) = rsAgencia!nMontoMin
'                grdAgencia.TextMatrix(i + 1, 4) = Format(rsAgencia!nMontoMax, "$ ##,##0.00")
'                matris(i + 1, 4) = rsAgencia!nMontoMax
'                rsAgencia.MoveNext
'            Next
'    End If
'
'End Sub
'Private Sub ConfigGridAgencia()
'    grdAgencia.Clear
'    grdAgencia.Rows = 2
'
'    With grdAgencia
'        .TextMatrix(0, 0) = "#"
'        .TextMatrix(0, 1) = "Cod"
'        .TextMatrix(0, 2) = "Agencia"
'        .TextMatrix(0, 3) = "Monto Min"
'        .TextMatrix(0, 4) = "Monto Max"
'
'        .ColWidth(0) = 300
'        .ColWidth(1) = 500
'        .ColWidth(2) = 3000
'        .ColWidth(3) = 1300
'        .ColWidth(4) = 1300
'
'        .ColAlignment(1) = flexAlignCenterCenter
'        .ColAlignment(3) = flexAlignRightCenter
'        .ColAlignment(4) = flexAlignRightCenter
'        .ColAlignmentFixed(1) = flexAlignCenterCenter
'        .ColAlignmentFixed(3) = flexAlignCenterCenter
'        .ColAlignmentFixed(4) = flexAlignCenterCenter
'
'
'    End With
'End Sub
'
'Private Sub grdAgencia_EnterCell()
'    With grdAgencia
'        If (.col > 2 And .col < (.Cols)) Then
'        .CellBackColor = &H80000018
'        .Tag = ""
'        End If
'    End With
'End Sub
'
'Private Sub grdAgencia_KeyDown(KeyCode As Integer, Shift As Integer)
'    With grdAgencia
'     Select Case KeyCode
'        Case 46 'si presiona tecla del
'            .Tag = grdAgencia
'
'            If (.col > 2 And .col < (.Cols)) Then
'                 matris(.row, .col) = "0"
'                 '.TextMatrix(.Row, .Col) = Format("0", "$##,##0.00")
'                 grdAgencia = Format("0", "$ ##,##0.00")
'            End If
'
'
'      End Select
'    End With
'End Sub
'
'Private Sub grdAgencia_KeyPress(KeyAscii As Integer)
'
'With grdAgencia
'    'si es enter
'    If KeyAscii = 13 Then
'            'si esta en las columnas 5 de edicion
'            If .col > 2 And .col < 4 And .TextMatrix(.row, .col) <> "" Then
'                    If .col = 3 Then
'                        If CInt(matris(.row, .col)) > 9999 Then
'                            .TextMatrix(.row, .col) = ""
'                            MsgBox "Solo esta permitido hasta $9999", vbInformation
'                            Exit Sub
'                        End If
'                    End If
'
'                    .CellBackColor = &H8000000E
'                    .row = .row
'                    .col = .col + 1
'                    .CellBackColor = &H80000018
'            'si esta en las columnas 5 de edicion
'            ElseIf .row <= .Rows - 1 And .TextMatrix(.row, .col) <> "" Then
'                    If .col = 4 Then
'                        If CLng(matris(.row, .col)) > 9999 Then
'                          MsgBox "Solo esta permitido hasta $9999", vbInformation
'                            Exit Sub
'                        ElseIf CLng(matris(.row, .col)) < CInt(matris(.row, .col - 1)) Then
'                          MsgBox "El Monto Maximo debe ser Mayor que el Monto Minimo", vbInformation
'                            Exit Sub
'                        End If
'                    End If
'                    .CellBackColor = &H8000000E
'
'                    If .row < .Rows - 1 Then 'si no esta en la ultima fila y columna
'                        .row = .row + 1
'                        .col = 3
'                        .CellBackColor = &H80000018
'                    Else
'                         Me.cmdGrabar.SetFocus
'                    End If
'
'            End If
'
'    ElseIf KeyAscii = 8 Then 'si es retroceso
'
'                If Len(.TextMatrix(.row, .col)) > 0 Then
'                        If Len(matris(.row, .col)) > 1 Then
'                            matris(.row, .col) = Mid(matris(.row, .col), 1, Len(matris(.row, .col)) - 1) 'retrocede 1 digito
'                        Else
'                            matris(.row, .col) = "0"
'                        End If
'                        grdAgencia = Format(matris(.row, .col), "$ ##,##0.00")
'                End If
'
'    ElseIf (.col > 2 And .col < (.Cols)) Then
'
'                If InStr("0123456789.", Chr(KeyAscii)) = 0 Then 'compara si no es un numero
'                        KeyAscii = 0
'                        .TextMatrix(.row, .col) = ""
'                Else
'
'                    matris(.row, .col) = matris(.row, .col) + Chr(KeyAscii)
'                    If CLng(matris(.row, .col)) < 10000 Then ' si es menor al monto del REU
'                        .TextMatrix(.row, .col) = Format(matris(.row, .col), "$ ##,##0.00")
'                        grdAgencia = .TextMatrix(.row, .col)
'                    Else
'                        .TextMatrix(.row, .col) = Format(matris(.row, .col), "$ ##,##0.00")
'                        MsgBox "Solo esta permitido ingresar hasta $9,999"
'                        matris(.row, .col) = Mid(matris(.row, .col), 1, Len(matris(.row, .col)) - 1)
'                        .TextMatrix(.row, .col) = Format(matris(.row, .col), "$ ##,##0.00")
'                    End If
'
'                End If
'
'    End If
'
'End With
'End Sub
'
'Private Sub grdAgencia_LeaveCell()
'    With grdAgencia
'        Dim s As String
'        If .col > 2 And .col < (.Cols) Then
'
'              If .col = 4 Then
'                If CLng(matris(.row, .col)) < CInt(matris(.row, .col - 1)) Then
'                   MsgBox "El Monto Maximo debe ser Mayor que el Monto Minimo", vbInformation
'                   matris(.row, .col) = "0"
'                   matris(.row, .col - 1) = "0"
'                   .TextMatrix(.row, .col - 1) = Format("0", "$ ##,##0.00")
'                   grdAgencia = Format("0", "$ ##,##0.00")
'                End If
'              End If
'              .CellBackColor = &H8000000E
'        End If
'
'    End With
'End Sub
'
