VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPPenalIncreMora 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Penalidad por Incremento de Mora"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   Icon            =   "frmCredBPPPenalIncreMora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Castigo por Incremento de Mora"
      TabPicture(0)   =   "frmCredBPPPenalIncreMora.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmbMeses"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmbAgencias"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdMostrar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdAgregar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdQuitar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdGuardar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCancelar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   5040
         TabIndex        =   3
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Left            =   4920
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox cmbAgencias 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox cmbMeses 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Filtro de Registro"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCredBPPPenalIncreMora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private i As Integer, j As Integer
'Private Sub cmdAgregar_Click()
'    flxPenalidad.AdicionaFila
'    flxPenalidad.SetFocus
'    SendKeys "{Enter}"
'End Sub
'
'Private Sub cmdCancelar_Click()
'    Unload Me
'End Sub
'
'Private Sub CmdGuardar_Click()
'On Error GoTo Error
'If ValidaDatos(1) Then
'    If MsgBox("Estas seguro de guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'    Dim lnDifer As Currency
'    Dim lnFilas As Integer
'    Dim lnDesde As Currency
'    Dim lnHasta As Currency
'    Dim lnMes As Integer
'    Dim lnAnio As Integer
'    Dim lsCodAge As String
'    Dim lnDscto As Integer
'
'        lnFilas = flxPenalidad.Rows - 1
'
'        lnMes = Right(cmbMeses.Text, 2)
'        lnAnio = uspAnio.valor
'        lsCodAge = Right(cmbAgencias.Text, 2)
'
'        For i = 1 To lnFilas
'            If i < lnFilas Then
'                lnDesde = flxPenalidad.TextMatrix(i + 1, 1)
'                lnHasta = flxPenalidad.TextMatrix(i, 2)
'
'                lnDifer = lnDesde - lnHasta
'
'                If lnDifer > 0.01 Then
'                    MsgBox "Sobrepaso la diferencia maxima entre : " & CStr(lnDesde) & " y " & CStr(lnHasta), vbExclamation, "Aviso"
'                    Exit Sub
'                End If
'            End If
'        Next
'
'        InsertaPenalidadXIncreMora lnMes, lnAnio, lsCodAge, lnFilas
'        MsgBox "Se registraron los datos correctamente", vbInformation, "Aviso"
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub cmdMostrar_Click()
'If ValidaDatos Then
'    Dim oBPP As COMNCredito.NCOMBPPR
'    Dim rsBPP As ADODB.Recordset
'    Dim lnMes As Integer
'    Dim lnAnio As Integer
'    Dim lsCodAge As String
'
'    lnMes = CInt(Right(cmbMeses.Text, 2))
'    lnAnio = CInt(uspAnio.valor)
'    lsCodAge = Trim(Right(cmbAgencias.Text, 2))
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'    Set rsBPP = oBPP.DevolverPenalidadXIncreMora(lnMes, lnAnio, lsCodAge)
'
'    LimpiaFlex flxPenalidad
'
'    If Not (rsBPP.EOF And rsBPP.BOF) Then
'        For i = 0 To rsBPP.RecordCount - 1
'            flxPenalidad.AdicionaFila
'            flxPenalidad.TextMatrix(i + 1, 1) = Format(rsBPP!nDesde, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxPenalidad.TextMatrix(i + 1, 2) = Format(rsBPP!nHasta, "###," & String(15, "#") & "#0." & String(2, "0"))
'            flxPenalidad.TextMatrix(i + 1, 3) = Format(rsBPP!nDescuento, "###," & String(15, "#") & "#0." & String(2, "0"))
'            rsBPP.MoveNext
'        Next i
'    End If
'End If
'End Sub
'
'Private Sub cmdQuitar_Click()
'    flxPenalidad.EliminaFila flxPenalidad.Rows - 1
'End Sub
'
'Private Sub Form_Load()
'    CargaCombos
'    uspAnio.valor = Year(gdFecSis)
'End Sub
'
'Private Sub CargaCombos()
'    CargaComboAgencias cmbAgencias
'    CargaComboMeses cmbMeses
'End Sub
'
'Private Sub InsertaPenalidadXIncreMora(ByVal pnMes As Integer, ByVal pnAnio As Integer, ByVal psCodAge As String, ByVal pnFilas As Integer)
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim lnDesde As Currency
'Dim lnHasta As Currency
'Dim lnDscto As Currency
'Dim lsFecha As String
'
'    lsFecha = Format(Now(), "YYYY-mm-DD HH:MM:SS")
'
'    Set oBPP = New COMNCredito.NCOMBPPR
'    Call oBPP.EliminaPenalidadXIncreMora(pnMes, pnAnio, psCodAge)
'
'    For i = 1 To pnFilas
'        lnDesde = flxPenalidad.TextMatrix(i, 1)
'        lnHasta = flxPenalidad.TextMatrix(i, 2)
'        lnDscto = flxPenalidad.TextMatrix(i, 3)
'
'        oBPP.InsertaPenalidadXIncreMora pnMes, pnAnio, psCodAge, lnDesde, lnHasta, lnDscto, gsCodUser, lsFecha
'
'    Next
'
'End Sub
'
'Private Function ValidaDatos(Optional pnTipo As Integer = 0) As Boolean
'
'   If Trim(cmbMeses.Text) = "" Then
'        MsgBox "Seleccione el Mes", vbInformation, "Aviso"
'        ValidaDatos = False
'        Exit Function
'    End If
'
'    If Trim(uspAnio.valor) = "" Or CDbl(uspAnio.valor) = 0 Then
'        MsgBox "Ingrese el Año", vbInformation, "Aviso"
'        ValidaDatos = False
'        Exit Function
'    End If
'
'    If Trim(cmbAgencias.Text) = "" Then
'        MsgBox "Seleccione la Agencia", vbInformation, "Aviso"
'        ValidaDatos = False
'        Exit Function
'    End If
'
'If pnTipo = 1 Then
'    For i = 0 To flxPenalidad.Rows - 2
'        For j = 1 To 3
'            If Trim(flxPenalidad.TextMatrix(i + 1, j)) = "" Then
'                MsgBox "Ingrese los datos Correctamente", vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'
'            If IsNumeric(Trim(flxPenalidad.TextMatrix(i + 1, j))) Then
'                If CDbl(Trim(flxPenalidad.TextMatrix(i + 1, j))) < 0 Or CDbl(Trim(flxPenalidad.TextMatrix(i + 1, j))) > 100 Then
'                    MsgBox "Ingrese los Valores Correctamente (0.00% - 100.00%) en la fila " & (i + 1), vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'            End If
'        Next j
'
'        If IsNumeric(Trim(flxPenalidad.TextMatrix(i + 1, 1))) And IsNumeric(Trim(flxPenalidad.TextMatrix(i + 1, 2))) Then
'            If CDbl(Trim(flxPenalidad.TextMatrix(i + 1, 1))) > CDbl(Trim(flxPenalidad.TextMatrix(i + 1, 2))) Then
'                MsgBox "El Valor de ''Desde'' no puede ser Mayor al Valor de ''Hasta'' en la fila " & (i + 1), vbInformation, "Aviso"
'                ValidaDatos = False
'                Exit Function
'            End If
'        End If
'    Next i
'End If
'
'ValidaDatos = True
'End Function
'
