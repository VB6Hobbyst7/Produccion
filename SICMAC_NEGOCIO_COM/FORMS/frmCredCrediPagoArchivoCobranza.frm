VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCredCrediPagoArchivoCobranza 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Archivo de Cobranza CrediPago"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraContenedor 
      Height          =   2685
      Index           =   3
      Left            =   7560
      TabIndex        =   9
      Top             =   0
      Width           =   2445
      Begin VB.ListBox List1 
         Height          =   2310
         Left            =   60
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   300
         Width           =   2295
      End
      Begin VB.OptionButton optSeleccionAg 
         Caption         =   "&Ninguno"
         Height          =   195
         Index           =   1
         Left            =   1500
         TabIndex        =   11
         Top             =   75
         Width           =   1020
      End
      Begin VB.OptionButton optSeleccionAg 
         Caption         =   "&Todos"
         Height          =   195
         Index           =   0
         Left            =   660
         TabIndex        =   10
         Top             =   60
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdRuta 
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
      Height          =   285
      Left            =   5250
      TabIndex        =   8
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox txtRuta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1020
      TabIndex        =   7
      Text            =   "D:\SICMACT"
      Top             =   5280
      Width           =   4245
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   375
      Left            =   7860
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar Archivo"
      Height          =   375
      Left            =   7860
      TabIndex        =   5
      Top             =   3660
      Width           =   1815
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7860
      TabIndex        =   4
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7515
      Begin VB.OptionButton optSeleccionCred 
         Caption         =   "&Todos"
         Height          =   195
         Index           =   0
         Left            =   5340
         TabIndex        =   2
         Top             =   60
         Width           =   765
      End
      Begin VB.OptionButton optSeleccionCred 
         Caption         =   "&Ninguno"
         Height          =   195
         Index           =   1
         Left            =   6180
         TabIndex        =   1
         Top             =   75
         Width           =   1020
      End
      Begin MSComctlLib.ListView lstCrediPago 
         Height          =   4515
         Left            =   60
         TabIndex        =   3
         Top             =   240
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   7964
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   285
      Left            =   60
      TabIndex        =   13
      Top             =   4860
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Caption         =   "&Guardar en"
      Height          =   285
      Left            =   60
      TabIndex        =   14
      Top             =   5280
      Width           =   825
   End
End
Attribute VB_Name = "frmCredCrediPagoArchivoCobranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fsCtaCrediPagoMN As String, fsCtaCrediPagoME As String
Dim ObjCredi As COMNCredito.NCOMCrediPago

Private Sub cmdGenerar_Click()

'Dim lsSQL As String
'Dim lrCob As New ADODB.Recordset
'Dim lsCabecera(2) As String
'Dim lsDetalle(2) As String
'Dim I As Integer, lnMoneda As Integer
'
'Dim lsCredito As String, lsCliente As String, lsFecVenc As String
'Dim lnMonto As Currency, lnMora As Currency
'Dim lsFecEmision As String
'
'Dim lnTotRegis As Integer, lnTotMonto As Currency

Dim ObjCredi As COMNCredito.NCOMCrediPago
Dim lsArcCob As String
Dim NumeroArchivo As Integer
Dim lnMoneda As Integer

Dim MatDatos As Variant
Dim MatCreditos() As Variant
Dim I As Integer

NumeroArchivo = FreeFile

'Set ObjCredi = New COMNCredito.NCOMCrediPago
' Verifica que no se haya generado el reporte
'If Not ObjCredi.VerificaReporteGenerado(gdFecSis) Then Exit Sub

lsArcCob = txtRuta.Text & "\CDPG.txt"

'lsFecEmision = Format(gdFecSis, "YYYYMMDD")

'For lnMoneda = 1 To 2 ' No olvidar cambiar a dos monedas
'    For I = 1 To Me.lstCrediPago.ListItems.Count
'        If lstCrediPago.ListItems(I).Checked = True And Mid(lstCrediPago.ListItems(I).Text, 6, 1) = lnMoneda Then
'            ' Carga las Variables
'            lsCredito = lstCrediPago.ListItems(I).Text
'            lsCliente = lstCrediPago.ListItems(I).SubItems(1)
'            lsFecVenc = Format(lstCrediPago.ListItems(I).SubItems(2), "YYYYMMDD")
'            lnMonto = CCur(lstCrediPago.ListItems(I).SubItems(3))
'            lnMora = CCur(lstCrediPago.ListItems(I).SubItems(4))
'            'lsMora = lstCrediPago.ListItems(i).SubItems(4)
'            lnTotRegis = lnTotRegis + 1
'            lnTotMonto = lnTotMonto + lstCrediPago.ListItems(I).SubItems(3)
'
'
'            'Inserta en la BD
'             Call ObjCredi.InsertaColocCrediPagoArcCobranza(gdFecSis, lsCredito, lstCrediPago.ListItems(I).SubItems(2), _
'             lstCrediPago.ListItems(I).SubItems(3), lstCrediPago.ListItems(I).SubItems(4), 0)
'
'            ' Cadena del Detalle
'            lsDetalle(lnMoneda) = lsDetalle(lnMoneda) & "DD" & "570" & Trim(Str(lnMoneda - 1)) & _
'                IIf(lnMoneda = 1, Mid(fsCtaCrediPagoMN, 5, 7), Mid(fsCtaCrediPagoME, 5, 7)) & _
'                FillNum(Trim(Str(lsCredito)), 14, "0") & ImpreFormat(Trim(lsCliente), 40, 0) & Space(30) & _
'                lsFecEmision & lsFecVenc & FillNum(ImpreFormat(EliminaPunto(lnMonto), 15, 0, False), 15, "0") & _
'                FillNum(ImpreFormat(EliminaPunto(lnMora), 15, 0, False), 15, "0") & "000000000" & Space(48) & Chr(13) & Chr(10)
'
'
'        End If
'    Next
'    ' Cadena de Cabecera
'    lsCabecera(lnMoneda) = "CC" & "570" & "0" & Mid(fsCtaCrediPagoMN, 5, 7) & "C" & ImpreFormat(Trim("CAJA MUNICIPAL DE TRUJILLO"), 40, 0) _
'                    & Format(gdFecSis, "YYYYMMDD") & FillNum(Trim(Str(lnTotRegis)), 9, "0") _
'                    & FillNum(ImpreFormat(EliminaPunto(lnTotMonto), 15, 0, False), 15, "0")
'Next

ReDim MatCreditos(Me.lstCrediPago.ListItems.Count, 5)
For lnMoneda = 1 To 2 ' No olvidar cambiar a dos monedas
    For I = 1 To Me.lstCrediPago.ListItems.Count
        If lstCrediPago.ListItems(I).Checked = True And CInt(Mid(lstCrediPago.ListItems(I).Text, 9, 1)) = lnMoneda Then
            MatCreditos(I - 1, 0) = lstCrediPago.ListItems(I).Text
            MatCreditos(I - 1, 1) = lstCrediPago.ListItems(I).SubItems(1)
            MatCreditos(I - 1, 2) = lstCrediPago.ListItems(I).SubItems(2)
            MatCreditos(I - 1, 3) = lstCrediPago.ListItems(I).SubItems(3)
            MatCreditos(I - 1, 4) = lstCrediPago.ListItems(I).SubItems(4)
        End If
    Next I
Next lnMoneda

'ReDim MatDatos(4)

Set ObjCredi = New COMNCredito.NCOMCrediPago
MatDatos = ObjCredi.GenerarArchivoCobranza(gdFecSis, MatCreditos, fsCtaCrediPagoMN, fsCtaCrediPagoME)
Set ObjCredi = Nothing

Open lsArcCob For Output As #NumeroArchivo
If LOF(1) > 0 Then
    If MsgBox("Existen Archivos de Cobranza Anteriores en el Directorio, Desea Remplazarlos ?" + Chr(13) + "Si Elige Si se Perderá la información contenida en los Archivos" + Chr(13) + "Si Elige No se adicionarán las Cuentas Seleccionadas", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
        Close #1
        Kill lsArcCob
        Open lsArcCob For Append As #1
    End If
End If

Print #NumeroArchivo, MatDatos(0) 'lsCabecera(1)
Print #NumeroArchivo, MatDatos(2) 'lsDetalle(1)
Print #NumeroArchivo, MatDatos(1) 'lsCabecera(2)
Print #NumeroArchivo, MatDatos(3) 'lsDetalle(2)

Close #NumeroArchivo   ' Cierra el archivo.
MsgBox "Se ha generado Archivo de Cobranza", vbInformation, "Aviso"
'Set ObjCredi = Nothing

End Sub

Private Sub cmdProcesar_Click()
'Dim lsSQL As String
Dim Flag As Boolean
Dim lrs As New ADODB.Recordset
Dim ObjCredi As COMNCredito.NCOMCrediPago
Dim Cuenta As String
Dim lnMax As Long
Dim lnX As Integer
Dim itmX As ListItem
Dim lnDiasAtraso As Integer
Dim I As Integer
Dim RecupAgencias As String
Dim ban As Boolean
'Dim nCred As COMNCredito.NCOMCredito
Dim MatMora() As Double

Set ObjCredi = New COMNCredito.NCOMCrediPago
'Set nCred = New COMNCredito.NCOMCredito

lstCrediPago.ListItems.Clear
ban = False
RecupAgencias = "("
For I = 0 To List1.ListCount - 1
    If List1.Selected(I) Then
        RecupAgencias = RecupAgencias & "'" & Left(List1.List(I), 2) & "',"
        ban = True
    End If
Next I
If Not ban Then Exit Sub
RecupAgencias = Mid(RecupAgencias, 1, Len(RecupAgencias) - 1) & ")"

'Set lrs = ObjCredi.ProcesaDatos(gdFecSis, RecupAgencias)
Set lrs = ObjCredi.ProcesarCrediPago(gdFecSis, RecupAgencias, MatMora)

If lrs.EOF And lrs.BOF Then
    'MsgBox "NO existes Creditos para cobranza en CrediPago", vbInformation, "Aviso"
    lnMax = 0
Else
    lnMax = lrs.RecordCount
    Me.PB.Max = lnMax + 1
    
    While Not lrs.EOF
        'LLenaLista
        Set itmX = lstCrediPago.ListItems.Add(, , lrs!cCtaCod)
            itmX.SubItems(1) = PstaNombre(lrs!cPersNombre)     'Nombre Cliente
            itmX.SubItems(2) = Format(lrs!dvenc, "dd/mm/yyyy")
            itmX.SubItems(3) = Format(lrs!Cuota, "#,##0.00")
                       'fgCalculaGastoACuota(lrs!cCodCta, fnConec, lrs!cNroCuo)
            '---no va ----> itmX.SubItems(4) = Format(fObtieneMontoCuotaPend(lrs!cCodCta, fnConec), "#,##0.00")
            
            'lnDiasAtraso = DateDiff("d", Format(lrs!dvenc, "yyyy/mm/dd"), Format(gdFecSis, "yyyy/mm/dd"))
            'itmX.SubItems(4) = Format(ObjCredi.ObtieneMoraProyectada(lrs!cCtaCod, lrs!nCuota, lnDiasAtraso, lrs!nNroCalen), "#,##0.00")
            itmX.SubItems(4) = Format(MatMora(lrs.Bookmark - 1), "#,##0.00")
            
            lrs.MoveNext
            PB.value = PB.value + 1
            PB.ToolTipText = "Registro : " + Str(PB.value)
    Wend
        lstCrediPago.SetFocus

End If
lrs.Close

'Set nCred = Nothing
Set lrs = Nothing
Set ObjCredi = Nothing
PB.value = 0
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()

    Dim rs As New ADODB.Recordset
    'Dim ObjAge As COMDConstantes.DCOMActualizaDatosArea
    'Set ObjAge = New COMDConstantes.DCOMActualizaDatosArea
    'Set ObjCredi = New COMNCredito.NCOMCrediPago
    
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    'CargaAgencias List1
    lstCrediPago.ColumnHeaders.Add , , "Credito", 1500
    lstCrediPago.ColumnHeaders.Add , , "Cliente ", 2400, lvwColumnLeft
    lstCrediPago.ColumnHeaders.Add , , "Vencimiento", 1200, lvwColumnLeft
    lstCrediPago.ColumnHeaders.Add , , "Cuota", 1000, lvwColumnRight
    lstCrediPago.ColumnHeaders.Add , , "Mora", 1000, lvwColumnRight
    
    '*************** Ubica Conexion Agencia España
    'lsAgConsolida = AgenciaJudicial  ' Agencia España
    
    'If Right(Trim(lsAgConsolida), 2) = Right(Trim(gsCodAge), 2) Then
    '   Set fnConecCP = dbCmact
    'Else
    '   If AbreConeccion(Right(Trim(lsAgConsolida), 2), True, False) Then
    '      Set fnConecCP = dbCmactN
    '   Else
    '      MsgBox "No se Pudo Conectar con la Agencia " & Mid(lsAgConsolida, 1, 2)
    '   End If
    'End If
    
    'Set rs = ObjAge.GetAgencias()
    
    Set ObjCredi = New COMNCredito.NCOMCrediPago
    Call ObjCredi.CargarObjetosArchivoCobranza(rs, fsCtaCrediPagoMN, fsCtaCrediPagoME)
    Set ObjCredi = Nothing
    
    While Not rs.EOF
        Me.List1.AddItem Trim(rs!Codigo) & " " & Trim(rs!Descripcion)
        If rs!Codigo = gsCodAge Then
            List1.Selected(List1.ListCount - 1) = True
        End If
        rs.MoveNext
    Wend
    'fsCtaCrediPagoMN = ObjCredi.GetNroCta
    'fsCtaCrediPagoME = ObjCredi.GetNroCta(False)
    
    txtRuta.Text = App.path & "\Spooler"
    Set rs = Nothing
    'Set ObjAge = Nothing
End Sub

Private Sub optSeleccionAg_Click(Index As Integer)
Dim I As Integer
Select Case Index
   Case 0
        For I = 1 To Me.List1.ListCount
            Me.List1.Selected(I - 1) = True
        Next I
    Case 1
        For I = 1 To Me.List1.ListCount
             Me.List1.Selected(I - 1) = False
        Next I
End Select
End Sub

Private Sub optSeleccionCred_Click(Index As Integer)
Dim I As Integer
Select Case Index
   Case 0
        For I = 1 To Me.lstCrediPago.ListItems.Count
            Me.lstCrediPago.ListItems(I).Checked = True
        Next I
    Case 1
        For I = 1 To Me.lstCrediPago.ListItems.Count
             Me.lstCrediPago.ListItems(I).Checked = False
        Next I
End Select
End Sub
