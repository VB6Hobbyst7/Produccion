VERSION 5.00
Begin VB.Form frmCapTarifarioGroupBitacora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bitacora de Cambios"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   Icon            =   "frmCapCapTarifarioGroupBitacora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin SICMACT.FlexEdit flxComision 
      Height          =   3135
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   6255
      _extentx        =   11033
      _extenty        =   5530
      cols0           =   5
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   "nro-Agencia-Tasa-Comis.-Fecha"
      encabezadosanchos=   "0-2000-1000-1000-1600"
      font            =   "frmCapCapTarifarioGroupBitacora.frx":030A
      font            =   "frmCapCapTarifarioGroupBitacora.frx":0336
      font            =   "frmCapCapTarifarioGroupBitacora.frx":0362
      font            =   "frmCapCapTarifarioGroupBitacora.frx":038E
      font            =   "frmCapCapTarifarioGroupBitacora.frx":03BA
      fontfixed       =   "frmCapCapTarifarioGroupBitacora.frx":03E6
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1  'True
      columnasaeditar =   "X-X-X-X-X"
      encabezadosalineacion=   "L-L-C-C-C"
      textarray0      =   "nro"
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "Mostrar"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox cmbAgencia 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   150
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Agencia:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "frmCapTarifarioGroupBitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMostrar_Click()
    CargarControles
End Sub

Private Sub Form_Load()
Dim lsCodAge As String

    CargaAgencias
    
End Sub

Public Sub CargaAgencias()
    Dim loCargaAg As COMDColocPig.DCOMColPFunciones
    Dim lrAgenc As ADODB.Recordset
    Set loCargaAg = New COMDColocPig.DCOMColPFunciones
        Set lrAgenc = loCargaAg.dObtieneAgencias(True)
    Set loCargaAg = Nothing
    If lrAgenc Is Nothing Then
        MsgBox " No se encuentran las Agencias ", vbInformation, " Aviso "
    Else
        Me.cmbAgencia.Clear
        cmbAgencia.AddItem "TODOS" & Space(50) & "00"
        With lrAgenc
            Do While Not .EOF
                cmbAgencia.AddItem Trim(!cAgeDescripcion) & Space(50) & !cAgeCod
                .MoveNext
            Loop
        End With
    End If
End Sub
Private Sub CargarControles()

Dim rsTmp As ADODB.Recordset
Set oCon = New COMNCaptaGenerales.NCOMCaptaDefinicion
lsCodAge = Right(cmbAgencia.Text, 2)
'Set rsInicial = Nothing

'cargando los grupos
Set rsTmp = oCon.ObtenerAgenciaGrupoTasaComisionXAgencia(lsCodAge)

CargarGrid rsTmp
Set oCon = Nothing

End Sub
Private Sub CargarGrid(ByVal prsGrupos As ADODB.Recordset)
Dim i As Integer

'Cargando el los combos de los grupos
'Dim rsGrupos As New ADODB.Recordset
'rsGrupos.Fields.Append "Grupo", adVarChar, 10
'rsGrupos.CursorLocation = adUseClient
'rsGrupos.CursorType = adOpenStatic
'rsGrupos.Open
'For i = 65 To 90
'    rsGrupos.AddNew
'    rsGrupos.Fields("Grupo") = Chr(i)
'Next i
'rsGrupos.MoveFirst
'grdGrupos.CargaCombo rsGrupos

'Cargando los valores de las agencias, grupos, tasas
i = 1
If Not prsGrupos Is Nothing Then
    If Not prsGrupos.EOF And Not prsGrupos.BOF Then
        If prsGrupos.RecordCount > 0 Then
            Do While Not prsGrupos.EOF And Not prsGrupos.BOF
                Set rsInicial = prsGrupos.Clone
                flxComision.AdicionaFila
                flxComision.TextMatrix(i, 0) = i
                'grdGrupos.TextMatrix(i, 1) = prsGrupos!nIdGrupo
                flxComision.TextMatrix(i, 1) = prsGrupos!cAgeDescripcion
                flxComision.TextMatrix(i, 2) = prsGrupos!cGrupoTasa
                flxComision.TextMatrix(i, 3) = prsGrupos!cGrupoComi
                flxComision.TextMatrix(i, 4) = prsGrupos!cFecha
                'grdGrupos.TextMatrix(i, 5) = prsGrupos!cAgeCod
                prsGrupos.MoveNext
                i = i + 1
            Loop
        End If
    End If
End If
End Sub

