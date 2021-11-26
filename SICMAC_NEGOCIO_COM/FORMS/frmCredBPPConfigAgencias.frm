VERSION 5.00
Begin VB.Form frmCredBPPConfigAgencias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Configuración de Agencias"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   Icon            =   "frmCredBPPConfigAgencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAgencia 
      Caption         =   "Configuracion de Zonas y Cant. de Comités"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   5160
         TabIndex        =   2
         Top             =   4920
         Width           =   1095
      End
      Begin SICMACT.FlexEdit feAgencias 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6120
         _extentx        =   10795
         _extenty        =   7858
         cols0           =   5
         highlight       =   1
         encabezadosnombres=   "#-Agencia-Zona-Cant. Comités-Aux"
         encabezadosanchos=   "0-2800-1500-1300-0"
         font            =   "frmCredBPPConfigAgencias.frx":030A
         font            =   "frmCredBPPConfigAgencias.frx":0332
         font            =   "frmCredBPPConfigAgencias.frx":035A
         font            =   "frmCredBPPConfigAgencias.frx":0382
         font            =   "frmCredBPPConfigAgencias.frx":03AA
         fontfixed       =   "frmCredBPPConfigAgencias.frx":03D2
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         tipobusqueda    =   3
         columnasaeditar =   "X-X-2-3-X"
         listacontroles  =   "0-0-3-0-0"
         encabezadosalineacion=   "C-L-L-R-C"
         formatosedit    =   "0-0-0-3-0"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbbuscaduplicadotext=   -1
         rowheight0      =   300
      End
   End
End
Attribute VB_Name = "frmCredBPPConfigAgencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private i As Integer
'
'Private Sub CargaAgencias()
'
'Dim oConst As COMDConstantes.DCOMConstantes
'Dim oAgencia As COMNCredito.NCOMBPPR
'Dim rsAgencias As ADODB.Recordset
'
'Set oAgencia = New COMNCredito.NCOMBPPR
'Set oConst = New COMDConstantes.DCOMConstantes
'
'Set rsAgencias = oAgencia.ObtenerAgenciaConfig
'
'LimpiaFlex feAgencias
'feAgencias.CargaCombo oConst.RecuperaConstantes(7073)
'
'If Not (rsAgencias.EOF And rsAgencias.BOF) Then
'    For i = 1 To rsAgencias.RecordCount
'        feAgencias.AdicionaFila
'        feAgencias.TextMatrix(i, 1) = Trim(rsAgencias!cAgeDescripcion)
'        feAgencias.TextMatrix(i, 2) = Trim(rsAgencias!Zona) & Space(75) & Trim(rsAgencias!CodZona)
'        feAgencias.TextMatrix(i, 3) = Trim(rsAgencias!nNComites)
'        feAgencias.TextMatrix(i, 4) = Trim(rsAgencias!cAgeCod)
'        rsAgencias.MoveNext
'
'    Next i
'End If
'Set rsAgencias = Nothing
'Set oAgencia = Nothing
'Set oConst = Nothing
'End Sub
'
'
'Private Sub cmdGuardar_Click()
'On Error GoTo Error
'If ValidaDatos Then
'    If MsgBox("Desea guardar los datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'        Dim oBPP As COMNCredito.NCOMBPPR
'        Set oBPP = New COMNCredito.NCOMBPPR
'
'        For i = 1 To feAgencias.Rows - 1
'            Call oBPP.OpeAgenciaConfig(feAgencias.TextMatrix(i, 4), CInt(Trim(Right(feAgencias.TextMatrix(i, 2), 4))), CLng(feAgencias.TextMatrix(i, 3)))
'        Next i
'
'        Call CargaAgencias
'        MsgBox "Datos Guardados Satisfactoriamente.", vbInformation, "Aviso"
'    End If
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub Form_Load()
'CargaAgencias
'End Sub
'
'Private Function ValidaDatos() As Boolean
'Dim oConst As COMDConstantes.DCOMConstantes
'Dim rsConst As ADODB.Recordset
'Dim nValor As Long
'
'ValidaDatos = True
'
'Set oConst = New COMDConstantes.DCOMConstantes
'Set rsConst = oConst.RecuperaConstantes(7066)
'nValor = 0
'If Not (rsConst.EOF And rsConst.BOF) Then
'    rsConst.MoveLast
'    nValor = CLng(rsConst!nConsValor)
'End If
'
'
'For i = 1 To feAgencias.Rows - 1
'    If Trim(feAgencias.TextMatrix(i, 2)) = "" Then
'        MsgBox "Seleccione la zona de la " & Trim(feAgencias.TextMatrix(i, 1)), vbInformation, "Aviso"
'        ValidaDatos = False
'        Exit Function
'    End If
'
'    If Trim(feAgencias.TextMatrix(i, 3)) = "" Then
'        MsgBox "Ingrese la Cantidad de Comités de la " & Trim(feAgencias.TextMatrix(i, 1)), vbInformation, "Aviso"
'        ValidaDatos = False
'        Exit Function
'    End If
'
'    If IsNumeric(Trim(feAgencias.TextMatrix(i, 3))) Then
'        If CDbl(Trim(feAgencias.TextMatrix(i, 3))) > 9999 Then
'            MsgBox "Valor No Valido en la Cantidad de Comités para la " & Trim(feAgencias.TextMatrix(i, 1)) & ".", vbInformation, "Aviso"
'            ValidaDatos = False
'            Exit Function
'        End If
'
'        If CLng(Trim(feAgencias.TextMatrix(i, 3))) <= 0 Then
'            MsgBox "Ingrese un valor Mayor a 0 en la Cantidad de Comités para la " & Trim(feAgencias.TextMatrix(i, 1)) & ".", vbInformation, "Aviso"
'            ValidaDatos = False
'            Exit Function
'        End If
'
'        If CLng(Trim(feAgencias.TextMatrix(i, 3))) > nValor Then
'            MsgBox "El numero máximo de Comités es " & nValor & ", favor de corregir el valor en la " & Trim(feAgencias.TextMatrix(i, 1)) & ".", vbInformation, "Aviso"
'            ValidaDatos = False
'            Exit Function
'        End If
'    End If
'Next i
'
'End Function
