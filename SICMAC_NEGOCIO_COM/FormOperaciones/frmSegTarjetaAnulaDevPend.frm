VERSION 5.00
Begin VB.Form frmSegTarjetaAnulaDevPend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nota de Débito - Aseguradora"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   Icon            =   "frmSegTarjetaAnulaDevPend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   " Listado de Pólizas anuladas luego de declarado en la trama "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7575
      Begin SICMACT.FlexEdit feLista 
         Height          =   2655
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   7335
         _extentx        =   12938
         _extenty        =   4683
         cols0           =   9
         fixedcols       =   2
         highlight       =   1
         allowuserresizing=   1
         rowsizingmode   =   1
         encabezadosnombres=   "Nro-Aux-#-Cliente-Fecha Afiliación-Fecha Anulación-Monto-cNumCert-cPersCod"
         encabezadosanchos=   "0-0-400-2900-1400-1400-900-0-0"
         font            =   "frmSegTarjetaAnulaDevPend.frx":030A
         font            =   "frmSegTarjetaAnulaDevPend.frx":0332
         font            =   "frmSegTarjetaAnulaDevPend.frx":035A
         font            =   "frmSegTarjetaAnulaDevPend.frx":0382
         font            =   "frmSegTarjetaAnulaDevPend.frx":03AA
         fontfixed       =   "frmSegTarjetaAnulaDevPend.frx":03D2
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-X-2-X-X-X-X-X-X"
         listacontroles  =   "0-0-4-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-C-L-C-C-R-C-C"
         formatosedit    =   "0-0-0-0-0-0-2-0-0"
         textarray0      =   "Nro"
         selectionmode   =   1
         lbeditarflex    =   -1  'True
         lbbuscaduplicadotext=   -1  'True
         rowheight0      =   300
         forecolorfixed  =   -2147483630
         cellforecolor   =   -2147483630
         cellbackcolor   =   -2147483633
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6480
      TabIndex        =   1
      Top             =   3360
      Width           =   1170
   End
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "Registrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5160
      TabIndex        =   0
      Top             =   3360
      Width           =   1170
   End
End
Attribute VB_Name = "frmSegTarjetaAnulaDevPend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmSegTarjetaAnulaDevPend
'** Descripción : Formulario para listar las devouciones pendientes de las anulaciones dentro
'**               de 15 dias pero en diferente mes a la afiliación segun Anexo 04-ERS068-2014
'** Creación : JUEZ, 20150510 09:00:00 AM
'**********************************************************************************************

Option Explicit

Dim oDSeg As COMDCaptaGenerales.DCOMSeguros
Dim rs As ADODB.Recordset

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
CargarLista
End Sub

Private Sub CargarLista()
Dim lnFila As Integer

Set oDSeg = New COMDCaptaGenerales.DCOMSeguros
    Set rs = oDSeg.RecuperaSegTarjetaAnulaDevPend(False, False)
Set oDSeg = Nothing

Call LimpiaFlex(feLista)

If rs.BOF And rs.EOF Then
    MsgBox "No existen registros", vbInformation, "Aviso"
    Exit Sub
End If

Do While Not rs.EOF
    feLista.AdicionaFila
    lnFila = feLista.row
    feLista.TextMatrix(lnFila, 1) = feLista.row
    feLista.TextMatrix(lnFila, 2) = "0"
    feLista.TextMatrix(lnFila, 3) = rs!cPersNombre
    feLista.TextMatrix(lnFila, 4) = rs!dFecAfiliacion
    feLista.TextMatrix(lnFila, 5) = rs!dFecAnula
    feLista.TextMatrix(lnFila, 6) = Format(rs!nMontoCom, "#,##0.00")
    feLista.TextMatrix(lnFila, 7) = rs!cNumCertificado
    feLista.TextMatrix(lnFila, 8) = rs!cPersCod
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
End Sub

Private Sub cmdRegistrar_Click()
Dim i As Integer
Dim bSelecciona As Boolean

bSelecciona = False

For i = 1 To feLista.Rows - 1
    If feLista.TextMatrix(i, 2) = "." Then
        Set oDSeg = New COMDCaptaGenerales.DCOMSeguros
            oDSeg.ActualizaSegTarjetaAnulaDevPendiente feLista.TextMatrix(i, 7), feLista.TextMatrix(i, 8), True, False
        Set oDSeg = Nothing
        bSelecciona = True
    End If
Next i

If Not bSelecciona Then
    MsgBox "Debe seleccionar al menos un registro", vbInformation, "Aviso"
Else
    MsgBox "Los registros fueron reconocidos para su devolución", vbInformation, "Aviso"
    CargarLista
End If
End Sub
