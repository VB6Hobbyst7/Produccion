VERSION 5.00
Begin VB.Form FrmMantTipoCambioNuevo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento Tipo Cambio Especial"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   10005
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraNuevo 
      Height          =   645
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   9780
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   360
         Left            =   2520
         TabIndex        =   8
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   8400
         TabIndex        =   7
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   360
         Left            =   135
         TabIndex        =   6
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   360
         Left            =   1320
         TabIndex        =   5
         Top             =   180
         Width           =   1155
      End
   End
   Begin VB.Frame FraNuevoTC 
      Caption         =   "Nuevo Tipo de Cambio Especial"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3135
      Left            =   5040
      TabIndex        =   2
      Top             =   360
      Width           =   4815
      Begin SICMACT.FlexEdit fgTipoCambioNuevo 
         Height          =   2640
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4500
         _extentx        =   7938
         _extenty        =   4657
         cols0           =   5
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "-Venta-Compra--"
         encabezadosanchos=   "2000-1200-1200-0-0"
         font            =   "FrmMantTipoCambioNuevo.frx":0000
         font            =   "FrmMantTipoCambioNuevo.frx":0028
         font            =   "FrmMantTipoCambioNuevo.frx":0050
         font            =   "FrmMantTipoCambioNuevo.frx":0078
         font            =   "FrmMantTipoCambioNuevo.frx":00A0
         fontfixed       =   "FrmMantTipoCambioNuevo.frx":00C8
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-1-2-X-X"
         textstylefixed  =   4
         listacontroles  =   "0-0-0-0-0"
         encabezadosalineacion=   "C-R-R-C-C"
         formatosedit    =   "0-2-2-0-0"
         cantdecimales   =   4
         lbeditarflex    =   -1  'True
         lbbuscaduplicadotext=   -1  'True
         colwidth0       =   1995
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
   End
   Begin VB.Frame FraTC 
      Caption         =   "Tipo de Cambio Anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4815
      Begin SICMACT.FlexEdit fgTipoCambio 
         Height          =   2640
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4500
         _extentx        =   7938
         _extenty        =   4657
         cols0           =   5
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "-Venta-Compra--"
         encabezadosanchos=   "2000-1200-1200-0-0"
         font            =   "FrmMantTipoCambioNuevo.frx":00EE
         font            =   "FrmMantTipoCambioNuevo.frx":0116
         font            =   "FrmMantTipoCambioNuevo.frx":013E
         font            =   "FrmMantTipoCambioNuevo.frx":0166
         font            =   "FrmMantTipoCambioNuevo.frx":018E
         fontfixed       =   "FrmMantTipoCambioNuevo.frx":01B6
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-1-2-X-X"
         textstylefixed  =   4
         listacontroles  =   "0-0-0-0-0"
         encabezadosalineacion=   "C-R-R-C-C"
         formatosedit    =   "0-2-2-0-0"
         cantdecimales   =   4
         lbbuscaduplicadotext=   -1  'True
         colwidth0       =   1995
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "FrmMantTipoCambioNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbNuevo As Boolean
'*** PEAC 20080819
Dim i As Integer
Dim j As Integer
'*** FIN PEAC


Private Sub cmdCancelar_Click()
Dim rs As ADODB.Recordset
Dim ObjTc As COMDConstSistema.DCOMTipoCambioEsp

Set ObjTc = New COMDConstSistema.DCOMTipoCambioEsp
Set rs = New ADODB.Recordset
    
    Set rs = ObjTc.GetEstructuraTC
    fgTipoCambioNuevo.Clear
    fgTipoCambioNuevo.FormaCabecera
    fgTipoCambioNuevo.Rows = 2
    Dim ban As Boolean
    ban = False
    If Not rs.EOF And Not rs.BOF Then
        While Not rs.EOF
            If ban Then
                fgTipoCambioNuevo.Rows = fgTipoCambioNuevo.Rows + 1
            End If
            ban = True
            fgTipoCambioNuevo.TextMatrix(fgTipoCambioNuevo.Rows - 1, 0) = "Hasta " & rs!nHasta
            fgTipoCambioNuevo.TextMatrix(fgTipoCambioNuevo.Rows - 1, 1) = ""
            fgTipoCambioNuevo.TextMatrix(fgTipoCambioNuevo.Rows - 1, 2) = ""
            fgTipoCambioNuevo.TextMatrix(fgTipoCambioNuevo.Rows - 1, 3) = rs!nIdRango
            fgTipoCambioNuevo.TextMatrix(fgTipoCambioNuevo.Rows - 1, 4) = rs!nIdRangoDet
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set rs = Nothing
    Me.FraNuevoTC.Enabled = False
    Me.cmdCancelar.Enabled = False
    Me.cmdNuevo.Enabled = True
    cmdGrabar.Enabled = False
    
End Sub



Private Sub cmdGrabar_Click()
Dim i As Integer
Dim ObjTc As COMDConstSistema.DCOMTipoCambioEsp
'Dim ObjHora As COMConecta.DCOMConecta
Dim FechaReg As Date

'Valida Espacios
For i = 1 To Me.fgTipoCambioNuevo.Rows - 1
    If Trim(fgTipoCambioNuevo.TextMatrix(i, 1)) = "" Then
        MsgBox "Falta Ingresar la Venta", vbInformation, "AVISO"
        fgTipoCambioNuevo.Row = i
        fgTipoCambioNuevo.Col = 1
        Exit Sub
    End If
    
    If fgTipoCambioNuevo.TextMatrix(i, 1) = 0 Then
        MsgBox "Falta Ingresar la Venta", vbInformation, "AVISO"
        fgTipoCambioNuevo.Row = i
        fgTipoCambioNuevo.Col = 1
        Exit Sub
    End If
    
    If Trim(fgTipoCambioNuevo.TextMatrix(i, 2)) = "" Then
        MsgBox "Falta Ingresar la Venta", vbInformation, "AVISO"
        fgTipoCambioNuevo.Row = i
        fgTipoCambioNuevo.Col = 2
        Exit Sub
    End If
    
    If fgTipoCambioNuevo.TextMatrix(i, 2) = 0 Then
        MsgBox "Falta Ingresar la Venta", vbInformation, "AVISO"
        fgTipoCambioNuevo.Row = i
        fgTipoCambioNuevo.Col = 2
        Exit Sub
    End If
    
Next i

If MsgBox("Esta seguro de guardar el Nuevo Tipo de Cambio Especial", vbYesNo + vbInformation, "AVISO") = vbNo Then
    Exit Sub
End If

Set ObjTc = New COMDConstSistema.DCOMTipoCambioEsp
'Set ObjHora = New COMConecta.DCOMConecta
'ObjHora.AbreConexion
'FechaReg = gdFecSis & " " & ObjHora.GetHoraServer

FechaReg = gdFecSis

For i = 1 To Me.fgTipoCambioNuevo.Rows - 1
 Call ObjTc.InsertaTipoCambio(FechaReg, CCur(fgTipoCambioNuevo.TextMatrix(i, 1)), CCur(fgTipoCambioNuevo.TextMatrix(i, 2)), CInt(fgTipoCambioNuevo.TextMatrix(i, 3)), CInt(fgTipoCambioNuevo.TextMatrix(i, 4)))
Next i

Set ObjTc = Nothing

cmdCancelar_Click
CargaTC_Especial
'ObjHora.CierraConexion
'Set ObjHora = Nothing

End Sub

Private Sub cmdNuevo_Click()
Me.cmdCancelar.Enabled = True
Me.cmdNuevo.Enabled = False
cmdGrabar.Enabled = True
FraNuevoTC.Enabled = True
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub fgTipoCambioNuevo_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)

    Select Case pnCol
        '*** PEAC 20080819 *************************
        Case 1
            If Val(fgTipoCambioNuevo.TextMatrix(pnRow, 1)) > 0 Then
                For i = 1 To fgTipoCambioNuevo.Rows - 1
                    fgTipoCambioNuevo.TextMatrix(i, 1) = fgTipoCambioNuevo.TextMatrix(pnRow, 1)
                Next i
            End If
        '*** FIN PEAC ******************************
        Case 2
            If Val(fgTipoCambioNuevo.TextMatrix(pnRow, 1)) < Val(fgTipoCambioNuevo.TextMatrix(pnRow, 2)) Then
                MsgBox "Tipo de Cambio Compra debe ser menor que el tipo de cambio venta", vbInformation, "Aviso"
                    '*** PEAC 20080819
                    For i = 1 To fgTipoCambioNuevo.Rows - 1
                            fgTipoCambioNuevo.TextMatrix(i, 2) = ""
                    Next i
                    '*** FIN PEAC
                Cancel = False
            '*** PEAC 20080819
            Else
                If Val(fgTipoCambioNuevo.TextMatrix(pnRow, 2)) > 0 Then
                    For i = 1 To fgTipoCambioNuevo.Rows - 1
                        fgTipoCambioNuevo.TextMatrix(i, 2) = fgTipoCambioNuevo.TextMatrix(pnRow, 2)
                    Next i
                End If
            '*** FIN PEAC
            End If
        Case Else
    End Select
End Sub

Sub CargaTC_Especial()
Dim rs As ADODB.Recordset
Dim ObjTc As COMDConstSistema.DCOMTipoCambioEsp
Dim ban As Boolean
Set ObjTc = New COMDConstSistema.DCOMTipoCambioEsp
Set rs = New ADODB.Recordset


Me.Icon = LoadPicture(App.path & gsRutaIcono)
    
    Set rs = ObjTc.GetTiposCambios
    fgTipoCambio.Clear
    fgTipoCambio.FormaCabecera
    fgTipoCambio.Rows = 2
    If Not rs.EOF And Not rs.BOF Then
        ban = False
        While Not rs.EOF
            If ban Then
                fgTipoCambio.Rows = fgTipoCambio.Rows + 1
            End If
            ban = True
            fgTipoCambio.TextMatrix(fgTipoCambio.Rows - 1, 0) = "Hasta " & rs!nHasta
            fgTipoCambio.TextMatrix(fgTipoCambio.Rows - 1, 1) = rs!nVenta
            fgTipoCambio.TextMatrix(fgTipoCambio.Rows - 1, 2) = rs!nCompra
            fgTipoCambio.TextMatrix(fgTipoCambio.Rows - 1, 3) = rs!nIdRango
            fgTipoCambio.TextMatrix(fgTipoCambio.Rows - 1, 4) = rs!nIdRangoDet
            rs.MoveNext
        Wend
    End If
    
    Set rs = ObjTc.GetEstructuraTC
    fgTipoCambioNuevo.Clear
    fgTipoCambioNuevo.FormaCabecera
    fgTipoCambioNuevo.Rows = 2
    
    ban = False
    If Not rs.EOF And Not rs.BOF Then
        While Not rs.EOF
            If ban Then
                fgTipoCambioNuevo.Rows = fgTipoCambioNuevo.Rows + 1
            End If
            ban = True
            fgTipoCambioNuevo.TextMatrix(fgTipoCambioNuevo.Rows - 1, 0) = "Hasta " & rs!nHasta
            fgTipoCambioNuevo.TextMatrix(fgTipoCambioNuevo.Rows - 1, 1) = ""
            fgTipoCambioNuevo.TextMatrix(fgTipoCambioNuevo.Rows - 1, 2) = ""
            fgTipoCambioNuevo.TextMatrix(fgTipoCambioNuevo.Rows - 1, 3) = rs!nIdRango
            fgTipoCambioNuevo.TextMatrix(fgTipoCambioNuevo.Rows - 1, 4) = rs!nIdRangoDet
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set rs = Nothing

End Sub
Private Sub Form_Load()
CargaTC_Especial
End Sub

