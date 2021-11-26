VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmValorizacionDiariaAnexo15ANew 
   Caption         =   "Valorización Diaria"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLTesoroTotal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   320
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "0.00"
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox txtCDBCRPTotal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   320
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "0.00"
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdBuscarValorizacion 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   2880
      TabIndex        =   36
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   10490
      TabIndex        =   23
      Top             =   3480
      Width           =   1000
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   10490
      TabIndex        =   22
      Top             =   2760
      Width           =   1000
   End
   Begin VB.Frame fraObligacionBN 
      Caption         =   "Obligaciones con el Banco de la Nación"
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
      Height          =   855
      Left            =   4200
      TabIndex        =   15
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtTipoCambSBS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Left            =   6480
         TabIndex        =   42
         Text            =   "0.00"
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox txtBancoMN 
         Height          =   375
         Left            =   720
         TabIndex        =   17
         Top             =   300
         Width           =   1815
      End
      Begin VB.TextBox txtBancoME 
         Height          =   375
         Left            =   3600
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label lblTpoCambSBS 
         Caption         =   "Tipo Camb. SBS"
         Height          =   375
         Left            =   5640
         TabIndex        =   41
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblSoles 
         Caption         =   "Soles:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   380
         Width           =   615
      End
      Begin VB.Label lblDolares 
         Caption         =   "Dólares:"
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   380
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   10260
      TabIndex        =   14
      Top             =   7560
      Width           =   1395
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   7560
      Width           =   1395
   End
   Begin VB.Frame frmPersona 
      Caption         =   "Valorizacion Diaria"
      Height          =   3495
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   11535
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   10350
         TabIndex        =   24
         Top             =   840
         Width           =   1000
      End
      Begin Sicmact.FlexEdit FlexValorizacion 
         Height          =   3075
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   10095
         _extentx        =   17806
         _extenty        =   5424
         cols0           =   9
         highlight       =   2
         allowuserresizing=   1
         encabezadosnombres=   "Nro-Fecha Emisión- Instrumento-Nemónico-Tasa Interés-Nro de Papeles-Precio Nominal-Precio Sucio-Valor Razonable"
         encabezadosanchos=   "400-1100-1050-1100-1100-1300-1300-1300-1300"
         font            =   "frmValorizacionDiariaAnexo15ANew.frx":0000
         font            =   "frmValorizacionDiariaAnexo15ANew.frx":0028
         font            =   "frmValorizacionDiariaAnexo15ANew.frx":0050
         font            =   "frmValorizacionDiariaAnexo15ANew.frx":0078
         font            =   "frmValorizacionDiariaAnexo15ANew.frx":00A0
         fontfixed       =   "frmValorizacionDiariaAnexo15ANew.frx":00C8
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-X-X-X-X-X-X-7-X"
         textstylefixed  =   4
         listacontroles  =   "0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-C-C-C-C-C-C-C"
         formatosedit    =   "0-0-0-2-2-3-2-2-0"
         cantdecimales   =   4
         textarray0      =   "Nro"
         lbeditarflex    =   -1  'True
         lbformatocol    =   -1  'True
         lbpuntero       =   -1  'True
         lbordenacol     =   -1  'True
         lbbuscaduplicadotext=   -1  'True
         colwidth0       =   405
         rowheight0      =   300
      End
   End
   Begin VB.Frame fraFechaValorizacion 
      Caption         =   "Fecha de Valorización"
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
      Height          =   705
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2485
      Begin MSMask.MaskEdBox txtFechaValorizacion 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
   End
   Begin VB.Frame fraDescripcion 
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   10215
      Begin VB.TextBox txtPrecValorRazon 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Height          =   350
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   33
         Tag             =   "2"
         Text            =   "0.00"
         Top             =   1240
         Width           =   1455
      End
      Begin VB.TextBox txtPrecSucio 
         Enabled         =   0   'False
         Height          =   350
         Left            =   6480
         TabIndex        =   32
         Top             =   1240
         Width           =   1450
      End
      Begin VB.TextBox txtPrecNom 
         Enabled         =   0   'False
         Height          =   350
         Left            =   4400
         TabIndex        =   31
         Top             =   1240
         Width           =   1575
      End
      Begin VB.TextBox txtNroPap 
         Enabled         =   0   'False
         Height          =   350
         Left            =   2250
         TabIndex        =   30
         Top             =   1240
         Width           =   1455
      End
      Begin VB.TextBox txtTasaInt 
         Enabled         =   0   'False
         Height          =   350
         Left            =   120
         TabIndex        =   29
         Top             =   1240
         Width           =   1455
      End
      Begin VB.TextBox txtNemonico 
         Enabled         =   0   'False
         Height          =   350
         Left            =   8280
         TabIndex        =   28
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtTipoInstr 
         Enabled         =   0   'False
         Height          =   350
         Left            =   5520
         TabIndex        =   27
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtValorRazon 
         Height          =   375
         Left            =   6720
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   2475
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtFechaEmision 
         Height          =   350
         Left            =   1680
         TabIndex        =   25
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFechaEmision 
         Caption         =   "Fecha Emisión:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   420
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Valor Razonable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8280
         TabIndex        =   12
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Instrumento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3600
         TabIndex        =   8
         Top             =   420
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Nemónico:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   7
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Tasa de Interés"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Nro de Papeles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Precio Nominal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Precio Sucio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   10320
      TabIndex        =   20
      Top             =   5400
      Width           =   1335
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   180
         TabIndex        =   35
         Top             =   1200
         Width           =   1000
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   180
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   180
         TabIndex        =   26
         Top             =   360
         Width           =   1000
      End
   End
   Begin VB.Label Label9 
      Caption         =   "LTESORO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   39
      Top             =   4860
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "CDBCRP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   38
      Top             =   4860
      Width           =   975
   End
End
Attribute VB_Name = "frmValorizacionDiariaAnexo15ANew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dFechaBN  As Date
Dim vDatos As Variant
Dim lNuevo   As Boolean
Dim ValorCelda As String
Dim rsValNew As New ADODB.Recordset
Dim oValor As New DAnexoRiesgos
Dim ix As Integer

Public Sub GeneraValorizacionDiaria(psOpeCod As String, pdFecha As Date)
     txtFechaValorizacion.Text = Format(pdFecha, "dd/MM/YYYY")
     Call CargarValorizacionDiaria(pdFecha)
     Call FormEstructura("FormNewNext") 'NAGL 20191214
     SumaBCRPLTesoro
     CentraForm Me
     Me.Show 1
End Sub

Private Sub CargarValorizacionDiaria(pdFecha As Date)
    Dim rs As New ADODB.Recordset
    Dim X As Integer
    
     Set rsValNew = oValor.CargaObligacionBN(pdFecha, "0")
     If Not (rsValNew.BOF And rsValNew.EOF) Then
         txtBancoMN = Format(rsValNew!mObligacionMN, "###,##0.00")
         txtBancoME = Format(rsValNew!mObligacionME, "###,##0.00")
     Else
         txtBancoMN = Format(0, "###,##0.00")
         txtBancoME = Format(0, "###,##0.00")
     End If
    txtTipoCambSBS = oValor.ObtieneTipoCambioContableSBS(pdFecha) 'NAGL 20190502 ERS006-2019
    Set rs = oValor.DevuelveValorizacionDiariaDet(pdFecha, "0")
    FlexValorizacion.Clear
    FormateaFlex FlexValorizacion

    If Not (rs.EOF And rs.BOF) Then
        For X = 1 To rs.RecordCount
            FlexValorizacion.AdicionaFila
            FlexValorizacion.TextMatrix(X, 1) = Format(rs!dFechaEmision, "dd/mm/yyyy")
            FlexValorizacion.TextMatrix(X, 2) = rs!cTipoInstrumento
            FlexValorizacion.TextMatrix(X, 3) = rs!cNemonico
            FlexValorizacion.TextMatrix(X, 4) = Format(rs!nTasaInteres, "#,##0.0000")
            FlexValorizacion.TextMatrix(X, 5) = Format(rs!nNroPapeles, "#,##0")
            FlexValorizacion.TextMatrix(X, 6) = Format(rs!nPrecioNominal, "#,##0")
            FlexValorizacion.TextMatrix(X, 7) = Format(rs!nPrecioSucio, "#,##0.0000")
            FlexValorizacion.TextMatrix(X, 8) = Format(rs!nValorRazonable, "#,##0.00")
            rs.MoveNext
        Next
    End If
End Sub

Private Sub cmdBuscarValorizacion_Click()
Dim rs As New ADODB.Recordset
Dim X As Integer
Dim pdFecha As Date

If txtFechaValorizacion.Text = "" Or txtFechaValorizacion.Text = "__/__/____" Then
   MsgBox "Debe Ingresar la Fecha de Valorización", vbOKOnly + vbInformation, "Atención"
   txtFechaValorizacion.SetFocus
ElseIf ValFecha(txtFechaValorizacion) = False Then
    txtFechaValorizacion.SetFocus
Else
   pdFecha = txtFechaValorizacion.Text
   If ValRegValorizacionDiaria(pdFecha) Then
         Set rsValNew = oValor.CargaObligacionBN(pdFecha, "1")
         If Not (rsValNew.BOF And rsValNew.EOF) Then
             txtBancoMN = Format(rsValNew!mObligacionMN, "###,##0.00")
             txtBancoME = Format(rsValNew!mObligacionME, "###,##0.00")
         Else
             txtBancoMN = Format(0, "###,##0.00")
             txtBancoME = Format(0, "###,##0.00")
         End If
         
        txtTipoCambSBS = oValor.ObtieneTipoCambioContableSBS(pdFecha) 'NAGL 20190502 ERS006-2019
        Set rs = oValor.DevuelveValorizacionDiariaDet(pdFecha, "1")
        FlexValorizacion.Clear
        FormateaFlex FlexValorizacion
        
        If Not (rs.EOF And rs.BOF) Then
            For X = 1 To rs.RecordCount
                FlexValorizacion.AdicionaFila
                FlexValorizacion.TextMatrix(X, 1) = Format(rs!dFechaEmision, "dd/mm/yyyy")
                FlexValorizacion.TextMatrix(X, 2) = rs!cTipoInstrumento
                FlexValorizacion.TextMatrix(X, 3) = rs!cNemonico
                FlexValorizacion.TextMatrix(X, 4) = Format(rs!nTasaInteres, "#,##0.0000")
                FlexValorizacion.TextMatrix(X, 5) = Format(rs!nNroPapeles, "#,##0")
                FlexValorizacion.TextMatrix(X, 6) = Format(rs!nPrecioNominal, "#,##0")
                FlexValorizacion.TextMatrix(X, 7) = Format(rs!nPrecioSucio, "#,##0.0000")
                FlexValorizacion.TextMatrix(X, 8) = Format(rs!nValorRazonable, "#,##0.00")
                rs.MoveNext
            Next
        End If
         SumaBCRPLTesoro
    End If
End If
End Sub

Private Function LlenarRsValorDiariaDet(ByVal feControl As FlexEdit) As ADODB.Recordset

 Dim rsVD As New ADODB.Recordset
 Dim nIndex As Integer
 
  If feControl.Rows >= 2 Then
         If feControl.TextMatrix(nIndex, 1) = "" Then
            Exit Function
        End If

        rsVD.CursorType = adOpenStatic
        rsVD.Fields.Append "dFechaEmision", adDate, adFldIsNullable
        rsVD.Fields.Append "cTipoInstrumento", adVarChar, 20, adFldIsNullable
        rsVD.Fields.Append "cNemonico", adVarChar, 20, adFldIsNullable
        rsVD.Fields.Append "nTasaInteres", adDouble, adFldIsNullable
        rsVD.Fields.Append "nNroPapeles", adInteger, adFldIsNullable
        rsVD.Fields.Append "nPrecioNominal", adDouble, adFldIsNullable
        rsVD.Fields.Append "nPrecioSucio", adDouble, adFldIsNullable
        rsVD.Fields.Append "nValorRazonable", adDouble, adFldIsNullable
        rsVD.Open

        For nIndex = 1 To feControl.Rows - 1
            rsVD.AddNew
            rsVD.Fields("dFechaEmision") = feControl.TextMatrix(nIndex, 1)
            rsVD.Fields("cTipoInstrumento") = feControl.TextMatrix(nIndex, 2)
            rsVD.Fields("cNemonico") = feControl.TextMatrix(nIndex, 3)
            rsVD.Fields("nTasaInteres") = feControl.TextMatrix(nIndex, 4)
            rsVD.Fields("nNroPapeles") = feControl.TextMatrix(nIndex, 5)
            rsVD.Fields("nPrecioNominal") = feControl.TextMatrix(nIndex, 6)
            rsVD.Fields("nPrecioSucio") = feControl.TextMatrix(nIndex, 7)
            rsVD.Fields("nValorRazonable") = feControl.TextMatrix(nIndex, 8)
            rsVD.Update
            rsVD.MoveFirst
        Next
    End If
    Set LlenarRsValorDiariaDet = rsVD
End Function

Private Sub cmdGuardar_Click()
    Dim oValor As New DAnexoRiesgos
    Dim lsItemValor As New ADODB.Recordset
    Dim i As Integer
    Dim registro As Boolean
    Dim lsMovNro As String
    Dim dFechaValor As Date
    Dim oCont As New NContFunciones
    
    If txtBancoMN.Text = "" Then
           MsgBox "Falta ingresar Moneda Nacional", vbInformation, "Aviso"
           Exit Sub
    End If
    If txtTipoCambSBS.Text = "" Or CDbl(txtTipoCambSBS.Text) = 0 Then
           MsgBox "Falta ingresar el Tipo de Cambio Contable", vbInformation, "Aviso"
           Exit Sub
    End If 'NAGL 20190502
    If Len(txtBancoMN.Text) > 12 Then
          MsgBox "El valor de la Obligación con el Banco de la Nación en MN es incorrecto", vbInformation, "Aviso"
         Exit Sub
    End If
    If Len(txtBancoME.Text) > 12 Then
         MsgBox "El valor de la Obligación con el Banco de la Nación en ME es incorrecto", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(txtTipoCambSBS) > 8 Then
         MsgBox "El valor del Tipo de Cambio Contable es incorrecto", vbInformation, "Aviso"
        Exit Sub
    End If 'NAGL 20190502
    
  If MsgBox("¿Desea Registrar estos Valores Representativos de Deuda.", vbInformation + vbYesNo, "Atención") = vbYes Then
        If (txtFechaValorizacion.Text = "" And txtFechaValorizacion.Text = "__/__/____") Then
                MsgBox "Falta ingresar la Fecha de Valorización", vbInformation, "Aviso"
        ElseIf (ValFecha(txtFechaValorizacion) = False) Then
              txtFechaValorizacion.SetFocus
        Else
                dFechaValor = Format(txtFechaValorizacion.Text, "dd/mm/yyyy")
                lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                'Call oValor.ControlValorDiaria(Format(txtFechaValorizacion.Text, "dd/mm/yyyy hh:mm:ss"), CDbl(txtBancoMN.Text), CDbl(txtBancoME.Text), LlenarRsValorDiariaDet(FlexValorizacion))
                Call oValor.GuardarTipoCambioSBS(Format(txtFechaValorizacion.Text, "dd/mm/yyyy"), CDbl(txtTipoCambSBS), lsMovNro) 'Agregado by NAGL Según ERS006-2019 20190502
                Call oValor.ControlValorDiaria(Format(txtFechaValorizacion.Text, "dd/mm/yyyy"), CDbl(txtBancoMN.Text), CDbl(txtBancoME.Text), lsMovNro, LlenarRsValorDiariaDet(FlexValorizacion))
                MsgBox "Los datos se guardaron satisfactoriamente.", vbOKOnly + vbInformation, "Atención"
                'Call LimpiarFormulario
                Unload Me
        End If
    End If
End Sub

Private Function ValRegValorizacionDiaria(pdFecha As Date) As Boolean
Dim DAnxVal As New DAnexoRiesgos
Dim psRegistro As String
 psRegistro = DAnxVal.ObtenerRegValorizacionDiaria(pdFecha)
    If psRegistro = "0" Then
            If MsgBox("No existen datos con la Fecha Registrada, Desea continuar?", vbInformation + vbYesNo, "Atención") = vbNo Then
                txtFechaValorizacion.SetFocus
                Exit Function
            End If
    End If
ValRegValorizacionDiaria = True
End Function 'NAGL 20170415

Private Sub CmdModificar_Click()
    If FlexValorizacion.TextMatrix(1, 0) = "" Then
        MsgBox "No existen Valorizaciones Representativos pendientes", vbInformation, "Aviso"
        Exit Sub
    End If
    
    txtFechaEmision = FlexValorizacion.TextMatrix(FlexValorizacion.row, 1)
    txtTipoInstr = FlexValorizacion.TextMatrix(FlexValorizacion.row, 2)
    txtNemonico = FlexValorizacion.TextMatrix(FlexValorizacion.row, 3)
    txtTasaInt.Text = FlexValorizacion.TextMatrix(FlexValorizacion.row, 4)
    txtNroPap.Text = FlexValorizacion.TextMatrix(FlexValorizacion.row, 5)
    txtPrecNom.Text = FlexValorizacion.TextMatrix(FlexValorizacion.row, 6)
    txtPrecSucio.Text = FlexValorizacion.TextMatrix(FlexValorizacion.row, 7)
    txtPrecValorRazon = FlexValorizacion.TextMatrix(FlexValorizacion.row, 8)
        
    Me.txtFechaEmision.Enabled = True
    Me.txtTipoInstr.Enabled = True
    Me.txtNemonico.Enabled = True
    Me.txtTasaInt.Enabled = True
    Me.txtNroPap.Enabled = True
    Me.txtPrecNom.Enabled = True
    Me.txtPrecSucio.Enabled = True
    
    txtFechaEmision.SetFocus
    cmdAceptar.Visible = True
    cmdAgregar.Visible = False
    Call FormEstructura 'NAGL 20191214
End Sub

Private Sub cmdAceptar_Click()
    If txtFechaEmision.Text = "" Or txtFechaEmision.Text = "__/__/____" Then
        MsgBox "Falta ingresar Fecha de Emisión", vbInformation, "Aviso"
        Exit Sub
    End If
    If txtTipoInstr.Text = "" Then
         MsgBox "Falta ingresar el Tipo de Instrumento", vbInformation, "Aviso"
         Exit Sub
    End If
    If txtNemonico.Text = "" Then
        MsgBox "Falta ingresar el Nemónico", vbInformation, "Aviso"
        Exit Sub
    End If
    If txtTasaInt.Text = "" Then
        MsgBox "Falta ingresar la Tasa de Intéres", vbInformation, "Aviso"
        Exit Sub
    End If
    If txtNroPap.Text = "" Then
        MsgBox "Falta ingresar el Nro de Papeles", vbInformation, "Aviso"
        Exit Sub
    End If
    If txtPrecNom.Text = "" Then
        MsgBox "Falta ingresar el Precio Nominal", vbInformation, "Aviso"
        Exit Sub
    End If
    If txtPrecSucio.Text = "" Then
        MsgBox "Falta ingresar el Precio Sucio", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(txtTasaInt.Text) > 12 Then
         MsgBox "El valor de la tasa de Interes es incorrecto", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(txtNroPap.Text) > 12 Then
         MsgBox "El valor del Nro de Papeles es incorrecto", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(txtPrecNom.Text) > 12 Then
         MsgBox "El valor del Precio Nominal es incorrecto", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(txtPrecSucio.Text) > 12 Then
         MsgBox "El valor del Precio Sucio es incorrecto", vbInformation, "Aviso"
        Exit Sub
    End If
    
    FlexValorizacion.TextMatrix(FlexValorizacion.row, 1) = Me.txtFechaEmision.Text
    FlexValorizacion.TextMatrix(FlexValorizacion.row, 2) = Me.txtTipoInstr.Text
    FlexValorizacion.TextMatrix(FlexValorizacion.row, 3) = Me.txtNemonico.Text
    FlexValorizacion.TextMatrix(FlexValorizacion.row, 4) = Me.txtTasaInt.Text
    FlexValorizacion.TextMatrix(FlexValorizacion.row, 5) = Me.txtNroPap.Text
    FlexValorizacion.TextMatrix(FlexValorizacion.row, 6) = Me.txtPrecNom.Text
    FlexValorizacion.TextMatrix(FlexValorizacion.row, 7) = Me.txtPrecSucio.Text
    FlexValorizacion.TextMatrix(FlexValorizacion.row, 8) = Me.txtPrecValorRazon.Text
    
    Call LimpiarFormulario
    Me.txtFechaEmision.Enabled = False
    Me.txtTipoInstr.Enabled = False
    Me.txtNemonico.Enabled = False
    Me.txtTasaInt.Enabled = False
    Me.txtNroPap.Enabled = False
    Me.txtPrecNom.Enabled = False
    Me.txtPrecSucio.Enabled = False
    Call FormEstructura("FormNewNext") 'NAGL 20191214
    SumaBCRPLTesoro
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub FlexValorizacion_RowColChange()
    If FlexValorizacion.col = 7 Then
            FlexValorizacion.AvanceCeldas = Vertical
        Else
            FlexValorizacion.AvanceCeldas = Horizontal
            'cmdGuardar.SetFocus
    End If
    SumaBCRPLTesoro
End Sub
Private Sub FlexValorizacion_EnterCell()
   ValorCelda = FlexValorizacion.TextMatrix(FlexValorizacion.row, FlexValorizacion.col)
End Sub

Private Sub FlexValorizacion_OnCellChange(pnRow As Long, pnCol As Long)
    Dim valorNew As Currency
    Dim ValorAnterior As Currency
    
    ValorAnterior = CDbl(ValorCelda)
     If (IsNumeric(FlexValorizacion.TextMatrix(pnRow, pnCol)) And Len(FlexValorizacion.TextMatrix(pnRow, pnCol)) < 9) Then
         valorNew = CDbl(FlexValorizacion.TextMatrix(pnRow, pnCol))
          If (valorNew < 0) Then
                MsgBox "No se puede asignar un valor negativo", vbInformation, "Aviso"
                FlexValorizacion.TextMatrix(pnRow, pnCol) = Format(ValorAnterior, "###,##0.0000")
                Exit Sub
          End If
         FlexValorizacion.TextMatrix(pnRow, pnCol + 1) = CDbl(FlexValorizacion.TextMatrix(pnRow, 5)) * CDbl(FlexValorizacion.TextMatrix(pnRow, 6)) / 100 * valorNew
         FlexValorizacion.TextMatrix(pnRow, pnCol + 1) = Format(FlexValorizacion.TextMatrix(pnRow, pnCol + 1), "#,##0.00")
    Else
        'MsgBox "No Valido", vbInformation, "Aviso"
        FlexValorizacion.TextMatrix(pnRow, pnCol) = Format(ValorAnterior, "###,##0.0000")
    End If
End Sub

Private Sub Form_Load()
    CentraForm Me
End Sub
Private Sub txtFechaValorizacion_GotFocus()
fEnfoque txtFechaValorizacion
End Sub

Private Sub txtFechaValorizacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      If ValFecha(txtFechaValorizacion) = True Then
         Me.txtBancoMN.SetFocus
      End If
End If
End Sub

Private Sub txtBancoMN_GotFocus()
fEnfoque txtBancoMN
End Sub

Private Sub txtBancoMN_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales4(txtBancoMN, KeyAscii)
    If KeyAscii = 13 Then
    txtBancoME.SetFocus
    txtBancoMN.Text = Format(txtBancoMN.Text, "#,##0.00")
    End If
End Sub

Private Sub txtBancoMN_LostFocus()
    If Len(Trim(txtBancoMN.Text)) = 0 Then
         txtBancoMN.Text = ""
    End If
    txtBancoMN.Text = Format(txtBancoMN.Text, "#,##0.00")
End Sub

Private Sub txtBancoME_GotFocus()
fEnfoque txtBancoME
End Sub

Private Sub txtBancoME_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales4(txtBancoME, KeyAscii)
    If KeyAscii = 13 Then
    txtTipoCambSBS.SetFocus 'NAGL 20190503 Cambío de cmdGuardar
    txtBancoME.Text = Format(txtBancoME.Text, "#,##0.00")
    End If
End Sub

Private Sub txtBancoME_LostFocus()
     If Len(Trim(txtBancoME.Text)) = 0 Then
         txtBancoME.Text = "0.00"
    End If
    txtBancoME.Text = Format(txtBancoME.Text, "#,##0.00")
End Sub

Private Sub txtTipoCambSBS_GotFocus()
    fEnfoque txtTipoCambSBS
End Sub 'NAGL 20190503

Private Sub txtTipoCambSBS_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales4(txtTipoCambSBS, KeyAscii)
    If KeyAscii = 13 Then
        cmdGuardar.SetFocus
    End If
End Sub 'NAGL 20190503

Private Sub txtPrecValorRazon_GotFocus()
 fEnfoque txtValorRazon
End Sub

Private Sub txtPrecValorRazon_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales4(txtPrecValorRazon, KeyAscii)
If KeyAscii = 13 Then
    If cmdAgregar.Visible = True Then
    cmdAgregar.SetFocus
    Else
    cmdAceptar.SetFocus
    End If
End If
End Sub

Private Sub txtPrecValorRazon_LostFocus()
If Val(txtPrecValorRazon) = 0 Then txtPrecValorRazon = 0
txtPrecValorRazon = Format(txtPrecValorRazon, "#,##0.00")
CalculaValorRazon
End Sub

Private Sub txtPrecSucio_GotFocus()
fEnfoque txtPrecSucio
End Sub

Private Sub txtPrecSucio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales4(txtPrecSucio, KeyAscii)
    If KeyAscii = 13 Then
    txtPrecValorRazon.SetFocus
    txtPrecSucio.Text = Format(txtPrecSucio.Text, "#,##0.0000")
    CalculaValorRazon
    End If
End Sub

Private Sub txtPrecSucio_LostFocus()
     If Len(Trim(txtPrecSucio.Text)) = 0 Then
         txtPrecSucio.Text = ""
    End If
    txtPrecSucio.Text = Format(txtPrecSucio.Text, "#,##0.0000")
End Sub

Private Sub txtPrecNom_GotFocus()
fEnfoque txtPrecNom
End Sub

Private Sub txtPrecNom_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtPrecNom, KeyAscii)
    If KeyAscii = 13 Then
        Me.txtPrecSucio.SetFocus
        CalculaValorRazon
    End If
End Sub

Private Sub txtPrecNom_LostFocus()
    Dim Cant As Integer
    Dim cant2 As Integer
    Dim sPSuc As String
    Dim sPSuc2 As String
    sPSuc = txtPrecNom.Text
    Cant = Len(sPSuc)
    sPSuc2 = Replace(sPSuc, ".", "")
    cant2 = Len(sPSuc2)
    
    If Len(Trim(txtPrecNom.Text)) = 0 Then
        txtPrecNom.Text = ""
    End If
    If Cant <> cant2 Then
        txtPrecNom.Text = Format(txtPrecNom.Text, "#,##0.00")
        ElseIf Cant = cant2 Then
        txtPrecNom.Text = Format(txtPrecNom.Text, "#,##0")
    End If
End Sub

Private Sub txtNroPap_GotFocus()
fEnfoque txtNroPap
End Sub

Private Sub txtNroPap_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales4(txtNroPap, KeyAscii)
    If KeyAscii = 13 Then
    Me.txtPrecNom.SetFocus
    txtNroPap.Text = Format(txtNroPap.Text, "#,##0")
    CalculaValorRazon
    End If
    KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtNroPap_LostFocus()
  If Len(Trim(txtNroPap.Text)) = 0 Then
         txtNroPap.Text = ""
  End If
   txtNroPap.Text = Format(txtNroPap.Text, "#,##0")
End Sub

Private Sub txtTasaInt_GotFocus()
fEnfoque txtTasaInt
End Sub

Private Sub txtTasaInt_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales4(txtTasaInt, KeyAscii)
If KeyAscii = 13 Then
   Me.txtNroPap.SetFocus
    txtTasaInt.Text = Format(txtTasaInt.Text, "#,##0.0000")
End If
End Sub

Private Sub txtTasaInt_LostFocus()
  If Len(Trim(txtTasaInt.Text)) = 0 Then
         txtTasaInt.Text = ""
  End If
   txtTasaInt.Text = Format(txtTasaInt.Text, "#,##0.0000")
End Sub

Private Sub txtNemonico_GotFocus()
fEnfoque txtNemonico
End Sub

Private Sub txtNemonico_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.txtTasaInt.SetFocus
End If
   KeyAscii = Letras(KeyAscii)
End Sub

Private Sub txtTipoInstr_GotFocus()
fEnfoque txtTipoInstr
End Sub

Private Sub txtTipoInstr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.txtNemonico.SetFocus
End If
    KeyAscii = Letras(KeyAscii)
End Sub

Private Sub txtFechaEmision_GotFocus()
fEnfoque txtFechaEmision
End Sub

Private Sub txtFechaEmision_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      If ValFecha(txtFechaEmision) Then
         Me.txtTipoInstr.SetFocus
      End If
End If
End Sub

Private Sub cmdAgregar_Click()
    If txtFechaEmision.Text = "" Or txtFechaEmision.Text = "__/__/____" Then
        MsgBox "Falta ingresar Fecha de Emisión", vbInformation, "Aviso"
        Exit Sub
    End If
    If txtTipoInstr.Text = "" Then
         MsgBox "Falta ingresar el Tipo de Instrumento", vbInformation, "Aviso"
         Exit Sub
    End If
    If txtNemonico.Text = "" Then
        MsgBox "Falta ingresar el Nemónico", vbInformation, "Aviso"
        Exit Sub
    End If
    If txtTasaInt.Text = "" Then
        MsgBox "Falta ingresar la Tasa de Intéres", vbInformation, "Aviso"
        Exit Sub
    End If
    If txtNroPap.Text = "" Then
        MsgBox "Falta ingresar el Nro de Papeles", vbInformation, "Aviso"
        Exit Sub
    End If
    If txtPrecNom.Text = "" Then
        MsgBox "Falta ingresar el Precio Nominal", vbInformation, "Aviso"
        Exit Sub
    End If
    If txtPrecSucio.Text = "" Then
        MsgBox "Falta ingresar el Precio Sucio", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(txtTasaInt.Text) > 12 Then
         MsgBox "El valor de la tasa de Interes es incorrecto", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(txtNroPap.Text) > 12 Then
         MsgBox "El valor del Nro de Papeles es incorrecto", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(txtPrecNom.Text) > 12 Then
         MsgBox "El valor del Precio Nominal es incorrecto", vbInformation, "Aviso"
        Exit Sub
    End If
    If Len(txtPrecSucio.Text) > 12 Then
         MsgBox "El valor del Precio Sucio es incorrecto", vbInformation, "Aviso"
        Exit Sub
    End If
        FlexValorizacion.AdicionaFila
        FlexValorizacion.TextMatrix(FlexValorizacion.Rows - 1, 1) = Me.txtFechaEmision.Text
        FlexValorizacion.TextMatrix(FlexValorizacion.Rows - 1, 2) = Me.txtTipoInstr.Text
        FlexValorizacion.TextMatrix(FlexValorizacion.Rows - 1, 3) = Me.txtNemonico.Text
        FlexValorizacion.TextMatrix(FlexValorizacion.Rows - 1, 4) = Me.txtTasaInt.Text
        FlexValorizacion.TextMatrix(FlexValorizacion.Rows - 1, 5) = Me.txtNroPap.Text
        FlexValorizacion.TextMatrix(FlexValorizacion.Rows - 1, 6) = Me.txtPrecNom.Text
        FlexValorizacion.TextMatrix(FlexValorizacion.Rows - 1, 7) = Me.txtPrecSucio.Text
        FlexValorizacion.TextMatrix(FlexValorizacion.Rows - 1, 8) = Me.txtPrecValorRazon.Text
    
    Call LimpiarFormulario
    If MsgBox("Desea agregar un nuevo instrumento...?", vbInformation + vbYesNo, "Atención") = vbNo Then
        Me.txtFechaEmision.Enabled = False
        Me.txtTipoInstr.Enabled = False
        Me.txtNemonico.Enabled = False
        Me.txtTasaInt.Enabled = False
        Me.txtNroPap.Enabled = False
        Me.txtPrecNom.Enabled = False
        Me.txtPrecSucio.Enabled = False
        Call FormEstructura("FormNewNext") 'NAGL 20191214
    End If 'NAGL 20191214 Agregó Condicional
    SumaBCRPLTesoro
End Sub

Private Sub cmdQuitar_Click()
    Call LimpiarFormulario
    'Me.txtFechaEmision.Enabled = False
    'Me.txtTipoInstr.Enabled = False
    'Me.txtNemonico.Enabled = False
    'Me.txtTasaInt.Enabled = False
    'Me.txtNroPap.Enabled = False
    'Me.txtPrecNom.Enabled = False
    'Me.txtPrecSucio.Enabled = False
    'Comentado by NAGL 20191216
End Sub

Private Sub LimpiarFormulario()
    txtFechaEmision.Text = "__/__/____"
    txtTipoInstr.Text = ""
    txtNemonico.Text = ""
    txtTasaInt.Text = ""
    txtNroPap.Text = ""
    txtPrecNom.Text = ""
    txtPrecSucio.Text = ""
    txtPrecValorRazon.Text = "0.00"
End Sub

Private Sub CalculaValorRazon()
    If Len(txtNroPap.Text) <> 0 Then
        txtPrecValorRazon = Format(CDbl(txtNroPap.Text), "#,#0.00")
    End If
    If Len(txtNroPap.Text) <> 0 And Len(txtPrecNom.Text) <> 0 Then
        txtPrecValorRazon = Format(CDbl(txtNroPap) * CDbl(txtPrecNom.Text), "#,#0.00")
    End If
    If (Len(txtNroPap.Text) = 0 And Len(txtPrecNom.Text) = 0 And Len(txtPrecSucio.Text) = 0) Then
    txtPrecValorRazon = Format(CDbl(txtPrecValorRazon), "#,#0.00")
    ElseIf (Len(txtNroPap.Text) <> 0 And Len(txtPrecNom.Text) <> 0 And Len(txtPrecSucio.Text) <> 0) Then
    txtPrecValorRazon = Format(CDbl(txtNroPap) * CDbl(txtPrecNom) * CDbl(txtPrecSucio) / 100, "#,#0.00")
    End If
End Sub

Public Sub SumaBCRPLTesoro()
Dim nSaldoCDBCRP As Double, nSaldoLTesoro As Double
Dim nNroRegVal As Integer
nSaldoCDBCRP = 0
nSaldoLTesoro = 0
pnRow = 1

If (FlexValorizacion.TextMatrix(FlexValorizacion.Rows - 1, 0) <> "") Then
    nNroRegVal = CInt(FlexValorizacion.TextMatrix(FlexValorizacion.Rows - 1, 0))
    If nNroRegVal <> 0 Then
         For ix = 1 To nNroRegVal
            If (Mid(CStr(FlexValorizacion.TextMatrix(pnRow, 2)), 1, 1) = "C") Then
                nSaldoCDBCRP = CDbl(FlexValorizacion.TextMatrix(pnRow, 8)) + nSaldoCDBCRP
                
            ElseIf (Mid(CStr(FlexValorizacion.TextMatrix(pnRow, 2)), 1, 1) = "L") Then
                nSaldoLTesoro = CDbl(FlexValorizacion.TextMatrix(pnRow, 8)) + nSaldoLTesoro
            End If
            pnRow = pnRow + 1
         Next ix
    Else
        nSaldoCDBCRP = Format(0, "#,#0.00")
        nSaldoLTesoro = Format(0, "#,#0.00")
    End If
    txtCDBCRPTotal.Text = Format(nSaldoCDBCRP, "#,#0.00")
    txtLTesoroTotal.Text = Format(nSaldoLTesoro, "#,#0.00")
Else
    txtCDBCRPTotal.Text = Format(0, "#,#0.00")
    txtLTesoroTotal.Text = Format(0, "#,#0.00")
End If
End Sub 'NAGL 20170423

Private Sub cmdNuevo_Click()
    cmdAceptar.Visible = False
    cmdAgregar.Visible = True
    cmdQuitar.Visible = True
    Call LimpiarFormulario
    Me.txtFechaEmision.Enabled = True
    Me.txtTipoInstr.Enabled = True
    Me.txtNemonico.Enabled = True
    Me.txtTasaInt.Enabled = True
    Me.txtNroPap.Enabled = True
    Me.txtPrecNom.Enabled = True
    Me.txtPrecSucio.Enabled = True
    Me.txtFechaEmision.SetFocus
    Call FormEstructura 'NAGL 20191214
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("¿Esta seguro que desea quitar el Valor Representativo?", vbYesNo + vbQuestion, "Atención") = vbNo Then Exit Sub
    Call FlexValorizacion.EliminaFila(FlexValorizacion.row)
    SumaBCRPLTesoro
End Sub

Public Sub FormEstructura(Optional psTipo As String)
If psTipo = "FormNewNext" Then
    frmValorizacionDiariaAnexo15ANew.Height = 5730
    cmdSalir.Top = 4800
    cmdGuardar.Top = 4800
Else
    frmValorizacionDiariaAnexo15ANew.Height = 8490
    cmdSalir.Top = 7560
    cmdGuardar.Top = 7560
End If
End Sub 'NAGL 20191214

