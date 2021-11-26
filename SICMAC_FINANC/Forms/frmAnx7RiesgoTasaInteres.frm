VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAnx7RiesgoTasaInteres 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anexo 07: Medición del Riesgo de Tasa de Interés"
   ClientHeight    =   3015
   ClientLeft      =   3165
   ClientTop       =   4515
   ClientWidth     =   9015
   Icon            =   "frmAnx7RiesgoTasaInteres.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConsolidar 
      Caption         =   "Consolidar"
      Height          =   345
      Left            =   5280
      TabIndex        =   15
      Top             =   2520
      Width           =   1155
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   7740
      TabIndex        =   11
      Top             =   2520
      Width           =   1155
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   345
      Left            =   6540
      TabIndex        =   2
      Top             =   2520
      Width           =   1155
   End
   Begin VB.Frame fraRep 
      Height          =   2325
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8775
      Begin VB.Frame Frame2 
         Caption         =   "Tipo Reporte"
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
         Height          =   765
         Left            =   4560
         TabIndex        =   17
         Top             =   240
         Width           =   4050
         Begin VB.OptionButton optAntiguoNuevo 
            Caption         =   "Nuevo"
            Height          =   495
            Index           =   1
            Left            =   2520
            TabIndex        =   19
            Top             =   170
            Width           =   1095
         End
         Begin VB.OptionButton optAntiguoNuevo 
            Caption         =   "Antiguo"
            Height          =   495
            Index           =   0
            Left            =   960
            TabIndex        =   18
            Top             =   170
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   1155
         Left            =   240
         TabIndex        =   6
         Top             =   1020
         Width           =   8370
         Begin VB.TextBox txtUNA 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5400
            TabIndex        =   20
            Text            =   "0"
            Top             =   720
            Width           =   1425
         End
         Begin VB.TextBox txtTipCambio 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1605
            TabIndex        =   8
            Text            =   "0"
            Top             =   225
            Width           =   1425
         End
         Begin VB.TextBox txtPatriEfec 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1605
            TabIndex        =   7
            Text            =   "0"
            Top             =   705
            Width           =   1425
         End
         Begin VB.Label lblUNA 
            Caption         =   "Ut.Neta.Anualizada"
            Height          =   255
            Left            =   3840
            TabIndex        =   21
            Top             =   780
            Width           =   1455
         End
         Begin VB.Label lblFechaTpoCambio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3480
            TabIndex        =   12
            Top             =   255
            Width           =   3375
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de Cambio"
            Height          =   285
            Left            =   225
            TabIndex        =   10
            Top             =   270
            Width           =   1275
         End
         Begin VB.Label Label4 
            Caption         =   "Patrimonio Efectivo"
            Height          =   255
            Left            =   45
            TabIndex        =   9
            Top             =   765
            Width           =   1455
         End
      End
      Begin VB.Frame fraMes 
         Caption         =   "Periodo"
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
         Height          =   765
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   4095
         Begin VB.ComboBox cboMes 
            Height          =   315
            ItemData        =   "frmAnx7RiesgoTasaInteres.frx":030A
            Left            =   2280
            List            =   "frmAnx7RiesgoTasaInteres.frx":0332
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   300
            Width           =   1455
         End
         Begin VB.TextBox txtAnio 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   630
            MaxLength       =   4
            TabIndex        =   0
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Año :"
            Height          =   195
            Left            =   180
            TabIndex        =   5
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Mes :"
            Height          =   195
            Left            =   1710
            TabIndex        =   4
            Top             =   390
            Width           =   390
         End
      End
   End
   Begin VB.Label lblAvance 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   4335
   End
End
Attribute VB_Name = "frmAnx7RiesgoTasaInteres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsFecha As String
Dim dPatrimonio As Date

Private Sub cboMes_Click()
    If Me.txtAnio.Text <> "" Then
        txtTipCambio = TipoCambioCierre(txtAnio, cboMes.ListIndex + 1)
                
If txtTipCambio.Text = "0" Then
            Me.lblFechaTpoCambio.ForeColor = &HFF&
            Me.lblFechaTpoCambio.Caption = "No Existe Tipo de Cambio"
            
        Else
            Me.lblFechaTpoCambio.ForeColor = &H0&
            Me.lblFechaTpoCambio.Caption = "al " + obtenerFechaLarga
        End If
'        Dim nPatrimonio As Currency
        If lsFecha <> "" And txtTipCambio.Text <> "0" Then
'
           dPatrimonio = DateAdd("d", -CInt(Mid(lsFecha, 1, 2)), CDate(lsFecha))
'            dPatrimonio = DateAdd("d", -1, DateAdd("d", -CInt(Mid(lsFecha, 1, 2)), CDate(lsFecha)))
'            nPatrimonio = RecuperaPatrimonioEfectivo(Mid(dPatrimonio, 4, 2), Mid(dPatrimonio, 7, 4))
'            Me.txtPatriEfec.Text = Format(nPatrimonio, "##,#00.00")
'        Else
'            Me.txtPatriEfec.Text = "0"
        End If
        
    Else
        Me.txtAnio.SetFocus
    End If
End Sub

Private Sub cboMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.txtAnio.Text <> "" Then
            txtTipCambio = TipoCambioCierre(txtAnio, cboMes.ListIndex + 1)
            If obtenerFechaLarga <> "" Then
                Me.lblFechaTpoCambio.Caption = "al " + obtenerFechaLarga
            End If 'NAGL 20171002
            'cmdGenerar.SetFocus
            Me.txtPatriEfec.SetFocus
        Else
            Me.txtAnio.SetFocus
        End If
    End If
End Sub

Private Sub cmdConsolidar_Click()
Dim objDAnexoRiesgos As DAnexoRiesgos
Dim oCtaIf As NCajaCtaIF
Dim rs As Recordset
Dim sMovNro As String
Me.MousePointer = vbHourglass
If ValidaDatos Then
    Exit Sub
End If
PB1.Min = 0
PB1.Max = 33
PB1.value = 0
PB1.Visible = True

PB1.value = 1

sMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
PB1.value = 2

Set objDAnexoRiesgos = New DAnexoRiesgos
Set oCtaIf = New NCajaCtaIF
Set rs = New Recordset
'creditos vigentes
PB1.value = 3
lblAvance.Caption = "Verificando Creditos Vigentes..."

Set rs = objDAnexoRiesgos.obtenerCreditosVigentesAnx7(CDate(lsFecha))
PB1.value = 4

If Not rs.EOF Or Not rs.BOF Then
    If MsgBox("Ya se ha consolidado Creditos Vigentes del mes de " + cboMes.Text + " del " + txtAnio.Text + ", Desea Reemplazarlo?", vbYesNo) = vbYes Then
        Set rs = Nothing
        PB1.value = 5
        Set rs = oCtaIf.GetCreditosVigentes(Format(lsFecha, "yyyymmdd"))
        PB1.value = 6
        lblAvance.Caption = "Consolidando Creditos Vigentes..."
        PB1.value = 7
        PB1.value = 8
        objDAnexoRiesgos.actualizarCreditosVigentesAnx7 CDate(lsFecha)
        PB1.value = 9
        If Not rs.EOF Or rs.BOF Then
            objDAnexoRiesgos.insertarCreditosVigentesAnx7 CDate(lsFecha), rs, sMovNro
        End If
        PB1.value = 10
        Set rs = Nothing
    End If
Else
        Set rs = Nothing
        PB1.value = 5
        Set rs = oCtaIf.GetCreditosVigentes(Format(lsFecha, "yyyymmdd"))
        PB1.value = 6
        lblAvance.Caption = "Consolidando Creditos Vigentes..."
        PB1.value = 7
        PB1.value = 8
        objDAnexoRiesgos.actualizarCreditosVigentesAnx7 CDate(lsFecha)
        PB1.value = 9
        If Not rs.EOF Or rs.BOF Then
            objDAnexoRiesgos.insertarCreditosVigentesAnx7 CDate(lsFecha), rs, sMovNro
        End If
        PB1.value = 10
        Set rs = Nothing

End If
'Refinanciados
PB1.value = 11
lblAvance.Caption = "Verificando Creditos Refinanciados..."

Set rs = objDAnexoRiesgos.obtenerCreditosRefinanciadosAnx7(CDate(lsFecha))
PB1.value = 12

If Not rs.EOF Or Not rs.BOF Then
    If MsgBox("Ya se ha consolidado Creditos Refinanciados del mes de " + cboMes.Text + " del " + txtAnio.Text + ", Desea Reemplazarlo?", vbYesNo) = vbYes Then
        Set rs = Nothing
        PB1.value = 13
        Set rs = oCtaIf.GetCreditosRefinanciados(Format(lsFecha, "yyyymmdd"))
        PB1.value = 14
        lblAvance.Caption = "Consolidando Creditos Refinanciados..."
        PB1.value = 15
        PB1.value = 16
        objDAnexoRiesgos.actualizarCreditosRefinanciadosAnx7 CDate(lsFecha)
         PB1.value = 17
        If Not rs.EOF Or rs.BOF Then
                objDAnexoRiesgos.insertarCreditosRefinanciadosAnx7 CDate(lsFecha), rs, sMovNro
        End If
        PB1.value = 18
        Set rs = Nothing
    End If
Else
        Set rs = Nothing
        PB1.value = 13
        Set rs = oCtaIf.GetCreditosRefinanciados(Format(lsFecha, "yyyymmdd"))
        PB1.value = 14
        lblAvance.Caption = "Consolidando Creditos Refinanciados..."
        PB1.value = 15
        PB1.value = 16
        objDAnexoRiesgos.actualizarCreditosRefinanciadosAnx7 CDate(lsFecha)
         PB1.value = 17
        If Not rs.EOF Or rs.BOF Then
            objDAnexoRiesgos.insertarCreditosRefinanciadosAnx7 CDate(lsFecha), rs, sMovNro
        End If
        PB1.value = 18
        Set rs = Nothing

End If
 Set rs = Nothing
 
 'Adeudados Moneda Nacional
 PB1.value = 19
lblAvance.Caption = "Verificando Adeudados..."

Set rs = objDAnexoRiesgos.obtenerAdeudadosAnx7(CDate(lsFecha), 1)
PB1.value = 20

If Not rs.EOF Or Not rs.BOF Then
    If MsgBox("Ya se ha consolidado Datos de Adeudados del mes de " + cboMes.Text + " del " + txtAnio.Text + ", Desea Reemplazarlo?", vbYesNo) = vbYes Then
        Set rs = Nothing
        PB1.value = 21
        Set rs = oCtaIf.GetCaleAdeudadosXTramos(Format(lsFecha, "yyyymmdd"), "1")
        PB1.value = 22
        lblAvance.Caption = "Consolidando Adeudados..."
        PB1.value = 23
        PB1.value = 24
        objDAnexoRiesgos.actualizarAdeudadosAnx7 CDate(lsFecha)
        PB1.value = 25
        If Not rs.EOF Or rs.BOF Then
            objDAnexoRiesgos.insertarAdeudadosAnx7 CDate(lsFecha), rs, sMovNro, 1
        End If
        PB1.value = 26
        Set rs = Nothing
    End If
Else
        Set rs = Nothing
        PB1.value = 21
        Set rs = oCtaIf.GetCaleAdeudadosXTramos(Format(lsFecha, "yyyymmdd"), "1")
        PB1.value = 22
        lblAvance.Caption = "Consolidando Adeudados..."
        PB1.value = 23
        PB1.value = 24
        objDAnexoRiesgos.actualizarAdeudadosAnx7 CDate(lsFecha)
        PB1.value = 25
        If Not rs.EOF Or rs.BOF Then
            objDAnexoRiesgos.insertarAdeudadosAnx7 CDate(lsFecha), rs, sMovNro, 1
        End If
        PB1.value = 26
        Set rs = Nothing

End If
 Set rs = Nothing
 
 'Adeudados Moneda Extranjera
 Set rs = objDAnexoRiesgos.obtenerAdeudadosAnx7(CDate(lsFecha), 2)
PB1.value = 27

If Not rs.EOF Or Not rs.BOF Then
    
        Set rs = Nothing
        PB1.value = 28
        Set rs = oCtaIf.GetCaleAdeudadosXTramos(Format(lsFecha, "yyyymmdd"), "2")
        PB1.value = 29
        lblAvance.Caption = "Consolidando Adeudados..."
        PB1.value = 30
        If Not rs.EOF Or rs.BOF Then
            objDAnexoRiesgos.insertarAdeudadosAnx7 CDate(lsFecha), rs, sMovNro, 2
        End If
        PB1.value = 31
        Set rs = Nothing
    
Else
        Set rs = Nothing
        PB1.value = 28
        Set rs = oCtaIf.GetCaleAdeudadosXTramos(Format(lsFecha, "yyyymmdd"), "2")
        PB1.value = 29
        lblAvance.Caption = "Consolidando Adeudados..."
        PB1.value = 30
        If Not rs.EOF Or rs.BOF Then
            objDAnexoRiesgos.insertarAdeudadosAnx7 CDate(lsFecha), rs, sMovNro, 2
        End If
        PB1.value = 31
        Set rs = Nothing

End If
 Set rs = Nothing
 
 PB1.value = 32
lblAvance.Caption = "Consolidado Finalizado"
Me.MousePointer = vbDefault
lblAvance.Caption = ""
PB1.value = 33
PB1.Visible = False
MsgBox "Se ha Consolidado la data del mes de" + cboMes.Text + " del " + txtAnio.Text, vbInformation
End Sub

Private Sub cmdGenerar_Click()
    If ValidaDatos Then
           Exit Sub
    End If
    'MIOL 20121130, SEGUN RQ12349 **************************
    'generarAnexo
    If Me.optAntiguoNuevo.Item(0).value = True Then
        generarAnexo
    ElseIf Me.optAntiguoNuevo.Item(1).value = True Then
        generarAnexoNuevo
    End If
    'END MIOL **********************************************
End Sub

Private Sub generarAnexo()
    Me.MousePointer = vbHourglass
        Dim sPathAnexo7 As String
       
        Dim fs As New Scripting.FileSystemObject
        Dim obj_excel As Object, Libro As Object, Hoja As Object
        
        On Error GoTo error_sub
          
        PB1.Min = 0
        PB1.Max = 30
        PB1.value = 0
        PB1.Visible = True
        sPathAnexo7 = App.path & "\Spooler\ANEXO_7_" + Format(lsFecha, "yyyymmdd") + ".xls"
        
        If fs.FileExists(sPathAnexo7) Then
            
            If ArchivoEstaAbierto(sPathAnexo7) Then
                If MsgBox("Debe Cerrar el Archivo:" + fs.GetFileName(sPathAnexo7) + " para continuar", vbRetryCancel) = vbCancel Then
                   Me.MousePointer = vbDefault
                   Exit Sub
                End If
                Me.MousePointer = vbHourglass
            End If
    
            fs.DeleteFile sPathAnexo7, True
        End If
              
        
        lblAvance.Caption = "Abriendo archivo a copiar"
        'sPathAnexo7 = App.path & "\Spooler\ANEXO_7_PLANTILLA.xls"
        sPathAnexo7 = App.path & "\FormatoCarta\ANEXO_7_PLANTILLA.xls"
        'lblAvance.Caption = "cargando archivo " + sPathAnexo7
        lblAvance.Caption = "cargando archivo..."
        If Len(Dir(sPathAnexo7)) = 0 Then
           MsgBox "No se Pudo Encontrar el Archivo:" & sPathAnexo7, vbCritical
           Me.MousePointer = vbDefault
           lblAvance.Caption = ""
            PB1.Visible = False
           Exit Sub
        End If
        
        Set obj_excel = CreateObject("Excel.Application")
        obj_excel.DisplayAlerts = False
        Set Libro = obj_excel.Workbooks.Open(sPathAnexo7)
        Set Hoja = Libro.ActiveSheet
        
        Dim celda As Excel.Range
        Dim oCtaCont As DbalanceCont
        Dim rsCtaCont As ADODB.Recordset
        
        Set oCtaCont = New DbalanceCont
        Set rsCtaCont = New ADODB.Recordset
       ' Fecha del ANEXO
        Set celda = obj_excel.Range("A4")
        celda.value = obtenerFechaLarga
         PB1.value = 1
        '****************************PATRIMONIO****************************************
         lblAvance.Caption = "Cargando Datos..."
        Set celda = obj_excel.Range("MN!S8")
        celda.value = dPatrimonio
        Set celda = obj_excel.Range("MN!S9")
        celda.value = Me.txtPatriEfec
        Set celda = obj_excel.Range("ME!S8")
        celda.value = dPatrimonio
        Set celda = obj_excel.Range("ME!S9")
        celda.value = Me.txtPatriEfec
        Set celda = obj_excel.Range("ME!T9")
        celda.value = Me.txtTipCambio
        '*****************************ACTIVOS*****************************************
        '(1)Activo Disponible MN:1111,1113,1116,1118;ME:1121,1123,1126,1128
         lblAvance.Caption = "Cargando Datos Activos..."
        Set celda = obj_excel.Range("MN!C11")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("1111,1113,1116,1118", CDate(lsFecha), "1", Me.txtTipCambio.Text) + (oCtaCont.ObtenerCtaContBalanceMensual("1112", CDate(lsFecha), "1", Me.txtTipCambio.Text) / 2)
        Set celda = obj_excel.Range("ME!C11")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("1121,1123,1126,1128", CDate(lsFecha), "2", Me.txtTipCambio.Text) + (oCtaCont.ObtenerCtaContBalanceMensual("1122", CDate(lsFecha), "2", Me.txtTipCambio.Text) / 2)
         PB1.value = 2
        
        '(2)Activo Disponible MN:1117;ME:1127
        Set celda = obj_excel.Range("MN!E11")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("1117", CDate(lsFecha), "1", Me.txtTipCambio.Text)
         Set celda = obj_excel.Range("ME!E11")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("1127", CDate(lsFecha), "2", Me.txtTipCambio.Text)
         PB1.value = 3
        '(3)Activo Disponible MN:1112;ME:1122
        Set celda = obj_excel.Range("MN!I11")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("1112", CDate(lsFecha), "1", Me.txtTipCambio.Text) / 2
        Set celda = obj_excel.Range("ME!I11")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("1122", CDate(lsFecha), "2", Me.txtTipCambio.Text) / 2
         PB1.value = 4
        '(4)Activo Creditos Otros MN:(1415,1416)+Ref;ME:(1425,1426)+Ref
        cargarActivosOtros oCtaCont.ObtenerCtaContBalanceMensual("1415,1416", CDate(lsFecha), "1", Me.txtTipCambio.Text), oCtaCont.ObtenerCtaContBalanceMensual("1425,1426", CDate(lsFecha), "2", Me.txtTipCambio.Text), obj_excel
         PB1.value = 5
        '(5)Activo Creditos Vigentes
         cargarActivosVigentes obj_excel
         PB1.value = 6
        '(6)Activo Creditos MN:1418+(1419 Repartido);ME:1428+(1429 Repartido)
         cargarActivosCreditos oCtaCont.ObtenerCtaContBalanceMensual("1418", CDate(lsFecha), "1", Me.txtTipCambio.Text), oCtaCont.ObtenerCtaContBalanceMensual("1428", CDate(lsFecha), "2", Me.txtTipCambio.Text), oCtaCont.ObtenerCtaContBalanceMensual("1419", CDate(lsFecha), "1", Me.txtTipCambio.Text), oCtaCont.ObtenerCtaContBalanceMensual("1429", CDate(lsFecha), "2", Me.txtTipCambio.Text), obj_excel
         PB1.value = 7
         'Solo para ME:13000
         Set celda = obj_excel.Range("ME!C13")
         celda.value = oCtaCont.ObtenerCtaContBalanceMensual("13", CDate(lsFecha), "2", Me.txtTipCambio.Text)
        
        '(7)Activo Cuentas por Cobrar MN:1500;ME:1500
        Set celda = obj_excel.Range("MN!I17")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("15", CDate(lsFecha), "1", Me.txtTipCambio.Text)
         Set celda = obj_excel.Range("ME!I17")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("15", CDate(lsFecha), "2", Me.txtTipCambio.Text)
         PB1.value = 8
        '(8)Activo Inversiones Permanentes MN:1700;ME:1700
        Set celda = obj_excel.Range("MN!P18")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("17", CDate(lsFecha), "1", Me.txtTipCambio.Text)
         Set celda = obj_excel.Range("ME!P18")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("17", CDate(lsFecha), "2", Me.txtTipCambio.Text)
         PB1.value = 9
        '(9)Activo Otras Cuentas Activas MN:16120102-16190201002+1900;ME:16220102-16290201002+1900
        Set celda = obj_excel.Range("MN!I23")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("16120102", CDate(lsFecha), "1", Me.txtTipCambio.Text) - Abs(oCtaCont.ObtenerCtaContBalanceMensual("161902010102", CDate(lsFecha), "1", Me.txtTipCambio.Text)) + oCtaCont.ObtenerCtaContBalanceMensual("19", CDate(lsFecha), "1", Me.txtTipCambio.Text)
         Set celda = obj_excel.Range("ME!I23")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("16220102", CDate(lsFecha), "2", Me.txtTipCambio.Text) - Abs(oCtaCont.ObtenerCtaContBalanceMensual("162902010102", CDate(lsFecha), "2", Me.txtTipCambio.Text)) + oCtaCont.ObtenerCtaContBalanceMensual("19", CDate(lsFecha), "2", Me.txtTipCambio.Text)
         PB1.value = 10
        '(10)Activo Otras Cuentas Activas MN:16-16120102-161902010102;ME:16-16220102-162902010102
        Set celda = obj_excel.Range("MN!J23")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("16", CDate(lsFecha), "1", Me.txtTipCambio.Text) - (Abs(oCtaCont.ObtenerCtaContBalanceMensual("16120102", CDate(lsFecha), "1", Me.txtTipCambio.Text)) - Abs(oCtaCont.ObtenerCtaContBalanceMensual("161902010102", CDate(lsFecha), "1", Me.txtTipCambio.Text)))
'         Set celda = obj_Excel.Range("ME!J23")
'        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("16", CDate(lsFecha), "2", Me.txtTipCambio.Text) - (Abs(oCtaCont.ObtenerCtaContBalanceMensual("16220102", CDate(lsFecha), "2", Me.txtTipCambio.Text)) - Abs(oCtaCont.ObtenerCtaContBalanceMensual("162902010102", CDate(lsFecha), "2", Me.txtTipCambio.Text)))
         PB1.value = 11
        '(12)Activo Otras Cuentas Activas MN:1811+1812-18190201;ME:1821+1822-18290201
        Set celda = obj_excel.Range("MN!P23")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("1811", CDate(lsFecha), "1", Me.txtTipCambio.Text) + oCtaCont.ObtenerCtaContBalanceMensual("1812", CDate(lsFecha), "1", Me.txtTipCambio.Text) - Abs(oCtaCont.ObtenerCtaContBalanceMensual("18190201", CDate(lsFecha), "1", Me.txtTipCambio.Text))
         Set celda = obj_excel.Range("ME!P23")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("1821", CDate(lsFecha), "2", Me.txtTipCambio.Text) + oCtaCont.ObtenerCtaContBalanceMensual("1822", CDate(lsFecha), "2", Me.txtTipCambio.Text) - Abs(oCtaCont.ObtenerCtaContBalanceMensual("18290201", CDate(lsFecha), "2", Me.txtTipCambio.Text))
         PB1.value = 12
        '(11)Activo Otras Cuentas Activas MN:18-(1811+1812-18190201);ME:18-(1821+1822-18290201)
        Set celda = obj_excel.Range("MN!K23")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("18", CDate(lsFecha), "1", Me.txtTipCambio.Text) - Abs(obj_excel.Range("MN!P23"))
         Set celda = obj_excel.Range("ME!K23")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("18", CDate(lsFecha), "2", Me.txtTipCambio.Text) - Abs(obj_excel.Range("ME!P23"))
         PB1.value = 13
        '*****************************PASIVOS*****************************************
        lblAvance.Caption = "Cargando Datos Pasivos..."
        '(15)Pasivos  Obligaciones a la Visita MN:2111;ME:2121
        Set celda = obj_excel.Range("MN!C28")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2111", CDate(lsFecha), "1", Me.txtTipCambio.Text)
         Set celda = obj_excel.Range("ME!C28")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2121", CDate(lsFecha), "2", Me.txtTipCambio.Text)
         PB1.value = 14
        '(16)Pasivos  Obligaciones con el Publico MN:2112;ME:2122
        Set celda = obj_excel.Range("MN!I29")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2112", CDate(lsFecha), "1", Me.txtTipCambio.Text)
         Set celda = obj_excel.Range("ME!I29")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2122", CDate(lsFecha), "2", Me.txtTipCambio.Text)
         PB1.value = 15
        '(17) y (18)Pasivos  Obligaciones por Cuenta de Ahorro (17)MN:PF-PFOtras+211305 ME:PF-PFOtras+212305 (17)MN:PF-PFOtras-2117 ME:PF-PFOtras-2127
        cargarPasivosCuentaPlazos oCtaCont.ObtenerCtaContBalanceMensual("211305", CDate(lsFecha), "1", Me.txtTipCambio.Text), oCtaCont.ObtenerCtaContBalanceMensual("212305", CDate(lsFecha), "2", Me.txtTipCambio.Text), oCtaCont.ObtenerCtaContBalanceMensual("2117", CDate(lsFecha), "1", Me.txtTipCambio.Text), oCtaCont.ObtenerCtaContBalanceMensual("2127", CDate(lsFecha), "2", Me.txtTipCambio.Text), obj_excel
         PB1.value = 16
        '(13)Pasivos  Obligaciones con el Publico MN:2114,2116,2118;ME:2124,2126,2128
        Set celda = obj_excel.Range("MN!C27")
        celda.value = obj_excel.Range("MN!C28") + obj_excel.Range("MN!C29") + obj_excel.Range("MN!C30") + oCtaCont.ObtenerCtaContBalanceMensual("2114,2116,2118", CDate(lsFecha), "1", Me.txtTipCambio.Text)
         Set celda = obj_excel.Range("ME!C27")
        celda.value = obj_excel.Range("ME!C28") + obj_excel.Range("ME!C29") + obj_excel.Range("ME!C30") + oCtaCont.ObtenerCtaContBalanceMensual("2124,2126,2128", CDate(lsFecha), "2", Me.txtTipCambio.Text)
         PB1.value = 17
        '(14)Pasivos  Obligaciones con el Publico MN:2117;ME:2127
        Set celda = obj_excel.Range("MN!I27")
        celda.value = obj_excel.Range("MN!I28") + obj_excel.Range("MN!I29") + obj_excel.Range("MN!I30") + oCtaCont.ObtenerCtaContBalanceMensual("2117", CDate(lsFecha), "1", Me.txtTipCambio.Text)
         Set celda = obj_excel.Range("ME!I27")
        celda.value = obj_excel.Range("ME!I28") + obj_excel.Range("ME!I29") + obj_excel.Range("ME!I30") + oCtaCont.ObtenerCtaContBalanceMensual("2127", CDate(lsFecha), "2", Me.txtTipCambio.Text)
         PB1.value = 18
        '(13)y (14)Pasivos  Obligaciones con el Publico Repartido
        cargarPasivosObligacionesPublico obj_excel
         PB1.value = 19
        '(19)Pasivos  Deposito del Sistema Financiero y OFI MN:PFOtras+2318 ME:PFOtras+2328
        Set celda = obj_excel.Range("MN!C32")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2318", CDate(lsFecha), "1", Me.txtTipCambio.Text) + obj_excel.Range("MN!C32")
        Set celda = obj_excel.Range("ME!C32")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2328", CDate(lsFecha), "2", Me.txtTipCambio.Text) + obj_excel.Range("ME!C32")
         PB1.value = 20
        '(20)Pasivos  Deposito del Sistema Financiero y OFI MN:PFOtras+2312 ME:PFOtras+2322
        Set celda = obj_excel.Range("MN!I32")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2312", CDate(lsFecha), "1", Me.txtTipCambio.Text) + obj_excel.Range("MN!I32")
        Set celda = obj_excel.Range("ME!I32")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2322", CDate(lsFecha), "2", Me.txtTipCambio.Text) + obj_excel.Range("ME!I32")
        PB1.value = 21
        '(21)Pasivos  Adeudados y Otras Actividades Financieras
        cargarCalecAdeudados obj_excel
        
        '(22)Pasivos  Adeudados y Otras Actividades Financieras MN:MN!J33+2616 ME:MN!J33+2626
        Set celda = obj_excel.Range("MN!J33")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2616", CDate(lsFecha), "1", Me.txtTipCambio.Text) + obj_excel.Range("MN!J33")
        Set celda = obj_excel.Range("ME!J33")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("2626", CDate(lsFecha), "2", Me.txtTipCambio.Text) + obj_excel.Range("ME!J33")
        PB1.value = 23
        '(23)Pasivos  Cuentas por Pagar MN:2500;ME:2500
        Set celda = obj_excel.Range("MN!D34")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("25", CDate(lsFecha), "1", Me.txtTipCambio.Text)
         Set celda = obj_excel.Range("ME!D34")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("25", CDate(lsFecha), "2", Me.txtTipCambio.Text)
        PB1.value = 24
    
        '(24)Pasivos  Otras Cuentas Pasivas MN:2900;ME:2900
        Set celda = obj_excel.Range("MN!E39")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("29", CDate(lsFecha), "1", Me.txtTipCambio.Text)
         Set celda = obj_excel.Range("ME!E39")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("29", CDate(lsFecha), "2", Me.txtTipCambio.Text)
        PB1.value = 25
        '(25)Pasivos  Otras Cuentas Pasivas MN:2700;ME:2700
        Set celda = obj_excel.Range("MN!F39")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("27", CDate(lsFecha), "1", Me.txtTipCambio.Text)
         Set celda = obj_excel.Range("ME!H39")
        celda.value = oCtaCont.ObtenerCtaContBalanceMensual("27", CDate(lsFecha), "2", Me.txtTipCambio.Text)
        PB1.value = 26
        'verifica si existe el archivo
        lblAvance.Caption = "Verificando Archivo..."
        sPathAnexo7 = App.path & "\Spooler\ANEXO_7_" + Format(lsFecha, "yyyymmdd") + ".xls"
        If fs.FileExists(sPathAnexo7) Then
            
            If ArchivoEstaAbierto(sPathAnexo7) Then
                MsgBox "Debe Cerrar el Archivo:" + fs.GetFileName(sPathAnexo7)
            End If
    '            Exit Sub
            'Set Libro = obj_Excel.Workbooks.Add
            fs.DeleteFile sPathAnexo7, True
        End If
        PB1.value = 27
        'guarda el archivo
        lblAvance.Caption = "Guardando Archivo..."
        Hoja.SaveAs sPathAnexo7
        'obj_Excel.Visible = True
        Libro.Close
        obj_excel.Quit
        PB1.value = 28
        Set Hoja = Nothing
        Set Libro = Nothing
        Set obj_excel = Nothing
        Me.MousePointer = vbDefault
        PB1.value = 29
            'abre y muestra el archivo
        lblAvance.Caption = "Abriendo Archivo..."
        Dim m_excel As New Excel.Application
        m_excel.Workbooks.Open (sPathAnexo7)
        m_excel.Visible = True
        PB1.value = 30
        PB1.Visible = False
        lblAvance.Caption = ""
Exit Sub
error_sub:
        MsgBox TextErr(Err.Description), vbInformation, "Aviso"
        Set Libro = Nothing
        Set obj_excel = Nothing
        Set Hoja = Nothing
        PB1.Visible = False
        lblAvance.Caption = ""
        Me.MousePointer = vbDefault
    
End Sub

'MIOL 20121130, SEGUN RQ12349 ********************************************
Private Sub generarAnexoNuevo()
    Me.MousePointer = vbHourglass
        Dim sPathAnexo7 As String
       
        Dim fs As New Scripting.FileSystemObject
        Dim obj_excel As Object, Libro As Object, Hoja As Object
        
        Dim convert As Double
        
        On Error GoTo error_sub
          
        PB1.Min = 0
        PB1.Max = 35
        PB1.value = 0
        PB1.Visible = True
        sPathAnexo7 = App.path & "\Spooler\NUEVO_ANEXO_7_" + Format(lsFecha, "yyyymmdd") + ".xls"
        
        If fs.FileExists(sPathAnexo7) Then
            
            If ArchivoEstaAbierto(sPathAnexo7) Then
                If MsgBox("Debe Cerrar el Archivo:" + fs.GetFileName(sPathAnexo7) + " para continuar", vbRetryCancel) = vbCancel Then
                   Me.MousePointer = vbDefault
                   Exit Sub
                End If
                Me.MousePointer = vbHourglass
            End If
    
            fs.DeleteFile sPathAnexo7, True
        End If
              
        
        lblAvance.Caption = "Abriendo archivo a copiar"
        sPathAnexo7 = App.path & "\FormatoCarta\PlantillaNuevoAnexo7.xls"

        lblAvance.Caption = "cargando archivo..."
        If Len(Dir(sPathAnexo7)) = 0 Then
           MsgBox "No se Pudo Encontrar el Archivo:" & sPathAnexo7, vbCritical
           Me.MousePointer = vbDefault
           lblAvance.Caption = ""
            PB1.Visible = False
           Exit Sub
        End If
        
        Set obj_excel = CreateObject("Excel.Application")
        obj_excel.DisplayAlerts = False
        Set Libro = obj_excel.Workbooks.Open(sPathAnexo7)
        Set Hoja = Libro.ActiveSheet
        
        Dim celda As Excel.Range
        Dim oCtaCont As DbalanceCont
        Dim rsCtaCont As ADODB.Recordset
        
        Set oCtaCont = New DbalanceCont
        Set rsCtaCont = New ADODB.Recordset
       ' Fecha del ANEXO
        Set celda = obj_excel.Range("ANEXO7AMN!B5")
        celda.value = obtenerFechaLarga
        Set celda = obj_excel.Range("ANEXO7AME!B5")
        celda.value = obtenerFechaLarga
        Set celda = obj_excel.Range("ANEXONRO7BMN!A5")
        celda.value = obtenerFechaLarga
        Set celda = obj_excel.Range("ANEXONRO7BME!A5")
        celda.value = obtenerFechaLarga
         PB1.value = 1
         '****************************PATRIMONIO****************************************
         lblAvance.Caption = "Cargando Datos..."
         Set celda = obj_excel.Range("ANEXO7AMN!M9")
         celda.value = Me.txtPatriEfec
         Set celda = obj_excel.Range("ANEXO7AMN!M10")
         celda.value = Me.txtTipCambio
         Set celda = obj_excel.Range("SECCIONCC1!D20")
         celda.value = Me.txtUNA
        '*********************************ACTIVOS**************************************
        
        'PASI20160127**
        '*****************************DISPONIBLES**************************************
        CargaActivosDisponiblesANX07 obj_excel
        PB1.value = 2
        '*************************INV. DISP. PARA LA VENTA*****************************
        CargaActivosInversionesDispxVentANX07 obj_excel
        PB1.value = 3
        '*************************CTAS X COBRAR SENSIBLES******************************
        CargaActivosCtasxCobrarSensiblesANX07 obj_excel
        PB1.value = 4
        'END PASI**
        
        '****************************CREDITOS VIGENTES*********************************
         cargarActivosVigentesANX7 obj_excel
         PB1.value = 5
         '****************************INTERES DEVENGADO********************************
         Set celda = obj_excel.Range("ANEXONRO7BMN!B22")
         convert = oCtaCont.ObtenerCtaContBalanceMensual("1418", CDate(lsFecha), "1", Me.txtTipCambio.Text)
         celda.value = convert
         PB1.value = 6
         Set celda = obj_excel.Range("ANEXONRO7BME!B22")
         convert = oCtaCont.ObtenerCtaContBalanceMensual("1428", CDate(lsFecha), "2", Me.txtTipCambio.Text)
         celda.value = convert
         PB1.value = 7
        '*********************************PASIVOS**************************************
        '****************************OBLIGACIONES X CTA AHORROS************************
         cargarPasivosObligxCtaAhorrosANX7 obj_excel, 1
         PB1.value = 8
         PB1.value = 9
         cargarPasivosObligxCtaAhorrosANX7 obj_excel, 2
         PB1.value = 10
        '****************************CTA PLAZO FIJO *********************************
         PB1.value = 11
         cargarPasivosCtaPlazoFijoANX7 obj_excel, 1
         PB1.value = 12
         cargarPasivosCtaPlazoFijoANX7 obj_excel, 2
         PB1.value = 13
        '****************************CTS*********************************************
         PB1.value = 14
         cargarPasivosCTSANX7 obj_excel, 1
         PB1.value = 15
         PB1.value = 16
         cargarPasivosCTSANX7 obj_excel, 2
         PB1.value = 17
        '*****************GASTOS POR PAGAR OBLIG CON EL PUBLICO *********************
         PB1.value = 18
         Set celda = obj_excel.Range("ANEXONRO7BMN!B38")
         convert = oCtaCont.ObtenerCtaContBalanceMensual("211803", CDate(lsFecha), "1", Me.txtTipCambio.Text)
         celda.value = convert
         PB1.value = 19
         Set celda = obj_excel.Range("ANEXONRO7BME!B38")
         convert = oCtaCont.ObtenerCtaContBalanceMensual("212803", CDate(lsFecha), "2", Me.txtTipCambio.Text)
         celda.value = convert
        '****************************DEPOSITOS IFIs**********************************
         PB1.value = 20
         cargarPasivosDepositosIFIsANX7 obj_excel, 1
         PB1.value = 21
         PB1.value = 22
         cargarPasivosDepositosIFIsANX7 obj_excel, 2
         PB1.value = 23
        '****************************DEPOSITOS PLAZO FIJO IFIs***********************
         PB1.value = 24
         cargarPasivosDepositosPFIFIsANX7 obj_excel, 1
         PB1.value = 25
         PB1.value = 26
         cargarPasivosDepositosPFIFIsANX7 obj_excel, 2
         PB1.value = 27
        '****************************GASTOS POR PAGAR *******************************
         PB1.value = 28
         Set celda = obj_excel.Range("ANEXONRO7BMN!B43")
         convert = oCtaCont.ObtenerCtaContBalanceMensual("2318", CDate(lsFecha), "1", Me.txtTipCambio.Text)
         celda.value = convert
         PB1.value = 29
         Set celda = obj_excel.Range("ANEXONRO7BME!B43")
         convert = oCtaCont.ObtenerCtaContBalanceMensual("2328", CDate(lsFecha), "2", Me.txtTipCambio.Text)
         celda.value = convert
         PB1.value = 30
         PB1.value = 31
        
        'verifica si existe el archivo
        lblAvance.Caption = "Verificando Archivo..."
        sPathAnexo7 = App.path & "\Spooler\NUEVOANEXO_7_" + Format(lsFecha, "yyyymmdd") + ".xls"
        If fs.FileExists(sPathAnexo7) Then
            
            If ArchivoEstaAbierto(sPathAnexo7) Then
                MsgBox "Debe Cerrar el Archivo:" + fs.GetFileName(sPathAnexo7)
            End If
    '            Exit Sub
            'Set Libro = obj_Excel.Workbooks.Add
            fs.DeleteFile sPathAnexo7, True
        End If
        PB1.value = 32
        'guarda el archivo
        lblAvance.Caption = "Guardando Archivo..."
        Hoja.SaveAs sPathAnexo7
        'obj_Excel.Visible = True
        Libro.Close
        obj_excel.Quit
        PB1.value = 33
        Set Hoja = Nothing
        Set Libro = Nothing
        Set obj_excel = Nothing
        Me.MousePointer = vbDefault
        PB1.value = 34
            'abre y muestra el archivo
        lblAvance.Caption = "Abriendo Archivo..."
        Dim m_excel As New Excel.Application
        m_excel.Workbooks.Open (sPathAnexo7)
        m_excel.Visible = True
        PB1.value = 35
        PB1.Visible = False
        lblAvance.Caption = ""
Exit Sub
error_sub:
        MsgBox TextErr(Err.Description), vbInformation, "Aviso"
        Set Libro = Nothing
        Set obj_excel = Nothing
        Set Hoja = Nothing
        PB1.Visible = False
        lblAvance.Caption = ""
        Me.MousePointer = vbDefault
    
End Sub

Private Sub cargarActivosVigentesANX7(ByVal pobj_Excel As Excel.Application)
     
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim objDAnexoRiesgos As DAnexoRiesgos
    Dim prs As ADODB.Recordset
    Dim nMoneda As Integer
    Dim nTabla As Integer
    Set objDAnexoRiesgos = New DAnexoRiesgos
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
    nMoneda = 0
 
        Set prs = oCtaIf.GetCreditosVigentesSBSanx7(Format(lsFecha, "yyyymmdd"))
        nMoneda = prs!cMoneda
        If Not prs.EOF Or prs.BOF Then
            Do While Not prs.EOF
                If prs!cMoneda = 1 Then
                    Select Case prs!cTpoCredCod
                        Case 1:
                                'CORPORATIVO
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!B18")
                                pcelda.value = IIf(nMoneda = 1, prs(2), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!C18")
                                pcelda.value = IIf(nMoneda = 1, prs(3), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!D18")
                                pcelda.value = IIf(nMoneda = 1, prs(4), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!E18")
                                pcelda.value = IIf(nMoneda = 1, prs(5), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!F18")
                                pcelda.value = IIf(nMoneda = 1, prs(6), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!G18")
                                pcelda.value = IIf(nMoneda = 1, prs(7), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!H18")
                                pcelda.value = IIf(nMoneda = 1, prs(8), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!I18")
                                pcelda.value = IIf(nMoneda = 1, prs(9), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!J18")
                                pcelda.value = IIf(nMoneda = 1, prs(10), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!K18")
                                pcelda.value = IIf(nMoneda = 1, prs(11), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!L18")
                                pcelda.value = IIf(nMoneda = 1, prs(12), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!M18")
                                pcelda.value = IIf(nMoneda = 1, prs(13), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!N18")
                                pcelda.value = IIf(nMoneda = 1, prs(14), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!O18")
                                pcelda.value = IIf(nMoneda = 1, prs(15), 0)
                        'PASI20160127
                        Case 2:
                                'GRANDES EMPRESAS
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!B19")
                                pcelda.value = IIf(nMoneda = 1, prs(2), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!C19")
                                pcelda.value = IIf(nMoneda = 1, prs(3), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!D19")
                                pcelda.value = IIf(nMoneda = 1, prs(4), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!E19")
                                pcelda.value = IIf(nMoneda = 1, prs(5), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!F19")
                                pcelda.value = IIf(nMoneda = 1, prs(6), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!G19")
                                pcelda.value = IIf(nMoneda = 1, prs(7), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!H19")
                                pcelda.value = IIf(nMoneda = 1, prs(8), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!I19")
                                pcelda.value = IIf(nMoneda = 1, prs(9), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!J19")
                                pcelda.value = IIf(nMoneda = 1, prs(10), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!K19")
                                pcelda.value = IIf(nMoneda = 1, prs(11), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!L19")
                                pcelda.value = IIf(nMoneda = 1, prs(12), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!M19")
                                pcelda.value = IIf(nMoneda = 1, prs(13), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!N19")
                                pcelda.value = IIf(nMoneda = 1, prs(14), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!O19")
                                pcelda.value = IIf(nMoneda = 1, prs(15), 0)
                        'end PASI
                        
                        Case 3:
                                'MEDIANA EMPRESA
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!B20")
                                pcelda.value = IIf(nMoneda = 1, prs(2), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!C20")
                                pcelda.value = IIf(nMoneda = 1, prs(3), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!D20")
                                pcelda.value = IIf(nMoneda = 1, prs(4), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!E20")
                                pcelda.value = IIf(nMoneda = 1, prs(5), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!F20")
                                pcelda.value = IIf(nMoneda = 1, prs(6), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!G20")
                                pcelda.value = IIf(nMoneda = 1, prs(7), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!H20")
                                pcelda.value = IIf(nMoneda = 1, prs(8), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!I20")
                                pcelda.value = IIf(nMoneda = 1, prs(9), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!J20")
                                pcelda.value = IIf(nMoneda = 1, prs(10), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!K20")
                                pcelda.value = IIf(nMoneda = 1, prs(11), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!L20")
                                pcelda.value = IIf(nMoneda = 1, prs(12), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!M20")
                                pcelda.value = IIf(nMoneda = 1, prs(13), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!N20")
                                pcelda.value = IIf(nMoneda = 1, prs(14), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!O20")
                                pcelda.value = IIf(nMoneda = 1, prs(15), 0)
                        Case 4:
                                'PEQUEÑA EMPRESA
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!B21")
                                pcelda.value = IIf(nMoneda = 1, prs(2), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!C21")
                                pcelda.value = IIf(nMoneda = 1, prs(3), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!D21")
                                pcelda.value = IIf(nMoneda = 1, prs(4), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!E21")
                                pcelda.value = IIf(nMoneda = 1, prs(5), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!F21")
                                pcelda.value = IIf(nMoneda = 1, prs(6), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!G21")
                                pcelda.value = IIf(nMoneda = 1, prs(7), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!H21")
                                pcelda.value = IIf(nMoneda = 1, prs(8), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!I21")
                                pcelda.value = IIf(nMoneda = 1, prs(9), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!J21")
                                pcelda.value = IIf(nMoneda = 1, prs(10), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!K21")
                                pcelda.value = IIf(nMoneda = 1, prs(11), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!L21")
                                pcelda.value = IIf(nMoneda = 1, prs(12), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!M21")
                                pcelda.value = IIf(nMoneda = 1, prs(13), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!N21")
                                pcelda.value = IIf(nMoneda = 1, prs(14), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!O21")
                                pcelda.value = IIf(nMoneda = 1, prs(15), 0)
                        Case 5:
                                'MICROEMPRESA
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!B14")
                                pcelda.value = IIf(nMoneda = 1, prs(2), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!C14")
                                pcelda.value = IIf(nMoneda = 1, prs(3), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!D14")
                                pcelda.value = IIf(nMoneda = 1, prs(4), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!E14")
                                pcelda.value = IIf(nMoneda = 1, prs(5), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!F14")
                                pcelda.value = IIf(nMoneda = 1, prs(6), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!G14")
                                pcelda.value = IIf(nMoneda = 1, prs(7), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!H14")
                                pcelda.value = IIf(nMoneda = 1, prs(8), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!I14")
                                pcelda.value = IIf(nMoneda = 1, prs(9), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!J14")
                                pcelda.value = IIf(nMoneda = 1, prs(10), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!K14")
                                pcelda.value = IIf(nMoneda = 1, prs(11), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!L14")
                                pcelda.value = IIf(nMoneda = 1, prs(12), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!M14")
                                pcelda.value = IIf(nMoneda = 1, prs(13), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!N14")
                                pcelda.value = IIf(nMoneda = 1, prs(14), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!O14")
                                pcelda.value = IIf(nMoneda = 1, prs(15), 0)
                        Case 7:
                                'CONSUMO
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!B15")
                                pcelda.value = IIf(nMoneda = 1, prs(2), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!C15")
                                pcelda.value = IIf(nMoneda = 1, prs(3), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!D15")
                                pcelda.value = IIf(nMoneda = 1, prs(4), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!E15")
                                pcelda.value = IIf(nMoneda = 1, prs(5), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!F15")
                                pcelda.value = IIf(nMoneda = 1, prs(6), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!G15")
                                pcelda.value = IIf(nMoneda = 1, prs(7), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!H15")
                                pcelda.value = IIf(nMoneda = 1, prs(8), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!I15")
                                pcelda.value = IIf(nMoneda = 1, prs(9), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!J15")
                                pcelda.value = IIf(nMoneda = 1, prs(10), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!K15")
                                pcelda.value = IIf(nMoneda = 1, prs(11), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!L15")
                                pcelda.value = IIf(nMoneda = 1, prs(12), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!M15")
                                pcelda.value = IIf(nMoneda = 1, prs(13), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!N15")
                                pcelda.value = IIf(nMoneda = 1, prs(14), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!O15")
                                pcelda.value = IIf(nMoneda = 1, prs(15), 0)
                        Case 8:
                                'HIPOTECARIO
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!B16")
                                pcelda.value = IIf(nMoneda = 1, prs(2), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!C16")
                                pcelda.value = IIf(nMoneda = 1, prs(3), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!D16")
                                pcelda.value = IIf(nMoneda = 1, prs(4), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!E16")
                                pcelda.value = IIf(nMoneda = 1, prs(5), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!F16")
                                pcelda.value = IIf(nMoneda = 1, prs(6), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!G16")
                                pcelda.value = IIf(nMoneda = 1, prs(7), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!H16")
                                pcelda.value = IIf(nMoneda = 1, prs(8), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!I16")
                                pcelda.value = IIf(nMoneda = 1, prs(9), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!J16")
                                pcelda.value = IIf(nMoneda = 1, prs(10), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!K16")
                                pcelda.value = IIf(nMoneda = 1, prs(11), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!L16")
                                pcelda.value = IIf(nMoneda = 1, prs(12), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!M16")
                                pcelda.value = IIf(nMoneda = 1, prs(13), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!N16")
                                pcelda.value = IIf(nMoneda = 1, prs(14), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!O16")
                                pcelda.value = IIf(nMoneda = 1, prs(15), 0)
                    End Select
                ElseIf prs!cMoneda = 2 Then
                    Select Case prs!cTpoCredCod
                        Case 1:
                               'EMPRESAS SIST.FINANCIERO NAC.
                               Set pcelda = pobj_Excel.Range("ANEXONRO7BME!B17")
                               pcelda.value = IIf(nMoneda = 1, prs(2), 0)
                               Set pcelda = pobj_Excel.Range("ANEXONRO7BME!C17")
                               pcelda.value = IIf(nMoneda = 1, prs(3), 0)
                               Set pcelda = pobj_Excel.Range("ANEXONRO7BME!D17")
                               pcelda.value = IIf(nMoneda = 1, prs(4), 0)
                               Set pcelda = pobj_Excel.Range("ANEXONRO7BME!E17")
                               pcelda.value = IIf(nMoneda = 1, prs(5), 0)
                               Set pcelda = pobj_Excel.Range("ANEXONRO7BME!F17")
                               pcelda.value = IIf(nMoneda = 1, prs(6), 0)
                               Set pcelda = pobj_Excel.Range("ANEXONRO7BME!G17")
                               pcelda.value = IIf(nMoneda = 1, prs(7), 0)
                               Set pcelda = pobj_Excel.Range("ANEXONRO7BME!H17")
                               pcelda.value = IIf(nMoneda = 1, prs(8), 0)
                               Set pcelda = pobj_Excel.Range("ANEXONRO7BME!I17")
                               pcelda.value = IIf(nMoneda = 1, prs(9), 0)
                               Set pcelda = pobj_Excel.Range("ANEXONRO7BME!J17")
                               pcelda.value = IIf(nMoneda = 1, prs(10), 0)
                               Set pcelda = pobj_Excel.Range("ANEXONRO7BME!K17")
                               pcelda.value = IIf(nMoneda = 1, prs(11), 0)
                               Set pcelda = pobj_Excel.Range("ANEXONRO7BME!L17")
                               pcelda.value = IIf(nMoneda = 1, prs(12), 0)
                               Set pcelda = pobj_Excel.Range("ANEXONRO7BME!M17")
                               pcelda.value = IIf(nMoneda = 1, prs(13), 0)
                               Set pcelda = pobj_Excel.Range("ANEXONRO7BME!N17")
                               pcelda.value = IIf(nMoneda = 1, prs(14), 0)
                               Set pcelda = pobj_Excel.Range("ANEXONRO7BME!O17")
                               pcelda.value = IIf(nMoneda = 1, prs(15), 0)
                               '*********Agregado by NAGL 20170927***************
                        Case 3:
                                'MEDIANA EMPRESA
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!B20")
                                pcelda.value = IIf(nMoneda = 1, prs(2), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!C20")
                                pcelda.value = IIf(nMoneda = 1, prs(3), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!D20")
                                pcelda.value = IIf(nMoneda = 1, prs(4), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!E20")
                                pcelda.value = IIf(nMoneda = 1, prs(5), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!F20")
                                pcelda.value = IIf(nMoneda = 1, prs(6), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!G20")
                                pcelda.value = IIf(nMoneda = 1, prs(7), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!H20")
                                pcelda.value = IIf(nMoneda = 1, prs(8), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!I20")
                                pcelda.value = IIf(nMoneda = 1, prs(9), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!J20")
                                pcelda.value = IIf(nMoneda = 1, prs(10), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!K20")
                                pcelda.value = IIf(nMoneda = 1, prs(11), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!L20")
                                pcelda.value = IIf(nMoneda = 1, prs(12), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!M20")
                                pcelda.value = IIf(nMoneda = 1, prs(13), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!N20")
                                pcelda.value = IIf(nMoneda = 1, prs(14), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!O20")
                                pcelda.value = IIf(nMoneda = 1, prs(15), 0)
                        Case 4:
                                'PEQUEÑA EMPRESA
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!B21")
                                pcelda.value = IIf(nMoneda = 1, prs(2), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!C21")
                                pcelda.value = IIf(nMoneda = 1, prs(3), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!D21")
                                pcelda.value = IIf(nMoneda = 1, prs(4), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!E21")
                                pcelda.value = IIf(nMoneda = 1, prs(5), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!F21")
                                pcelda.value = IIf(nMoneda = 1, prs(6), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!G21")
                                pcelda.value = IIf(nMoneda = 1, prs(7), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!H21")
                                pcelda.value = IIf(nMoneda = 1, prs(8), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!I21")
                                pcelda.value = IIf(nMoneda = 1, prs(9), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!J21")
                                pcelda.value = IIf(nMoneda = 1, prs(10), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!K21")
                                pcelda.value = IIf(nMoneda = 1, prs(11), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!L21")
                                pcelda.value = IIf(nMoneda = 1, prs(12), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!M21")
                                pcelda.value = IIf(nMoneda = 1, prs(13), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!N21")
                                pcelda.value = IIf(nMoneda = 1, prs(14), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!O21")
                                pcelda.value = IIf(nMoneda = 1, prs(15), 0)
                        Case 5:
                                'MICROEMPRESA
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!B14")
                                pcelda.value = IIf(nMoneda = 1, prs(2), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!C14")
                                pcelda.value = IIf(nMoneda = 1, prs(3), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!D14")
                                pcelda.value = IIf(nMoneda = 1, prs(4), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!E14")
                                pcelda.value = IIf(nMoneda = 1, prs(5), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!F14")
                                pcelda.value = IIf(nMoneda = 1, prs(6), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!G14")
                                pcelda.value = IIf(nMoneda = 1, prs(7), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!H14")
                                pcelda.value = IIf(nMoneda = 1, prs(8), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!I14")
                                pcelda.value = IIf(nMoneda = 1, prs(9), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!J14")
                                pcelda.value = IIf(nMoneda = 1, prs(10), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!K14")
                                pcelda.value = IIf(nMoneda = 1, prs(11), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!L14")
                                pcelda.value = IIf(nMoneda = 1, prs(12), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!M14")
                                pcelda.value = IIf(nMoneda = 1, prs(13), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!N14")
                                pcelda.value = IIf(nMoneda = 1, prs(14), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!O14")
                                pcelda.value = IIf(nMoneda = 1, prs(15), 0)
                        Case 7:
                                'CONSUMO
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!B15")
                                pcelda.value = IIf(nMoneda = 1, prs(2), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!C15")
                                pcelda.value = IIf(nMoneda = 1, prs(3), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!D15")
                                pcelda.value = IIf(nMoneda = 1, prs(4), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!E15")
                                pcelda.value = IIf(nMoneda = 1, prs(5), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!F15")
                                pcelda.value = IIf(nMoneda = 1, prs(6), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!G15")
                                pcelda.value = IIf(nMoneda = 1, prs(7), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!H15")
                                pcelda.value = IIf(nMoneda = 1, prs(8), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!I15")
                                pcelda.value = IIf(nMoneda = 1, prs(9), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!J15")
                                pcelda.value = IIf(nMoneda = 1, prs(10), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!K15")
                                pcelda.value = IIf(nMoneda = 1, prs(11), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!L15")
                                pcelda.value = IIf(nMoneda = 1, prs(12), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!M15")
                                pcelda.value = IIf(nMoneda = 1, prs(13), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!N15")
                                pcelda.value = IIf(nMoneda = 1, prs(14), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!O15")
                                pcelda.value = IIf(nMoneda = 1, prs(15), 0)
                        Case 8:
                                'HIPOTECARIO
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!B16")
                                pcelda.value = IIf(nMoneda = 1, prs(2), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!C16")
                                pcelda.value = IIf(nMoneda = 1, prs(3), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!D16")
                                pcelda.value = IIf(nMoneda = 1, prs(4), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!E16")
                                pcelda.value = IIf(nMoneda = 1, prs(5), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!F16")
                                pcelda.value = IIf(nMoneda = 1, prs(6), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!G16")
                                pcelda.value = IIf(nMoneda = 1, prs(7), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!H16")
                                pcelda.value = IIf(nMoneda = 1, prs(8), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!I16")
                                pcelda.value = IIf(nMoneda = 1, prs(9), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!J16")
                                pcelda.value = IIf(nMoneda = 1, prs(10), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!K16")
                                pcelda.value = IIf(nMoneda = 1, prs(11), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!L16")
                                pcelda.value = IIf(nMoneda = 1, prs(12), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!M16")
                                pcelda.value = IIf(nMoneda = 1, prs(13), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!N16")
                                pcelda.value = IIf(nMoneda = 1, prs(14), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!O16")
                                pcelda.value = IIf(nMoneda = 1, prs(15), 0)
                    End Select
                End If
                prs.MoveNext
            Loop
        End If
End Sub

Private Sub cargarPasivosCtaPlazoFijoANX7(ByVal pobj_Excel As Excel.Application, ByVal pcMoneda As String)
     
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim objDAnexoRiesgos As DAnexoRiesgos
    Dim prs As ADODB.Recordset
    Dim cMoneda As String
    Dim nTabla As Integer
    Set objDAnexoRiesgos = New DAnexoRiesgos
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
     
        Set prs = oCtaIf.GetCtaPlazoFijoSBSanx7(Format(lsFecha, "yyyymmdd"), pcMoneda)
     
        If Not prs.EOF Or prs.BOF Then
                If pcMoneda = 1 Then
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!B34")
                                pcelda.value = IIf(pcMoneda = 1, prs(0), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!C34")
                                pcelda.value = IIf(pcMoneda = 1, prs(1), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!D34")
                                pcelda.value = IIf(pcMoneda = 1, prs(2), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!E34")
                                pcelda.value = IIf(pcMoneda = 1, prs(3), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!F34")
                                pcelda.value = IIf(pcMoneda = 1, prs(4), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!G34")
                                pcelda.value = IIf(pcMoneda = 1, prs(5), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!H34")
                                pcelda.value = IIf(pcMoneda = 1, prs(6), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!I34")
                                pcelda.value = IIf(pcMoneda = 1, prs(7), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!J34")
                                pcelda.value = IIf(pcMoneda = 1, prs(8), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!K34")
                                pcelda.value = IIf(pcMoneda = 1, prs(9), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!L34")
                                pcelda.value = IIf(pcMoneda = 1, prs(10), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!M34")
                                pcelda.value = IIf(pcMoneda = 1, prs(11), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!N34")
                                pcelda.value = IIf(pcMoneda = 1, prs(12), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!O34")
                                pcelda.value = IIf(pcMoneda = 1, prs(13), 0)
                ElseIf pcMoneda = 2 Then
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!B34")
                                pcelda.value = IIf(pcMoneda = 2, prs(0), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!C34")
                                pcelda.value = IIf(pcMoneda = 2, prs(1), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!D34")
                                pcelda.value = IIf(pcMoneda = 2, prs(2), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!E34")
                                pcelda.value = IIf(pcMoneda = 2, prs(3), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!F34")
                                pcelda.value = IIf(pcMoneda = 2, prs(4), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!G34")
                                pcelda.value = IIf(pcMoneda = 2, prs(5), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!H34")
                                pcelda.value = IIf(pcMoneda = 2, prs(6), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!I34")
                                pcelda.value = IIf(pcMoneda = 2, prs(7), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!J34")
                                pcelda.value = IIf(pcMoneda = 2, prs(8), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!K34")
                                pcelda.value = IIf(pcMoneda = 2, prs(9), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!L34")
                                pcelda.value = IIf(pcMoneda = 2, prs(10), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!M34")
                                pcelda.value = IIf(pcMoneda = 2, prs(11), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!N34")
                                pcelda.value = IIf(pcMoneda = 2, prs(11), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!O34")
                                pcelda.value = IIf(pcMoneda = 2, prs(13), 0)
                End If
        End If
End Sub
Private Sub cargarPasivosObligxCtaAhorrosANX7(ByVal pobj_Excel As Excel.Application, ByVal cMoneda As String)
     
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim objDAnexoRiesgos As DAnexoRiesgos
    Dim prs As ADODB.Recordset
    Dim nFilas As Integer
    Dim nTabla As Integer
    Set objDAnexoRiesgos = New DAnexoRiesgos
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
 
        Set prs = oCtaIf.GetObligxCtaAhorrosSBSanx7(Format(lsFecha, "yyyymmdd"), cMoneda)
        nFilas = 2
        If Not prs.EOF Or prs.BOF Then
            If cMoneda = 1 Then
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("ObligxCtaAhoMN!A" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("ObligxCtaAhoMN!B" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(1), 0)
                    nFilas = nFilas + 1
                    prs.MoveNext
                Loop
            ElseIf cMoneda = 2 Then
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("ObligxCtaAhoME!A" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("ObligxCtaAhoME!B" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(1), 0)
                    nFilas = nFilas + 1
                    prs.MoveNext
                Loop
            End If
        End If
        Set prs = Nothing
End Sub

Private Sub cargarPasivosDepositosIFIsANX7(ByVal pobj_Excel As Excel.Application, ByVal cMoneda As String)
     
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim objDAnexoRiesgos As DAnexoRiesgos
    Dim prs As ADODB.Recordset
    Dim nFilas As Integer
    Dim nTabla As Integer
    Set objDAnexoRiesgos = New DAnexoRiesgos
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
 
        Set prs = oCtaIf.GetDepositosIFIsSBSanx7(Format(lsFecha, "yyyymmdd"), cMoneda)
        nFilas = 2
        If Not prs.EOF Or prs.BOF Then
            If cMoneda = 1 Then
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("DepositosIFIsMN!A" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("DepositosIFIsMN!B" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(1), 0)
                    nFilas = nFilas + 1
                    prs.MoveNext
                Loop
            ElseIf cMoneda = 2 Then
                Do While Not prs.EOF
                    Set pcelda = pobj_Excel.Range("DepositosIFIsME!A" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("DepositosIFIsME!B" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(1), 0)
                    nFilas = nFilas + 1
                    prs.MoveNext
                Loop
            End If
        End If
    Set prs = Nothing
End Sub

Private Sub cargarPasivosDepositosPFIFIsANX7(ByVal pobj_Excel As Excel.Application, ByVal pcMoneda As String)
     
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim objDAnexoRiesgos As DAnexoRiesgos
    Dim prs As ADODB.Recordset
    Dim nFilas As Integer
    Dim nTabla As Integer
    Set objDAnexoRiesgos = New DAnexoRiesgos
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
 
        Set prs = oCtaIf.GetDepositosPFIFIsSBSanx7(Format(lsFecha, "yyyymmdd"), pcMoneda)
       
        If Not prs.EOF Or prs.BOF Then
                If pcMoneda = 1 Then
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!B42")
                                pcelda.value = IIf(pcMoneda = 1, prs(0), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!C42")
                                pcelda.value = IIf(pcMoneda = 1, prs(1), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!D42")
                                pcelda.value = IIf(pcMoneda = 1, prs(2), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!E42")
                                pcelda.value = IIf(pcMoneda = 1, prs(3), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!F42")
                                pcelda.value = IIf(pcMoneda = 1, prs(4), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!G42")
                                pcelda.value = IIf(pcMoneda = 1, prs(5), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!H42")
                                pcelda.value = IIf(pcMoneda = 1, prs(6), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!I42")
                                pcelda.value = IIf(pcMoneda = 1, prs(7), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!J42")
                                pcelda.value = IIf(pcMoneda = 1, prs(8), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!K42")
                                pcelda.value = IIf(pcMoneda = 1, prs(9), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!L42")
                                pcelda.value = IIf(pcMoneda = 1, prs(10), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!M42")
                                pcelda.value = IIf(pcMoneda = 1, prs(11), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!N42")
                                pcelda.value = IIf(pcMoneda = 1, prs(12), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!O42")
                                pcelda.value = IIf(pcMoneda = 1, prs(13), 0)
                ElseIf pcMoneda = 2 Then
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!B42")
                                pcelda.value = IIf(pcMoneda = 2, prs(0), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!C42")
                                pcelda.value = IIf(pcMoneda = 2, prs(1), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!D42")
                                pcelda.value = IIf(pcMoneda = 2, prs(2), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!E42")
                                pcelda.value = IIf(pcMoneda = 2, prs(3), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!F42")
                                pcelda.value = IIf(pcMoneda = 2, prs(4), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!G42")
                                pcelda.value = IIf(pcMoneda = 2, prs(5), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!H42")
                                pcelda.value = IIf(pcMoneda = 2, prs(6), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!I42")
                                pcelda.value = IIf(pcMoneda = 2, prs(7), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!J42")
                                pcelda.value = IIf(pcMoneda = 2, prs(8), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!K42")
                                pcelda.value = IIf(pcMoneda = 2, prs(9), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!L42")
                                pcelda.value = IIf(pcMoneda = 2, prs(10), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!M42")
                                pcelda.value = IIf(pcMoneda = 2, prs(11), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!N42")
                                pcelda.value = IIf(pcMoneda = 2, prs(11), 0)
                                Set pcelda = pobj_Excel.Range("ANEXONRO7BME!O42")
                                pcelda.value = IIf(pcMoneda = 2, prs(13), 0)
                End If
        End If
End Sub

Private Sub cargarPasivosCTSANX7(ByVal pobj_Excel As Excel.Application, ByVal cMoneda As String)
     
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim objDAnexoRiesgos As DAnexoRiesgos
    Dim prs As ADODB.Recordset
    Dim nFilas As Integer
    Dim nTabla As Integer
    Set objDAnexoRiesgos = New DAnexoRiesgos
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
 
        Set prs = oCtaIf.GetCTSSBSanx7(Format(lsFecha, "yyyymmdd"), cMoneda)
        nFilas = 2
        If Not prs.EOF Or prs.BOF Then
            If cMoneda = 1 Then
                Do While Not prs.EOF
                    'Máximo Retiro Esperado
                    Set pcelda = pobj_Excel.Range("CTSDISPMN!A" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("CTSDISPMN!B" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(1), 0)
                    'Saldo Restringido
                    Set pcelda = pobj_Excel.Range("CTSRESTMN!A" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("CTSRESTMN!B" & nFilas)
                    pcelda.value = IIf(cMoneda = 1, prs(2), 0)
                    nFilas = nFilas + 1
                    prs.MoveNext
                Loop
            ElseIf cMoneda = 2 Then
                Do While Not prs.EOF
                    'Máximo Retiro Esperado
                    Set pcelda = pobj_Excel.Range("CTSDISPME!A" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("CTSDISPME!B" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(1), 0)
                    'Saldo Restringido
                    Set pcelda = pobj_Excel.Range("CTSRESTME!A" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(0), 0)
                    Set pcelda = pobj_Excel.Range("CTSRESTME!B" & nFilas)
                    pcelda.value = IIf(cMoneda = 2, prs(2), 0)
                    nFilas = nFilas + 1
                    prs.MoveNext
                Loop
            End If
        End If
    Set prs = Nothing
End Sub
'END MIOL ****************************************************************
Private Sub cargarActivosCreditos(ByVal pCtaContMN1415 As Currency, ByVal pCtaContME1415 As Currency, ByVal pCtaContMN1419 As Currency, ByVal pCtaContME1419 As Currency, ByVal pobj_Excel As Excel.Application)
     
    Dim pcelda As Excel.Range
      
       'MONEDA NACIONAL
            Set pcelda = pobj_Excel.Range("MN!C14")
            pcelda.value = pobj_Excel.Range("MN!C15") + pobj_Excel.Range("MN!C16") + pCtaContMN1415 + (pCtaContMN1419 * 0.03)
            Set pcelda = pobj_Excel.Range("MN!D14")
            pcelda.value = pobj_Excel.Range("MN!D15") + pobj_Excel.Range("MN!D16") + (pCtaContMN1419 * 0.02)
            Set pcelda = pobj_Excel.Range("MN!E14")
            pcelda.value = pobj_Excel.Range("MN!E15") + pobj_Excel.Range("MN!E16") + (pCtaContMN1419 * 0.07)
            Set pcelda = pobj_Excel.Range("MN!F14")
            pcelda.value = pobj_Excel.Range("MN!F15") + pobj_Excel.Range("MN!F16") + (pCtaContMN1419 * 0.16)
            Set pcelda = pobj_Excel.Range("MN!G14")
            pcelda.value = pobj_Excel.Range("MN!G15") + pobj_Excel.Range("MN!G16") + (pCtaContMN1419 * 0.09)
            Set pcelda = pobj_Excel.Range("MN!H14")
            pcelda.value = pobj_Excel.Range("MN!H15") + pobj_Excel.Range("MN!H16") + (pCtaContMN1419 * 0.12)
            Set pcelda = pobj_Excel.Range("MN!I14")
            pcelda.value = pobj_Excel.Range("MN!I15") + pobj_Excel.Range("MN!I16") + (pCtaContMN1419 * 0.51)
            Set pcelda = pobj_Excel.Range("MN!J14")
            pcelda.value = pobj_Excel.Range("MN!J15") + pobj_Excel.Range("MN!J16")
            Set pcelda = pobj_Excel.Range("MN!K14")
            pcelda.value = pobj_Excel.Range("MN!K15") + pobj_Excel.Range("MN!K16")
            Set pcelda = pobj_Excel.Range("MN!L14")
            pcelda.value = pobj_Excel.Range("MN!L15") + pobj_Excel.Range("MN!L16")
            Set pcelda = pobj_Excel.Range("MN!M14")
            pcelda.value = pobj_Excel.Range("MN!M15") + pobj_Excel.Range("MN!M16")
            Set pcelda = pobj_Excel.Range("MN!N14")
            pcelda.value = pobj_Excel.Range("MN!N15") + pobj_Excel.Range("MN!N16")
            Set pcelda = pobj_Excel.Range("MN!O14")
            pcelda.value = pobj_Excel.Range("MN!O15") + pobj_Excel.Range("MN!O16")
        'MONEDA EXTRAJERA
            Set pcelda = pobj_Excel.Range("ME!C14")
            pcelda.value = pobj_Excel.Range("ME!C15") + pobj_Excel.Range("ME!C16") + pCtaContME1415 + (pCtaContME1419 * 0.03)
            Set pcelda = pobj_Excel.Range("ME!D14")
            pcelda.value = pobj_Excel.Range("ME!D15") + pobj_Excel.Range("ME!D16") + (pCtaContME1419 * 0.02)
            Set pcelda = pobj_Excel.Range("ME!E14")
            pcelda.value = pobj_Excel.Range("ME!E15") + pobj_Excel.Range("ME!E16") + (pCtaContME1419 * 0.07)
            Set pcelda = pobj_Excel.Range("ME!F14")
            pcelda.value = pobj_Excel.Range("ME!F15") + pobj_Excel.Range("ME!F16") + (pCtaContME1419 * 0.16)
            Set pcelda = pobj_Excel.Range("ME!G14")
            pcelda.value = pobj_Excel.Range("ME!G15") + pobj_Excel.Range("ME!G16") + (pCtaContME1419 * 0.09)
            Set pcelda = pobj_Excel.Range("ME!H14")
            pcelda.value = pobj_Excel.Range("ME!H15") + pobj_Excel.Range("ME!H16") + (pCtaContME1419 * 0.12)
            Set pcelda = pobj_Excel.Range("ME!I14")
            pcelda.value = pobj_Excel.Range("ME!I15") + pobj_Excel.Range("ME!I16") + (pCtaContME1419 * 0.51)
            Set pcelda = pobj_Excel.Range("ME!J14")
            pcelda.value = pobj_Excel.Range("ME!J15") + pobj_Excel.Range("ME!J16")
            Set pcelda = pobj_Excel.Range("ME!K14")
            pcelda.value = pobj_Excel.Range("ME!K15") + pobj_Excel.Range("ME!K16")
            Set pcelda = pobj_Excel.Range("ME!L14")
            pcelda.value = pobj_Excel.Range("ME!L15") + pobj_Excel.Range("ME!L16")
            Set pcelda = pobj_Excel.Range("ME!M14")
            pcelda.value = pobj_Excel.Range("ME!M15") + pobj_Excel.Range("ME!M16")
            Set pcelda = pobj_Excel.Range("ME!N14")
            pcelda.value = pobj_Excel.Range("ME!N15") + pobj_Excel.Range("ME!N16")
            Set pcelda = pobj_Excel.Range("ME!O14")
            pcelda.value = pobj_Excel.Range("ME!O15") + pobj_Excel.Range("ME!O16")
  
End Sub
Private Sub cargarActivosVigentes(ByVal pobj_Excel As Excel.Application)
     
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim objDAnexoRiesgos As DAnexoRiesgos
    Dim prs As ADODB.Recordset
    Dim nMoneda As Integer
    Dim nTabla As Integer
    Set objDAnexoRiesgos = New DAnexoRiesgos
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
    nMoneda = 0
    Set prs = objDAnexoRiesgos.obtenerCreditosVigentesAnx7(CDate(lsFecha))
    If prs.EOF Or prs.BOF Then
        nTabla = 2
        Set prs = oCtaIf.GetCreditosVigentes(Format(lsFecha, "yyyymmdd"))
        If Not prs.EOF Or prs.BOF Then
            nMoneda = prs(prs.Fields.Count - 1)
        End If
    Else
        nTabla = 1
        nMoneda = prs!nMoneda
    End If
    
    'nMoneda = 0
'    If Not prs.EOF Or prs.BOF Then
'        nMoneda = prs(prs.Fields.Count - 1)
'    End If
     'MONEDA NACIONAL
            Set pcelda = pobj_Excel.Range("MN!C15")
            pcelda.value = IIf(nMoneda = 1, prs(0), 0)
            Set pcelda = pobj_Excel.Range("MN!D15")
            pcelda.value = IIf(nMoneda = 1, prs(1), 0)
            Set pcelda = pobj_Excel.Range("MN!E15")
            pcelda.value = IIf(nMoneda = 1, prs(2), 0)
            Set pcelda = pobj_Excel.Range("MN!F15")
            pcelda.value = IIf(nMoneda = 1, prs(3), 0)
            Set pcelda = pobj_Excel.Range("MN!G15")
            pcelda.value = IIf(nMoneda = 1, prs(4), 0)
            Set pcelda = pobj_Excel.Range("MN!H15")
            pcelda.value = IIf(nMoneda = 1, prs(5), 0)
            Set pcelda = pobj_Excel.Range("MN!I15")
            pcelda.value = IIf(nMoneda = 1, prs(6), 0)
            Set pcelda = pobj_Excel.Range("MN!J15")
            pcelda.value = IIf(nMoneda = 1, prs(7), 0)
            Set pcelda = pobj_Excel.Range("MN!K15")
            pcelda.value = IIf(nMoneda = 1, prs(8), 0)
            Set pcelda = pobj_Excel.Range("MN!L15")
            pcelda.value = IIf(nMoneda = 1, prs(9), 0)
            Set pcelda = pobj_Excel.Range("MN!M15")
            pcelda.value = IIf(nMoneda = 1, prs(10), 0)
            Set pcelda = pobj_Excel.Range("MN!N15")
            pcelda.value = IIf(nMoneda = 1, prs(11), 0)
            Set pcelda = pobj_Excel.Range("MN!O15")
            pcelda.value = IIf(nMoneda = 1, prs(12), 0)
            
        If nMoneda = 1 And nTabla = 1 Then
            prs.MoveNext
             If Not prs.EOF Or prs.BOF Then
                 nMoneda = prs!nMoneda
             End If
        Else
             prs.MoveNext
             If Not prs.EOF Or prs.BOF Then
                nMoneda = prs(prs.Fields.Count - 1)
             End If
        End If
        
            'MONEDA EXTRAJERA
            
            Set pcelda = pobj_Excel.Range("ME!C15")
            pcelda.value = IIf(nMoneda = 2, prs(0), 0)
            Set pcelda = pobj_Excel.Range("ME!D15")
            pcelda.value = IIf(nMoneda = 2, prs(1), 0)
            Set pcelda = pobj_Excel.Range("ME!E15")
            pcelda.value = IIf(nMoneda = 2, prs(2), 0)
            Set pcelda = pobj_Excel.Range("ME!F15")
            pcelda.value = IIf(nMoneda = 2, prs(3), 0)
            Set pcelda = pobj_Excel.Range("ME!G15")
            pcelda.value = IIf(nMoneda = 2, prs(4), 0)
            Set pcelda = pobj_Excel.Range("ME!H15")
            pcelda.value = IIf(nMoneda = 2, prs(5), 0)
            Set pcelda = pobj_Excel.Range("ME!I15")
            pcelda.value = IIf(nMoneda = 2, prs(6), 0)
            Set pcelda = pobj_Excel.Range("ME!J15")
            pcelda.value = IIf(nMoneda = 2, prs(7), 0)
            Set pcelda = pobj_Excel.Range("ME!K15")
            pcelda.value = IIf(nMoneda = 2, prs(8), 0)
            Set pcelda = pobj_Excel.Range("ME!L15")
            pcelda.value = IIf(nMoneda = 2, prs(9), 0)
            Set pcelda = pobj_Excel.Range("ME!M15")
            pcelda.value = IIf(nMoneda = 2, prs(10), 0)
            Set pcelda = pobj_Excel.Range("ME!N15")
            pcelda.value = IIf(nMoneda = 2, prs(11), 0)
            Set pcelda = pobj_Excel.Range("ME!O15")
            pcelda.value = IIf(nMoneda = 2, prs(12), 0)
 
End Sub
Private Sub cargarPasivosObligacionesPublico(ByVal pobj_Excel As Excel.Application)
     
    Dim pcelda As Excel.Range
     'MONEDA NACIONAL
            Set pcelda = pobj_Excel.Range("MN!D27")
            pcelda.value = pobj_Excel.Range("MN!D28") + pobj_Excel.Range("MN!D29") + pobj_Excel.Range("MN!D30")
            Set pcelda = pobj_Excel.Range("MN!E27")
            pcelda.value = pobj_Excel.Range("MN!E28") + pobj_Excel.Range("MN!E29") + pobj_Excel.Range("MN!E30")
            Set pcelda = pobj_Excel.Range("MN!F27")
            pcelda.value = pobj_Excel.Range("MN!F28") + pobj_Excel.Range("MN!F29") + pobj_Excel.Range("MN!F30")
            Set pcelda = pobj_Excel.Range("MN!G27")
            pcelda.value = pobj_Excel.Range("MN!G28") + pobj_Excel.Range("MN!G29") + pobj_Excel.Range("MN!G30")
            Set pcelda = pobj_Excel.Range("MN!H27")
            pcelda.value = pobj_Excel.Range("MN!H28") + pobj_Excel.Range("MN!H29") + pobj_Excel.Range("MN!H30")
            Set pcelda = pobj_Excel.Range("MN!J27")
            pcelda.value = pobj_Excel.Range("MN!J28") + pobj_Excel.Range("MN!J29") + pobj_Excel.Range("MN!J30")
            Set pcelda = pobj_Excel.Range("MN!K27")
            pcelda.value = pobj_Excel.Range("MN!K28") + pobj_Excel.Range("MN!K29") + pobj_Excel.Range("MN!K30")
            Set pcelda = pobj_Excel.Range("MN!L27")
            pcelda.value = pobj_Excel.Range("MN!L28") + pobj_Excel.Range("MN!L29") + pobj_Excel.Range("MN!L30")
            Set pcelda = pobj_Excel.Range("MN!M27")
            pcelda.value = pobj_Excel.Range("MN!M28") + pobj_Excel.Range("MN!M29") + pobj_Excel.Range("MN!M30")
            Set pcelda = pobj_Excel.Range("MN!N27")
            pcelda.value = pobj_Excel.Range("MN!N28") + pobj_Excel.Range("MN!N29") + pobj_Excel.Range("MN!N30")
            Set pcelda = pobj_Excel.Range("MN!O27")
            pcelda.value = pobj_Excel.Range("MN!O28") + pobj_Excel.Range("MN!O29") + pobj_Excel.Range("MN!O30")
            
              
            'MONEDA EXTRAJERA
            
            Set pcelda = pobj_Excel.Range("ME!D27")
            pcelda.value = pobj_Excel.Range("ME!D28") + pobj_Excel.Range("ME!D29") + pobj_Excel.Range("ME!D30")
            Set pcelda = pobj_Excel.Range("ME!E27")
            pcelda.value = pobj_Excel.Range("ME!E28") + pobj_Excel.Range("ME!E29") + pobj_Excel.Range("ME!E30")
            Set pcelda = pobj_Excel.Range("ME!F27")
            pcelda.value = pobj_Excel.Range("ME!F28") + pobj_Excel.Range("ME!F29") + pobj_Excel.Range("ME!F30")
            Set pcelda = pobj_Excel.Range("ME!G27")
            pcelda.value = pobj_Excel.Range("ME!G28") + pobj_Excel.Range("ME!G29") + pobj_Excel.Range("ME!G30")
            Set pcelda = pobj_Excel.Range("ME!H27")
            pcelda.value = pobj_Excel.Range("ME!H28") + pobj_Excel.Range("ME!H29") + pobj_Excel.Range("ME!H30")
            Set pcelda = pobj_Excel.Range("ME!J27")
            pcelda.value = pobj_Excel.Range("ME!J28") + pobj_Excel.Range("ME!J29") + pobj_Excel.Range("ME!J30")
            Set pcelda = pobj_Excel.Range("ME!K27")
            pcelda.value = pobj_Excel.Range("ME!K28") + pobj_Excel.Range("ME!K29") + pobj_Excel.Range("ME!K30")
            Set pcelda = pobj_Excel.Range("ME!L27")
            pcelda.value = pobj_Excel.Range("ME!L28") + pobj_Excel.Range("ME!L29") + pobj_Excel.Range("ME!L30")
            Set pcelda = pobj_Excel.Range("ME!M27")
            pcelda.value = pobj_Excel.Range("ME!M28") + pobj_Excel.Range("ME!M29") + pobj_Excel.Range("ME!M30")
            Set pcelda = pobj_Excel.Range("ME!N27")
            pcelda.value = pobj_Excel.Range("ME!N28") + pobj_Excel.Range("ME!N29") + pobj_Excel.Range("ME!N30")
            Set pcelda = pobj_Excel.Range("ME!O27")
            pcelda.value = pobj_Excel.Range("ME!O28") + pobj_Excel.Range("ME!O29") + pobj_Excel.Range("ME!O30")
 
End Sub
Private Sub cargarActivosOtros(ByVal pCtaContMN As Currency, ByVal pCtaContME As Currency, ByVal pobj_Excel As Excel.Application)
     
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim objDAnexoRiesgos As DAnexoRiesgos
    Dim prs As ADODB.Recordset
    Dim nMoneda As Integer
    Dim nTabla As Integer
    Set objDAnexoRiesgos = New DAnexoRiesgos
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
    nMoneda = 0
    Set prs = objDAnexoRiesgos.obtenerCreditosRefinanciadosAnx7(CDate(lsFecha))
    If prs.EOF Or prs.BOF Then
        nTabla = 2
        Set prs = oCtaIf.GetCreditosRefinanciados(Format(lsFecha, "yyyymmdd"))
        If Not prs.EOF Or prs.BOF Then
            nMoneda = prs(prs.Fields.Count - 1)
        
        End If
    Else
        nTabla = 1
        nMoneda = prs!nMoneda
    End If
    
    'nMoneda = 0
'    If Not prs.EOF Or prs.BOF Then
'        nMoneda = prs(prs.Fields.Count - 1)
'
'    End If
     'MONEDA NACIONAL
            Set pcelda = pobj_Excel.Range("MN!C16")
            pcelda.value = IIf(nMoneda = 1, prs(0), 0) + (pCtaContMN * 0.03)
            Set pcelda = pobj_Excel.Range("MN!D16")
            pcelda.value = IIf(nMoneda = 1, prs(1), 0) + (pCtaContMN * 0.02)
            Set pcelda = pobj_Excel.Range("MN!E16")
            pcelda.value = IIf(nMoneda = 1, prs(2), 0) + (pCtaContMN * 0.07)
            Set pcelda = pobj_Excel.Range("MN!F16")
            pcelda.value = IIf(nMoneda = 1, prs(3), 0) + (pCtaContMN * 0.16)
            Set pcelda = pobj_Excel.Range("MN!G16")
            pcelda.value = IIf(nMoneda = 1, prs(4), 0) + (pCtaContMN * 0.09)
            Set pcelda = pobj_Excel.Range("MN!H16")
            pcelda.value = IIf(nMoneda = 1, prs(5), 0) + (pCtaContMN * 0.12)
            Set pcelda = pobj_Excel.Range("MN!I16")
            pcelda.value = IIf(nMoneda = 1, prs(6), 0) + (pCtaContMN * 0.51)
            Set pcelda = pobj_Excel.Range("MN!J16")
            pcelda.value = IIf(nMoneda = 1, prs(7), 0)
            Set pcelda = pobj_Excel.Range("MN!K16")
            pcelda.value = IIf(nMoneda = 1, prs(8), 0)
            Set pcelda = pobj_Excel.Range("MN!L16")
            pcelda.value = IIf(nMoneda = 1, prs(9), 0)
            Set pcelda = pobj_Excel.Range("MN!M16")
            pcelda.value = IIf(nMoneda = 1, prs(10), 0)
            Set pcelda = pobj_Excel.Range("MN!N16")
            pcelda.value = IIf(nMoneda = 1, prs(11), 0)
            Set pcelda = pobj_Excel.Range("MN!O16")
            pcelda.value = IIf(nMoneda = 1, prs(12), 0)
        
        If nMoneda = 1 And nTabla = 1 Then
            prs.MoveNext
            If Not prs.EOF Or prs.BOF Then
                nMoneda = prs!nMoneda
            End If
        Else
            prs.MoveNext
            If Not prs.EOF Or prs.BOF Then
                nMoneda = prs(prs.Fields.Count - 1)
            End If
        End If
            
            'MONEDA EXTRAJERA
            
            Set pcelda = pobj_Excel.Range("ME!C16")
            pcelda.value = IIf(nMoneda = 2, prs(0), 0) + (pCtaContME * 0.03)
            Set pcelda = pobj_Excel.Range("ME!D16")
            pcelda.value = IIf(nMoneda = 2, prs(1), 0) + (pCtaContME * 0.02)
            Set pcelda = pobj_Excel.Range("ME!E16")
            pcelda.value = IIf(nMoneda = 2, prs(2), 0) + (pCtaContME * 0.07)
            Set pcelda = pobj_Excel.Range("ME!F16")
            pcelda.value = IIf(nMoneda = 2, prs(3), 0) + (pCtaContME * 0.16)
            Set pcelda = pobj_Excel.Range("ME!G16")
            pcelda.value = IIf(nMoneda = 2, prs(4), 0) + (pCtaContME * 0.09)
            Set pcelda = pobj_Excel.Range("ME!H16")
            pcelda.value = IIf(nMoneda = 2, prs(5), 0) + (pCtaContME * 0.12)
            Set pcelda = pobj_Excel.Range("ME!I16")
            pcelda.value = IIf(nMoneda = 2, prs(6), 0) + (pCtaContME * 0.51)
            Set pcelda = pobj_Excel.Range("ME!J16")
            pcelda.value = IIf(nMoneda = 2, prs(7), 0)
            Set pcelda = pobj_Excel.Range("ME!K16")
            pcelda.value = IIf(nMoneda = 2, prs(8), 0)
            Set pcelda = pobj_Excel.Range("ME!L16")
            pcelda.value = IIf(nMoneda = 2, prs(9), 0)
            Set pcelda = pobj_Excel.Range("ME!M16")
            pcelda.value = IIf(nMoneda = 2, prs(10), 0)
            Set pcelda = pobj_Excel.Range("ME!N16")
            pcelda.value = IIf(nMoneda = 2, prs(11), 0)
            Set pcelda = pobj_Excel.Range("ME!O16")
            pcelda.value = IIf(nMoneda = 2, prs(12), 0)
    
End Sub
Private Sub cargarPasivosCuentaPlazos(ByVal pCtaContMN As Currency, ByVal pCtaContME As Currency, ByVal pCtaContMN2 As Currency, ByVal pCtaContME2 As Currency, ByVal pobj_Excel As Excel.Application)
     
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim rsPF As ADODB.Recordset
    Dim rsPFOtros As ADODB.Recordset
    Dim nPFOtrosMoneda As Integer
    Dim nPFMoneda As Integer
    nPFOtrosMoneda = 0
    nPFMoneda = 0
    
    Set rsPF = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
    Set rsPF = oCtaIf.GetPlazoFijoAnexo7(Format(lsFecha, "yyyymmdd"), Me.txtTipCambio.Text)
    Set rsPFOtros = oCtaIf.GetPlazoFijosDeOtros(Format(lsFecha, "yyyymmdd"), Me.txtTipCambio.Text)
    
    If Not (rsPF.EOF Or rsPF.BOF) Then
        nPFMoneda = rsPF!Moneda
    End If
    If Not (rsPFOtros.EOF Or rsPFOtros.BOF) Then
        nPFOtrosMoneda = rsPF!Moneda
    End If
   
       'MONEDA NACIONAL
            Set pcelda = pobj_Excel.Range("MN!C30")
            pcelda.value = IIf(nPFMoneda = 1, rsPF(0), 0) - IIf(nPFOtrosMoneda = 1, rsPFOtros(0), 0) + pCtaContMN
            Set pcelda = pobj_Excel.Range("MN!C32")
            pcelda.value = IIf(nPFOtrosMoneda = 1, rsPFOtros(0), 0)
            
            Set pcelda = pobj_Excel.Range("MN!D30")
            pcelda.value = IIf(nPFMoneda = 1, rsPF(1), 0) - IIf(nPFOtrosMoneda = 1, rsPFOtros(1), 0)
            Set pcelda = pobj_Excel.Range("MN!D32")
            pcelda.value = IIf(nPFOtrosMoneda = 1, rsPFOtros(1), 0)
            
            Set pcelda = pobj_Excel.Range("MN!E30")
            pcelda.value = IIf(nPFMoneda = 1, rsPF(2), 0) - IIf(nPFOtrosMoneda = 1, rsPFOtros(2), 0)
            Set pcelda = pobj_Excel.Range("MN!E32")
            pcelda.value = IIf(nPFOtrosMoneda = 1, rsPFOtros(2), 0)
            
            Set pcelda = pobj_Excel.Range("MN!F30")
            pcelda.value = IIf(nPFMoneda = 1, rsPF(3), 0) - IIf(nPFOtrosMoneda = 1, rsPFOtros(3), 0)
            Set pcelda = pobj_Excel.Range("MN!F32")
            pcelda.value = IIf(nPFOtrosMoneda = 1, rsPFOtros(3), 0)
            
            Set pcelda = pobj_Excel.Range("MN!G30")
            pcelda.value = IIf(nPFMoneda = 1, rsPF(4), 0) - IIf(nPFOtrosMoneda = 1, rsPFOtros(4), 0)
            Set pcelda = pobj_Excel.Range("MN!G32")
            pcelda.value = IIf(nPFOtrosMoneda = 1, rsPFOtros(4), 0)
            
            Set pcelda = pobj_Excel.Range("MN!H30")
            pcelda.value = IIf(nPFMoneda = 1, rsPF(5), 0) - IIf(nPFOtrosMoneda = 1, rsPFOtros(5), 0)
            Set pcelda = pobj_Excel.Range("MN!H32")
            pcelda.value = IIf(nPFOtrosMoneda = 1, rsPFOtros(5), 0)
            
            Set pcelda = pobj_Excel.Range("MN!I30")
            pcelda.value = IIf(nPFMoneda = 1, rsPF(6), 0) - IIf(nPFOtrosMoneda = 1, rsPFOtros(6), 0) - pCtaContMN2
            Set pcelda = pobj_Excel.Range("MN!I32")
            pcelda.value = IIf(nPFOtrosMoneda = 1, rsPFOtros(6), 0)
            
            Set pcelda = pobj_Excel.Range("MN!J30")
            pcelda.value = IIf(nPFMoneda = 1, rsPF(7), 0) - IIf(nPFOtrosMoneda = 1, rsPFOtros(7), 0)
            Set pcelda = pobj_Excel.Range("MN!J32")
            pcelda.value = IIf(nPFOtrosMoneda = 1, rsPFOtros(7), 0)
            
            Set pcelda = pobj_Excel.Range("MN!K30")
            pcelda.value = IIf(nPFMoneda = 1, rsPF(8), 0) - IIf(nPFOtrosMoneda = 1, rsPFOtros(8), 0)
            Set pcelda = pobj_Excel.Range("MN!K32")
            pcelda.value = IIf(nPFOtrosMoneda = 1, rsPFOtros(8), 0)
            
            Set pcelda = pobj_Excel.Range("MN!L30")
            pcelda.value = IIf(nPFMoneda = 1, rsPF(9), 0) - IIf(nPFOtrosMoneda = 1, rsPFOtros(9), 0)
            Set pcelda = pobj_Excel.Range("MN!L32")
            pcelda.value = IIf(nPFOtrosMoneda = 1, rsPFOtros(9), 0)
            
            Set pcelda = pobj_Excel.Range("MN!M30")
            pcelda.value = IIf(nPFMoneda = 1, rsPF(10), 0) - IIf(nPFOtrosMoneda = 1, rsPFOtros(10), 0)
            Set pcelda = pobj_Excel.Range("MN!M32")
            pcelda.value = IIf(nPFOtrosMoneda = 1, rsPFOtros(10), 0)
            
            Set pcelda = pobj_Excel.Range("MN!N30")
            pcelda.value = IIf(nPFMoneda = 1, rsPF(11), 0) - IIf(nPFOtrosMoneda = 1, rsPFOtros(11), 0)
            Set pcelda = pobj_Excel.Range("MN!N32")
            pcelda.value = IIf(nPFOtrosMoneda = 1, rsPFOtros(11), 0)
            
            Set pcelda = pobj_Excel.Range("MN!O30")
            pcelda.value = IIf(nPFMoneda = 1, rsPF(12), 0) - IIf(nPFOtrosMoneda = 1, rsPFOtros(12), 0)
            Set pcelda = pobj_Excel.Range("MN!O32")
            pcelda.value = IIf(nPFOtrosMoneda = 1, rsPFOtros(12), 0)
            
   
            If nPFMoneda = 1 Then
                 rsPF.MoveNext
                 If Not (rsPF.EOF Or rsPF.BOF) Then
                    nPFMoneda = rsPF!Moneda
                End If
            End If
            If nPFOtrosMoneda = 1 Then
                 rsPFOtros.MoveNext
                 If Not (rsPFOtros.EOF Or rsPFOtros.BOF) Then
                    nPFOtrosMoneda = rsPF!Moneda
                End If
            End If
        'MONEDA EXTRANJERA
            Set pcelda = pobj_Excel.Range("ME!C30")
            pcelda.value = IIf(nPFMoneda = 2, rsPF(0), 0) - IIf(nPFOtrosMoneda = 2, rsPFOtros(0), 0) + pCtaContME
            Set pcelda = pobj_Excel.Range("ME!C32")
            pcelda.value = IIf(nPFOtrosMoneda = 2, rsPFOtros(0), 0)
            
            Set pcelda = pobj_Excel.Range("ME!D30")
            pcelda.value = IIf(nPFMoneda = 2, rsPF(1), 0) - IIf(nPFOtrosMoneda = 2, rsPFOtros(1), 0)
            Set pcelda = pobj_Excel.Range("ME!D32")
            pcelda.value = IIf(nPFOtrosMoneda = 2, rsPFOtros(1), 0)
            
            Set pcelda = pobj_Excel.Range("ME!E30")
            pcelda.value = IIf(nPFMoneda = 2, rsPF(2), 0) - IIf(nPFOtrosMoneda = 2, rsPFOtros(2), 0)
            Set pcelda = pobj_Excel.Range("ME!E32")
            pcelda.value = IIf(nPFOtrosMoneda = 2, rsPFOtros(2), 0)
            
            Set pcelda = pobj_Excel.Range("ME!F30")
            pcelda.value = IIf(nPFMoneda = 2, rsPF(3), 0) - IIf(nPFOtrosMoneda = 2, rsPFOtros(3), 0)
            Set pcelda = pobj_Excel.Range("ME!F32")
            pcelda.value = IIf(nPFOtrosMoneda = 2, rsPFOtros(3), 0)
            
            Set pcelda = pobj_Excel.Range("ME!G30")
            pcelda.value = IIf(nPFMoneda = 2, rsPF(4), 0) - IIf(nPFOtrosMoneda = 2, rsPFOtros(4), 0)
            Set pcelda = pobj_Excel.Range("ME!G32")
            pcelda.value = IIf(nPFOtrosMoneda = 2, rsPFOtros(4), 0)
            
            Set pcelda = pobj_Excel.Range("ME!H30")
            pcelda.value = IIf(nPFMoneda = 2, rsPF(5), 0) - IIf(nPFOtrosMoneda = 2, rsPFOtros(5), 0)
            Set pcelda = pobj_Excel.Range("ME!H32")
            pcelda.value = IIf(nPFOtrosMoneda = 2, rsPFOtros(5), 0)
            
            Set pcelda = pobj_Excel.Range("ME!I30")
            pcelda.value = IIf(nPFMoneda = 2, rsPF(6), 0) - IIf(nPFOtrosMoneda = 2, rsPFOtros(6), 0) - pCtaContME2
            Set pcelda = pobj_Excel.Range("ME!I32")
            pcelda.value = IIf(nPFOtrosMoneda = 2, rsPFOtros(6), 0)
            
            Set pcelda = pobj_Excel.Range("ME!J30")
            pcelda.value = IIf(nPFMoneda = 2, rsPF(7), 0) - IIf(nPFOtrosMoneda = 2, rsPFOtros(7), 0)
            Set pcelda = pobj_Excel.Range("ME!J32")
            pcelda.value = IIf(nPFOtrosMoneda = 2, rsPFOtros(7), 0)
            
            Set pcelda = pobj_Excel.Range("ME!K30")
            pcelda.value = IIf(nPFMoneda = 2, rsPF(8), 0) - IIf(nPFOtrosMoneda = 2, rsPFOtros(8), 0)
            Set pcelda = pobj_Excel.Range("ME!K32")
            pcelda.value = IIf(nPFOtrosMoneda = 2, rsPFOtros(8), 0)
            
            Set pcelda = pobj_Excel.Range("ME!L30")
            pcelda.value = IIf(nPFMoneda = 2, rsPF(9), 0) - IIf(nPFOtrosMoneda = 2, rsPFOtros(9), 0)
            Set pcelda = pobj_Excel.Range("ME!L32")
            pcelda.value = IIf(nPFOtrosMoneda = 2, rsPFOtros(9), 0)
            
            Set pcelda = pobj_Excel.Range("ME!M30")
            pcelda.value = IIf(nPFMoneda = 2, rsPF(10), 0) - IIf(nPFOtrosMoneda = 2, rsPFOtros(10), 0)
            Set pcelda = pobj_Excel.Range("ME!M32")
            pcelda.value = IIf(nPFOtrosMoneda = 2, rsPFOtros(10), 0)
            
            Set pcelda = pobj_Excel.Range("ME!N30")
            pcelda.value = IIf(nPFMoneda = 2, rsPF(11), 0) - IIf(nPFOtrosMoneda = 2, rsPFOtros(11), 0)
            Set pcelda = pobj_Excel.Range("ME!N32")
            pcelda.value = IIf(nPFOtrosMoneda = 2, rsPFOtros(11), 0)
            
            Set pcelda = pobj_Excel.Range("ME!O30")
            pcelda.value = IIf(nPFMoneda = 2, rsPF(12), 0) - IIf(nPFOtrosMoneda = 2, rsPFOtros(12), 0)
            Set pcelda = pobj_Excel.Range("ME!O32")
            pcelda.value = IIf(nPFOtrosMoneda = 2, rsPFOtros(12), 0)
       
  
End Sub
Private Sub cargarCalecAdeudados(ByVal pobj_Excel As Excel.Application)
     
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim prs As ADODB.Recordset
    Dim objDAnexoRiesgos As DAnexoRiesgos
    
    Set objDAnexoRiesgos = New DAnexoRiesgos
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
    
    
    Set prs = objDAnexoRiesgos.obtenerAdeudadosAnx7(CDate(lsFecha), 1)
    If Not prs.EOF Or prs.BOF Then
        Set prs = oCtaIf.GetCaleAdeudadosXTramos(Format(lsFecha, "yyyymmdd"), "1")
    End If
    
    
    If Not prs.EOF Or prs.BOF Then
       
     'MONEDA NACIONAL
            Set pcelda = pobj_Excel.Range("MN!C33")
            pcelda.value = prs(0)
            Set pcelda = pobj_Excel.Range("MN!D33")
            pcelda.value = prs(1)
            Set pcelda = pobj_Excel.Range("MN!E33")
            pcelda.value = prs(2)
            Set pcelda = pobj_Excel.Range("MN!F33")
            pcelda.value = prs(3)
            Set pcelda = pobj_Excel.Range("MN!G33")
            pcelda.value = prs(4)
            Set pcelda = pobj_Excel.Range("MN!H33")
            pcelda.value = prs(5)
            Set pcelda = pobj_Excel.Range("MN!I33")
            pcelda.value = prs(6)
            Set pcelda = pobj_Excel.Range("MN!J33")
            pcelda.value = prs(7)
            Set pcelda = pobj_Excel.Range("MN!K33")
            pcelda.value = prs(8)
            Set pcelda = pobj_Excel.Range("MN!L33")
            pcelda.value = prs(9)
            Set pcelda = pobj_Excel.Range("MN!M33")
            pcelda.value = prs(10)
            Set pcelda = pobj_Excel.Range("MN!N33")
            pcelda.value = prs(11)
            Set pcelda = pobj_Excel.Range("MN!O33")
            pcelda.value = prs(12)
            
      End If
    
    'Set prs = oCtaIf.GetCaleAdeudadosXTramos(Format(lsFecha, "yyyymmdd"), "2")
    Set prs = objDAnexoRiesgos.obtenerAdeudadosAnx7(CDate(lsFecha), 2)
    If Not prs.EOF Or prs.BOF Then
        Set prs = oCtaIf.GetCaleAdeudadosXTramos(Format(lsFecha, "yyyymmdd"), "2")
    End If
    
    If Not prs.EOF Or prs.BOF Then
            'MONEDA EXTRAJERA
            
            Set pcelda = pobj_Excel.Range("ME!C33")
            pcelda.value = prs(0)
            Set pcelda = pobj_Excel.Range("ME!D33")
            pcelda.value = prs(1)
            Set pcelda = pobj_Excel.Range("ME!E33")
            pcelda.value = prs(2)
            Set pcelda = pobj_Excel.Range("ME!F33")
            pcelda.value = prs(3)
            Set pcelda = pobj_Excel.Range("ME!G33")
            pcelda.value = prs(4)
            Set pcelda = pobj_Excel.Range("ME!H33")
            pcelda.value = prs(5)
            Set pcelda = pobj_Excel.Range("ME!I33")
            pcelda.value = prs(6)
            Set pcelda = pobj_Excel.Range("ME!J33")
            pcelda.value = prs(7)
            Set pcelda = pobj_Excel.Range("ME!K33")
            pcelda.value = prs(8)
            Set pcelda = pobj_Excel.Range("ME!L33")
            pcelda.value = prs(9)
            Set pcelda = pobj_Excel.Range("ME!M33")
            pcelda.value = prs(10)
            Set pcelda = pobj_Excel.Range("ME!N33")
            pcelda.value = prs(11)
            Set pcelda = pobj_Excel.Range("ME!O33")
            pcelda.value = prs(12)
    End If
End Sub
Private Function ValidaDatos() As Boolean
    ValidaDatos = False
    If Val(IIf(Me.txtAnio.Text = "", 0, Me.txtAnio.Text)) > Val(Year(Date)) Or Me.txtAnio.Text = "" Then
        Me.MousePointer = vbDefault
        MsgBox "Ingrese un Año Valido", vbInformation
        ValidaDatos = True
        Exit Function
    End If
    If Me.cboMes.ListIndex = -1 Then
        Me.MousePointer = vbDefault
        MsgBox "Seleccione el Mes", vbInformation
        ValidaDatos = True
        Exit Function
    End If
    If Me.txtTipCambio = "0" Then
        Me.MousePointer = vbDefault
        MsgBox "El Tipo de Cambio Debe ser mayor a cero(0),Ingrese un Periodo Valido", vbInformation
        ValidaDatos = True
        Exit Function
    End If
    If Me.txtPatriEfec = "0" Or Me.txtPatriEfec = "" Then
        Me.MousePointer = vbDefault
        MsgBox "Ingrese el Patrimonio", vbInformation
        ValidaDatos = True
        Exit Function
    End If
End Function

Private Sub cmdSalir_Click()
    Unload Me
End Sub



Private Sub Form_Load()
Me.optAntiguoNuevo.Item(0).value = True
End Sub
Private Sub txtAnio_Change()
    If Len(txtAnio.Text) = 4 Then
        cboMes.SetFocus
    End If
End Sub
Private Function obtenerFechaLarga() As String
Dim sFecha  As Date
If cboMes.ListIndex <> -1 Then
    sFecha = CDate("01/" & Format(cboMes.ListIndex + 1, "00") & "/" & Trim(Me.txtAnio.Text))
    sFecha = DateAdd("m", 1, sFecha)
    sFecha = sFecha - 1
    lsFecha = sFecha
    obtenerFechaLarga = Mid(CStr(sFecha), 1, 2) + " de " + cboMes.Text + " del " + Me.txtAnio.Text
Else
    obtenerFechaLarga = ""
End If 'NAGL 20171002
End Function
Private Function ArchivoEstaAbierto(ByVal Ruta As String) As Boolean
On Error GoTo HayErrores
Dim f As Integer
   f = FreeFile
   Open Ruta For Append As f
   Close f
   ArchivoEstaAbierto = False
   Exit Function
HayErrores:
   If Err.Number = 70 Then
      ArchivoEstaAbierto = True
   Else
      Err.Raise Err.Number
   End If
End Function

Private Sub txtPatriEfec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub

Private Sub txtPatriEfec_LostFocus()
    Me.txtPatriEfec.Text = Format(Me.txtPatriEfec.Text, "##,##0.00")
End Sub

Private Sub txtUNA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub

Private Sub txtUNA_LostFocus()
    Me.txtUNA.Text = Format(Me.txtUNA.Text, "##,##0.00")
End Sub
'PASI20160127**
Private Sub CargaActivosDisponiblesANX07(ByVal pobj_Excel As Excel.Application)
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim prs As ADODB.Recordset
    Dim nTabla As Integer
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
    
    
    Set prs = oCtaIf.CargaActivosDisponiblesANX07(Format(lsFecha, "yyyymmdd"), "1")
    
    If Not prs.EOF Or prs.BOF Then
        Do While Not prs.EOF
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!B10")
            pcelda.value = prs(0)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!C10")
            pcelda.value = prs(1)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!D10")
            pcelda.value = prs(2)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!E10")
            pcelda.value = prs(3)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!F10")
            pcelda.value = prs(4)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!G10")
            pcelda.value = prs(5)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!H10")
            pcelda.value = prs(6)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!I10")
            pcelda.value = prs(7)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!J10")
            pcelda.value = prs(8)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!K10")
            pcelda.value = prs(9)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!L10")
            pcelda.value = prs(10)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!M10")
            pcelda.value = prs(11)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!N10")
            pcelda.value = prs(12)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!O10")
            pcelda.value = prs(13)
            prs.MoveNext
        Loop
    End If
    Set prs = oCtaIf.CargaActivosDisponiblesANX07(Format(lsFecha, "yyyymmdd"), "2")
    
    If Not prs.EOF Or prs.BOF Then
        Do While Not prs.EOF
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!B10")
            pcelda.value = prs(0)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!C10")
            pcelda.value = prs(1)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!D10")
            pcelda.value = prs(2)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!E10")
            pcelda.value = prs(3)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!F10")
            pcelda.value = prs(4)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!G10")
            pcelda.value = prs(5)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!H10")
            pcelda.value = prs(6)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!I10")
            pcelda.value = prs(7)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!J10")
            pcelda.value = prs(8)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!K10")
            pcelda.value = prs(9)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!L10")
            pcelda.value = prs(10)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!M10")
            pcelda.value = prs(11)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!N10")
            pcelda.value = prs(12)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!O10")
            pcelda.value = prs(13)
            prs.MoveNext
        Loop
    End If
End Sub
Private Sub CargaActivosInversionesDispxVentANX07(ByVal pobj_Excel As Excel.Application)
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim prs As ADODB.Recordset
    Dim nTabla As Integer
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
    
    Set prs = oCtaIf.CargaActivosInversionesDispxVentANX07(Format(lsFecha, "yyyymmdd"), "1")
    
    If Not prs.EOF Or prs.BOF Then
        Do While Not prs.EOF
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!B12")
            pcelda.value = prs(0)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!C12")
            pcelda.value = prs(1)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!D12")
            pcelda.value = prs(2)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!E12")
            pcelda.value = prs(3)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!F12")
            pcelda.value = prs(4)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!G12")
            pcelda.value = prs(5)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!H12")
            pcelda.value = prs(6)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!I12")
            pcelda.value = prs(7)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!J12")
            pcelda.value = prs(8)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!K12")
            pcelda.value = prs(9)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!L12")
            pcelda.value = prs(10)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!M12")
            pcelda.value = prs(11)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!N12")
            pcelda.value = prs(12)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!O12")
            pcelda.value = prs(13)
            prs.MoveNext
        Loop
    End If
    Set prs = oCtaIf.CargaActivosInversionesDispxVentANX07(Format(lsFecha, "yyyymmdd"), "2")
    
    If Not prs.EOF Or prs.BOF Then
        Do While Not prs.EOF
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!B12")
            pcelda.value = prs(0)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!C12")
            pcelda.value = prs(1)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!D12")
            pcelda.value = prs(2)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!E12")
            pcelda.value = prs(3)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!F12")
            pcelda.value = prs(4)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!G12")
            pcelda.value = prs(5)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!H12")
            pcelda.value = prs(6)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!I12")
            pcelda.value = prs(7)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!J12")
            pcelda.value = prs(8)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!K12")
            pcelda.value = prs(9)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!L12")
            pcelda.value = prs(10)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!M12")
            pcelda.value = prs(11)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!N12")
            pcelda.value = prs(12)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!O12")
            pcelda.value = prs(13)
            prs.MoveNext
        Loop
    End If
    
End Sub
Private Sub CargaActivosCtasxCobrarSensiblesANX07(ByVal pobj_Excel As Excel.Application)
    Dim pcelda As Excel.Range
    Dim oCtaIf As NCajaCtaIF
    Dim prs As ADODB.Recordset
    Dim nTabla As Integer
    Set prs = New ADODB.Recordset
    Set oCtaIf = New NCajaCtaIF
    
    Set prs = oCtaIf.CargaActivosCtasxCobrarSensiblesANX07(Format(lsFecha, "yyyymmdd"), "1")
    
    If Not prs.EOF Or prs.BOF Then
        Do While Not prs.EOF
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!B23")
            pcelda.value = prs(0)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!C23")
            pcelda.value = prs(1)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!D23")
            pcelda.value = prs(2)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!E23")
            pcelda.value = prs(3)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!F23")
            pcelda.value = prs(4)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!G23")
            pcelda.value = prs(5)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!H23")
            pcelda.value = prs(6)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!I23")
            pcelda.value = prs(7)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!J23")
            pcelda.value = prs(8)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!K23")
            pcelda.value = prs(9)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!L23")
            pcelda.value = prs(10)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!M23")
            pcelda.value = prs(11)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!N23")
            pcelda.value = prs(12)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BMN!O23")
            pcelda.value = prs(13)
            prs.MoveNext
        Loop
    End If
    Set prs = oCtaIf.CargaActivosCtasxCobrarSensiblesANX07(Format(lsFecha, "yyyymmdd"), "2")
    
    If Not prs.EOF Or prs.BOF Then
        Do While Not prs.EOF
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!B23")
            pcelda.value = prs(0)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!C23")
            pcelda.value = prs(1)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!D23")
            pcelda.value = prs(2)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!E23")
            pcelda.value = prs(3)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!F23")
            pcelda.value = prs(4)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!G23")
            pcelda.value = prs(5)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!H23")
            pcelda.value = prs(6)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!I23")
            pcelda.value = prs(7)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!J23")
            pcelda.value = prs(8)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!K23")
            pcelda.value = prs(9)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!L23")
            pcelda.value = prs(10)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!M23")
            pcelda.value = prs(11)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!N23")
            pcelda.value = prs(12)
            Set pcelda = pobj_Excel.Range("ANEXONRO7BME!O23")
            pcelda.value = prs(13)
            prs.MoveNext
        Loop
    End If
End Sub
'END PASI**
