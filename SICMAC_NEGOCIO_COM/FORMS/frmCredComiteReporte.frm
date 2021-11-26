VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCredComiteReporte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Actas de Comite"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   Icon            =   "frmCredComiteReporte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.TextBox txtacta 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton btn_salir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton btn_generar 
         Caption         =   "&Generar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   3120
         Width           =   1095
      End
      Begin VB.ComboBox cboagencia 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1800
         Width           =   3135
      End
      Begin VB.ComboBox cbocomite 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2520
         Width           =   3135
      End
      Begin MSMask.MaskEdBox txNotaIni 
         Height          =   300
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtNotaFin 
         Height          =   300
         Left            =   2640
         TabIndex        =   12
         Top             =   480
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   8
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtdur 
         Height          =   300
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   8
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Acta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   14
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "hh:mm:ss"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "H. Inicio :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Día :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Comite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Agencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmCredComiteReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btn_generar_Click()
  Dim oCOMNCredDoc As COMNCredito.NCOMCredDoc
  Dim MatProductos As String
  Set oCOMNCredDoc = New COMNCredito.NCOMCredDoc
    If Me.txtacta.Text <> "" And Me.txNotaIni <> "" And Me.txtNotaFin <> "" And Me.txtdur.Text <> "" And Me.cboagencia.ListIndex <> -1 And Me.cbocomite.ListIndex <> -1 Then
            If IsDate(txNotaIni) And IsDate(Me.txtNotaFin) And IsDate(Me.txtdur) Then
            MatProductos = "0,101,102,103,201,202,204,301,302,303,304,320,401,403,423"
            'Call ImprimeActaComiteCredAprox(CDate(Me.txNotaIni.Text), CDate(Me.txtNotaFin.Text), CDate(Me.txtdur.Text), CInt(Right((Me.cboagencia.Text), 2)), gsCodUser, Me.txtacta.Text, gdFecSis, Left((Me.cboagencia.Text), 20), MatProductos, Right((Me.cbocomite.Text), 2), gsNomCmac) 'Comentado x JGPA20180523
            '*** JGPA20180523 según ACTA 027-2018
            Call ImprimeActaComiteCredAprox(CDate(Me.txNotaIni.Text), CDate(Me.txtNotaFin.Text), CDate(Me.txtdur.Text), CInt(Right((Me.cboagencia.Text), 3)), gsCodUser, Me.txtacta.Text, gdFecSis, Left((Me.cboagencia.Text), 20), MatProductos, Right((Me.cbocomite.Text), 3), gsNomCmac)
            '*** End JGPA20180523
            Else
            MsgBox "Los Valores indicados en los textos no son Correctos"
            End If
    Else
            MsgBox "Complete los Datos para Generar el Reporte"
    End If
End Sub

Private Sub btn_salir_Click()
Unload Me
End Sub

Private Sub cboAgencia_Click()
CargarComites
End Sub

Private Sub Form_Load()
CargaAgencia
Me.txtNotaFin.Text = Format(Time(), "hh:mm:ss")
Me.txtdur.Text = Format(0, "hh:mm:ss")
Me.txNotaIni.Text = gdFecSis
End Sub


Private Sub CargaAgencia()
'Dim loCargaAg As COMDColocPig.DCOMColPFunciones
Dim loCargaAg As COMDConstantes.DCOMAgencias 'FRHU 20150326
Dim lrAgenc As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    'FRHU 20150326
    'Set loCargaAg = New COMDColocPig.DCOMColPFunciones
    'Set lrAgenc = loCargaAg.dObtieneAgencias(True)
    Set loCargaAg = New COMDConstantes.DCOMAgencias
    Set lrAgenc = loCargaAg.ObtieneAgencias()
    'FIN FRHU 20150326
    Set loCargaAg = Nothing
    Call llenar_cbo_agencia(lrAgenc, cboagencia)
    Exit Sub

ERRORCargaControles:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Sub llenar_cbo_agencia(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
pcboObjeto.Clear
Do While Not pRs.EOF
    'pcboObjeto.AddItem Trim(pRs!cAgeDescripcion) & Space(100) & Trim(str(pRs!cAgeCod))
    pcboObjeto.AddItem Trim(pRs!cConsDescripcion) & Space(100) & Trim(str(pRs!nConsValor)) 'FRHU 20150326
    pRs.MoveNext
Loop
pRs.Close
End Sub

Public Sub CargarComites()
Dim rs As ADODB.Recordset
Dim oGen As COMDCredito.DCOMCreditos

    On Error GoTo ERRORCargarComites
    
    Set oGen = New COMDCredito.DCOMCreditos
        Set rs = New ADODB.Recordset
        Set rs = oGen.CargaComiteAgencia(CInt(Right(cboagencia.Text, 2)))
    Set oGen = Nothing
    
     If Not (rs.EOF And rs.BOF) Then
        Call llenar_cbo_comite(rs, Me.cbocomite)
    Else
        MsgBox ("Esta agencia no tiene registrado Comite, Cancele y Registre !!")
        Me.cbocomite.Clear
        Exit Sub
    End If
    Set rs = Nothing
    Exit Sub
ERRORCargarComites:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Sub llenar_cbo_comite(pRs As ADODB.Recordset, pcboObjeto As ComboBox)
pcboObjeto.Clear
Do While Not pRs.EOF
    pcboObjeto.AddItem Trim(pRs!nom_comite) & Space(100) & Trim(str(pRs!id_comite))
    pRs.MoveNext
Loop
pRs.Close
End Sub

Public Function ImprimeActaComiteCredAprox( _
ByVal pdFecIni As Date, ByVal pdHora As Date, ByVal pdur As Date, _
ByVal psCodAge As Integer, ByVal psCodUser As String, ByVal pacta As String, _
ByVal pdFecSis As Date, _
ByVal psNomAge As String, _
ByVal pMatProd As String, _
ByVal pcomite As Integer, _
Optional ByVal psNomCmac As String = "") As String

Dim oDCred As COMDCredito.DCOMCredDoc
Dim oGen As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset
Dim Ra As ADODB.Recordset
Dim Rx As ADODB.Recordset
Dim sCadImp As String
Dim vProd As String
Dim vEstado As String
Dim sLetraUlt As String
Dim xrs As String
Dim i As Integer, j As Integer
Dim Hora As String

    Set oDCred = New COMDCredito.DCOMCredDoc 'DCredDoc
    Set oGen = New COMDCredito.DCOMCreditos
    Set R = oDCred.RecuperaActaComiteCredApro1(psCodAge, pdFecIni, pMatProd, pcomite)
    Set Ra = oGen.CargaPersonasComite(pcomite) 'madm 20100908
    Set Rx = oGen.CargaPersonasJefeyResponsable(pcomite, psCodAge) 'madm 20100908
    Set oDCred = Nothing
    Set oGen = Nothing

    If R.RecordCount = 0 Then
        MsgBox "No existen Datos para este Reporte.", vbInformation, "Atención"
        Exit Function
    End If
        
    Dim nom As String
    Dim comite As String
    Dim dirage As String
    Dim xnom As String
    i = 0
    Do While Not Ra.EOF
         If i = 0 Then
            nom = Ra!Nombre
         Else
            nom = nom & ", " & Ra!Nombre
         End If
         Ra.MoveNext
         i = i + 1
        If Ra.EOF Then
            Exit Do
        End If
    comite = Ra!nom_comite
    dirage = Ra!cAgeDireccion
    xnom = Ra!cAgeDescripcion
    Loop
        
    'presentar las filas de responsables
     If Not Rx.EOF Then
      xrs = Rx!Nombre
      xrs = xrs & "(" & IIf(Rx!cRHCargoDescripcion = "", "Coordinador de Créditos", Rx!cRHCargoDescripcion) & ")"
    End If
    
    Dim ApExcel As Variant
    Set ApExcel = CreateObject("Excel.application")
    
    'Agrega un nuevo Libro
    ApExcel.Workbooks.Add
   
   
    ApExcel.Cells(2, 2).Formula = "ACTAS DE REUNION Nº " & pacta & " DEL " & comite & " PERTENECIENTES A LA " & xnom
    ApExcel.Cells(3, 2).Formula = "DE LA CMAC MAYNAS DEL " & Format(pdFecIni, "dd/MM/YYYY")
    ApExcel.Cells(2, 2).Font.Bold = True
    ApExcel.Cells(3, 2).Font.Bold = True
    ApExcel.Range("B2:B3").HorizontalAlignment = 3
    ApExcel.Range("B2", "g2").MergeCells = True
    ApExcel.Range("B3", "g3").MergeCells = True
  
    ApExcel.Cells(5, 1).Formula = "En la Ciudad de Iquitos siendo las " & Format(pdHora, "hh:mm AMPM") & " del " & Format(pdFecIni, "dd/mm/yyyy ") & " en su local institucional sitio en " & dirage & ", con la finalidad de llevar a cabo la"
    
    If Len(nom) <= 80 Then
    ApExcel.Cells(6, 1).Formula = "reunión del Comité de Créditos, se reunieron : " & xrs & " ," & nom & ", "
    ApExcel.Cells(7, 1).Formula = "habiéndose aprobado los siguientes créditos: "
    ApExcel.Range("a5:a6").HorizontalAlignment = 2
    ApExcel.Range("a5", "h5").MergeCells = True
    ApExcel.Range("a6", "h6").MergeCells = True
    ApExcel.Range("a7", "h7").MergeCells = True
    i = 8
    ElseIf Len(nom) > 80 And Len(nom) <= 180 Then
    ApExcel.Cells(6, 1).Formula = "reunión del Comité de Créditos, se reunieron : " & xrs & " , " & nom & ", "
    ApExcel.Cells(8, 1).Formula = "habiéndose aprobado los siguientes créditos: "
    ApExcel.Range("a5:h5").HorizontalAlignment = 2
    ApExcel.Range("a6:h7").HorizontalAlignment = 2
    ApExcel.Range("a6:a7").VerticalAlignment = 4
    ApExcel.Range("a5", "h5").MergeCells = True
    ApExcel.Range("a6", "h7").MergeCells = True
    ApExcel.Range("a8", "h8").MergeCells = True
    i = 9
    ElseIf Len(nom) > 180 And Len(nom) <= 280 Then
    ApExcel.Cells(6, 1).Formula = "reunión del Comité de Créditos, se reunieron : " & xrs & " , " & nom & ", "
    ApExcel.Cells(9, 1).Formula = "habiéndose aprobado los siguientes créditos: "
    ApExcel.Range("a5:h5").HorizontalAlignment = 2
    ApExcel.Range("a6:h8").HorizontalAlignment = 2
    ApExcel.Range("a5:a8").VerticalAlignment = 4
    ApExcel.Range("a5", "h5").MergeCells = True
    ApExcel.Range("a6", "h8").MergeCells = True
    ApExcel.Range("a9", "h9").MergeCells = True
    i = 10
    ElseIf Len(nom) > 280 And Len(nom) <= 360 Then
    ApExcel.Cells(6, 1).Formula = "reunión del Comité de Créditos, se reunieron : " & xrs & " , " & nom & ", "
    ApExcel.Cells(10, 1).Formula = "habiéndose aprobado los siguientes créditos: "
    ApExcel.Range("a5:h5").HorizontalAlignment = 2
    ApExcel.Range("a6:h9").HorizontalAlignment = 2
    ApExcel.Range("a5:a8").VerticalAlignment = 4
    ApExcel.Range("a5", "h5").MergeCells = True
    ApExcel.Range("a6", "h9").MergeCells = True
    ApExcel.Range("a10", "h10").MergeCells = True
    i = 11
    ElseIf Len(nom) > 360 And Len(nom) <= 420 Then
    ApExcel.Cells(6, 1).Formula = "reunión del Comité de Créditos, se reunieron : " & xrs & " , " & nom
    ApExcel.Cells(11, 1).Formula = "habiéndose aprobado los siguientes créditos: "
    ApExcel.Range("a5:h5").HorizontalAlignment = 2
    ApExcel.Range("a6:h10").HorizontalAlignment = 2
    ApExcel.Range("a5:a8").VerticalAlignment = 4
    ApExcel.Range("a5", "h5").MergeCells = True
    ApExcel.Range("a6", "h10").MergeCells = True
    ApExcel.Range("a11", "h11").MergeCells = True
    i = 12
    Else
    ApExcel.Cells(6, 1).Formula = "reunión del Comité de Créditos, se reunieron : " & xrs & " , " & nom
    ApExcel.Cells(12, 1).Formula = "habiéndose aprobado los siguientes créditos: "
    ApExcel.Range("a5:h5").HorizontalAlignment = 2
    ApExcel.Range("a6:h11").HorizontalAlignment = 2
    ApExcel.Range("a5:a8").VerticalAlignment = 4
    ApExcel.Range("a5", "h5").MergeCells = True
    ApExcel.Range("a6", "h11").MergeCells = True
    ApExcel.Range("a12", "h12").MergeCells = True
    i = 13
    End If
    
    vEstado = ""
    Do While Not R.EOF
        If vEstado <> R!nPrdEstado Then
            i = i + 2
            ApExcel.Cells(i, 1).Formula = "CREDITOS : " & R!cPrdEstado
            ApExcel.Cells(i, 1).Font.Bold = True
            ''''ApExcel.Range("A" & i, "A" & i).HorizontalAlignment = 2
            i = i + 1
            'ALPA 20081007***********************************************************************
            If R!nPrdEstado = 2002 Then
                ApExcel.Cells(i, 1).Formula = "Nº"
                ApExcel.Cells(i, 2).Formula = "Cuenta"
                ApExcel.Cells(i - 1, 3).Formula = "Nombre de"
                ApExcel.Cells(i, 3).Formula = "Cliente"
'                ApExcel.Cells(i - 1, 4).Formula = "Fecha de"
'                ApExcel.Cells(i, 4).Formula = "aprobación"
                ApExcel.Cells(i - 1, 4).Formula = "Monto"
                ApExcel.Cells(i, 4).Formula = "aprobado"
                ApExcel.Cells(i, 5).Formula = "Cuotas"
'                ApExcel.Cells(i - 1, 7).Formula = "Tipo"
'                ApExcel.Cells(i, 7).Formula = "plazo"
                ApExcel.Cells(i - 1, 6).Formula = "Plazo/"
                ApExcel.Cells(i, 6).Formula = "Fecha fija"
                ApExcel.Cells(i, 7).Formula = "Linea"
                ApExcel.Cells(i, 8).Formula = "Analista"
                sLetraUlt = "H"
            ElseIf R!nPrdEstado = 2003 Then
                ApExcel.Cells(i, 1).Formula = "Nº"
                ApExcel.Cells(i, 2).Formula = "Cuenta"
                ApExcel.Cells(i - 1, 3).Formula = "Nombre de"
                ApExcel.Cells(i, 3).Formula = "Cliente"
'                ApExcel.Cells(i - 1, 4).Formula = "Fecha "
'                ApExcel.Cells(i, 4).Formula = "rechazo"
                ApExcel.Cells(i - 1, 4).Formula = "Monto"
                ApExcel.Cells(i, 4).Formula = "rechazo"
                ApExcel.Cells(i, 5).Formula = "Cuotas"
'                ApExcel.Cells(i - 1, 7).Formula = "Tipo"
'                ApExcel.Cells(i, 7).Formula = "plazo"
                ApExcel.Cells(i - 1, 6).Formula = "Plazo/"
                ApExcel.Cells(i, 6).Formula = "Fecha fija"
                ApExcel.Cells(i, 7).Formula = "Linea"
                ApExcel.Cells(i, 8).Formula = "Analista"
                ApExcel.Cells(i, 9).Formula = "Motivo"
                sLetraUlt = "I"
            ElseIf R!nPrdEstado = 2080 Then
                  ApExcel.Cells(i, 1).Formula = "Nº"
                ApExcel.Cells(i, 2).Formula = "Cuenta"
                ApExcel.Cells(i, 3).Formula = "Nombre de"
                ApExcel.Cells(i, 3).Formula = "Cliente"
'                ApExcel.Cells(i - 1, 4).Formula = "Fecha "
'                ApExcel.Cells(i, 4).Formula = "retiro"
                ApExcel.Cells(i - 1, 4).Formula = "Monto"
                ApExcel.Cells(i, 4).Formula = "retiro"
                ApExcel.Cells(i, 5).Formula = "Cuotas"
'                ApExcel.Cells(i - 1, 7).Formula = "Tipo"
'                ApExcel.Cells(i, 7).Formula = "plazo"
                ApExcel.Cells(i - 1, 6).Formula = "Plazo/"
                ApExcel.Cells(i, 6).Formula = "Fecha fija"
                ApExcel.Cells(i, 7).Formula = "Linea"
                ApExcel.Cells(i, 8).Formula = "Analista"
                ApExcel.Cells(i, 9).Formula = "Motivo"
                sLetraUlt = "I"
            End If
            '*******************************************************************************************
            'devCelda
            ApExcel.Range("A" & (i - 1), sLetraUlt & i).Interior.Color = RGB(10, 190, 160)
            ApExcel.Range("A" & (i - 1), sLetraUlt & i).Font.Bold = True
            ApExcel.Range("A" & (i - 1), sLetraUlt & i).HorizontalAlignment = 3
            vProd = ""
        End If
        vEstado = R!nPrdEstado
        i = i + 1
                ApExcel.Cells(i, 1).Formula = "PRODUCTO : " & R!Producto
             
             ApExcel.Cells(i, 1).Font.Bold = True
               vProd = R!Producto
             
            
             j = 0
             Do While R!Producto = vProd
                 j = j + 1
                 i = i + 1
                     
                     ApExcel.Cells(i, 1).Formula = j
                     ApExcel.Cells(i, 2).Formula = "'" & R!cCtaCod
                     ApExcel.Cells(i, 3).Formula = R!Cliente
'                     ApExcel.Cells(i, 4).Formula = Format(R!dPrdEstado, "mm/dd/yyyy")
                     ApExcel.Cells(i, 4).Formula = R!nMonto
                     ApExcel.Cells(i, 5).Formula = R!nCuotas
'                     ApExcel.Cells(i, 7).Formula = R!cConsDescripcion
                     ApExcel.Cells(i, 6).Formula = R!Plazo
                     ApExcel.Cells(i, 7).Formula = R!Descri_Linea
                     ApExcel.Cells(i, 8).Formula = R!Analista
                     ApExcel.Cells(i, 9).Formula = R!MotivoRechazo
                     
                     ApExcel.Range("D" & Trim(str(i)) & ":" & "D" & Trim(str(i))).NumberFormat = "#,##0.00"
                     ApExcel.Range("A" & Trim(str(i)) & ":" & sLetraUlt & Trim(str(i))).Borders.LineStyle = 1
                     
                     R.MoveNext
                     If R.EOF Then
                         Exit Do
                     End If
                                               
                 Loop
            Loop
            'imprime hora
                    Dim Total As Date
                    Total = pdHora + pdur
                    ApExcel.Cells(i + 2, 1).Formula = "Siendo las " & Format(Total, "hh:mm AMPM") & " se levantó la reunión."
            'imprime relacion analistas
                    Dim fil As Integer
                    Dim c, f As String
                    fil = i + 5
                    c = "c" & fil
                    f = "e" & fil
                    ApExcel.Cells(fil, 3).Formula = "Analistas de Créditos"
                    'ApExcel.Range(c:f).HorizontalAlignment = 3
                    ApExcel.Range(c, f).MergeCells = True
    
            'presentar las filas de analistas
                     Ra.MoveFirst
                     Dim X As Integer
                     Dim e As Integer
                     X = 3
                     e = 0
                     i = i + 9
                     Do While Not Ra.EOF
                             ApExcel.Cells(i, X).Formula = Ra!Nombre
                             e = e + 1
                             If e Mod 2 Then
                             X = X + 4 'mueva de 2 en dos las personas
                             Else
                             X = 3
                             i = i + 3
                             End If
                             
                             Ra.MoveNext
                             
                            If Ra.EOF Then
                                Exit Do
                            End If
                        Loop
            
             'presentar las filas de responsables
                     Rx.MoveFirst
                     Dim X1 As Integer
                     Dim e1 As Integer
                     X1 = 3
                     e1 = 0
                     i = i + 3
                     Do While Not Rx.EOF
                                ApExcel.Cells(i, X1).Formula = Rx!Nombre
                                ApExcel.Cells(i + 1, X1).Formula = IIf(Rx!cRHCargoDescripcion = "", "Coordinador de Créditos", Rx!cRHCargoDescripcion)
                             
                             e1 = e1 + 1
                             
                             If e1 Mod 2 Then
                                X1 = X1 + 4 'mueva de 2 en dos las personas
                             Else
                                X1 = 3
                                i = i + 5
                             End If
                             
                             Rx.MoveNext
                             
                            If Rx.EOF Then
                                Exit Do
                            End If
                        Loop
    Ra.Close
    R.Close
    Set Ra = Nothing
    Set R = Nothing
    
    ApExcel.Cells.Select
    ApExcel.Cells.EntireColumn.AutoFit
    ApExcel.Columns("A:A").ColumnWidth = 6#
    ApExcel.Range("A2").Select
    ApExcel.Range("A8:A10000").HorizontalAlignment = 2
    ApExcel.Range("A2:A2").HorizontalAlignment = 2
    ApExcel.Range("H1:H1").HorizontalAlignment = 4
    ApExcel.Range("H2:H2").HorizontalAlignment = 4
    ApExcel.Visible = True
    Set ApExcel = Nothing

End Function


