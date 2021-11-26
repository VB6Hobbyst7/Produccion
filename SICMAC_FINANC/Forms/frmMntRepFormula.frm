VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMntRepFormula 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estructura de Reportes en Base a Fórmulas"
   ClientHeight    =   4590
   ClientLeft      =   660
   ClientTop       =   2760
   ClientWidth     =   8715
   Icon            =   "frmMntRepFormula.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkOrdenaCodigo 
      Caption         =   "Ordena Codigo"
      Height          =   285
      Left            =   5145
      TabIndex        =   20
      Top             =   150
      Width           =   1665
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   60
      TabIndex        =   15
      Top             =   -60
      Width           =   4950
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   210
         Width           =   4695
      End
   End
   Begin VB.Frame Fradatos 
      Enabled         =   0   'False
      Height          =   1500
      Left            =   45
      TabIndex        =   0
      Top             =   4590
      Visible         =   0   'False
      Width           =   8595
      Begin VB.CheckBox chkEditable 
         Caption         =   "Resaltado"
         Height          =   225
         Left            =   7215
         TabIndex        =   21
         Top             =   180
         Width           =   1935
      End
      Begin VB.TextBox txtObserva 
         Height          =   300
         Left            =   3510
         MaxLength       =   254
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1005
         Width           =   3345
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   330
         Left            =   6975
         TabIndex        =   6
         Top             =   1005
         Width           =   1470
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   330
         Left            =   6960
         TabIndex        =   5
         Top             =   450
         Width           =   1470
      End
      Begin VB.TextBox txtFormula 
         Height          =   300
         Left            =   120
         MaxLength       =   254
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1005
         Width           =   3345
      End
      Begin VB.TextBox TxtDescrip 
         Height          =   300
         Left            =   2610
         MaxLength       =   254
         TabIndex        =   2
         Top             =   450
         Width           =   4215
      End
      Begin VB.TextBox TxtCodigo 
         Height          =   300
         Left            =   120
         MaxLength       =   19
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   450
         Width           =   2430
      End
      Begin VB.Label label4 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   3495
         TabIndex        =   19
         Top             =   795
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Formula"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   795
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   2610
         TabIndex        =   8
         Top             =   210
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   210
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4005
      Left            =   45
      TabIndex        =   10
      Top             =   525
      Width           =   8595
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   330
         Left            =   4680
         TabIndex        =   18
         Top             =   3510
         Width           =   1470
      End
      Begin MSDataGridLib.DataGrid GrdRep 
         Height          =   3135
         Left            =   210
         TabIndex        =   17
         Top             =   240
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   5530
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "cCodigo"
            Caption         =   "Código"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cDescrip"
            Caption         =   "Descripción"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cFormula"
            Caption         =   "Formula"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "cOpeCod"
            Caption         =   "cOpeCod"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "cObserva"
            Caption         =   "cObserva"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   4
            BeginProperty Column00 
               Alignment       =   2
               ColumnAllowSizing=   -1  'True
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   6
               ColumnWidth     =   3495.118
            EndProperty
            BeginProperty Column02 
               DividerStyle    =   6
               ColumnWidth     =   5999.812
            EndProperty
            BeginProperty Column03 
               DividerStyle    =   6
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   7109.858
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   330
         Left            =   6960
         TabIndex        =   14
         Top             =   3510
         Width           =   1470
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   330
         Left            =   3210
         TabIndex        =   13
         Top             =   3510
         Width           =   1470
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   330
         Left            =   1710
         TabIndex        =   12
         Top             =   3510
         Width           =   1470
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   330
         Left            =   210
         TabIndex        =   11
         Top             =   3525
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frmMntRepFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nAccion As Integer
Dim sOpeCod As String

Dim rsRep   As ADODB.Recordset
Dim clsRep  As DRepFormula

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If TxtCodigo = "" Then
   MsgBox "Código no Válido...!", vbInformation, "¡Aviso!"
   TxtCodigo.SetFocus
   Exit Function
End If
If TxtDescrip = "" Then
   MsgBox "Descripción no puede ser vacia...!", vbInformation, "¡Aviso!"
   TxtDescrip.SetFocus
   Exit Function
End If
ValidaDatos = True
End Function

Private Sub cboTipo_Click()
FiltraDatos
End Sub

Private Sub FiltraDatos()
Dim prs As ADODB.Recordset
Screen.MousePointer = 11
If cboTipo.ListIndex < 0 Then
   sOpeCod = ""
Else
   sOpeCod = Trim(Right(cboTipo, 100))
End If
rsRep.Filter = "cOpeCod = '" & sOpeCod & "'"
If Me.chkOrdenaCodigo.value = 1 Then
    rsRep.Sort = "cCodigo"
End If
Set GrdRep.DataSource = rsRep
Screen.MousePointer = 0
'EJVG20130429 ***
If sOpeCod = gContRepBaseNotasEstadoSitFinan Or sOpeCod = gContRepBaseNotasEstadoResultado Then
    Call frmNIIFNotasEstadoConfig.Inicio(sOpeCod, Trim(Left(cboTipo.Text, 100)))
    HabilitaControles (False)
ElseIf sOpeCod = gContRepEstadoSitFinanEEFF1 Or sOpeCod = gContRepEstadoSitFinanEEFF2 Or sOpeCod = gContRepEstadoSitFinanEEFF3 Then
    Call frmNIIFBaseFormulasEEFF.Inicio(sOpeCod, Trim(Left(cboTipo.Text, 100)))
    HabilitaControles (False)
ElseIf sOpeCod = 760114 Or sOpeCod = 760115 Then '*** PEAC 20131126
    Call frmMantRepoFormulaBCR.Inicio(sOpeCod)
    HabilitaControles (False)
ElseIf sOpeCod = gContRepBaseNotasComplementariaInfoAnual Then 'EJVG20140318
    frmNIIFNotasComplementariaConfig.Show 1
    HabilitaControles (False)
Else
    HabilitaControles (True)
End If
'END EJVG *******
'FRHU 20131223 RQ13657
If sOpeCod = gContRepEstadoFlujoEfectivo Then
    Call frmConfigEstadoFlujoEfectivo.Inicio(sOpeCod, Trim(Left(cboTipo.Text, 100)))
End If
If sOpeCod = gContRepHojaTrabajoFlujoEfectivo Then
    Call frmConfigHojaTrabajoFE.Inicio(sOpeCod, Trim(Left(cboTipo.Text, 100)))
End If
'FIN FRHU 20131223 RQ13657
End Sub

Private Sub cmdImprimir_Click()
 Dim fs              As Scripting.FileSystemObject
    Dim xlAplicacion    As Excel.Application
    Dim xlLibro         As Excel.Workbook
    Dim xlHoja1         As Excel.Worksheet
    Dim lbExisteHoja    As Boolean
    Dim liLineas        As Integer
    Dim I               As Integer
    Dim glsArchivo      As String
    Dim lsNomHoja       As String
    Dim HojasExcel      As Integer 'numero de hojas de Excel a usar para mostrar las Cuentas contables

    Dim RSTEMP As New ADODB.Recordset
    'Dim clsBuscar As New ClassDescObjeto

    'Set rsTemp = clsCtaCont.CargaCtaCont("SubString(cCtaContCod,3,1)='" & Trim(Str(lnIndex)) & "' or Len(rtrim(cCtaContCod))<3", "CtaCont", adLockOptimistic)
       
    If rsRep Is Nothing Then
        MsgBox "No exite informacion para imprimir", vbInformation, "Aviso"
        Exit Sub
    End If

    glsArchivo = "Reporte_Cuentas_Formulas" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.path & "\SPOOLER\" & glsArchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & glsArchivo)
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
        Set xlHoja1 = xlLibro.Worksheets.Add
    
        xlHoja1.PageSetup.CenterHorizontally = True
        xlHoja1.PageSetup.Zoom = 60
        xlHoja1.PageSetup.Orientation = xlLandscape
    
        lbExisteHoja = False
        lsNomHoja = "CuentasContables_Formulas"
        For Each xlHoja1 In xlLibro.Worksheets
            If xlHoja1.Name = lsNomHoja Then
                xlHoja1.Activate
                lbExisteHoja = True
                Exit For
            End If
        Next
        If lbExisteHoja = False Then
            Set xlHoja1 = xlLibro.Worksheets.Add
            xlHoja1.Name = lsNomHoja
        End If

            xlAplicacion.Range("A1:A1").ColumnWidth = 10
            xlAplicacion.Range("B1:B1").ColumnWidth = 51 '10
            xlAplicacion.Range("c1:c1").ColumnWidth = 100 '15
            
          
            xlAplicacion.Range("A1:Z100").Font.Size = 9

            xlHoja1.Cells(1, 1) = gsNomCmac
            xlHoja1.Cells(2, 2) = "L I S T A D O   D E   C U E N T A S   C O N T A B L E S  E N  B A S E   A  F O R M U L A S"
            xlHoja1.Cells(3, 2) = "INFORMACION  AL  " & Format(gdFecSis, "dd/mm/yyyy")

            xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 3)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 3)).Merge True
            xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 3)).Merge True
            xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 3)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 3)).HorizontalAlignment = xlCenter

            liLineas = 6

            xlHoja1.Cells(liLineas, 1) = "Codigo"
            xlHoja1.Cells(liLineas, 2) = "Descripcion"
            xlHoja1.Cells(liLineas, 3) = "Formula"
         

            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 3)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 3)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 3)).Borders.LineStyle = 1
            xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, 3)).Interior.Color = RGB(159, 206, 238)
            liLineas = liLineas + 1

       
         rsRep.MoveFirst
         Do Until rsRep.EOF
             
                xlHoja1.Cells(liLineas, 1) = IIf(rsRep(0) = "", "", rsRep(0))
                xlHoja1.Cells(liLineas, 2) = IIf(rsRep(2) = "", "", rsRep(2))
                xlHoja1.Cells(liLineas, 3) = IIf(rsRep(3) = "", "", rsRep(3))
                liLineas = liLineas + 1
                'Me.pgbCtaCont.value = rsTemp.Bookmark
                rsRep.MoveNext

         Loop
         'Me.pgbCtaCont.Visible = False

        ExcelCuadro xlHoja1, 2, 6, 3, liLineas - 1
        xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
        ExcelEnd App.path & "\Spooler\" & glsArchivo, xlAplicacion, xlLibro, xlHoja1
        'Cierra el libro de trabajo
        'xlLibro.Close
        ' Cierra Microsoft Excel con el método Quit.
        'xlAplicacion.Quit
        'Libera los objetos.
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
        Call CargaArchivo(glsArchivo, App.path & "\SPOOLER\")



End Sub

Private Sub Form_Load()
frmMdiMain.Enabled = False
Set clsRep = New DRepFormula
Set rsRep = clsRep.CargaRepFormula(, , , adLockOptimistic)

Dim clsOpe As New DOperacion
Dim rsOpe  As New ADODB.Recordset

CentraForm Me
If Not rsRep.EOF Then
   Set rsOpe = clsOpe.CargaOpeTpo(Mid(rsRep!cOpeCod, 1, 4) & "_[1234567890]", True, , , , " or (copecod in (Select distinct copecod from RepBaseFormula)) ")
   RSLlenaCombo rsOpe, Me.cboTipo
End If
RSClose rsOpe
Set clsOpe = Nothing

If cboTipo.ListCount > 0 Then
   cboTipo.ListIndex = 0
End If
FiltraDatos
End Sub

Private Sub cmdAceptar_Click()
Dim nRedefinir As Boolean
 
On Error GoTo ErrAceptar
nRedefinir = False

If Not ValidaDatos Then
   Exit Sub
End If
If MsgBox(" ¿ Seguro que desea grabar Datos ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
   Exit Sub
End If

If nAccion = 1 Then 'Insertar
    If clsRep.ExisteCodigo(sOpeCod, Trim(TxtCodigo.Text)) = True Then
        If MsgBox("¡El Código ya existe!" & Chr(13) & Chr(13) & "Desea redefinir los números ya ingresados a partir de este nuevo?", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
            Exit Sub
        Else
            If MsgBox("La opción que ha presionado es irreversible!!!" & Chr(13) & Chr(13) & "Desea Continuar?", vbQuestion + vbYesNo, "Advertencia!!!") = vbYes Then
                nRedefinir = True
            Else
                Exit Sub
            End If
        End If
    Else
        nRedefinir = False
    End If
End If

gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   Select Case nAccion
       Case 1 'Insertar
           If nRedefinir = True Then
            clsRep.ReorganizaOrdenRep sOpeCod, Trim(TxtCodigo.Text)
        End If
           
           clsRep.InsertaRepFormula Trim(TxtCodigo.Text), sOpeCod, TxtDescrip, txtFormula, txtObserva, gsMovNro, Me.chkEditable.value
       Case 2 'Actualizar
           clsRep.ActualizaRepFormula Trim(TxtCodigo), sOpeCod, TxtDescrip, txtFormula, txtObserva, gsMovNro, Me.chkEditable.value
           
   End Select
   Set rsRep = clsRep.CargaRepFormula(, , , adLockOptimistic)
   FiltraDatos
   rsRep.Find "cCodigo = '" & TxtCodigo & "'"
   ActivaIngreso False
   GrdRep.SetFocus
Exit Sub
ErrAceptar:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub ActivaIngreso(lActiva As Boolean)
   If Not lActiva Then
      Height = 5000
   Else
      Height = 6330
   End If
   If nAccion = 1 Then
      TxtCodigo.Text = ""
      TxtDescrip.Text = ""
      txtFormula.Text = ""
      txtObserva.Text = ""
   Else
      TxtCodigo = rsRep!cCodigo
      TxtDescrip = rsRep!cDescrip
      txtFormula = rsRep!cFormula
      txtObserva = "" & rsRep!cObserva
      txtObserva = "" & rsRep!cObserva
      Me.chkEditable.value = IIf(rsRep!bIngresoManual, 1, 0)
   End If
   Fradatos.Enabled = lActiva
   Fradatos.Visible = lActiva
   'grdrep.enabled=not lactiva
   
   Frame1.Enabled = Not lActiva
   Frame2.Enabled = Not lActiva
   
   TxtCodigo.Enabled = lActiva
End Sub

Private Sub cmdAgregar_Click()
    nAccion = 1
    TxtCodigo.Text = clsRep.GeneraCodigoFila(sOpeCod)
    ActivaIngreso True
    TxtCodigo.SetFocus
End Sub

Private Sub cmdCancelar_Click()
    ActivaIngreso False
    GrdRep.SetFocus
End Sub

Private Sub cmdEliminar_Click()
If Not rsRep.EOF Then
   If MsgBox(" ¿ Seguro que desea eliminar Fila ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
      Exit Sub
   End If
   clsRep.EliminaRepFormula rsRep!cCodigo, rsRep!cOpeCod
   rsRep.Delete adAffectCurrent
   GrdRep.SetFocus
End If
End Sub

Private Sub CmdModificar_Click()
    nAccion = 2
    ActivaIngreso True
    TxtCodigo.Enabled = False
    TxtDescrip.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rsRep
Set clsRep = Nothing
frmMdiMain.Enabled = True
End Sub
 

Private Sub grdRep_DblClick()
    CmdModificar_Click
End Sub

Private Sub grdRep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdModificar_Click
    End If
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   TxtDescrip.SetFocus
End If
End Sub

Private Sub TxtDescrip_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFormula.SetFocus
    End If
End Sub

Private Sub txtFormula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtObserva.SetFocus
    End If
End Sub

Private Sub txtObserva_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub
'EJVG20130429 ***
Private Sub HabilitaControles(ByVal pbHabilita As Boolean)
    GrdRep.Enabled = pbHabilita
    cmdAgregar.Enabled = pbHabilita
    cmdEliminar.Enabled = pbHabilita
    cmdModificar.Enabled = pbHabilita
    cmdAceptar.Enabled = pbHabilita
    cmdCancelar.Enabled = pbHabilita
    cmdImprimir.Enabled = pbHabilita
End Sub
'END EJVG *******
