VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPreparaSorteo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PREPARACION DE SORTEO"
   ClientHeight    =   5340
   ClientLeft      =   3210
   ClientTop       =   2040
   ClientWidth     =   8070
   Icon            =   "frmPreparaSorteo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   8070
   Begin VB.Frame fraContenedor 
      Height          =   5205
      Index           =   0
      Left            =   75
      TabIndex        =   4
      Top             =   45
      Width           =   7875
      Begin VB.Frame FraInfo 
         BorderStyle     =   0  'None
         Height          =   1740
         Left            =   90
         TabIndex        =   17
         Top             =   210
         Width           =   7740
         Begin VB.CommandButton cmdNuevo 
            Caption         =   "&Nuevo"
            Height          =   390
            Left            =   4410
            TabIndex        =   26
            Top             =   1275
            Width           =   780
         End
         Begin VB.TextBox txtEstado 
            Alignment       =   2  'Center
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
            Height          =   315
            Left            =   6180
            TabIndex        =   25
            Top             =   75
            Width           =   1320
         End
         Begin VB.TextBox TxtNumSorteo 
            Alignment       =   2  'Center
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
            Height          =   330
            Left            =   690
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   60
            Width           =   1650
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   390
            Left            =   6885
            TabIndex        =   23
            Top             =   1275
            Width           =   780
         End
         Begin VB.CommandButton cmdGrabar 
            Caption         =   "&Grabar"
            Enabled         =   0   'False
            Height          =   390
            Left            =   6075
            TabIndex        =   22
            Top             =   1275
            Width           =   780
         End
         Begin VB.Frame fraAgencia 
            Caption         =   "Alcance "
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
            Height          =   690
            Left            =   15
            TabIndex        =   20
            Top             =   1005
            Width           =   4365
            Begin VB.ComboBox cboAlcance 
               Height          =   315
               Left            =   135
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   240
               Width           =   4155
            End
         End
         Begin VB.TextBox txtObs 
            Enabled         =   0   'False
            Height          =   345
            Left            =   1485
            MaxLength       =   500
            TabIndex        =   19
            Top             =   525
            Width           =   6000
         End
         Begin VB.CommandButton cmdEditar 
            Caption         =   "&Editar"
            Enabled         =   0   'False
            Height          =   390
            Left            =   5250
            TabIndex        =   18
            Top             =   1275
            Width           =   780
         End
         Begin MSMask.MaskEdBox TxtFecha 
            Height          =   315
            Left            =   3030
            TabIndex        =   27
            Top             =   75
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtHora 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "H:mm:ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   4
            EndProperty
            Height          =   315
            Left            =   4725
            TabIndex        =   28
            Top             =   75
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Estado"
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
            Index           =   5
            Left            =   5565
            TabIndex        =   33
            Top             =   135
            Width           =   585
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Número"
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
            Index           =   7
            Left            =   30
            TabIndex        =   32
            Top             =   135
            Width           =   660
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Hora"
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
            Index           =   0
            Left            =   4305
            TabIndex        =   31
            Top             =   135
            Width           =   405
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Fecha"
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
            Index           =   6
            Left            =   2460
            TabIndex        =   30
            Top             =   135
            Width           =   540
         End
         Begin VB.Label lblEtiqueta 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   30
            TabIndex        =   29
            Top             =   615
            Width           =   1395
         End
      End
      Begin VB.CommandButton cmdEstado 
         Caption         =   "En Proceso ..."
         Height          =   360
         Left            =   120
         TabIndex        =   14
         Top             =   4785
         Width           =   1260
      End
      Begin VB.Frame fraProcesos 
         Height          =   2775
         Left            =   105
         TabIndex        =   6
         Top             =   1950
         Width           =   7695
         Begin VB.CheckBox chkProcesoOtros 
            Caption         =   "Procesar Otras Agencias"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3870
            TabIndex        =   34
            Top             =   300
            Visible         =   0   'False
            Width           =   2460
         End
         Begin VB.CheckBox chkCliente 
            Caption         =   "Por Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4905
            TabIndex        =   16
            Top             =   645
            Value           =   1  'Checked
            Width           =   1065
         End
         Begin VB.CheckBox chkCupon 
            Caption         =   "Por Cupón"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3870
            TabIndex        =   15
            Top             =   645
            Width           =   975
         End
         Begin VB.CommandButton cmdImprimirTIcketCTS 
            Caption         =   "Imprimir Cupones x Cta. CTS"
            Height          =   360
            Left            =   90
            TabIndex        =   13
            Top             =   1470
            Width           =   3645
         End
         Begin VB.CommandButton cmdImprimirBatch 
            Caption         =   "Imprimir Cupones  Batch"
            Height          =   360
            Left            =   105
            TabIndex        =   12
            Top             =   1875
            Width           =   3645
         End
         Begin VB.CommandButton cmdPrepararPlanilla 
            Caption         =   "Planilla para el Sorteo"
            Height          =   360
            Left            =   105
            TabIndex        =   1
            Top             =   660
            Width           =   3645
         End
         Begin VB.CommandButton cmProcesar 
            Caption         =   "Procesar Cuentas para Sorteo"
            Height          =   360
            Left            =   105
            TabIndex        =   0
            Top             =   255
            Width           =   3645
         End
         Begin VB.CommandButton cmdImprimirTicket 
            Caption         =   "Imprimir Cupones x Cta. Plazo FIjo"
            Height          =   360
            Left            =   105
            TabIndex        =   2
            Top             =   1080
            Width           =   3645
         End
         Begin VB.CommandButton cmdImpConsolidados 
            Caption         =   "Listados Consolidados"
            Height          =   360
            Left            =   105
            TabIndex        =   3
            Top             =   2295
            Width           =   3645
         End
         Begin VB.Frame fraImpresion 
            Caption         =   "Impresión"
            Height          =   1050
            Left            =   5370
            TabIndex        =   8
            Top             =   1200
            Width           =   2220
            Begin VB.OptionButton optImpresion 
               Caption         =   "Archivo"
               Height          =   270
               Index           =   2
               Left            =   495
               TabIndex        =   11
               Top             =   720
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.OptionButton optImpresion 
               Caption         =   "Impresora"
               Height          =   270
               Index           =   1
               Left            =   495
               TabIndex        =   10
               Top             =   480
               Width           =   990
            End
            Begin VB.OptionButton optImpresion 
               Caption         =   "Pantalla"
               Height          =   195
               Index           =   0
               Left            =   495
               TabIndex        =   9
               Top             =   285
               Value           =   -1  'True
               Width           =   960
            End
         End
         Begin VB.CommandButton cmdAgencia 
            Caption         =   "A&gencias..."
            Height          =   345
            Left            =   5385
            TabIndex        =   7
            Top             =   2340
            Width           =   1140
         End
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   390
         Left            =   6780
         TabIndex        =   5
         Top             =   4785
         Width           =   990
      End
   End
   Begin SICMACT.Usuario User 
      Left            =   150
      Top             =   4785
      _extentx        =   820
      _extenty        =   820
   End
   Begin MSComDlg.CommonDialog dlgGrabar 
      Left            =   120
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPreparaSorteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gsMontoxCuponS As Double
Dim gsMontoxCuponD As Double
Dim gsMinPlazoOtorgado As Long
Dim gsLimMaxOtorgadoS As Double
Dim gsLimMaxOtorgadoD As Double
Public gsAlcance As String
Dim bNuevo As Boolean

Private Sub CargaParametros()
   Dim oSorteo As DSorteo, rsDatos As Recordset
   Set oSorteo = New DSorteo
      Set rsDatos = oSorteo.GetParametrosSorteo
      While Not rsDatos.EOF
            
            Select Case rsDatos!nparamcod
               Case 2001
                     gsMontoxCuponS = rsDatos!cparamvalor
               
               Case 2002
                    gsMontoxCuponD = rsDatos!cparamvalor
                    
               Case 2003
                    gsMinPlazoOtorgado = rsDatos!cparamvalor
                    
               Case 2004
                    gsLimMaxOtorgadoS = rsDatos!cparamvalor
               
               Case 2005
                    gsLimMaxOtorgadoD = rsDatos!cparamvalor
            
            End Select
            rsDatos.MoveNext
      Wend
   
   
   
End Sub

Private Sub CargaAgencias(ByVal cAlcance As String)
Dim clsAge As DAgencias, rsTemp As ADODB.Recordset, i As Integer
cboAlcance.Clear
   Set clsAge = New DAgencias
   Set rsTemp = clsAge.RecuperaAgencias()
   i = 0
     If cAlcance = "00" Then
        cboAlcance.AddItem "***GENERAL***" & Space(100 - Len("***GENERAL***")) & "00"
     Else
        cboAlcance.AddItem "<SELECCIONE AGENCIA>" & Space(100 - Len("<SELECCIONE AGENCIA>")) & "00"
        While Not rsTemp.EOF
                If rsTemp!cAgeCod <> "13" And rsTemp!cAgeCod <> "03" And rsTemp!cAgeCod <> "06" And rsTemp!cAgeCod <> "04" And rsTemp!cAgeCod <> "05" And rsTemp!cAgeCod <> "01" And rsTemp!cAgeCod <> "14" Then
                    cboAlcance.AddItem Trim(rsTemp!cAgeDescripcion) & Space(100 - Len(Trim(rsTemp!cAgeDescripcion))) & rsTemp!cAgeCod
                  
                ElseIf rsTemp!cAgeCod = "01" Then
                    cboAlcance.AddItem "Ica-Parcona" & Space(100 - Len(Trim(rsTemp!cAgeDescripcion))) & "01"
                    
                ElseIf rsTemp!cAgeCod = "13" Then 'And RSTEMP!cAgeCod <> "03" And RSTEMP!cAgeCod <> "06" Then
                    cboAlcance.AddItem "San Vic-Imperial-Mala" & Space(100 - Len(Trim(rsTemp!cAgeDescripcion))) & "13"
                    'RSTEMP.MoveNext
                
                ElseIf rsTemp!cAgeCod = "04" Then 'And RSTEMP!cAgeCod <> "05" Then
                    cboAlcance.AddItem "Nasca-Palpa" & Space(100 - Len(Trim(rsTemp!cAgeDescripcion))) & "04"
                    'RSTEMP.MoveNext
                
                End If
                  rsTemp.MoveNext
        Wend
     End If
   If cboAlcance.ListCount > 0 Then cboAlcance.ListIndex = 0
   Set rsTemp = Nothing
 
End Sub

Private Function GeneraNumeroSorteo(ByVal cAlcance As String, ByVal cAnio As String) As String
Dim oSorteo As DSorteo
   Set oSorteo = New DSorteo
        GeneraNumeroSorteo = oSorteo.GeneraNumSorteo(cAlcance, cAnio)
   Set oSorteo = Nothing
End Function

Private Sub cboAlcance_Click()
Dim oSorteo As DSorteo, rsDatos As Recordset
Set oSorteo = New DSorteo


If Me.cboAlcance.ListCount > 0 Then
  If cboAlcance.ListIndex > 0 Then
    gsAlcance = Right(cboAlcance.Text, 2)
    'Me.TxtNumSorteo.Text = gsAlcance & Mid(Trim(TxtNumSorteo.Text), 3)
    If cmdNuevo.Enabled = False And cmdNuevo.Caption = "&Nuevo" Then
    
        If oSorteo.GetSorteoEstados(gsAlcance, Year(gdFecSis), "I") = False Then
            TxtNumSorteo.Text = ""
            TxtFecha.Text = "00/00/0000"
            TxtHora.Text = "00:00"
            txtEstado.Text = ""
            txtObs.Text = ""
            TxtNumSorteo.Text = GeneraNumeroSorteo(gsAlcance, Year(gdFecSis))
            
        Else
              MsgBox "Ya Existe un sorteo Iniciado", vbOKOnly + vbInformation, "Aviso"

              Set rsDatos = oSorteo.GetSorteo("I", gsAlcance)
                If Not (rsDatos.EOF Or rsDatos.BOF) Then
                    TxtNumSorteo.Text = ""
                    TxtFecha.Text = "00/00/0000"
                    TxtHora.Text = "00:00"
                    txtEstado.Text = ""
                    txtObs.Text = ""
                     CargaDatos rsDatos, True
                     cmdEditar.Enabled = True

                 End If
    
    
    
        End If
    Else
'                Set rsDatos = oSorteo.GetSorteo("I", gsAlcance)
'                If Not (rsDatos.EOF Or rsDatos.BOF) Then
'                     CargaDatos rsDatos
'                     cmdEditar.Enabled = True
'
'                End If
    
     
    End If
  End If
End If

End Sub

Private Sub chkCliente_Click()
    If chkCliente.value = vbChecked Then
        chkCupon.value = vbUnchecked
    Else
        chkCupon.value = vbChecked
    End If
    
End Sub

Private Sub chkCupon_Click()

If chkCupon.value = vbChecked Then
        chkCliente.value = vbUnchecked
Else
        chkCliente.value = vbChecked
End If

End Sub

Private Sub cmdAgencia_Click()
    frmSelectAgencias.Inicio Me
    frmSelectAgencias.Show 1
End Sub

Private Sub cmdCancelar_Click()
 
 Dim oSorteo As DSorteo, rsDatos As Recordset
 Set oSorteo = New DSorteo
    
           
       Set rsDatos = oSorteo.GetSorteo("I", gsAlcance)
                       
        LimpiaDatos
        If Not rsDatos.EOF Then
            CargaDatos rsDatos
        End If

 
  TxtFecha.Enabled = False
  TxtHora.Enabled = False
  txtObs.Enabled = False
  
 If gsAlcance <> "00" Then
   fraAgencia.Enabled = False
 End If
     
 cmdGrabar.Enabled = False
 cmdCancelar.Enabled = False
 cmdNuevo.Enabled = True
 cmdEditar.Enabled = True
 
 If Trim(txtEstado.Text) <> "" Then
        fraProcesos.Enabled = True
        cmdEstado.Enabled = True
 End If
End Sub

Private Sub cmdEditar_Click()
 TxtFecha.Enabled = True
 TxtHora.Enabled = True
 txtObs.Enabled = True
    bNuevo = False
  cmdNuevo.Enabled = False
  cmdGrabar.Enabled = True
  cmdCancelar.Enabled = True
  fraProcesos.Enabled = False
  cmdEstado.Enabled = False
  cmdEditar.Enabled = False
End Sub

Private Sub cmdEstado_Click()
Dim oSorteo As DSorteo
Set oSorteo = New DSorteo

 If MsgBox("ADVERTENCIA!!!: Si Realiza este proceso no podrá modificar la información del Sorteo" & vbCrLf & "Esta seguro de cambiar el estado del Sorteo??? ", vbYesNo, "AVISO") = vbYes Then
  Call oSorteo.ActualizaSorteo(Trim(TxtNumSorteo.Text), Trim(txtObs.Text), 0, 0, 0, "P", Format(TxtFecha.Text & " " & Trim(TxtHora.Text) & ":00", "yyyy-mm-dd hh:mm:ss"), sMovNro)
  Call oSorteo.ActualizaCuentasCanceladas(Left(Trim(TxtNumSorteo.Text), 6))
  Call oSorteo.ActualizaCuentasAnuladas(Left(Trim(TxtNumSorteo.Text), 6))
  Call oSorteo.InsertaTempPortable(Left(Trim(TxtNumSorteo.Text), 6))
  fraProcesos.Enabled = False
  
  Me.cmProcesar.Enabled = False
  
 End If
  
End Sub

Private Sub cmdGrabar_Click()
Dim oSorteo As DSorteo, sMovNro As String, clsMov As NContFunciones

  Set oSorteo = New DSorteo
  
  
   Set clsMov = New NContFunciones
   sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
     Set clsMov = Nothing
        
  If IsDate(TxtFecha.Text) = False Then
    MsgBox "Ingrese una fecha válida", vbOKOnly + vbInformation, "Aviso"
    Exit Sub
  End If
        
  If CDate(TxtFecha.Text) < gdFecSis Then
      MsgBox "La fecha ingresa es menor a la fecha actual del sistema.", vbOKOnly + vbInformation, "Aviso"
      Exit Sub
  End If
  
  If TxtHora.Text = "00:00" Then
      MsgBox "Debe indicar una hora válida.", vbOKOnly + vbInformation, "Aviso"
      Exit Sub
  End If
  
  If gsAlcance <> "00" Then
      If cboAlcance.ListIndex = 0 Then
        MsgBox "Debe indicar una AGENCIA válida.", vbOKOnly + vbInformation, "Aviso"
        Exit Sub
      End If
  End If
  
  If oSorteo.GetSorteoEstados(gsAlcance, Year(gdFecSis), "I") = True Then
      MsgBox "No se pude iniciar nuevo sorteo, ya existe uno iniciado.", vbOKOnly + vbExclamation, "AVISO"
      Exit Sub
  End If
  
  If bNuevo Then
     Call oSorteo.InsertaSorteo(Trim(TxtNumSorteo.Text), Trim(txtObs.Text), 0, 0, 0, "I", Format(TxtFecha.Text & " " & Trim(TxtHora.Text) & ":00", "yyyy-mm-dd hh:mm:ss"), sMovNro)
     txtEstado.Text = "Iniciado"
     
  Else
     Call oSorteo.ActualizaSorteo(Trim(TxtNumSorteo.Text), Trim(txtObs.Text), 0, 0, 0, "I", Format(TxtFecha.Text & " " & Trim(TxtHora.Text) & ":00", "yyyy-mm-dd hh:mm:ss"), sMovNro)
     
  End If
  
  TxtFecha.Enabled = False
  TxtHora.Enabled = False
  txtObs.Enabled = False
  
  If gsAlcance <> "00" Then
   fraAgencia.Enabled = False
  End If
  
  cmdGrabar.Enabled = False
  cmdCancelar.Enabled = False
  cmdNuevo.Enabled = True
  cmdEditar.Enabled = True
'  cmdNuevo.Caption = "&Editar"
  
 If Trim(txtEstado.Text) <> "" Then
        fraProcesos.Enabled = True
        cmdEstado.Enabled = True
 End If

  
  
End Sub

Private Sub cmdImpConsolidados_Click()

If MsgBox("Este proceso podría tardar varios minutos." & vbCrLf & "Esta seguro de procesarlo.", vbYesNo + vbQuestion, "Aviso") = vbYes Then

    Screen.MousePointer = vbHourglass
    Dim lnAge As Integer
    Dim oPrevio As Previo.clsPrevio, lscadena As String, cCodAgeAnt As String
    Dim orep As nCaptaReportes
    Set orep = New nCaptaReportes
    Set oPrevio = New Previo.clsPrevio
     lscadena = ""
     cCodAgeAnt = ""
    
     If gsAlcance = "00" Then
       
            For lnAge = 1 To frmSelectAgencias.List1.ListCount
                 If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
                   lscadena = lscadena & orep.CONSOLIDADOSORTEO(Left(Trim(TxtNumSorteo.Text), 6), Left(frmSelectAgencias.List1.List(lnAge - 1), 2), cCodAgeAnt, False)
                   cCodAgeAnt = Left(frmSelectAgencias.List1.List(lnAge - 1), 2)
                End If
            Next lnAge
        
       
     Else
           lscadena = orep.CONSOLIDADOSORTEO(Left(Trim(TxtNumSorteo.Text), 6), gsAlcance, gsAlcance)

     End If
    
         If Trim(lscadena) = "" Then
            MsgBox "No hay información para esta consulta", vbOKOnly + vbInformation, "AVISO"
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
    
    
        If Me.optImpresion(0).value = True Then
            Set loPrevio = New Previo.clsPrevio
                loPrevio.Show lscadena, "PLANILLA PARA EL SORTEO " & Trim(TxtNumSorteo.Text), True
            Set loPrevio = Nothing
        Else
            dlgGrabar.CancelError = True
            dlgGrabar.InitDir = App.path
            dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
            dlgGrabar.ShowSave
            If dlgGrabar.FileName <> "" Then
               Open dlgGrabar.FileName For Output As #1
                Print #1, lscadena
                Close #1
            End If
        End If
       Screen.MousePointer = vbDefault
        
End If

End Sub

Private Sub cmdImprimirBatch_Click()

  
  frmImprimirCuponesBatch.Show 1
  frmImprimirCuponesBatch.gsAlcance = Left(Trim(TxtNumSorteo.Text), 2)
  Call frmImprimirCuponesBatch.CargaMatriz(Left(Trim(TxtNumSorteo.Text), 2))
  frmImprimirCuponesBatch.gsAlcance = Left(Trim(TxtNumSorteo.Text), 2)


'If MsgBox("Este proceso podría tardar varios minutos." & vbCrLf & "Esta seguro de procesarlo.", vbYesNo + vbQuestion, "Aviso") = vbYes Then
'    Screen.MousePointer = vbHourglass
'    Dim lnAge As Integer
'    Dim oPrevio As Previo.clsPrevio, lscadena As String
'    Dim orep As nCaptaReportes
'    Set orep = New nCaptaReportes
'    Set oPrevio = New Previo.clsPrevio
'     lscadena = ""
'
'
'     If gsAlcance = "00" Then
'
'            For lnAge = 1 To frmSelectAgencias.List1.ListCount
'                 If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
'                   lscadena = lscadena & orep.CUPONESBATCH(Left(Trim(TxtNumSorteo.Text), 6), Left(frmSelectAgencias.List1.List(lnAge - 1), 2))
'
'                End If
'            Next lnAge
'
'
'     Else
'           lscadena = orep.CUPONESBATCH(Trim(TxtNumSorteo.Text), gsAlcance)
'
'     End If
'
'        If Me.optImpresion(0).value = True Then
'            Set loPrevio = New Previo.clsPrevio
'                loPrevio.Show lscadena, "Cupones" & Trim(TxtNumSorteo.Text), True
'            Set loPrevio = Nothing
'        Else
'            dlgGrabar.CancelError = True
'            dlgGrabar.InitDir = App.path
'            dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
'            dlgGrabar.ShowSave
'            If dlgGrabar.FileName <> "" Then
'               Open dlgGrabar.FileName For Output As #1
'                Print #1, lscadena
'                Close #1
'            End If
'        End If
'       Screen.MousePointer = vbDefault
'
'End If
End Sub

Private Sub cmdImprimirTicket_Click()
 
  frmImprimirCupones.Inicia Producto.gCapPlazoFijo, Left(Trim(TxtNumSorteo.Text), 2)
  frmImprimirCupones.Show 1
  frmImprimirCupones.gsAlcance = Left(Trim(TxtNumSorteo.Text), 2)
  Call frmImprimirCupones.CargaMatriz(Left(Trim(TxtNumSorteo.Text), 2))
  frmImprimirCupones.gsAlcance = Left(Trim(TxtNumSorteo.Text), 2)

  
End Sub

Private Sub cmdImprimirTIcketCTS_Click()

 frmImprimirCupones.Inicia Producto.gCapCTS, Left(Trim(TxtNumSorteo.Text), 2)
 frmImprimirCupones.gsAlcance = Left(Trim(TxtNumSorteo.Text), 2)
 frmImprimirCupones.Show 1
 frmImprimirCupones.gsAlcance = Left(Trim(TxtNumSorteo.Text), 2)
 
End Sub

Private Sub cmdNuevo_Click()
Dim oSorteo As DSorteo
Set oSorteo = New DSorteo


If cmdNuevo.Caption = "&Nuevo" Then

' If Cargousu(Vusuario) = "006001" Then
' 'And User.AreaCod = "026"
'
'    If oSorteo.GetSorteoEstados(gsAlcance, Year(gdFecSis), "I") = True Then
'      MsgBox "No se pude iniciar nuevo sorteo, ya existe uno iniciado.", vbOKOnly + vbExclamation, "AVISO"
'      Exit Sub
'    End If
'
' End If
 bNuevo = True
    LimpiaDatos
    TxtFecha.Enabled = True
    TxtHora.Enabled = True
    txtObs.Enabled = True
    
    
    
    TxtFecha.SetFocus
    TxtNumSorteo.Text = GeneraNumeroSorteo(gsAlcance, Year(gdFecSis))
Else
    TxtFecha.Enabled = True
    TxtHora.Enabled = True
    txtObs.Enabled = True
End If
  
  
  If gsAlcance <> "00" Then
        fraAgencia.Enabled = True
  End If
  
  
  cmdNuevo.Enabled = False
  cmdGrabar.Enabled = True
  cmdCancelar.Enabled = True
  fraProcesos.Enabled = False
  cmdEstado.Enabled = False
  
End Sub

Private Sub cmdPrepararPlanilla_Click()

If MsgBox("Este proceso podría tardar varios minutos." & vbCrLf & "Esta seguro de procesarlo.", vbYesNo + vbQuestion, "Aviso") = vbYes Then

    Screen.MousePointer = vbHourglass

        Dim lnAge As Integer
        Dim oPrevio As Previo.clsPrevio, lscadena As String, cCodAgeAnt As String
        Dim orep As nCaptaReportes
        Set orep = New nCaptaReportes
        Set oPrevio = New Previo.clsPrevio
         lscadena = ""
         cCodAgeAnt = ""
        
         If gsAlcance = "00" Then
           
                For lnAge = 1 To frmSelectAgencias.List1.ListCount
                    If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
                        lscadena = lscadena & orep.PLANILLACONSOLSORTEO(Left(Trim(TxtNumSorteo.Text), 6), Left(frmSelectAgencias.List1.List(lnAge - 1), 2), cCodAgeAnt, IIf(chkCupon.value = vbChecked, 1, 2))
                        cCodAgeAnt = Left(frmSelectAgencias.List1.List(lnAge - 1), 2)
                    End If
                Next lnAge
           
         Else
                lscadena = orep.PLANILLACONSOLSORTEO(Left(Trim(TxtNumSorteo.Text), 6), gsCodAge, , IIf(chkCupon.value = vbChecked, 1, 2))

         End If
                
          If Trim(lscadena) = "" Then
            MsgBox "No hay información para esta consulta", vbOKOnly + vbInformation, "AVISO"
              Screen.MousePointer = vbDefault
            Exit Sub
         End If
    
                
                
            If Me.optImpresion(0).value = True Then
                Set oPrevio = New Previo.clsPrevio
                    oPrevio.Show lscadena, "PLANILLA PARA EL SORTEO " & Trim(TxtNumSorteo.Text), True
                Set oPrevio = Nothing
            ElseIf Me.optImpresion(1).value = True Then
                Set oPrevio = New Previo.clsPrevio
                    oPrevio.PrintSpool sLpt, lscadena, True
                Set oPrevio = Nothing
                
            Else
                dlgGrabar.CancelError = True
                dlgGrabar.InitDir = App.path
                dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
                dlgGrabar.ShowSave
                If dlgGrabar.FileName <> "" Then
                   Open dlgGrabar.FileName For Output As #1
                    Print #1, lscadena
                    Close #1
                End If
            End If
        '     oPrevio.Show lsCadena, Caption, True, 66
             Set oPrevio = Nothing
             
           Screen.MousePointer = vbDefault
 End If

End Sub

Private Sub cmdSalir_Click()
 Unload Me
End Sub
Private Sub CargaDatos(ByVal rsDatos As Recordset, Optional ByVal bAlcance As Boolean = False)

   
  
            TxtNumSorteo.Text = rsDatos!cnumsorteo
            TxtFecha.Text = Format(rsDatos!dFecha, "dd/mm/yyyy")
            TxtHora.Text = Format(rsDatos!dhora, "hh:mm")
            txtEstado.Text = rsDatos!sEstado
            txtObs.Text = rsDatos!CDESCRIPCION
            If bAlcance = False Then
                Call CargaAgencias(gsAlcance)
                If rsDatos!cAlcance <> "00" Then
                   For i = 0 To cboAlcance.ListCount - 1
                     If Right(cboAlcance.List(i), 2) = rsDatos!cAlcance Then
                       cboAlcance.ListIndex = i
                       Exit Sub
                     End If
                    Next i
                End If
            End If
'    cmdNuevo.Caption = "&Editar"
End Sub
Private Sub LimpiaDatos()
    TxtNumSorteo.Text = ""
    TxtFecha.Text = "00/00/0000"
    TxtHora.Text = "00:00"
    txtEstado.Text = ""
    txtObs.Text = ""
    Call CargaAgencias(gsAlcance)


End Sub

Public Sub Inicia(ByVal cAlcance As String)
Dim oSorteo As DSorteo, rsDatos As Recordset
 Set oSorteo = New DSorteo
    
    If cAlcance <> "00" Then
        If cAlcance = "13" Or cAlcance = "03" Or cAlcance = "06" Then
          gsAlcance = "13"
        ElseIf cAlcance = "04" Or cAlcance = "05" Then
          gsAlcance = "04"
        ElseIf cAlcance = "01" Or cAlcance = "14" Then
          gsAlcance = "01"
        ElseIf Not (cAlcance = "13" And cAlcance = "03" And cAlcance = "06" And cAlcance = "04" And cAlcance = "05" And cAlcance = "01" And cAlcance = "14") Then
          gsAlcance = cAlcance
        End If
    Else
        gsAlcance = cAlcance
    End If
        
      Set rsDatos = oSorteo.GetSorteo("I", gsAlcance)
      User.Inicio gsCodUser
      
      
      If User.AreaCod = "026" Then
            cmdEstado.Visible = False
            Me.cmdPrepararPlanilla.Enabled = False
            Me.cmdImpConsolidados.Enabled = False
            Me.cmdImprimirBatch.Enabled = False
            Me.cmProcesar.Enabled = False
'            Me.cmdNuevo.Enabled = False
'            Me.cmdCancelar.Enabled = False
            Me.FraInfo.Enabled = False
      End If
        
        LimpiaDatos
        If Not (rsDatos.EOF Or rsDatos.BOF) Then
            CargaDatos rsDatos
            cmdEditar.Enabled = True
            
        End If
            CargaParametros
        
        If gsAlcance = "00" Then
            cmdAgencia.Visible = True
            Me.chkProcesoOtros.Visible = True
            
        Else
            cmdAgencia.Visible = False
            Me.chkProcesoOtros.Visible = False
        End If
    
        Me.Show
        Exit Sub
              
        
End Sub

Private Sub cmProcesar_Click()
    Dim oSorteo As DSorteo, clsMov As NContFunciones, sMovNro As String

        Set oSorteo = New DSorteo
        Set clsMov = New NContFunciones
        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set clsMov = Nothing
  Screen.MousePointer = vbHourglass
  
  If Me.chkProcesoOtros.Visible = False Then
    If CInt(Right(TxtNumSorteo, 2)) = 1 Then
      Call oSorteo.ProcesarPCtasSorteoPrimer(Trim(TxtNumSorteo.Text), gsMontoxCuponS, gsMontoxCuponD, gsMinPlazoOtorgado, gsLimMaxOtorgadoS, gsLimMaxOtorgadoD, gsAlcance, sMovNro)
      
    Else
      Call oSorteo.ProcesarPCtasSorteoOtros(Trim(TxtNumSorteo.Text), gsMontoxCuponS, gsMontoxCuponD, gsMinPlazoOtorgado, gsLimMaxOtorgadoS, gsLimMaxOtorgadoD, gsAlcance, sMovNro)
      
    End If
  Else
    If CInt(Right(TxtNumSorteo, 2)) = 1 And Me.chkProcesoOtros.value = vbUnchecked Then
        Call oSorteo.ProcesarPCtasSorteoPrimer(Trim(TxtNumSorteo.Text), gsMontoxCuponS, gsMontoxCuponD, gsMinPlazoOtorgado, gsLimMaxOtorgadoS, gsLimMaxOtorgadoD, gsAlcance, sMovNro)
        
    ElseIf CInt(Right(TxtNumSorteo, 2)) = 1 And Me.chkProcesoOtros.value = vbChecked Then
        Call oSorteo.ProcesarPCtasSorteoGenOA(Trim(TxtNumSorteo.Text), gsMontoxCuponS, gsMontoxCuponD, gsMinPlazoOtorgado, gsLimMaxOtorgadoS, gsLimMaxOtorgadoD, gsAlcance, sMovNro)
      
    End If
  
  End If
  
  MsgBox "Proceso Terminado ...", vbInformation + vbOKOnly, "Aviso"
  Screen.MousePointer = vbDefault
  

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then SendKeys "{TAB}"
 
End Sub

Private Sub Form_Load()
 
 Me.KeyPreview = True
 Me.Left = 2925
 Me.Top = 1875
 
End Sub


