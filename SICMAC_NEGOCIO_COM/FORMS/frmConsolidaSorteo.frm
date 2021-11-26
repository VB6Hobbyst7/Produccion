VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmConsolidaSorteo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONSOLIDACION DE SORTEO"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   Icon            =   "frmConsolidaSorteo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraContenedor 
      Height          =   4950
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   7680
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
         Left            =   6270
         TabIndex        =   15
         Top             =   225
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
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   210
         Width           =   1650
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   390
         Left            =   6525
         TabIndex        =   13
         Top             =   4410
         Width           =   990
      End
      Begin VB.Frame fraContenedor 
         Height          =   2520
         Index           =   1
         Left            =   105
         TabIndex        =   5
         Top             =   1815
         Width           =   7470
         Begin VB.CommandButton cmdListadoCuponesAnulados 
            Caption         =   "Listado de Cupones Anuladas"
            Height          =   360
            Left            =   255
            TabIndex        =   25
            Top             =   1905
            Width           =   3645
         End
         Begin VB.CommandButton cmdListadoCuponesCancelados 
            Caption         =   "Listado de Cupones Cancelados"
            Height          =   360
            Left            =   255
            TabIndex        =   24
            Top             =   1500
            Width           =   3645
         End
         Begin VB.CommandButton cmdListado 
            Caption         =   "Consolidados de Resultados"
            Height          =   360
            Left            =   255
            TabIndex        =   23
            Top             =   1095
            Width           =   3645
         End
         Begin VB.CommandButton cmdAgencia 
            Caption         =   "A&gencias..."
            Height          =   345
            Left            =   4905
            TabIndex        =   12
            Top             =   1440
            Width           =   1140
         End
         Begin VB.Frame fraImpresion 
            Caption         =   "Impresión"
            Height          =   1050
            Left            =   4890
            TabIndex        =   8
            Top             =   300
            Visible         =   0   'False
            Width           =   2220
            Begin VB.OptionButton optImpresion 
               Caption         =   "Pantalla"
               Height          =   195
               Index           =   0
               Left            =   495
               TabIndex        =   11
               Top             =   285
               Value           =   -1  'True
               Width           =   960
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
               Caption         =   "Archivo"
               Height          =   270
               Index           =   2
               Left            =   495
               TabIndex        =   9
               Top             =   720
               Width           =   990
            End
         End
         Begin VB.CommandButton cmProcesar 
            Caption         =   "Actualizar Informacion Sorteo"
            Height          =   360
            Left            =   255
            TabIndex        =   7
            Top             =   285
            Width           =   3645
         End
         Begin VB.CommandButton cmdPrepararPlanilla 
            Caption         =   "Planilla de Ganadores"
            Height          =   360
            Left            =   255
            TabIndex        =   6
            Top             =   690
            Width           =   3645
         End
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
         Left            =   105
         TabIndex        =   3
         Top             =   1140
         Width           =   4815
         Begin VB.ComboBox cboAlcance 
            Height          =   315
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   4575
         End
      End
      Begin VB.TextBox txtObs 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1575
         MaxLength       =   500
         TabIndex        =   2
         Top             =   675
         Width           =   6000
      End
      Begin VB.CommandButton cmdEstado 
         Caption         =   "Cerrar Proceso"
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   4455
         Width           =   1260
      End
      Begin MSMask.MaskEdBox TxtFecha 
         Height          =   315
         Left            =   3120
         TabIndex        =   16
         Top             =   225
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
         Left            =   4815
         TabIndex        =   17
         Top             =   225
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
         Left            =   5655
         TabIndex        =   22
         Top             =   285
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
         Left            =   120
         TabIndex        =   21
         Top             =   285
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
         Left            =   4395
         TabIndex        =   20
         Top             =   285
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
         Left            =   2550
         TabIndex        =   19
         Top             =   285
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
         Left            =   120
         TabIndex        =   18
         Top             =   765
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmConsolidaSorteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gsAlcance As String
Private Sub LimpiaDatos()
    TxtNumSorteo.Text = ""
    txtFecha.Text = "00/00/0000"
    TxtHora.Text = "00:00"
    txtEstado.Text = ""
    txtObs.Text = ""
    Call CargaAgencias(gsAlcance)


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
                    
                ElseIf rsTemp!cAgeCod = "13" Then 'And RSTEMP!cAgeCod <> "03" And RSTEMP!cAgeCod <> "06" Then
                    cboAlcance.AddItem "San Vic-Imperial-Mala" & Space(100 - Len(Trim(rsTemp!cAgeDescripcion))) & "13"
                
                ElseIf rsTemp!cAgeCod = "04" Then 'And RSTEMP!cAgeCod <> "05" Then
                    cboAlcance.AddItem "Nasca-Palpa" & Space(100 - Len(Trim(rsTemp!cAgeDescripcion))) & "04"
                    
                ElseIf rsTemp!cAgeCod = "01" Then 'And RSTEMP!cAgeCod <> "05" Then
                    cboAlcance.AddItem "Ica-Parcona" & Space(100 - Len(Trim(rsTemp!cAgeDescripcion))) & "01"
                
                End If
                rsTemp.MoveNext
        Wend
     End If
   If cboAlcance.ListCount > 0 Then cboAlcance.ListIndex = 0
   Set rsTemp = Nothing
 
End Sub
Private Sub CargaDatos(ByVal rsDatos As Recordset)

    Call CargaAgencias(gsAlcance)
  
            TxtNumSorteo.Text = rsDatos!cnumsorteo
            txtFecha.Text = Format(rsDatos!dfecha, "dd/mm/yyyy")
            TxtHora.Text = Format(rsDatos!dhora, "hh:mm")
            txtEstado.Text = rsDatos!sEstado
            txtObs.Text = rsDatos!CDESCRIPCION
            If rsDatos!cAlcance <> "00" Then
               For i = 0 To cboAlcance.ListCount - 1
                 If Right(cboAlcance.List(i), 2) = rsDatos!cAlcance Then
                   cboAlcance.ListIndex = i
                   Exit Sub
                 End If
                Next i
            End If
  
    
End Sub

Public Sub Inicia(ByVal cAlcance As String)
Dim oSorteo As DSorteo, rsDatos As Recordset
 Set oSorteo = New DSorteo
    
    gsAlcance = cAlcance
        
       Set rsDatos = oSorteo.GetSorteo("P", cAlcance)
        
       
        
        LimpiaDatos
        If Not rsDatos.EOF Then
            CargaDatos rsDatos
             
        End If
'            CargaParametros
        
        If gsAlcance = "00" Then
            cmdAgencia.Visible = True
        Else
            cmdAgencia.Visible = False
        End If
    
        Me.Show
        Exit Sub
              
        
End Sub

Private Sub cmdAgencia_Click()
    frmSelectAgencias.Inicio Me
    frmSelectAgencias.Show 1
End Sub

Private Sub cmdEstado_Click()
Dim oSorteo As DSorteo
Set oSorteo = New DSorteo

If MsgBox("ADVERTENCIA!!!: Esta seguro de cerrar este sorteo.??? ", vbYesNo + vbQuestion, "AVISO") = vbYes Then
  Call oSorteo.ActualizaSorteo(Trim(TxtNumSorteo.Text), Trim(txtObs.Text), 0, 0, 0, "C", Format(txtFecha.Text & " " & Trim(TxtHora.Text) & ":00", "yyyy-mm-dd hh:mm:ss"), sMovNro)
  Me.fraContenedor(1).Enabled = False
  Me.cmdEstado.Enabled = False
End If
End Sub




Private Sub cmdListado_Click()
Dim lnAge As Integer

Dim oPrevio As previo.clsprevio, lscadena As String
Dim orep As nCaptaReportes
Set orep = New nCaptaReportes
Set oPrevio = New previo.clsprevio
 Dim cCodAgeAnt As String, contal As Integer

 lscadena = ""
 cCodAgeAnt = ""


Screen.MousePointer = vbHourglass

On Error GoTo ErrMsg

 If gsAlcance = "00" Then
   
        For lnAge = 1 To frmSelectAgencias.List1.ListCount
            If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
                lscadena = lscadena & orep.CONSOLIDADOSORTEO(Left(Trim(TxtNumSorteo.Text), 6), Left(frmSelectAgencias.List1.List(lnAge - 1), 2), cCodAgeAnt, True, contal)
                cCodAgeAnt = Left(frmSelectAgencias.List1.List(lnAge - 1), 2)
                    
            End If
        Next lnAge
    
   
 Else
    lscadena = orep.CONSOLIDADOSORTEO(Left(Trim(TxtNumSorteo.Text), 6), gsCodAge, cCodAgeAnt, True)

 End If
        
    If lscadena = "" Then
        MsgBox "No hay información para esta consulta", vbOKOnly + vbInformation, "AVISO"
          Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
     'oPrevio.Show lscadena, Caption, True, 66
     oPrevio.Show lscadena, Caption, True, 66, gImpresora

     Set oPrevio = Nothing
     
 Screen.MousePointer = vbDefault
Exit Sub

ErrMsg:
   MsgBox "Error: " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "AVISO"
   Screen.MousePointer = vbDefault
     
End Sub

Private Sub cmdListadoCuponesAnulados_Click()
 Dim lnAge As Integer

Dim oPrevio As previo.clsprevio, lscadena As String
Dim orep As nCaptaReportes
Set orep = New nCaptaReportes
Set oPrevio = New previo.clsprevio
 Dim cCodAgeAnt As String

 lscadena = ""
 cCodAgeAnt = ""
 
 
Screen.MousePointer = vbHourglass

On Error GoTo ErrMsg

 If gsAlcance = "00" Then
   
        For lnAge = 1 To frmSelectAgencias.List1.ListCount
            If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
                lscadena = lscadena & orep.CONSOLESTADO(Left(Trim(TxtNumSorteo.Text), 6), Left(frmSelectAgencias.List1.List(lnAge - 1), 2), cCodAgeAnt, "A")
                cCodAgeAnt = Left(frmSelectAgencias.List1.List(lnAge - 1), 2)
                    
            End If
        Next lnAge
    
   
 Else
                lscadena = orep.CONSOLESTADO(Left(Trim(TxtNumSorteo.Text), 6), gsCodAge, cCodAgeAnt, "A")

 End If
 
    If lscadena = "" Then
        MsgBox "No hay información para esta consulta", vbOKOnly + vbInformation, "AVISO"
          Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    
     'oPrevio.Show lscadena, Caption, True, 66
     oPrevio.Show lscadena, Caption, True, 66, gImpresora

     Set oPrevio = Nothing
     
  Screen.MousePointer = vbDefault
Exit Sub

ErrMsg:
   MsgBox "Error: " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "AVISO"
   Screen.MousePointer = vbDefault

End Sub

Private Sub cmdListadoCuponesCancelados_Click()
Dim lnAge As Integer

Dim oPrevio As previo.clsprevio, lscadena As String
Dim orep As nCaptaReportes
Set orep = New nCaptaReportes
Set oPrevio = New previo.clsprevio
 Dim cCodAgeAnt As String

 lscadena = ""
 cCodAgeAnt = ""
 
 
Screen.MousePointer = vbHourglass

On Error GoTo ErrMsg

 If gsAlcance = "00" Then
   
        For lnAge = 1 To frmSelectAgencias.List1.ListCount
            If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
               lscadena = lscadena & orep.CONSOLESTADO(Left(Trim(TxtNumSorteo.Text), 6), Left(frmSelectAgencias.List1.List(lnAge - 1), 2), cCodAgeAnt, "C")
                cCodAgeAnt = Left(frmSelectAgencias.List1.List(lnAge - 1), 2)
                    
            End If
        Next lnAge
    
   
 Else
   lscadena = orep.CONSOLESTADO(Left(Trim(TxtNumSorteo.Text), 6), gsCodAge, cCodAgeAnt, "C")

 End If
        
    If lscadena = "" Then
        MsgBox "No hay información para esta consulta", vbOKOnly + vbInformation, "AVISO"
          Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
     'oPrevio.Show lscadena, Caption, True, 66
     oPrevio.Show lscadena, Caption, True, 66, gImpresora

     Set oPrevio = Nothing
     
  Screen.MousePointer = vbDefault
Exit Sub

ErrMsg:
   MsgBox "Error: " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "AVISO"
   Screen.MousePointer = vbDefault
     
End Sub

Private Sub cmdPrepararPlanilla_Click()
Dim lnAge As Integer

Dim oPrevio As previo.clsprevio, lscadena As String
Dim orep As nCaptaReportes
Set orep = New nCaptaReportes
Set oPrevio = New previo.clsprevio
 Dim cCodAgeAnt As String

Screen.MousePointer = vbHourglass

On Error GoTo ErrMsg
 cCodAgeAnt = ""

 If gsAlcance = "00" Then
   
        For lnAge = 1 To frmSelectAgencias.List1.ListCount
            If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
                lscadena = lscadena & orep.PLANILLAGANADORES(Left(Trim(TxtNumSorteo.Text), 6), Left(frmSelectAgencias.List1.List(lnAge - 1), 2), cCodAgeAnt)
                cCodAgeAnt = Left(frmSelectAgencias.List1.List(lnAge - 1), 2)
                    
            End If
        Next lnAge
    
   
 Else
    lscadena = orep.PLANILLAGANADORES(Left(Trim(TxtNumSorteo.Text), 6), gsCodAge, cCodAgeAnt)

 End If
    If lscadena = "" Then
        MsgBox "No hay información para esta consulta", vbOKOnly + vbInformation, "AVISO"
          Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
     'oPrevio.Show lscadena, Caption, True, 66
     oPrevio.Show lscadena, Caption, True, 66, gImpresora

     Set oPrevio = Nothing
     
Screen.MousePointer = vbDefault
Exit Sub

ErrMsg:
   MsgBox "Error: " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "AVISO"
   Screen.MousePointer = vbDefault
     
     
End Sub

Private Sub cmdSalir_Click()
 Unload Me
 
End Sub

Private Sub cmProcesar_Click()
Dim cn As ADODB.Connection
Set cn = New ADODB.Connection
Dim sCadena As String, rsTemp As Recordset, ssql As String
Dim oSorteo As DSorteo
    Set oSorteo = New DSorteo
    
On Error GoTo ErrMsg
Screen.MousePointer = vbHourglass
sCadena = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\SORTEO\" & Left(Me.TxtNumSorteo.Text, 2) & "\dbpruebas.mdb;Persist Security Info=False"
cn.ConnectionString = sCadena
cn.Open
ssql = " SELECT * FROM  TEMPSORTEO where bganador=true"
Set rsTemp = New Recordset
Set rsTemp = cn.Execute(ssql)
        
  While Not rsTemp.EOF
    With rsTemp
       Call oSorteo.ActualizaCtasSorteo(Left(!cnumsorteo, 6), !cCtaCod, !bEntregado, !bGanador, !Nroganador, !bcancelar)
    End With
       rsTemp.MoveNext
  Wend
   rsTemp.Close
   Set rsTemp = Nothing
   cn.Close
   Set cn = Nothing
Screen.MousePointer = vbDefault
Exit Sub
ErrMsg:
    Screen.MousePointer = vbDefault
    MsgBox "Error:" & Err.Description
  
End Sub

