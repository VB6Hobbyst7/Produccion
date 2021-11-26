VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReporteSorteo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "INFORMACION  DE SORTEO"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "frmReporteSorteo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraContenedor 
      Height          =   4260
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8025
      Begin VB.Frame Frame1 
         Caption         =   "Busqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Left            =   105
         TabIndex        =   24
         Top             =   135
         Width           =   7785
         Begin VB.CheckBox chkTipo 
            Caption         =   "Agencias"
            Height          =   255
            Index           =   1
            Left            =   6210
            TabIndex        =   33
            Top             =   360
            Value           =   1  'Checked
            Width           =   1260
         End
         Begin VB.CheckBox chkTipo 
            Caption         =   "General"
            Height          =   255
            Index           =   0
            Left            =   4995
            TabIndex        =   31
            Top             =   360
            Width           =   1035
         End
         Begin VB.ComboBox cboNumSorteo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   420
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   720
            Width           =   6750
         End
         Begin VB.CheckBox chkEstado 
            Caption         =   "Iniciado"
            Height          =   255
            Index           =   0
            Left            =   885
            TabIndex        =   27
            Top             =   345
            Value           =   1  'Checked
            Width           =   1035
         End
         Begin VB.CheckBox chkEstado 
            Caption         =   "Cerrado"
            Height          =   255
            Index           =   2
            Left            =   3135
            TabIndex        =   26
            Top             =   345
            Width           =   1005
         End
         Begin VB.CheckBox chkEstado 
            Caption         =   "Proceso"
            Height          =   255
            Index           =   1
            Left            =   2025
            TabIndex        =   25
            Top             =   345
            Width           =   1050
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Tipo:"
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
            Index           =   2
            Left            =   4425
            TabIndex        =   32
            Top             =   345
            Width           =   480
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Estado:"
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
            Index           =   3
            Left            =   120
            TabIndex        =   29
            Top             =   345
            Width           =   720
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
            TabIndex        =   28
            Top             =   840
            Width           =   660
         End
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   390
         Left            =   6870
         TabIndex        =   12
         Top             =   3795
         Width           =   990
      End
      Begin VB.Frame fraContenedor 
         Height          =   2355
         Index           =   1
         Left            =   105
         TabIndex        =   5
         Top             =   1380
         Width           =   7785
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
            Left            =   4095
            TabIndex        =   35
            Top             =   165
            Width           =   975
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
            Left            =   5130
            TabIndex        =   34
            Top             =   165
            Value           =   1  'Checked
            Width           =   1065
         End
         Begin VB.CommandButton cmdPlanillaInicial 
            Caption         =   "Planilla para el Sorteo"
            Height          =   360
            Left            =   255
            TabIndex        =   23
            Top             =   195
            Width           =   3645
         End
         Begin VB.CommandButton cmdListadoCuponesAnulados 
            Caption         =   "Listado de Cupones Anuladas"
            Height          =   360
            Left            =   255
            TabIndex        =   22
            Top             =   1845
            Width           =   3645
         End
         Begin VB.CommandButton cmdListadoCuponesCancelados 
            Caption         =   "Listado de Cupones Cancelados"
            Height          =   360
            Left            =   255
            TabIndex        =   21
            Top             =   1440
            Width           =   3645
         End
         Begin VB.CommandButton cmdListado 
            Caption         =   "Consolidados de Resultados"
            Height          =   360
            Left            =   255
            TabIndex        =   20
            Top             =   1035
            Width           =   3645
         End
         Begin VB.CommandButton cmdAgencia 
            Caption         =   "A&gencias..."
            Height          =   345
            Left            =   5250
            TabIndex        =   11
            Top             =   1890
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Frame fraImpresion 
            Caption         =   "Impresión"
            Height          =   1050
            Left            =   5235
            TabIndex        =   7
            Top             =   750
            Visible         =   0   'False
            Width           =   2220
            Begin VB.OptionButton optImpresion 
               Caption         =   "Pantalla"
               Height          =   195
               Index           =   0
               Left            =   495
               TabIndex        =   10
               Top             =   285
               Value           =   -1  'True
               Width           =   960
            End
            Begin VB.OptionButton optImpresion 
               Caption         =   "Impresora"
               Height          =   270
               Index           =   1
               Left            =   495
               TabIndex        =   9
               Top             =   480
               Width           =   990
            End
            Begin VB.OptionButton optImpresion 
               Caption         =   "Archivo"
               Height          =   270
               Index           =   2
               Left            =   495
               TabIndex        =   8
               Top             =   720
               Width           =   990
            End
         End
         Begin VB.CommandButton cmdPrepararPlanilla 
            Caption         =   "Planilla de Ganadores"
            Height          =   360
            Left            =   255
            TabIndex        =   6
            Top             =   630
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
         Top             =   2355
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
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   345
         Left            =   1575
         MaxLength       =   500
         TabIndex        =   2
         Top             =   2925
         Width           =   6000
      End
      Begin VB.CommandButton cmdEstado 
         Caption         =   "Cerrar Proceso"
         Height          =   360
         Left            =   75
         TabIndex        =   1
         Top             =   3795
         Visible         =   0   'False
         Width           =   1260
      End
      Begin MSMask.MaskEdBox TxtFecha 
         Height          =   315
         Left            =   720
         TabIndex        =   14
         Top             =   2475
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
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
         Left            =   2730
         TabIndex        =   15
         Top             =   2475
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   6225
         TabIndex        =   13
         Top             =   2475
         Width           =   1320
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
         Left            =   5610
         TabIndex        =   19
         Top             =   2535
         Width           =   585
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
         Left            =   2310
         TabIndex        =   18
         Top             =   2535
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
         Left            =   120
         TabIndex        =   17
         Top             =   2535
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
         TabIndex        =   16
         Top             =   3015
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmReporteSorteo"
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
    'Call CargaAgencias(gsAlcance)


End Sub
Private Sub CargaSorteos()
Dim oSorteo As DSorteo, rsTemp As ADODB.Recordset, i As Integer
Dim cEstado As String, cTipo As String, clsAge As DAgencias

cboNumSorteo.Clear
   Set clsAge = New DAgencias
   Set oSorteo = New DSorteo
   If chkEstado(0).value = vbChecked Then
        cEstado = "I"
   ElseIf chkEstado(1).value = vbChecked Then
        cEstado = "P"
   ElseIf chkEstado(2).value = vbChecked Then
        cEstado = "C"
   End If
   
   If chkTipo(0).value = vbChecked Then
        cTipo = "00"
   Else
        cTipo = "XX"
   End If
   
   If chkTipo(1).value = vbChecked Then
     Set rsTemp = oSorteo.GetSorteo(cEstado, "", True)
   Else
     Set rsTemp = oSorteo.GetSorteo(cEstado, "00", False)
   End If
   
   i = 0
   
   
   
     If cTipo = "00" Then
        cboNumSorteo.AddItem "<SELECCIONE SORTEO>" & Space(100 - Len("<SELECCIONE SORTEO>")) & "00"
        
        While Not rsTemp.EOF
            cboNumSorteo.AddItem Trim(rsTemp!cnumsorteo) & Space(2) & "***GENERAL***"
            rsTemp.MoveNext
            
        Wend
        
     Else
        cboNumSorteo.AddItem "<SELECCIONE SORTEO>" & Space(100 - Len("<SELECCIONE SORTEO>")) & "00"
        While Not rsTemp.EOF
        
                If Left(rsTemp!cnumsorteo, 2) <> "13" And Left(rsTemp!cnumsorteo, 2) <> "01" And Left(rsTemp!cnumsorteo, 2) <> "04" Then
                    cboNumSorteo.AddItem Trim(rsTemp!cnumsorteo) & Space(2) & clsAge.NombreAgencia(Left(rsTemp!cnumsorteo, 2))
                ElseIf Left(rsTemp!cnumsorteo, 2) = "01" Then
                    cboNumSorteo.AddItem Trim(rsTemp!cnumsorteo) & Space(2) & "Ica-Parcona"
                ElseIf Left(rsTemp!cnumsorteo, 2) = "13" Then
                    cboNumSorteo.AddItem Trim(rsTemp!cnumsorteo) & Space(2) & "San Vic-Imperial-Mala"
                ElseIf Left(rsTemp!cnumsorteo, 2) = "04" Then
                    cboNumSorteo.AddItem Trim(rsTemp!cnumsorteo) & Space(2) & "Nasca-Palpa"
                End If
                
                rsTemp.MoveNext
                
         Wend
     End If
     
   If cboNumSorteo.ListCount > 0 Then cboNumSorteo.ListIndex = 0
   Set rsTemp = Nothing
 
End Sub
Private Sub CargaDatos(ByVal rsDatos As Recordset)

    'Call CargaAgencias(gsAlcance)
  
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
          '  CargaParametros
        
'        If gsAlcance = "00" Then
'            cmdAgencia.Visible = True
'        Else
'            cmdAgencia.Visible = False
'        End If
'
        Me.Show
        Exit Sub
              
        
End Sub

Private Sub cboNumSorteo_Click()
If cboNumSorteo.ListCount > 0 Then
    If cboNumSorteo.ListIndex > 0 Then
        gsAlcance = Left(cboNumSorteo.Text, 2)
        If gsAlcance = "00" Then
            cmdAgencia.Visible = True
        Else
            cmdAgencia.Visible = False
        End If
    Else
        gsAlcance = ""
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

Private Sub chkEstado_Click(Index As Integer)
Select Case Index
   Case 0
            If chkEstado(0).value = vbChecked Then
                chkEstado(1).value = vbUnchecked
                chkEstado(2).value = vbUnchecked
                
            Else
                chkEstado(1).value = vbChecked
                chkEstado(2).value = vbUnchecked
                
            End If
   
   Case 1
            If chkEstado(1).value = vbChecked Then
                chkEstado(2).value = vbUnchecked
                chkEstado(0).value = vbUnchecked
                
            Else
                chkEstado(2).value = vbChecked
                chkEstado(0).value = vbUnchecked
                
            End If
   
   Case 2
            If chkEstado(2).value = vbChecked Then
                chkEstado(0).value = vbUnchecked
                chkEstado(1).value = vbUnchecked
                
            Else
                chkEstado(0).value = vbChecked
                chkEstado(1).value = vbUnchecked
                
            End If
   
End Select
CargaSorteos

End Sub

Private Sub chkTipo_Click(Index As Integer)

Select Case Index
    Case 0
           If chkTipo(0).value = vbChecked Then
                chkTipo(1).value = vbUnchecked
           Else
                chkTipo(1).value = vbChecked
           End If
    
    Case 1
           If chkTipo(1).value = vbChecked Then
                chkTipo(0).value = vbUnchecked
           Else
                chkTipo(0).value = vbChecked
           End If
    
End Select
CargaSorteos

End Sub





Private Sub cmdAgencia_Click()
    frmSelectAgencias.Inicio Me
    frmSelectAgencias.Show 1
End Sub

Private Sub cmdBuscar_Click()

End Sub

Private Sub cmdEstado_Click()
Dim oSorteo As DSorteo
Set oSorteo = New DSorteo

If MsgBox("ADVERTENCIA!!!: Esta seguro de cerrar este sorteo.??? ") = vbYes Then
'  Call oSorteo.ActualizaSorteo(Trim(TxtNumSorteo.Text), Trim(txtObs.Text), 0, 0, 0, "C", Format(TxtFecha.Text & " " & Trim(TxtHora.Text) & ":00", "yyyy-mm-dd hh:mm:ss"), sMovNro)
  
End If
End Sub




Private Sub cmdListado_Click()
Dim lnAge As Integer

If cboNumSorteo.ListCount > 0 Then
    If cboNumSorteo.ListIndex = 0 Then
        MsgBox "Seleccione un sorteo de la lista", vbOKOnly + vbInformation, "Aviso"
        Exit Sub
    End If
End If



Dim oPrevio As previo.clsprevio, lscadena As String
Dim orep As nCaptaReportes
Set orep = New nCaptaReportes
Set oPrevio = New previo.clsprevio
 Dim cCodAgeAnt As String, contal As Integer

 lscadena = ""
 cCodAgeAnt = ""

 If gsAlcance = "00" Then
   
        For lnAge = 1 To frmSelectAgencias.List1.ListCount
            If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
                lscadena = lscadena & orep.CONSOLIDADOSORTEO(Left(cboNumSorteo.Text, 6), Left(frmSelectAgencias.List1.List(lnAge - 1), 2), cCodAgeAnt, True, contal)
                cCodAgeAnt = Left(frmSelectAgencias.List1.List(lnAge - 1), 2)
                    
            End If
        Next lnAge
    
   
 Else
    lscadena = orep.CONSOLIDADOSORTEO(Left(cboNumSorteo.Text, 6), gsAlcance, cCodAgeAnt, True)

 End If
     If Trim(lscadena) = "" Then
            MsgBox "No hay información para esta consulta", vbOKOnly + vbInformation, "AVISO"
              Screen.MousePointer = vbDefault
            Exit Sub
     End If
    
    
     oPrevio.Show lscadena, Caption, True, 66, gImpresora

     Set oPrevio = Nothing
End Sub

Private Sub cmdListadoCuponesAnulados_Click()
 Dim lnAge As Integer
 
 If cboNumSorteo.ListCount > 0 Then
    If cboNumSorteo.ListIndex = 0 Then
        MsgBox "Seleccione un sorteo de la lista", vbOKOnly + vbInformation, "Aviso"
        Exit Sub
    End If
End If

 

Dim oPrevio As previo.clsprevio, lscadena As String
Dim orep As nCaptaReportes
Set orep = New nCaptaReportes
Set oPrevio = New previo.clsprevio
 Dim cCodAgeAnt As String

 lscadena = ""
 cCodAgeAnt = ""

 If gsAlcance = "00" Then
        For lnAge = 1 To frmSelectAgencias.List1.ListCount
            If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
                lscadena = lscadena & orep.CONSOLESTADO(Left(cboNumSorteo.Text, 6), Left(frmSelectAgencias.List1.List(lnAge - 1), 2), cCodAgeAnt, "A")
                cCodAgeAnt = Left(frmSelectAgencias.List1.List(lnAge - 1), 2)
            End If
        Next lnAge
 Else
        lscadena = orep.CONSOLESTADO(Left(cboNumSorteo.Text, 6), gsAlcance, cCodAgeAnt, "A")

 End If
        
    
    
     oPrevio.Show lscadena, Caption, True, 66, gImpresora

     Set oPrevio = Nothing

End Sub

Private Sub cmdListadoCuponesCancelados_Click()
Dim lnAge As Integer

If cboNumSorteo.ListCount > 0 Then
    If cboNumSorteo.ListIndex = 0 Then
        MsgBox "Seleccione un sorteo de la lista", vbOKOnly + vbInformation, "Aviso"
        Exit Sub
    End If
End If


Dim oPrevio As previo.clsprevio, lscadena As String
Dim orep As nCaptaReportes
Set orep = New nCaptaReportes
Set oPrevio = New previo.clsprevio
 Dim cCodAgeAnt As String

 lscadena = ""
 cCodAgeAnt = ""

 If gsAlcance = "00" Then
   
        For lnAge = 1 To frmSelectAgencias.List1.ListCount
            If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
               lscadena = lscadena & orep.CONSOLESTADO(Left(cboNumSorteo.Text, 6), Left(frmSelectAgencias.List1.List(lnAge - 1), 2), cCodAgeAnt, "C")
                cCodAgeAnt = Left(frmSelectAgencias.List1.List(lnAge - 1), 2)
                    
            End If
        Next lnAge
    
   
 Else
   lscadena = orep.CONSOLESTADO(Left(cboNumSorteo.Text, 6), gsAlcance, cCodAgeAnt, "C")

 End If
        
     If Trim(lscadena) = "" Then
            MsgBox "No hay información para esta consulta", vbOKOnly + vbInformation, "AVISO"
              Screen.MousePointer = vbDefault
            Exit Sub
     End If
    
     oPrevio.Show lscadena, Caption, True, 66, gImpresora

     Set oPrevio = Nothing
End Sub

Private Sub cmdPlanillaInicial_Click()
If cboNumSorteo.ListCount > 0 Then
    If cboNumSorteo.ListIndex = 0 Then
        MsgBox "Seleccione un sorteo de la lista", vbOKOnly + vbInformation, "Aviso"
        Exit Sub
    End If
End If

If MsgBox("Este proceso podría tardar varios minutos." & vbCrLf & "Esta seguro de procesarlo.", vbYesNo + vbQuestion, "Aviso") = vbYes Then

    Screen.MousePointer = vbHourglass

        Dim lnAge As Integer
        Dim oPrevio As previo.clsprevio, lscadena As String, cCodAgeAnt As String
        Dim orep As nCaptaReportes
        Set orep = New nCaptaReportes
        Set oPrevio = New previo.clsprevio
         lscadena = ""
         cCodAgeAnt = ""
        
         If gsAlcance = "00" Then
           
                For lnAge = 1 To frmSelectAgencias.List1.ListCount
                    If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
                        lscadena = lscadena & orep.PLANILLACONSOLSORTEO(Left(cboNumSorteo.Text, 6), Left(frmSelectAgencias.List1.List(lnAge - 1), 2), cCodAgeAnt, IIf(chkCupon.value = vbChecked, 1, 2))
                        cCodAgeAnt = Left(frmSelectAgencias.List1.List(lnAge - 1), 2)
                    End If
                Next lnAge
           
         Else
                lscadena = orep.PLANILLACONSOLSORTEO(Left(cboNumSorteo.Text, 6), gsAlcance, , IIf(chkCupon.value = vbChecked, 1, 2))

         End If
                
          If Trim(lscadena) = "" Then
            MsgBox "No hay información para esta consulta", vbOKOnly + vbInformation, "AVISO"
              Screen.MousePointer = vbDefault
            Exit Sub
         End If
    
                
                
            If Me.optImpresion(0).value = True Then
                Set oPrevio = New previo.clsprevio
                    oPrevio.Show lscadena, "PLANILLA PARA EL SORTEO " & Left(cboNumSorteo.Text, 6), True, , gImpresora
                Set oPrevio = Nothing
            ElseIf Me.optImpresion(1).value = True Then
                Set oPrevio = New previo.clsprevio
                    oPrevio.PrintSpool sLpt, lscadena, True
                Set oPrevio = Nothing
                
            Else
                dlgGrabar.CancelError = True
                dlgGrabar.InitDir = App.path
                dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
                dlgGrabar.ShowSave
                If dlgGrabar.Filename <> "" Then
                   Open dlgGrabar.Filename For Output As #1
                    Print #1, lscadena
                    Close #1
                End If
            End If
        '     oPrevio.Show lsCadena, Caption, True, 66
             Set oPrevio = Nothing
             
           Screen.MousePointer = vbDefault
 End If

End Sub

Private Sub cmdPrepararPlanilla_Click()
Dim lnAge As Integer

If cboNumSorteo.ListCount > 0 Then
    If cboNumSorteo.ListIndex = 0 Then
        MsgBox "Seleccione un sorteo de la lista", vbOKOnly + vbInformation, "Aviso"
        Exit Sub
    End If
End If




Dim oPrevio As previo.clsprevio, lscadena As String
Dim orep As nCaptaReportes
Set orep = New nCaptaReportes
Set oPrevio = New previo.clsprevio
 Dim cCodAgeAnt As String

 cCodAgeAnt = ""

 If gsAlcance = "00" Then
   
        For lnAge = 1 To frmSelectAgencias.List1.ListCount
            If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
                lscadena = lscadena & orep.PLANILLAGANADORES(Left(cboNumSorteo.Text, 6), Left(frmSelectAgencias.List1.List(lnAge - 1), 2), cCodAgeAnt)
                cCodAgeAnt = Left(frmSelectAgencias.List1.List(lnAge - 1), 2)
                    
            End If
        Next lnAge
    
   
 Else
    lscadena = orep.PLANILLAGANADORES(Left(cboNumSorteo.Text, 6), gsAlcance, cCodAgeAnt)

 End If
      If Trim(lscadena) = "" Then
            MsgBox "No hay información para esta consulta", vbOKOnly + vbInformation, "AVISO"
              Screen.MousePointer = vbDefault
            Exit Sub
     End If
    
     oPrevio.Show lscadena, Caption, True, 66, gImpresora
     

     Set oPrevio = Nothing
End Sub

Private Sub cmdSalir_Click()
 Unload Me
 
End Sub



Private Sub TxtNumSorteo_Change()

End Sub

Private Sub Form_Load()
CargaSorteos
End Sub

