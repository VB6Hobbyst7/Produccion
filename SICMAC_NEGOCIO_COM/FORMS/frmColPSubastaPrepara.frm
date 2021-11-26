VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmColPSubastaPrepara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Preparación de la Subasta"
   ClientHeight    =   4830
   ClientLeft      =   975
   ClientTop       =   2385
   ClientWidth     =   7110
   Icon            =   "frmColPSubastaPrepara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dlgGrabar 
      Left            =   4200
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Subasta"
      Height          =   4545
      Left            =   240
      TabIndex        =   13
      Top             =   105
      Width           =   6675
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5430
         TabIndex        =   26
         Top             =   3930
         Width           =   975
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   2880
         TabIndex        =   8
         Top             =   1485
         Width           =   975
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   5295
         TabIndex        =   3
         Top             =   300
         Width           =   1200
      End
      Begin VB.TextBox txtNumSubasta 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   750
         TabIndex        =   0
         Top             =   315
         Width           =   645
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5400
         TabIndex        =   10
         Top             =   1485
         Width           =   975
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4200
         TabIndex        =   9
         Top             =   1485
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Height          =   1695
         Left            =   165
         TabIndex        =   19
         Top             =   1995
         Width           =   6360
         Begin VB.Frame Frame7 
            Caption         =   "Impresión"
            Height          =   1050
            Left            =   4800
            TabIndex        =   30
            Top             =   150
            Width           =   1395
            Begin VB.OptionButton optImpresion 
               Caption         =   "Archivo"
               Height          =   270
               Index           =   2
               Left            =   165
               TabIndex        =   33
               Top             =   720
               Width           =   990
            End
            Begin VB.OptionButton optImpresion 
               Caption         =   "Pantalla"
               Height          =   195
               Index           =   0
               Left            =   165
               TabIndex        =   32
               Top             =   240
               Value           =   -1  'True
               Width           =   960
            End
            Begin VB.OptionButton optImpresion 
               Caption         =   "Impresora"
               Height          =   285
               Index           =   1
               Left            =   165
               TabIndex        =   31
               Top             =   450
               Width           =   990
            End
         End
         Begin VB.CommandButton cmdImpAvisSuba 
            Caption         =   "Cartas de Aviso de Subasta"
            Height          =   360
            Left            =   360
            TabIndex        =   29
            Top             =   270
            Width           =   4215
         End
         Begin VB.CommandButton cmdAgencia 
            Caption         =   "A&gencias..."
            Height          =   345
            Left            =   5010
            TabIndex        =   28
            Top             =   1260
            Width           =   1020
         End
         Begin VB.CommandButton cmdImpListSubasta 
            Caption         =   "Listado de Contratos para Subasta"
            Enabled         =   0   'False
            Height          =   360
            Left            =   360
            TabIndex        =   12
            Top             =   1155
            Width           =   4215
         End
         Begin VB.CommandButton cmdImpPlanSubasta 
            Caption         =   "Planilla para Subasta"
            Enabled         =   0   'False
            Height          =   360
            Left            =   360
            TabIndex        =   11
            Top             =   720
            Width           =   4215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Precios del Oro "
         Height          =   600
         Left            =   135
         TabIndex        =   14
         Top             =   735
         Width           =   6375
         Begin VB.TextBox txtPreOro21 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   5430
            MaxLength       =   6
            TabIndex        =   7
            Top             =   210
            Width           =   750
         End
         Begin VB.TextBox txtPreOro18 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   3795
            MaxLength       =   6
            TabIndex        =   6
            Top             =   210
            Width           =   750
         End
         Begin VB.TextBox txtPreOro16 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2250
            MaxLength       =   6
            TabIndex        =   5
            Top             =   225
            Width           =   750
         End
         Begin VB.TextBox txtPreOro14 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   615
            MaxLength       =   6
            TabIndex        =   4
            Top             =   225
            Width           =   750
         End
         Begin VB.Label Label5 
            Caption         =   "14 K :"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   255
            Width           =   540
         End
         Begin VB.Label Label5 
            Caption         =   "16 K :"
            Height          =   225
            Index           =   2
            Left            =   1755
            TabIndex        =   17
            Top             =   255
            Width           =   525
         End
         Begin VB.Label Label5 
            Caption         =   "18 K :"
            Height          =   225
            Index           =   3
            Left            =   3315
            TabIndex        =   16
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label5 
            Caption         =   "21 K :"
            Height          =   225
            Index           =   4
            Left            =   4950
            TabIndex        =   15
            Top             =   240
            Width           =   540
         End
      End
      Begin MSMask.MaskEdBox txtFecSubasta 
         Height          =   315
         Left            =   2115
         TabIndex        =   1
         Top             =   300
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHorSubasta 
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
         Left            =   3825
         TabIndex        =   2
         Top             =   285
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin RichTextLib.RichTextBox rtfImp 
         Height          =   360
         Left            =   6150
         TabIndex        =   24
         Top             =   4080
         Visible         =   0   'False
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   635
         _Version        =   393217
         TextRTF         =   $"frmColPSubastaPrepara.frx":030A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtfCartas 
         Height          =   360
         Left            =   5325
         TabIndex        =   25
         Top             =   4080
         Visible         =   0   'False
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   635
         _Version        =   393217
         TextRTF         =   $"frmColPSubastaPrepara.frx":038B
      End
      Begin MSComctlLib.ProgressBar prgList 
         Height          =   330
         Left            =   150
         TabIndex        =   27
         Top             =   3960
         Visible         =   0   'False
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label3 
         Caption         =   "Estado :"
         Height          =   255
         Index           =   1
         Left            =   4665
         TabIndex        =   23
         Top             =   345
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Número :"
         Height          =   255
         Left            =   105
         TabIndex        =   22
         Top             =   330
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "Hora :"
         Height          =   255
         Index           =   0
         Left            =   3375
         TabIndex        =   21
         Top             =   330
         Width           =   525
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha :"
         Height          =   255
         Left            =   1545
         TabIndex        =   20
         Top             =   345
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmColPSubastaPrepara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************
'* MANTENIMIENTO DE CONTRATO PIGNORATICIO
'Archivo:  frmColPSubastaPreparacion.frm
'LAYG   :  15/06/2001.
'Resumen:  Nos permite ingresar o actualizar los precios del oro con
'          que van a ser procesados los listado y planillas para subasta
Option Explicit

Dim pCorte As Variant
Dim pPrevioMax As Double
Dim pLineasMax As Double
Dim pHojaFiMax As Integer

Dim fnTasaImpuesto As Double
Dim fnTasaPreparacionRemate As Double
Dim fnTasaIGV As Double
Dim pVerCodAnt As Boolean

Dim fsSubastaCadaAgencia As String
Dim fnJoyasDet As Integer

Dim vNumAvisos As Integer
Dim MuestraImpresion As Boolean
Dim vRTFImp As String, vBuffer As String
Dim vCont As Double
Dim vNomAge As String

Private Sub cmdAgencia_Click()
    frmSelectAgencias.Inicio Me
    frmSelectAgencias.Show 1
End Sub
Private Sub HabilitaControles(ByVal pbEditar As Boolean, ByVal pbGrabar As Boolean, ByVal pbSalir As Boolean, _
    ByVal pbCancelar As Boolean, ByVal pbFecSubasta As Boolean, ByVal pbHorSubasta As Boolean, _
    ByVal pbPreOro14 As Boolean, ByVal pbPreOro16 As Boolean, ByVal pbPreOro18 As Boolean, ByVal pbPreOro21 As Boolean, _
    ByVal pbImpAvisSuba As Boolean, ByVal pbImpPlanSubasta As Boolean, ByVal pbImpListSubasta As Boolean)

    cmdEditar.Enabled = pbEditar
    cmdGrabar.Enabled = pbGrabar
    cmdSalir.Enabled = pbSalir
    cmdCancelar.Enabled = pbCancelar
    txtFecSubasta.Enabled = pbFecSubasta
    txtHorSubasta.Enabled = pbHorSubasta
    txtPreOro14.Enabled = pbPreOro14
    txtPreOro16.Enabled = pbPreOro16
    txtPreOro18.Enabled = pbPreOro18
    txtPreOro21.Enabled = pbPreOro21
    cmdImpAvisSuba.Enabled = pbImpAvisSuba
    cmdImpPlanSubasta.Enabled = pbImpPlanSubasta
    cmdImpListSubasta.Enabled = pbImpListSubasta
   
End Sub
'Permite no reconocer los datos ingresados
Private Sub cmdCancelar_Click()

On Error GoTo ControlError

Call HabilitaControles(True, False, True, False, False, False, False, False, False, False, True, False, False)
Limpiar
VeriDatRem
If txtPreOro14 > 0 And txtPreOro16 > 0 And txtPreOro18 > 0 And txtPreOro21 > 0 Then
    cmdImpListSubasta.Enabled = True
    cmdImpPlanSubasta.Enabled = True
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Permite editar los campos editables de una subasta
Private Sub CmdEditar_Click()
    Call HabilitaControles(False, True, False, True, True, True, _
        True, True, True, True, False, False, False)
    txtFecSubasta.SetFocus
End Sub

'permite grabar los cambios ingresados
Private Sub cmdGrabar_Click()

On Error GoTo ControlError
Dim loGrabSub As COMNColoCPig.NCOMColPRecGar

Call HabilitaControles(True, False, True, False, False, False, False, False, False, False, True, True, True)

If txtPreOro16 > 0 And txtPreOro18 > 0 And txtPreOro21 > 0 Then
    cmdImpListSubasta.Enabled = True
    cmdImpPlanSubasta.Enabled = True
End If

Set loGrabSub = New COMNColoCPig.NCOMColPRecGar
    Call loGrabSub.nRecGarGrabaDatosPreparaCredPignoraticio("S", Me.txtNumSubasta, gColPRecGarEstNoIniciado, Format(Me.txtFecSubasta.Text, "mm/dd/yyyy hh:mm"), fsSubastaCadaAgencia, Val(Me.txtPreOro14), Val(Me.txtPreOro16), Val(Me.txtPreOro18), Val(Me.txtPreOro21), , , , False)
Set loGrabSub = Nothing
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub cmdImpAvisSuba_Click()
On Error GoTo ControlError
Dim loImprime As COMNColoCPig.NCOMColPRecGar
Dim lsCadImprimir  As String
Dim lsmensaje As String
Dim loPrevio As previo.clsPrevio

Dim lnAge As Integer
    
    lsCadImprimir = ""
    rtfCartas.FileName = App.path & cPlantillaAvisoSubasta
    
    For lnAge = 1 To frmSelectAgencias.List1.ListCount
        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
            
            Set loImprime = New COMNColoCPig.NCOMColPRecGar
                lsCadImprimir = lsCadImprimir & loImprime.nImprimeAvisoSubasta(rtfCartas.Text, Format(Me.txtFecSubasta.Text, "mm/dd/yyyy"), _
                        Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, gdFecSis, lsmensaje, gImpresora)
                If Trim(lsmensaje) <> "" Then
                     MsgBox lsmensaje, vbInformation, "Aviso"
                     Exit Sub
                End If
            Set loImprime = Nothing
                
        End If
    Next lnAge
    
    If Me.optImpresion(0).value = True Then
        Set loPrevio = New previo.clsPrevio
            loPrevio.Show lsCadImprimir, "Cartas Aviso de Subasta ", False
        Set loPrevio = Nothing
    Else
        dlgGrabar.CancelError = True
        dlgGrabar.InitDir = App.path
        dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
        dlgGrabar.ShowSave
        If dlgGrabar.FileName <> "" Then
           Open dlgGrabar.FileName For Output As #1
            Print #1, vBuffer
            Close #1
        End If
    End If

Exit Sub

ControlError:   ' Rutina de control de errores.
    If Err.Number = 32755 Then
        MsgBox " Grabación Cancelada ", vbInformation, " Aviso "
    Else
        MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
            " Avise al Area de Sistemas ", vbInformation, " Aviso "
    End If
    
End Sub

Private Sub cmdSalir_Click()
Unload frmSelectAgencias
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
 CargaParametros
 Limpiar
End Sub

Private Sub Limpiar()

Dim loDatos As COMNColoCPig.NCOMColPRecGar
Dim lrDatosSub As ADODB.Recordset
Dim lsUltSubasta As String
Dim lsmensaje As String

txtPreOro14 = Format(0, "#0.00")
txtPreOro16 = Format(0, "#0.00")
txtPreOro18 = Format(0, "#0.00")
txtPreOro21 = Format(0, "#0.00")

Set lrDatosSub = New ADODB.Recordset
Set loDatos = New COMNColoCPig.NCOMColPRecGar
    lsUltSubasta = loDatos.nObtieneNroUltimoProceso("S", fsSubastaCadaAgencia, lsmensaje)
    If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        Exit Sub
    End If
    Set lrDatosSub = loDatos.nObtieneDatosProcesoRGCredPig("S", lsUltSubasta, fsSubastaCadaAgencia, lsmensaje)
    If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        Exit Sub
    End If
Set loDatos = Nothing
If lrDatosSub Is Nothing Then Exit Sub
'Mostrar Datos
Me.txtNumSubasta = lrDatosSub!cNroProceso
Me.txtFecSubasta = Format(lrDatosSub!dProceso, "dd/mm/yyyy")
Me.txtHorSubasta = Format(lrDatosSub!dProceso, "hh:mm")
Me.txtPreOro14 = Format(lrDatosSub!nPrecioK14, "#0.00")
Me.txtPreOro16 = Format(lrDatosSub!nPrecioK16, "#0.00")
Me.txtPreOro18 = Format(lrDatosSub!nPrecioK18, "#0.00")
Me.txtPreOro21 = Format(lrDatosSub!nPrecioK21, "#0.00")

If lrDatosSub!nRGEstado = gColPRecGarEstNoIniciado Then
    txtEstado = "NO INICIADO"
    If Val(txtPreOro14) > 0 And Val(txtPreOro16) > 0 And Val(txtPreOro18) > 0 _
        And Val(txtPreOro21) > 0 Then
        cmdImpListSubasta.Enabled = True
        cmdImpPlanSubasta.Enabled = True
    End If
ElseIf lrDatosSub!nRGEstado = gColPRecGarEstIniciado Then
    txtEstado = "INICIADO"
    cmdEditar.Enabled = False
Else
    MsgBox " No existe el remate generado", vbCritical, " Error de Sistema "
    cmdEditar.Enabled = False
    txtNumSubasta = ""
    txtFecSubasta = Format("01/01/2000", "dd/mm/yyyy")
    txtHorSubasta = Format("00:00", "hh:mm")
End If
Set lrDatosSub = Nothing

End Sub

'Valida el campo txtFecSubasta
Private Sub txtFecSubasta_GotFocus()
fEnfoque txtFecSubasta
End Sub
Private Sub txtFecSubasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtHorSubasta.SetFocus
End If
End Sub
Private Sub txtFecSubasta_LostFocus()
If Not ValFecha(txtFecSubasta) Then
    txtFecSubasta.SetFocus
ElseIf DateDiff("d", txtFecSubasta, gdFecSis) > 0 Then
    MsgBox " Fecha no debe ser anterior a la fecha actual", vbInformation, " Aviso "
    txtFecSubasta.SetFocus
End If
'VeriDatRem
End Sub

Private Sub txtFecSubasta_Validate(Cancel As Boolean)
If Not ValFecha(txtFecSubasta) Then
    Cancel = True
ElseIf DateDiff("d", txtFecSubasta, gdFecSis) > 0 Then
    MsgBox " Fecha no debe ser anterior a la fecha actual", vbInformation, " Aviso "
    Cancel = True
End If

End Sub

'Valida el campo txtHorSubasta
Private Sub txtHorSubasta_GotFocus()
fEnfoque txtHorSubasta
End Sub
Private Sub txtHorSubasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPreOro14.SetFocus
End If
End Sub
Private Sub txtHorSubasta_LostFocus()
If Not ValidaHora(txtHorSubasta) Then
    txtHorSubasta.SetFocus
End If
'VeriDatRem
End Sub

'Valida el campo txtpreoro14
Private Sub txtPreOro14_GotFocus()
fEnfoque txtPreOro14
End Sub
Private Sub txtPreOro14_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreOro14, KeyAscii)
If KeyAscii = 13 Then
    txtPreOro14 = Format(txtPreOro14, "#0.00")
    txtPreOro16.SetFocus
End If
End Sub
Private Sub txtPreOro14_LostFocus()
    VeriPreOro
End Sub

'Valida el campo txtpreoro16
Private Sub txtPreOro16_GotFocus()
fEnfoque txtPreOro16
End Sub
Private Sub txtPreOro16_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreOro16, KeyAscii)
If KeyAscii = 13 Then
    txtPreOro16 = Format(txtPreOro16, "#0.00")
    txtPreOro18.SetFocus
End If
End Sub
Private Sub txtPreOro16_LostFocus()
    VeriPreOro
End Sub

'Valida el campo txtpreoro18
Private Sub txtPreOro18_GotFocus()
fEnfoque txtPreOro18
End Sub
Private Sub txtPreOro18_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreOro18, KeyAscii)
If KeyAscii = 13 Then
    txtPreOro18 = Format(txtPreOro18, "#0.00")
    txtPreOro21.SetFocus
End If
End Sub
Private Sub txtPreOro18_LostFocus()
    VeriPreOro
End Sub

'Valida el campo txtpreoro21
Private Sub txtPreOro21_GotFocus()
fEnfoque txtPreOro21
End Sub
Private Sub txtPreOro21_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreOro21, KeyAscii)
If KeyAscii = 13 Then
    txtPreOro21 = Format(txtPreOro21, "#0.00")
    'cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
End If
End Sub
Private Sub txtPreOro21_LostFocus()
    VeriPreOro
End Sub

' Permite activar la opción de procesar solo cuando están ingresados los campos
' fecha y hora de Subasta
Public Sub VeriDatRem()
    If Len(txtNumSubasta) > 0 And Len(txtFecSubasta) = 10 And Len(txtHorSubasta) = 5 Then
    End If
End Sub

' Permite activar la opción de grabar solo cuando están ingresados los campos
' del precio del oro
Public Sub VeriPreOro()
    If Val(txtPreOro14) > 0 And Val(txtPreOro16) > 0 And Val(txtPreOro18) > 0 _
        And Val(txtPreOro21) > 0 Then
        cmdGrabar.Enabled = True
    End If
End Sub

'Permite imprimir en pantalla o directo a la impresora la planilla de subasta
Private Sub cmdImpPlanSubasta_Click()
On Error GoTo ControlError
Dim loImprime As COMNColoCPig.NCOMColPRecGar
Dim lsCadImprimir  As String
Dim lsmensaje As String
Dim loPrevio As previo.clsPrevio

Dim lnAge As Integer
   
    lsCadImprimir = ""
    
    For lnAge = 1 To frmSelectAgencias.List1.ListCount
        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
            
            Set loImprime = New COMNColoCPig.NCOMColPRecGar
                lsCadImprimir = lsCadImprimir & loImprime.nImprimePlanillaParaSubasta(Format(Me.txtFecSubasta.Text, "mm/dd/yyyy"), _
                        Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, gdFecSis, _
                        fnTasaIGV, CCur(Val(Me.txtPreOro14.Text)), CCur(Val(Me.txtPreOro16.Text)), _
                        CCur(Val(Me.txtPreOro18.Text)), CCur(Val(Me.txtPreOro21.Text)), gsNomCmac, gsNomAge, gsCodUser, Me.txtNumSubasta.Text, lsmensaje, gImpresora)
                        If Trim(lsmensaje) <> "" Then
                             MsgBox lsmensaje, vbInformation, "Aviso"
                             Exit Sub
                        End If
            Set loImprime = Nothing
                
        End If
    Next lnAge
    
    If Me.optImpresion(0).value = True Then
        Set loPrevio = New previo.clsPrevio
            loPrevio.Show lsCadImprimir, "Cartas Aviso de Vencimiento", True
        Set loPrevio = Nothing
    Else
        dlgGrabar.CancelError = True
        dlgGrabar.InitDir = App.path
        dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
        dlgGrabar.ShowSave
        If dlgGrabar.FileName <> "" Then
           Open dlgGrabar.FileName For Output As #1
            Print #1, vBuffer
            Close #1
        End If
    End If

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
           " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

'Permite imprimir los listados de subasta
Private Sub cmdImpListSubasta_Click()

On Error GoTo ControlError
Dim loImprime As COMNColoCPig.NCOMColPRecGar
Dim lsCadImprimir  As String
Dim lsmensaje As String
Dim loPrevio As previo.clsPrevio

Dim lnAge As Integer
   
    lsCadImprimir = ""
    
    For lnAge = 1 To frmSelectAgencias.List1.ListCount
        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
            
            Set loImprime = New COMNColoCPig.NCOMColPRecGar
                lsCadImprimir = lsCadImprimir & loImprime.nImprimeListadoParaSubasta(Format(Me.txtFecSubasta.Text, "mm/dd/yyyy"), _
                        Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), 66, gdFecSis, _
                        fnTasaIGV, CCur(Val(Me.txtPreOro14.Text)), CCur(Val(Me.txtPreOro16.Text)), _
                        CCur(Val(Me.txtPreOro18.Text)), CCur(Val(Me.txtPreOro21.Text)), gsNomCmac, gsNomAge, gsCodUser, Me.txtNumSubasta.Text, lsmensaje, gImpresora)
                        If Trim(lsmensaje) <> "" Then
                             MsgBox lsmensaje, vbInformation, "Aviso"
                             Exit Sub
                        End If
            Set loImprime = Nothing
                
        End If
    Next lnAge
    
    If Me.optImpresion(0).value = True Then
        Set loPrevio = New previo.clsPrevio
            loPrevio.Show lsCadImprimir, "Cartas Aviso de Vencimiento", False
        Set loPrevio = Nothing
    Else
        dlgGrabar.CancelError = True
        dlgGrabar.InitDir = App.path
        dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
        dlgGrabar.ShowSave
        If dlgGrabar.FileName <> "" Then
           Open dlgGrabar.FileName For Output As #1
            Print #1, vBuffer
            Close #1
        End If
    End If

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
           " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
'Cabecera de las Impresiones
Private Sub Cabecera(ByVal vOpt As String, ByVal vPagina As Integer, Optional ByVal pPagCorta As Boolean = True)
    Dim vTitulo As String
    Dim vSubTit As String
    Dim vArea As String * 30
    Dim vNroLineas As Integer
    Select Case vOpt
        Case "ListSuba"
            vTitulo = "LISTADO GENERAL DE CONTRATOS PARA LA SUBASTA N°: " & Format(txtNumSubasta, "@@@@")
        Case "PlanSuba"
            vTitulo = "PLANILLA DE PRENDAS PARA LA SUBASTA N°: " & Format(txtNumSubasta, "@@@@") & " DEL " & txtFecSubasta
    End Select
    vSubTit = "(Previo a la Subasta)"
    vArea = "Crédito Pignoraticio"
    vNroLineas = IIf(pPagCorta = True, 105, 162)
    'Centra Título
    vTitulo = String(Round((vNroLineas - Len(Trim(vTitulo))) / 2), " ") & vTitulo
    'Centra SubTítulo
    vSubTit = String(Round(((vNroLineas - 60) - Len(Trim(vSubTit))) / 2), " ") & vSubTit & String(Round(((vNroLineas - 60) - Len(Trim(vSubTit))) / 2), " ")

    vRTFImp = vRTFImp & pCorte
    vRTFImp = vRTFImp & Space(1) & ImpreFormat(vNomAge, 25, 0) & Space(vNroLineas - 40) & "Página: " & Format(vPagina, "@@@@") & pCorte
    vRTFImp = vRTFImp & Space(1) & vTitulo & pCorte
    vRTFImp = vRTFImp & Space(1) & vArea & vSubTit & Space(10) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & pCorte
    vRTFImp = vRTFImp & String(vNroLineas, "-") & pCorte
    Select Case vOpt
        Case "ListSuba"
            vRTFImp = vRTFImp & Space(1) & "ITEM    CONTRATO     PZ     DESCRIPCION                    14Kl.    16Kl.    18Kl.    21Kl.      PRECIO" & pCorte
            vRTFImp = vRTFImp & Space(1) & "                                                                                                  BASE" & pCorte
        Case "PlanSuba"
            vRTFImp = vRTFImp & Space(1) & "ITEM    CONTRATO         FECHA      FECHA           NOMBRE CLIENTE              GRAMOS          PRECIO" & pCorte
            vRTFImp = vRTFImp & Space(1) & "                        VENCIMI.   ADJUDIC.                                                      BASE" & pCorte
    End Select
    vRTFImp = vRTFImp & String(vNroLineas, "-") & pCorte
End Sub

Private Sub CargaParametros()

Dim loParam As COMDColocPig.DCOMColPCalculos
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim lnProcesoCadaAgencia As Integer

Set loParam = New COMDColocPig.DCOMColPCalculos
    fnTasaImpuesto = loParam.dObtieneColocParametro(gConsColPTasaImpuesto)
    fnTasaPreparacionRemate = loParam.dObtieneColocParametro(gConsColPTasaPreparaRemate)
    fnTasaIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)
Set loParam = Nothing


Set loConstSis = New COMDConstSistema.NCOMConstSistema
    fnJoyasDet = loConstSis.LeeConstSistema(109) ' Joyas en Detalle
    lnProcesoCadaAgencia = loConstSis.LeeConstSistema(121)  ' gConstSistPigRemateCadaAg
    If lnProcesoCadaAgencia = 1 Then  ' En cada agencia
        fsSubastaCadaAgencia = gsCodCMAC & gsCodAge
    Else
        fsSubastaCadaAgencia = gsCodCMAC & "00"
    End If
Set loConstSis = Nothing

    pPrevioMax = 5000
    pLineasMax = 56
    pHojaFiMax = 66
End Sub

Private Function FormatoContratro(pContrato As String) As String
  FormatoContratro = Mid(pContrato, 1, 2) & "-" & Mid(pContrato, 3, 4) & "-" & Mid(pContrato, 7, 5) & "-" & Mid(pContrato, 12, 1)
End Function

