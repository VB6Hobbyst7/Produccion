VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmColPRepoTransac 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio : Listados Diarios"
   ClientHeight    =   4590
   ClientLeft      =   2775
   ClientTop       =   1935
   ClientWidth     =   7440
   HelpContextID   =   200
   Icon            =   "frmColPRepoTransac.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboUsuario 
      Height          =   315
      ItemData        =   "frmColPRepoTransac.frx":030A
      Left            =   3465
      List            =   "frmColPRepoTransac.frx":0311
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4065
      Visible         =   0   'False
      Width           =   1365
   End
   Begin RichTextLib.RichTextBox rtfImp 
      Height          =   315
      Left            =   7110
      TabIndex        =   7
      Top             =   4035
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   556
      _Version        =   393217
      TextRTF         =   $"frmColPRepoTransac.frx":0324
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
   Begin VB.Frame fraContenedor 
      Height          =   3300
      Index           =   1
      Left            =   210
      TabIndex        =   6
      Top             =   540
      Width           =   7005
      Begin VB.CommandButton cmdBoveda 
         Caption         =   "&Bovedas ..."
         Height          =   375
         Left            =   5670
         TabIndex        =   14
         Top             =   2790
         Width           =   1005
      End
      Begin VB.CommandButton cmdAgencia 
         Caption         =   "A&gencias ..."
         Height          =   375
         Left            =   5670
         TabIndex        =   11
         Top             =   2340
         Width           =   1005
      End
      Begin VB.Frame fraImpresion 
         Caption         =   "Impresión"
         Height          =   885
         Left            =   5475
         TabIndex        =   8
         Top             =   1350
         Width           =   1350
         Begin VB.OptionButton optImpresion 
            Caption         =   "Impresora"
            Height          =   270
            Index           =   1
            Left            =   180
            TabIndex        =   5
            Top             =   540
            Width           =   990
         End
         Begin VB.OptionButton optImpresion 
            Caption         =   "Pantalla"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   4
            Top             =   315
            Value           =   -1  'True
            Width           =   960
         End
      End
      Begin VB.ListBox lstListados 
         Height          =   2985
         ItemData        =   "frmColPRepoTransac.frx":039E
         Left            =   120
         List            =   "frmColPRepoTransac.frx":03A0
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   210
         Width           =   5235
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4005
      Width           =   1065
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6060
      TabIndex        =   3
      Top             =   4005
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar prgList 
      Height          =   330
      Left            =   210
      TabIndex        =   10
      Top             =   4050
      Visible         =   0   'False
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSMask.MaskEdBox txtFecha 
      Height          =   315
      Left            =   1380
      TabIndex        =   12
      Top             =   135
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Fecha :"
      Height          =   225
      Left            =   540
      TabIndex        =   13
      Top             =   165
      Width           =   585
   End
   Begin VB.Label lblUsuario 
      Caption         =   "Usuario"
      Height          =   195
      Left            =   3495
      TabIndex        =   9
      Top             =   3840
      Visible         =   0   'False
      Width           =   690
   End
End
Attribute VB_Name = "frmColPRepoTransac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modificacion de Bases: CASL 04.12.2000
'---------------------------------------'

'REPORTES.
'Archivo:              frmListados.frm
'Fecha de creación   : -------------
'Fecha de modificación   : 30/06/1999.
'Resumen:
'   El Proceso de REPORTES DEL DIA nos permite generar impresiones de acuerdo a
'   la opción escogida entre todas las opciones permitidas.

'Variables el formulario
Option Compare Text
Option Explicit
Dim RegCredPrend As New ADODB.Recordset
Dim RegTransDiariaPrend As New ADODB.Recordset
Dim sSql As String
Dim pbListGene As Boolean
Dim MuestraImpresion As Boolean
Dim pPrevioMax As Double
Dim pLineasMax As Double
Dim pHojaFiMax As Integer
Dim vNameForm As String
Dim vRTFImp As String
Dim vCont As Double
Dim vNomAge As String
Dim vPage As Integer
Dim vBoveda As String

Private Sub cboUsuario_Click()
Dim I As Long
    Dim lbBan As Boolean
    lbBan = False
    For I = 0 To Me.cboUsuario.ListCount - 1
        If Me.cboUsuario.List(I) = Me.cboUsuario.Text Then
            lbBan = True
            I = Me.cboUsuario.ListCount - 1
        End If
    Next I
    If Not lbBan Or Len(Me.cboUsuario.Text) = 0 Then
        Exit Sub
    End If
End Sub

Private Sub cmdAgencia_Click()
frmPigAgencias.Inicio frmListados
frmPigAgencias.Show 1
End Sub

Private Sub cmdBoveda_Click()
frmPigAgenciaBoveda.Inicio frmListados
frmPigAgenciaBoveda.Show 1
End Sub

'Permite seleccionar la opción a imprimir y además si se visualiza en pantalla
' o va directo a la impresora
Private Sub cmdImprimir_Click()
Dim x As Integer
On Error GoTo ControlError
pPrevioMax = 4000
vRTFImp = ""
vPage = 0
'Carga Boveda Seleccionada
vBoveda = CargaBovedaSelec

MuestraImpresion = False
'Verifica si está preparada la impresora cuando se envia directo a ella
If optImpresion(1).Value = True Then
    If Not ImpreSensa Then Exit Sub
End If
With lstListados
    If .Selected(0) = True Then  'Préstamos No Desembolsados
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vRTFImp = vRTFImp & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    Call ImprimeContRegi(dbCmact)
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        Call ImprimeContRegi(dbCmactN)
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        'ImprimeContRegi
        If Not MuestraImpresion Then
            If Len(vRTFImp) > 0 Then
                vRTFImp = Left(vRTFImp, Len(vRTFImp) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    If .Selected(1) = True Then  'Prendas Nuevas en condición de diferidas
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vRTFImp = vRTFImp & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    Call ImprimePrenNuevCondDife(dbCmact)
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        Call ImprimePrenNuevCondDife(dbCmactN)
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        'ImprimePrenNuevCondDife 'Reporte drpPrenDife
        If Not MuestraImpresion Then
            If Len(vRTFImp) > 0 Then
                vRTFImp = Left(vRTFImp, Len(vRTFImp) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    If .Selected(2) = True Then  'Préstamos Nuevos
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vRTFImp = vRTFImp & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    Call ImprimePresNuev(dbCmact)
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        Call ImprimePresNuev(dbCmactN)
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        'ImprimePresNuev 'Reporte drpPrenDeseNuev
        If Not MuestraImpresion Then
            If Len(vRTFImp) > 0 Then
                vRTFImp = Left(vRTFImp, Len(vRTFImp) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    If .Selected(3) = True Then  'Préstamo Rescatados
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vRTFImp = vRTFImp & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    Call ImprimePresResc(dbCmact)
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        Call ImprimePresResc(dbCmactN)
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        'ImprimePresResc 'Reporte drpPrenDifeDeta
        If Not MuestraImpresion Then
            If Len(vRTFImp) > 0 Then
                vRTFImp = Left(vRTFImp, Len(vRTFImp) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    If .Selected(4) = True Then  'Créditos Renovados
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vRTFImp = vRTFImp & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    Call ImprimeCredReno(dbCmact)
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        Call ImprimeCredReno(dbCmactN)
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        'ImprimeCredReno 'Reporte drpPrenReno
        If Not MuestraImpresion Then
            If Len(vRTFImp) > 0 Then
                vRTFImp = Left(vRTFImp, Len(vRTFImp) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    If .Selected(5) = True Then  'Prendas Nuevas
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vRTFImp = vRTFImp & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    Call ImprimePrenNuev(dbCmact)
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        Call ImprimePrenNuev(dbCmactN)
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        'ImprimePrenNuev 'Reporte drpPrenDese
        If Not MuestraImpresion Then
            If Len(vRTFImp) > 0 Then
                vRTFImp = Left(vRTFImp, Len(vRTFImp) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    
    'Reporte duplicado
'''    If .Selected(6) = True Then  'Prendas Devueltas
'''        For X = 1 To frmPigAgencias.List1.ListCount
'''            If frmPigAgencias.List1.Selected(X - 1) = True Then
'''                If MuestraImpresion Then vRTFImp = vRTFImp & gPrnSaltoPagina
'''                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(X - 1), 1, 2) Then
'''                    vNomAge = gsNomAge
'''                    Call ImprimePrenDevu(dbCmact)
'''                Else
'''                    If AbreConeccion(Mid(frmPigAgencias.List1.List(X - 1), 1, 2) & "XXXXXXXXXX") Then
'''                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralcom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
'''                        Call ImprimePrenDevu(dbCmactN)
'''                    End If
'''                    CierraConeccion
'''                End If
'''            End If
'''        Next X
'''        'ImprimePrenDevu 'Reporte drpPrenCanc
'''        If Not MuestraImpresion Then
'''            If Len(vRTFImp) > 0 Then
'''                vRTFImp = Left(vRTFImp, Len(vRTFImp) - 1)
'''                MuestraImpresion = True
'''            End If
'''        End If
'''    End If
    If .Selected(6) = True Then  'Contratos Anulados
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vRTFImp = vRTFImp & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    Call ImprimeContAnul(dbCmact)
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        Call ImprimeContAnul(dbCmactN)
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        'ImprimeContAnul 'Reporte drpPrenAnul
        If Not MuestraImpresion Then
            If Len(vRTFImp) > 0 Then
                vRTFImp = Left(vRTFImp, Len(vRTFImp) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    If .Selected(7) = True Then  'Prendas Rescatadas en Condición de Diferidas
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vRTFImp = vRTFImp & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    Call ImprimePrenRescCondDife(dbCmact)
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        Call ImprimePrenRescCondDife(dbCmactN)
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        'ImprimePrenRescCondDife 'Reporte drpPrenCancF
        If Not MuestraImpresion Then
            If Len(vRTFImp) > 0 Then
                vRTFImp = Left(vRTFImp, Len(vRTFImp) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    If .Selected(8) = True Then  'Contrato en condición de Diferido
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vRTFImp = vRTFImp & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    Call ImprimeContCondDife(dbCmact)
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        Call ImprimeContCondDife(dbCmactN)
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        'ImprimeContCondDife
        If Not MuestraImpresion Then
            If Len(vRTFImp) > 0 Then
                vRTFImp = Left(vRTFImp, Len(vRTFImp) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    If .Selected(9) = True Then  'Operaciones Extornadas
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vRTFImp = vRTFImp & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    Call ImprimeOperExto(dbCmact)
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        Call ImprimeOperExto(dbCmactN)
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        'ImprimeOperExto
        If Not MuestraImpresion Then
            If Len(vRTFImp) > 0 Then
                vRTFImp = Left(vRTFImp, Len(vRTFImp) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
    If .Selected(10) = True Then  'Pago de Sobrantes
        For x = 1 To frmPigAgencias.List1.ListCount
            If frmPigAgencias.List1.Selected(x - 1) = True Then
                If MuestraImpresion Then vRTFImp = vRTFImp & gPrnSaltoPagina
                If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
                    vNomAge = gsNomAge
                    Call ImprimePagoSobr(dbCmact)
                Else
                    If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
                        vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
                        Call ImprimePagoSobr(dbCmactN)
                    End If
                    CierraConeccion
                End If
            End If
        Next x
        'ImprimePagoSobr
        If Not MuestraImpresion Then
            If Len(vRTFImp) > 0 Then
                vRTFImp = Left(vRTFImp, Len(vRTFImp) - 1)
                MuestraImpresion = True
            End If
        End If
    End If
End With
'Envia a la impresion Previa
If optImpresion(0).Value = True And Len(Trim(vRTFImp)) > 0 Then
    rtfImp.Text = vRTFImp
    frmPrevio.Previo rtfImp, " Impresiones Generales ", False, pHojaFiMax
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Permite Selecionar o quitar la seleción de las opciones de los listados
Private Sub cmdSeleccion_Click(Index As Integer)
If Index = 0 Then
    OpcListado lstListados, True
Else
    OpcListado lstListados, False
End If
End Sub

'Permite salir del formulario actual
Private Sub cmdSalir_Click()
Unload frmPigAgencias
Unload Me
End Sub

'Permite inicializar el formulario
Private Sub Form_Load()
Dim RegUsu As New ADODB.Recordset
AbreConexion
txtFecha.Text = gdFecSis
gcIntCentra = CentraSdi(Me)
CargaParametros
vNameForm = "Crédito Pignoraticio : Listados Diarios"
With lstListados
    .Clear
    .AddItem "Listado de Contratos Registrados"
    .AddItem "Listado de Prendas Nuevas en condición de Diferidas"
    .AddItem "Listado de Préstamos Nuevos Desembolsados"
    .AddItem "Listado de Préstamos Cancelados"
    .AddItem "Listado de Créditos Renovados"
    .AddItem "Listado de Prendas Nuevas"
'    .AddItem "Listado de Prendas Devueltas"
    .AddItem "Listado de Contratos Anulados"
    .AddItem "Listado de Prendas Rescatadas en Condición de Diferidas"
    .AddItem "Listado de Contratos en Condición de Diferido (solo del momento)"
    .AddItem "Listado de Operaciones Extornadas"
    .AddItem "Listado de Pago de Sobrantes"
End With
If pbListGene Then
    'Carga Usuarios en Combo
    cboUsuario.Visible = True
    lblUsuario.Visible = True
    sSql = "SELECT * FROM " & gcCentralCom & "Usuario Usuario INNER JOIN " & gcCentralCom & "GRUPOUSU GRUPOUSU ON USUARIO.cCodUsu = GRUPOUSU.cCodUsu AND GRUPOUSU.cCodGrp = '0008'"
    RegUsu.Open sSql, dbCmact, adOpenForwardOnly, adLockOptimistic, adCmdText
    GrabaCombo RegUsu, Me.cboUsuario, "cCodUsu"
    cboUsuario.ListIndex = 0
    RegUsu.Close
    Set RegUsu = Nothing
End If
End Sub

'Cabecera de las Impresiones
Private Sub Cabecera(ByVal vOpt As String, ByVal vPagina As Integer)
    Dim vTitulo As String
    Dim vSubTit As String
    'Dim vSpaTit As Integer
    'Dim vSpaSub As Integer
    Dim vArea As String * 30
    Dim vNroLineas As Integer
    vSubTit = ""
    Select Case vOpt
        Case "ContRegi"
            vTitulo = "LISTADO DE LOS CONTRATOS REGISTRADOS  DEL : " & txtFecha.Text
        Case "PrNuCoDi"
            vTitulo = "LISTADO DE PRENDAS NUEVAS EN CONDICION  DE DIFERIDAS  DEL : " & txtFecha
            vSubTit = "(Para la  devolución  de  sus Joyas)"
        Case "ContAnul"
            vTitulo = "LISTADO DE CONTRATOS ANULADOS DEL : " & txtFecha
        Case "PresNuev"
            vTitulo = "LISTADO DE LOS PRESTAMOS NUEVOS DESEMBOLSADOS DEL : " & txtFecha
        Case "PrReCoDi"
            vTitulo = "LISTADO DE PRENDAS RESCATADAS, CONDICION DE DIFERIDAS DEL : " & txtFecha
            'vSubTit = "(Entregadas dentro y fuera del Plazo)" & gPrnSaltoLinea
        Case "CredReno"
            vTitulo = "LISTADO DE LOS CREDITOS RENOVADOS DEL : " & txtFecha
        Case "PrenDevu"
            vTitulo = "LISTADO DE PRENDAS DEVUELTAS DEL : " & txtFecha
            'vSubTit = "(Entregadas dentro del Plazo)" & gPrnSaltoLinea
        Case "PrenNuev"
            vTitulo = "LISTADO DE PRENDAS NUEVAS DEL : " & txtFecha
        Case "PresResc"
            vTitulo = "LISTADO  DE  LOS PRESTAMOS CANCELADOS DEL : " & txtFecha
        Case "ContCondDife"
            vTitulo = "LISTADO DE CONTRATOS EN CONDICION DE DIFERIDOS AL : " & gdFecSis
        Case "OperExto"
            vTitulo = "LISTADO DE OPERACIONES EXTORNADAS DEL : " & txtFecha
        Case "PagoSobr"
            vTitulo = "LISTADO DE  SOBRANTES PAGADOS DEL : " & txtFecha
    End Select
    
    vArea = "Crédito Pignoraticio"
    vNroLineas = 110
    'Centra Título
    vTitulo = String(Round((vNroLineas - Len(Trim(vTitulo))) / 2, 0) - 18, " ") & vTitulo & String(Round((vNroLineas - Len(Trim(vTitulo))) / 2, 0) - 18, " ")
    'Centra SubTítulo
    vSubTit = String(Round((vNroLineas - Len(Trim(vSubTit))) / 2, 0) - 20, " ") & vSubTit & String(Round((vNroLineas - Len(Trim(vSubTit))) / 2, 0) - 20, " ")
    
    'vRTFImp = vRTFImp & gPrnSaltoLinea
    vRTFImp = vRTFImp & Space(1) & ImpreFormat(vNomAge, 20, 0) & vTitulo & Space(2) & "Página: " & Format(vPagina, "@@@@") & gPrnSaltoLinea
    vRTFImp = vRTFImp & Space(1) & ImpreFormat(vArea, 20, 0) & vSubTit & Space(5) & Format(gdFecSis & " " & Time, "dd/mm/yyyy hh:mm") & gPrnSaltoLinea
    vRTFImp = vRTFImp & String(vNroLineas, "-") & gPrnSaltoLinea
    Select Case vOpt
        Case "ContRegi"
            vRTFImp = vRTFImp & Space(1) & "ITEM   PLAZO   CONTRATO      PRESTAMO     INTERES    COSTO DE    COSTO DE    IMPUESTO      NETO A  USER" & gPrnSaltoLinea
            vRTFImp = vRTFImp & Space(1) & "                                                     TASACION    CUSTODIA                  PAGAR" & gPrnSaltoLinea
        Case "PrNuCoDi"
            vRTFImp = vRTFImp & Space(1) & "ITEM     CONTRATO    CONTRATO ANT.     PLAZO     VALOR TASACION       ORO NETO   BLOQ  BOVEDA " & gPrnSaltoLinea
        Case "ContAnul"
            vRTFImp = vRTFImp & Space(1) & "ITEM    CONTRATO            CLIENTE                         PLAZO        PRESTAMO" & gPrnSaltoLinea
        Case "PresNuev"
            vRTFImp = vRTFImp & Space(1) & "ITEM   PLAZO   CONTRATO      PRESTAMO     INTERES    COSTO DE    COSTO DE    IMPUESTO       NETO   USER" & gPrnSaltoLinea
            vRTFImp = vRTFImp & Space(1) & "                                                     TASACION    CUSTODIA                  PAGADO" & gPrnSaltoLinea
        Case "PrReCoDi"
            vRTFImp = vRTFImp & Space(1) & "ITEM  PLAZO   CONTRATO       VALOR TASACION          ORO NETO  BOVEDA " & gPrnSaltoLinea
        Case "CredReno"
            vRTFImp = vRTFImp & Space(1) & "ITEM PLAZO  CONTRATO   RN    PRESTAMO    MONTO   PRESTAMO   INTERES  IMPUESTO  COSTO DE  COSTO DE   NETO" & gPrnSaltoLinea
            vRTFImp = vRTFImp & Space(1) & "                                        AMORTIZ  REDUCIDO                      CUSTODIA  PRE-REM.  COBRADO" & gPrnSaltoLinea
        Case "PrenDevu"
            vRTFImp = vRTFImp & Space(1) & "ITEM   PLAZO   CONTRATO       VALOR TASACION        ORO NETO" & gPrnSaltoLinea
        Case "PrenNuev"
            vRTFImp = vRTFImp & Space(1) & "ITEM   PLAZO   CONTRATO       VALOR TASACION        ORO NETO" & gPrnSaltoLinea
        Case "PresResc"
            vRTFImp = vRTFImp & Space(1) & "ITEM PLAZ0  CONTRATO       PRESTAMO      SALDO      INTERES   IMPUESTO   COSTO DE    COSTO DE     NETO   BOV" & gPrnSaltoLinea
            vRTFImp = vRTFImp & Space(1) & "                                        ANTERIOR                         CUSTODIA    PRE-REM.    COBRADO   " & gPrnSaltoLinea
        Case "ContCondDife"
            vRTFImp = vRTFImp & Space(1) & "ITEM    CONTRATO      PLAZO    VALOR TASACION         ORO NETO    FEC.CANCEL.  DIFERIDO  BLOQ  BOVEDA " & gPrnSaltoLinea
        Case "OperExto"
            vRTFImp = vRTFImp & Space(1) & "ITEM     CONTRATO        FECHA/HORA                OPERACION                    MONTO       USUARIO" & gPrnSaltoLinea
        Case "PagoSobr"
            vRTFImp = vRTFImp & Space(1) & "ITEM     CONTRATO        FECHA/HORA           MONTO      USUARIO" & gPrnSaltoLinea
    End Select
    vRTFImp = vRTFImp & String(vNroLineas, "-") & gPrnSaltoLinea
End Sub

'Contratos Anulados
Private Sub ImprimeContAnul(vConexion As ADODB.Connection)
    'Dim vNombre As String * 37
    Dim RegTmp As New ADODB.Recordset
    Dim vIndice As Integer  'contador de Item
    Dim vLineas As Integer
    Dim vCodAnt As String
    MousePointer = 11
    MuestraImpresion = True
    'vRTFImp = ""
    'Estado Anulado : A  ; Falta limitar por fecha
    'sSql = "SELECT cdp.cCodCta, p.cnompers, cdp.nplazo, cdp.nprestamo " & _
        " FROM ConsdiarioPrend AS CDP, " & gcCentralpers & "Persona AS P, PersCuenta AS PC" & _
        " WHERE cdp.ccodcta = pc.ccodcta AND  pc.ccodpers = p.ccodpers AND " & _
        " cdp.cestado = 'A' AND cflag IS NULL " & _
        " AND DATEDIFF(dd,cdp.dFecha,'" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 "
        
    sSql = " Select a.ccodcta, d.cnompers, b.nplazo, a.nprestamo " & _
        " From CredPrenda A Inner Join TransPrenda B On a.ccodcta = b.ccodcta " & _
        " Inner Join PersCuenta C On b.ccodcta = c.ccodcta " & _
        " Inner Join " & gcCentralPers & "Persona D On c.ccodpers = d.ccodpers " & _
        " Where b.ccodtran = '" & gsAnuContrato & "' And " & _
        " DATEDIFF(dd,b.dFecha,'" & Format(txtFecha, "mm/dd/yyyy") & "') = 0 "
        
        If pbListGene = False Then      'Verifica el listado (General o no)
            sSql = sSql & " AND b.ccodusu = '" & gsCodUser & "' "
        ElseIf cboUsuario <> "<Consolidado>" Then
            sSql = sSql & " AND b.ccodusu = '" & Trim(cboUsuario.Text) & "' "
        End If
        If Len(Trim(vBoveda)) > 0 Then
            sSql = sSql & " AND A.cAgeBoveda in " & vBoveda & " "
        End If
        sSql = sSql & " ORDER BY a.cCodCta"
    RegTmp.Open sSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
    If (RegTmp.BOF Or RegTmp.EOF) Then
        RegTmp.Close
        Set RegTmp = Nothing
        MsgBox " No existen Contratos Anulados ", vbInformation, " Aviso "
        MuestraImpresion = False
        MousePointer = 0
        Exit Sub
    Else
        vPage = vPage + 1: vCont = 0
        prgList.Min = 0
        prgList.Max = RegTmp.RecordCount
        If optImpresion(0).Value = True Then
            If prgList.Max > pPrevioMax Then
                RegTmp.Close
                Set RegTmp = Nothing
                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
                MousePointer = 0
                Exit Sub
            Else
                pPrevioMax = pPrevioMax - prgList.Max
            End If
            Cabecera "ContAnul", vPage
        Else
            ImpreBegin False, pHojaFiMax
            vRTFImp = ""
            Cabecera "ContAnul", vPage
            Print #ArcSal, ImpreCarEsp(vRTFImp);
            vRTFImp = ""
        End If
        prgList.Visible = True
        vIndice = 1
        vLineas = 7
        With RegTmp
            Do While Not .EOF
                If !cCodCta = vCodAnt Then
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & ImpreFormat(PstaNombre(!cNomPers, False), 37, 22) & gPrnSaltoLinea
                    Else
                        Print #ArcSal, ImpreFormat(PstaNombre(!cNomPers, False), 37, 22)
                    End If
                Else
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & ImpreFormat(vIndice, 5, 0) & Space(2) & FormatoContratro(!cCodCta) & Space(1) & _
                            ImpreFormat(PstaNombre(!cNomPers, False), 36, 2) & ImpreFormat(!nPlazo, 4, 0) & _
                            ImpreFormat(!nPrestamo, 12) & gPrnSaltoLinea
                    Else
                        Print #ArcSal, ImpreFormat(vIndice, 5, 0) & Space(2) & FormatoContratro(!cCodCta) & Space(1) & _
                            ImpreFormat(PstaNombre(!cNomPers, False), 36, 2) & ImpreFormat(!nPlazo, 4, 0) & _
                            ImpreFormat(!nPrestamo, 12)
                    End If
                    vIndice = vIndice + 1
                    vCodAnt = !cCodCta
                End If
                vLineas = vLineas + 1
                If vLineas > pLineasMax Then
                    vPage = vPage + 1
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoPagina
                        Cabecera "ContAnul", vPage
                    Else
                        ImpreNewPage
                        vRTFImp = ""
                        Cabecera "ContAnul", vPage
                        Print #ArcSal, ImpreCarEsp(vRTFImp);
                        vRTFImp = ""
                    End If
                    vLineas = 7
                End If
                vCont = vCont + 1
                prgList.Value = vCont
                Me.Caption = "Registro Nro.: " & vCont
                .MoveNext
            Loop
            If Not optImpresion(0).Value = True Then ImpreEnd
        End With
        RegTmp.Close
        Set RegTmp = Nothing
        prgList.Visible = False
        prgList.Value = 0
    End If
    Me.Caption = vNameForm
    MousePointer = 0
End Sub

'Prendas Nuevas en condición de diferidas
Private Sub ImprimePrenNuevCondDife(vConexion As ADODB.Connection)
    Dim RegTmp As New ADODB.Recordset
    Dim vIndice As Integer  'contador de Item
    Dim vLineas As Integer
    Dim lsBloq As String * 5
    'Dim vPage As Integer
    Dim vSumValTasac As Double 'suma de los valores de tasacion
    Dim vSumOroNeto As Double  'suma del oro neto
    MousePointer = 11
    MuestraImpresion = True
    'gdFecSis = "27/11/1999"
    'vRTFImp = ""
    'Estado Diferido - 2  ; Falta limitar por fecha
    'SSQL = "SELECT cdp.cCodCta, cdp.nplazo, cp.nValTasac, cp.nOroNeto , rc.cCodAnt " & _
        " FROM ConsdiarioPrend CDP INNER JOIN CredPrenda CP ON cdp.ccodcta = cp.ccodcta " & _
        " LEFT JOIN RelConNueAntPrend RC ON cp.ccodcta = rc.cCodNue " & _
        " WHERE cdp.cestado = '2' AND cflag IS NULL " & _
        " AND DATEDIFF(dd ,cdp.dFecha,'" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 "
        
    sSql = "SELECT tp.cCodCta, tp.nplazo, cp.nValTasac, cp.nOroNeto , rc.cCodAnt, cp.cAgeBoveda " & _
        " FROM TransPrenda TP INNER JOIN CredPrenda CP ON tp.ccodcta = cp.ccodcta " & _
        " LEFT JOIN RelConNueAntPrend RC ON cp.ccodcta = rc.cCodNue " & _
        " WHERE tp.cCodTran IN ('" & gsCanNorPrestamo & "','" & gsCanMorPrestamo & "','" & gsCanNorEnOtAgPrestamo & "','" & gsCanMorEnOtAgPrestamo & "','" & gsCanNorEnOtCjPrestamo & "','" & gsCanMorEnOtCjPrestamo & "') AND tp.cflag IS NULL " & _
        " AND DATEDIFF(dd ,tp.dFecha,'" & Format(txtFecha, "mm/dd/yyyy") & "') = 0 "
        
        If pbListGene = False Then      'Verifica el listado (General o no)
            sSql = sSql & " AND tp.ccodusu = '" & gsCodUser & "' "
        ElseIf cboUsuario <> "<Consolidado>" Then
            'SSQL = SSQL & " AND tp.ccodusu = '" & Trim(cboUsuario.Text) & "' "
            ' *** Cambio para la Operaciones con Interconexion
            sSql = sSql & " AND (tp.cCodusu = '" & Trim(cboUsuario.Text) & "'  or tp.cCodUsuRem = '" & Trim(cboUsuario.Text) & "' ) "
        End If
        If Len(Trim(vBoveda)) > 0 Then
            sSql = sSql & " AND CP.cAgeBoveda in " & vBoveda & " "
        End If
        sSql = sSql & " ORDER BY tp.cCodCta"
    RegTmp.Open sSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
    If (RegTmp.BOF Or RegTmp.EOF) Then
        RegTmp.Close
        Set RegTmp = Nothing
        MsgBox " No existen Prendas Nuevas en condición de diferidas ", vbInformation, " Aviso "
        MuestraImpresion = False
        MousePointer = 0
        Exit Sub
    Else
        vPage = vPage + 1: vCont = 0
        prgList.Min = 0
        prgList.Max = RegTmp.RecordCount
        If optImpresion(0).Value = True Then
            If prgList.Max > pPrevioMax Then
                RegTmp.Close
                Set RegTmp = Nothing
                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
                MousePointer = 0
                Exit Sub
            Else
                pPrevioMax = pPrevioMax - prgList.Max
            End If
            Cabecera "PrNuCoDi", vPage
        Else
            ImpreBegin False, pHojaFiMax
            vRTFImp = ""
            Cabecera "PrNuCoDi", vPage
            Print #ArcSal, ImpreCarEsp(vRTFImp);
            vRTFImp = ""
        End If
        prgList.Visible = True
        vIndice = 1
        vLineas = 7
        vSumValTasac = Format(0, "#0.00")
        vSumOroNeto = Format(0, "#0.00")
        With RegTmp
            .MoveFirst
            'lsBloq = IIf(IsCtaBlo(!cCodCta, vConexion), "BLOQ", "   ")
            Do While Not .EOF
                lsBloq = IIf(IsCtaBlo(!cCodCta, vConexion), "BLOQ", "   ")
                If optImpresion(0).Value = True Then
                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 5, 0) & Space(2) & FormatoContratro(!cCodCta) & Space(1) & ImpreFormat(!cCodAnt & "", 12, 2) & _
                        ImpreFormat(!nPlazo, 6, 0) & ImpreFormat(!nvaltasac, 15) & _
                        ImpreFormat(!noroneto, 13) & Space(3) & lsBloq & Space(4) & Mid(!cAgeBoveda, 4, 2) & gPrnSaltoLinea
                Else
                    Print #ArcSal, ImpreFormat(vIndice, 5, 0) & Space(2) & FormatoContratro(!cCodCta) & Space(1) & ImpreFormat(!cCodAnt, 12, 2) & _
                        ImpreFormat(!nPlazo, 6, 0) & ImpreFormat(!nvaltasac, 15) & _
                        ImpreFormat(!noroneto, 13) & Space(3) & lsBloq & Space(4) & Mid(!cAgeBoveda, 4, 2)
                End If
                vSumValTasac = vSumValTasac + Val(Format(IIf(!nvaltasac = Null, 0, Abs(!nvaltasac)), "#0.00"))
                vSumOroNeto = vSumOroNeto + Val(Format(IIf(!noroneto = Null, 0, Abs(!noroneto)), "#0.00"))
                vIndice = vIndice + 1
                vLineas = vLineas + 1
                If vLineas > pLineasMax Then
                    vPage = vPage + 1
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoPagina
                        Cabecera "PrNuCoDi", vPage
                    Else
                        ImpreNewPage
                        vRTFImp = ""
                        Cabecera "PrNuCoDi", vPage
                        Print #ArcSal, ImpreCarEsp(vRTFImp);
                        vRTFImp = ""
                    End If
                    vLineas = 7
                End If
                vCont = vCont + 1
                prgList.Value = vCont
                Me.Caption = "Registro Nro.: " & vCont
                .MoveNext
            Loop
            If optImpresion(0).Value = True Then
                vRTFImp = vRTFImp & gPrnSaltoLinea
                vRTFImp = vRTFImp & ImpreFormat("RESUMEN", 8, 25) & ImpreFormat(vSumValTasac, 25) & _
                    ImpreFormat(vSumOroNeto, 13) & gPrnSaltoLinea
            Else
                Print #ArcSal, ""
                Print #ArcSal, ImpreFormat("RESUMEN", 8, 25) & ImpreFormat(vSumValTasac, 25) & _
                    ImpreFormat(vSumOroNeto, 13)
                ImpreEnd
            End If
        End With
        RegTmp.Close
        Set RegTmp = Nothing
        prgList.Visible = False
        prgList.Value = 0
    End If
    Me.Caption = vNameForm
    MousePointer = 0
End Sub

'Préstamos Para Desembolsar
Private Sub ImprimeContRegi(vConexion As ADODB.Connection)
    Dim vPlazo As Integer
    Dim vIndice As Integer
    Dim vLineas As Integer
    Dim vNeto As Double
    
    Dim vSumPrestamo As Double
    Dim vSumInteres As Double
    Dim vSumCostCust As Double
    Dim vSumCostTasa As Double
    Dim vSumImpuesto As Double
    Dim vSumNeto As Double
    
    Dim vTotPrestamo As Double
    Dim vTotInteres As Double
    Dim vTotCostCust As Double
    Dim vTotCostTasa As Double
    Dim vTotImpuesto As Double
    Dim vTotNeto As Double
    MousePointer = 11
    MuestraImpresion = True
    'vRTFImp = ""
    'Estado No Desembolsado - 0  ; Falta limitar por fecha
    'SSQL = "SELECT cCodCta, nplazo, nPrestamo, nInteres, nCostoTasac, nCostoCusto, nImpuesto, ccodusu " & _
        " FROM ConsdiarioPrend " & _
        " WHERE cestado = '0' AND cflag IS NULL " & _
        " AND DATEDIFF(dd,dFecha,'" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 "
    sSql = "SELECT T.cCodCta, T.nplazo, T.nMontoCred, T.nInteres, T.nCostTasac, T.nCostCust, T.nImpuesto, T.ccodusu, CP.cAgeBoveda " & _
        " FROM TransPrenda T INNER JOIN CredPrenda CP ON T.cCodCta = CP.cCodCta " & _
        " WHERE T.cCodTran = '" & gsRegContrato & "' AND T.cflag IS NULL " & _
        " AND DATEDIFF(dd,T.dFecha,'" & Format(txtFecha, "mm/dd/yyyy") & "') = 0 "
        
        If pbListGene = False Then      'Verifica el listado (General o no)
            sSql = sSql & " AND T.ccodusu = '" & gsCodUser & "' "
        ElseIf cboUsuario <> "<Consolidado>" Then
            sSql = sSql & " AND T.ccodusu = '" & Trim(cboUsuario.Text) & "' "
        End If
        If Len(Trim(vBoveda)) > 0 Then
            sSql = sSql & " AND CP.cAgeBoveda in " & vBoveda & " "
        End If
        sSql = sSql & " ORDER BY T.nPlazo, T.cCodCta"
    RegTransDiariaPrend.Open sSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
    If (RegTransDiariaPrend.BOF Or RegTransDiariaPrend.EOF) Then
        RegTransDiariaPrend.Close
        Set RegTransDiariaPrend = Nothing
        MsgBox " No existen Contratos Registrados ", vbInformation, " Aviso "
        MuestraImpresion = False
        MousePointer = 0
        Exit Sub
    Else
        vPage = vPage + 1: vCont = 0
        prgList.Min = 0
        prgList.Max = RegTransDiariaPrend.RecordCount
        If optImpresion(0).Value = True Then
            If prgList.Max > pPrevioMax Then
                RegTransDiariaPrend.Close
                Set RegTransDiariaPrend = Nothing
                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
                MousePointer = 0
                Exit Sub
            Else
                pPrevioMax = pPrevioMax - prgList.Max
            End If
            Cabecera "ContRegi", vPage
        Else
            ImpreBegin False, pHojaFiMax
            vRTFImp = ""
            Cabecera "ContRegi", vPage
            Print #ArcSal, ImpreCarEsp(vRTFImp);
            vRTFImp = ""
        End If
        prgList.Visible = True
        vIndice = 1
        vLineas = 7
        vTotPrestamo = Format(0, "#0.00"): vTotInteres = Format(0, "#0.00")
        vTotCostCust = Format(0, "#0.00"): vTotCostTasa = Format(0, "#0.00")
        vTotImpuesto = Format(0, "#0.00"): vTotNeto = Format(0, "#.00")
        vSumPrestamo = Format(0, "#0.00"): vSumInteres = Format(0, "#0.00")
        vSumCostCust = Format(0, "#0.00"): vSumCostTasa = Format(0, "#0.00")
        vSumImpuesto = Format(0, "#0.00"): vSumNeto = Format(0, "#.00")
        With RegTransDiariaPrend
            vPlazo = !nPlazo
            .MoveFirst
            Do While Not .EOF
                If Val(!nPlazo) <> Val(vPlazo) Then
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoLinea
                        vRTFImp = vRTFImp & ImpreFormat("Resumen", 8, 8) & ImpreFormat(vPlazo, 6, 0) & _
                            ImpreFormat(vSumPrestamo, 13) & ImpreFormat(vSumInteres, 9) & _
                            ImpreFormat(vSumCostTasa, 9) & ImpreFormat(vSumCostCust, 9) & _
                            ImpreFormat(vSumImpuesto, 9) & ImpreFormat(vSumNeto, 9) & gPrnSaltoLinea
                        vRTFImp = vRTFImp & gPrnSaltoLinea
                    Else
                        Print #ArcSal, ""
                        Print #ArcSal, ImpreFormat("Resumen", 8, 8) & ImpreFormat(vPlazo, 6, 0) & _
                            ImpreFormat(vSumPrestamo, 13) & ImpreFormat(vSumInteres, 9) & _
                            ImpreFormat(vSumCostTasa, 9) & ImpreFormat(vSumCostCust, 9) & _
                            ImpreFormat(vSumImpuesto, 9) & ImpreFormat(vSumNeto, 9)
                        Print #ArcSal, ""
                    End If
                    vLineas = vLineas + 3
                    
                    vTotPrestamo = vTotPrestamo + vSumPrestamo: vTotInteres = vTotInteres + vSumInteres
                    vTotCostCust = vTotCostCust + vSumCostCust: vTotCostTasa = vTotCostTasa + vSumCostTasa
                    vTotImpuesto = vTotImpuesto + vSumImpuesto: vTotNeto = vTotNeto + vSumNeto
                    vSumPrestamo = Format(0, "#0.00"): vSumInteres = Format(0, "#0.00")
                    vSumCostCust = Format(0, "#0.00"): vSumCostTasa = Format(0, "#0.00")
                    vSumImpuesto = Format(0, "#0.00"): vSumNeto = Format(0, "#0.00")
                End If
                vNeto = IIf(IsNull(!nMontoCred), 0, Abs(!nMontoCred)) - IIf(IsNull(!nInteres), 0, Abs(!nInteres)) - IIf(IsNull(!nImpuesto), 0, Abs(!nImpuesto)) - _
                        IIf(IsNull(!nCostCust), 0, Abs(!nCostCust)) - IIf(IsNull(!nCostTasac), 0, Abs(!nCostTasac))
                If optImpresion(0).Value = True Then
                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 5, 0) & ImpreFormat(!nPlazo, 4, 0) & Space(1) & _
                        FormatoContratro(!cCodCta) & Space(1) & ImpreFormat(Abs(!nMontoCred), 9) & _
                        ImpreFormat(!nInteres, 9) & ImpreFormat(!nCostTasac, 9) & _
                        ImpreFormat(!nCostCust, 9) & ImpreFormat(!nImpuesto, 9) & _
                        ImpreFormat(vNeto, 9) & ImpreFormat(!cCodUsu, 4, 3) & ImpreFormat("", 0, 0) & gPrnSaltoLinea
                Else
                    Print #ArcSal, ImpreFormat(vIndice, 5, 0) & ImpreFormat(!nPlazo, 4, 0) & Space(1) & _
                        FormatoContratro(!cCodCta) & Space(1) & ImpreFormat(Abs(!nMontoCred), 9) & _
                        ImpreFormat(!nInteres, 9) & ImpreFormat(!nCostTasac, 9) & _
                        ImpreFormat(!nCostCust, 9) & ImpreFormat(!nImpuesto, 9) & _
                        ImpreFormat(vNeto, 9) & ImpreFormat(!cCodUsu, 4, 3) & ImpreFormat("", 0, 0)
                End If
                vSumPrestamo = vSumPrestamo + IIf(IsNull(!nMontoCred), 0, Abs(!nMontoCred)): vSumInteres = vSumInteres + IIf(IsNull(!nInteres), 0, Abs(!nInteres))
                vSumCostCust = vSumCostCust + IIf(IsNull(!nCostCust), 0, Abs(!nCostCust)): vSumCostTasa = vSumCostTasa + IIf(IsNull(!nCostTasac), 0, Abs(!nCostTasac))
                vSumImpuesto = vSumImpuesto + IIf(IsNull(!nImpuesto), 0, Abs(!nImpuesto)): vSumNeto = vSumNeto + vNeto
                vPlazo = !nPlazo
                vIndice = vIndice + 1
                vLineas = vLineas + 1
                If vLineas > pLineasMax Then
                    vPage = vPage + 1
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoPagina
                        Cabecera "ContRegi", vPage
                    Else
                        ImpreNewPage
                        vRTFImp = ""
                        Cabecera "ContRegi", vPage
                        Print #ArcSal, ImpreCarEsp(vRTFImp);
                        vRTFImp = ""
                    End If
                    vLineas = 7
                End If
                vCont = vCont + 1
                prgList.Value = vCont
                Me.Caption = "Registro Nro.: " & vCont
                .MoveNext
            Loop
            vTotPrestamo = vTotPrestamo + vSumPrestamo: vTotInteres = vTotInteres + vSumInteres
            vTotCostCust = vTotCostCust + vSumCostCust: vTotCostTasa = vTotCostTasa + vSumCostTasa
            vTotImpuesto = vTotImpuesto + vSumImpuesto: vTotNeto = vTotNeto + vSumNeto

            If optImpresion(0).Value = True Then
                vRTFImp = vRTFImp & gPrnSaltoLinea
                vRTFImp = vRTFImp & ImpreFormat("Resumen", 8, 8) & ImpreFormat(vPlazo, 6, 0) & _
                    ImpreFormat(vSumPrestamo, 13) & ImpreFormat(vSumInteres, 9) & _
                    ImpreFormat(vSumCostTasa, 9) & ImpreFormat(vSumCostCust, 9) & _
                    ImpreFormat(vSumImpuesto, 9) & ImpreFormat(vSumNeto, 9) & gPrnSaltoLinea
                
                vRTFImp = vRTFImp & gPrnSaltoLinea
                vRTFImp = vRTFImp & ImpreFormat("RESUMEN", 8, 8) & ImpreFormat(vTotPrestamo, 19) & _
                    ImpreFormat(vTotInteres, 9) & ImpreFormat(vTotCostTasa, 9) & _
                    ImpreFormat(vTotCostCust, 9) & ImpreFormat(vTotImpuesto, 9) & _
                    ImpreFormat(vTotNeto, 9) & gPrnSaltoLinea
            Else
                Print #ArcSal, ""
                Print #ArcSal, ImpreFormat("Resumen", 8, 8) & ImpreFormat(vPlazo, 6, 0) & _
                    ImpreFormat(vSumPrestamo, 13) & ImpreFormat(vSumInteres, 9) & _
                    ImpreFormat(vSumCostTasa, 9) & ImpreFormat(vSumCostCust, 9) & _
                    ImpreFormat(vSumImpuesto, 9) & ImpreFormat(vSumNeto, 9)
                
                Print #ArcSal, ""
                Print #ArcSal, ImpreFormat("RESUMEN", 8, 8) & ImpreFormat(vTotPrestamo, 19) & _
                    ImpreFormat(vTotInteres, 9) & ImpreFormat(vTotCostTasa, 9) & _
                    ImpreFormat(vTotCostCust, 9) & ImpreFormat(vTotImpuesto, 9) & _
                    ImpreFormat(vTotNeto, 9)
                ImpreEnd
            End If
        End With
        RegTransDiariaPrend.Close
        Set RegTransDiariaPrend = Nothing
        prgList.Visible = False
        prgList.Value = 0
    End If
    Me.Caption = vNameForm
    MousePointer = 0
End Sub

'Préstamos Nuevos
Private Sub ImprimePresNuev(vConexion As ADODB.Connection)
    Dim vPlazo As Integer
    Dim vIndice As Integer
    Dim vLineas As Integer
    'Dim vPage As Integer
    Dim vNeto As Double
    
    Dim vSumPrestamo As Double
    Dim vSumInteres As Double
    Dim vSumCostCust As Double
    Dim vSumCostTasa As Double
    Dim vSumImpuesto As Double
    Dim vSumNeto As Double
    
    Dim vTotPrestamo As Double
    Dim vTotInteres As Double
    Dim vTotCostCust As Double
    Dim vTotCostTasa As Double
    Dim vTotImpuesto As Double
    Dim vTotNeto As Double
    MousePointer = 11
    MuestraImpresion = True
    'vRTFImp = ""
    'Estado Desembolsado - 1  ; Falta limitar por fecha
    'SSQL = "SELECT cCodCta, nplazo, nPrestamo, nInteres, nCostoTasac, nCostoCusto, nImpuesto, ccodusu " & _
        " FROM ConsdiarioPrend WHERE cestado = '1' AND cflag IS NULL " & _
        " AND DATEDIFF(dd,dFecha,'" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 "
        
    sSql = "SELECT T.cCodCta, T.nplazo, T.nMontoCred, T.nInteres, T.nCostTasac, T.nCostCust, T.nImpuesto, T.ccodusu, CP.cAgeBoveda " & _
        " FROM TransPrenda T INNER JOIN CredPrenda CP ON T.cCodCta = CP.cCodCta WHERE T.cCodTran = '" & gsDesPrestamo & "' AND T.cflag IS NULL " & _
        " AND DATEDIFF(dd,T.dFecha,'" & Format(txtFecha, "mm/dd/yyyy") & "') = 0 "
        
        If pbListGene = False Then      'Verifica el listado (General o no)
            sSql = sSql & " AND T.cCodusu = '" & gsCodUser & "' "
        ElseIf cboUsuario <> "<Consolidado>" Then
            sSql = sSql & " AND T.cCodUsu = '" & Trim(cboUsuario.Text) & "' "
        End If
        If Len(Trim(vBoveda)) > 0 Then
            sSql = sSql & " AND CP.cAgeBoveda in " & vBoveda & " "
        End If
        sSql = sSql & " ORDER BY T.nPlazo, T.cCodCta"
    RegTransDiariaPrend.Open sSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
    If (RegTransDiariaPrend.BOF Or RegTransDiariaPrend.EOF) Then
        RegTransDiariaPrend.Close
        Set RegTransDiariaPrend = Nothing
        MsgBox " No existen Préstamos Nuevos ", vbInformation, " Aviso "
        MuestraImpresion = False
        MousePointer = 0
        Exit Sub
    Else
        vPage = vPage + 1: vCont = 0
        prgList.Min = 0
        prgList.Max = RegTransDiariaPrend.RecordCount
        If optImpresion(0).Value = True Then
            If prgList.Max > pPrevioMax Then
                RegTransDiariaPrend.Close
                Set RegTransDiariaPrend = Nothing
                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
                MousePointer = 0
                Exit Sub
            Else
                pPrevioMax = pPrevioMax - prgList.Max
            End If
            Cabecera "PresNuev", vPage
        Else
            ImpreBegin False, pHojaFiMax
            vRTFImp = ""
            Cabecera "PresNuev", vPage
            Print #ArcSal, ImpreCarEsp(vRTFImp);
            vRTFImp = ""
        End If
        prgList.Visible = True
        vIndice = 1
        vLineas = 7
        vTotPrestamo = Format(0, "#0.00"): vTotInteres = Format(0, "#0.00")
        vTotCostCust = Format(0, "#0.00"): vTotCostTasa = Format(0, "#0.00")
        vTotImpuesto = Format(0, "#0.00"): vTotNeto = Format(0, "#.00")
        vSumPrestamo = Format(0, "#0.00"): vSumInteres = Format(0, "#0.00")
        vSumCostCust = Format(0, "#0.00"): vSumCostTasa = Format(0, "#0.00")
        vSumImpuesto = Format(0, "#0.00"): vSumNeto = Format(0, "#.00")
        With RegTransDiariaPrend
            vPlazo = !nPlazo
            .MoveFirst
            Do While Not .EOF
                If Val(!nPlazo) <> Val(vPlazo) Then
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoLinea
                        vRTFImp = vRTFImp & ImpreFormat("Resumen", 8, 8) & ImpreFormat(vPlazo, 7, 0) & _
                            ImpreFormat(vSumPrestamo, 12) & ImpreFormat(vSumInteres, 9) & _
                            ImpreFormat(vSumCostTasa, 9) & ImpreFormat(vSumCostCust, 9) & _
                            ImpreFormat(vSumImpuesto, 9) & ImpreFormat(vSumNeto, 9) & gPrnSaltoLinea
                        vRTFImp = vRTFImp & gPrnSaltoLinea
                    Else
                        Print #ArcSal, ""
                        Print #ArcSal, ImpreFormat("Resumen", 8, 8) & ImpreFormat(vPlazo, 7, 0) & _
                            ImpreFormat(vSumPrestamo, 12) & ImpreFormat(vSumInteres, 9) & _
                            ImpreFormat(vSumCostTasa, 9) & ImpreFormat(vSumCostCust, 9) & _
                            ImpreFormat(vSumImpuesto, 9) & ImpreFormat(vSumNeto, 9)
                        Print #ArcSal, ""
                    End If
                    vLineas = vLineas + 3
                    
                    vTotPrestamo = vTotPrestamo + vSumPrestamo: vTotInteres = vTotInteres + vSumInteres
                    vTotCostCust = vTotCostCust + vSumCostCust: vTotCostTasa = vTotCostTasa + vSumCostTasa
                    vTotImpuesto = vTotImpuesto + vSumImpuesto: vTotNeto = vTotNeto + vSumNeto
                    vSumPrestamo = Format(0, "#0.00"): vSumInteres = Format(0, "#0.00")
                    vSumCostCust = Format(0, "#0.00"): vSumCostTasa = Format(0, "#0.00")
                    vSumImpuesto = Format(0, "#0.00"): vSumNeto = Format(0, "#0.00")
                End If
                vNeto = IIf(IsNull(!nMontoCred), 0, Abs(!nMontoCred)) - IIf(IsNull(!nInteres), 0, Abs(!nInteres)) - IIf(IsNull(!nImpuesto), 0, Abs(!nImpuesto)) - _
                        IIf(IsNull(!nCostCust), 0, Abs(!nCostCust)) - IIf(IsNull(!nCostTasac), 0, Abs(!nCostTasac))
                If optImpresion(0).Value = True Then
                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 5, 0) & ImpreFormat(!nPlazo, 4, 0) & Space(1) & _
                        FormatoContratro(!cCodCta) & Space(1) & ImpreFormat(Abs(!nMontoCred), 9) & _
                        ImpreFormat(!nInteres, 9) & ImpreFormat(!nCostTasac, 9) & _
                        ImpreFormat(!nCostCust, 9) & ImpreFormat(!nImpuesto, 9) & _
                        ImpreFormat(vNeto, 9) & ImpreFormat(!cCodUsu, 4, 2) & gPrnSaltoLinea
                Else
                    Print #ArcSal, ImpreFormat(vIndice, 5, 0) & ImpreFormat(!nPlazo, 4, 0) & Space(1) & _
                        FormatoContratro(!cCodCta) & Space(1) & ImpreFormat(Abs(!nMontoCred), 9) & _
                        ImpreFormat(!nInteres, 9) & ImpreFormat(!nCostTasac, 9) & _
                        ImpreFormat(!nCostCust, 9) & ImpreFormat(!nImpuesto, 9) & _
                        ImpreFormat(vNeto, 9) & ImpreFormat(!cCodUsu, 4, 2)
                End If
                vSumPrestamo = vSumPrestamo + IIf(IsNull(!nMontoCred), 0, Abs(!nMontoCred)): vSumInteres = vSumInteres + IIf(IsNull(!nInteres), 0, Abs(!nInteres))
                vSumCostCust = vSumCostCust + IIf(IsNull(!nCostCust), 0, Abs(!nCostCust)): vSumCostTasa = vSumCostTasa + IIf(IsNull(!nCostTasac), 0, Abs(!nCostTasac))
                vSumImpuesto = vSumImpuesto + IIf(IsNull(!nImpuesto), 0, Abs(!nImpuesto)): vSumNeto = vSumNeto + vNeto
                vPlazo = !nPlazo
                vIndice = vIndice + 1
                vLineas = vLineas + 1
                If vLineas > pLineasMax Then
                    vPage = vPage + 1
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoPagina
                        Cabecera "PresNuev", vPage
                    Else
                        ImpreNewPage
                        vRTFImp = ""
                        Cabecera "PresNuev", vPage
                        Print #ArcSal, ImpreCarEsp(vRTFImp);
                        vRTFImp = ""
                    End If
                    vLineas = 7
                End If
                vCont = vCont + 1
                prgList.Value = vCont
                Me.Caption = "Registro Nro.: " & vCont
                .MoveNext
            Loop
            vTotPrestamo = vTotPrestamo + vSumPrestamo: vTotInteres = vTotInteres + vSumInteres
            vTotCostCust = vTotCostCust + vSumCostCust: vTotCostTasa = vTotCostTasa + vSumCostTasa
            vTotImpuesto = vTotImpuesto + vSumImpuesto: vTotNeto = vTotNeto + vSumNeto
            If optImpresion(0).Value = True Then
                vRTFImp = vRTFImp & gPrnSaltoLinea
                vRTFImp = vRTFImp & ImpreFormat("Resumen", 8, 8) & ImpreFormat(vPlazo, 7, 0) & _
                    ImpreFormat(vSumPrestamo, 12) & ImpreFormat(vSumInteres, 9) & _
                    ImpreFormat(vSumCostTasa, 9) & ImpreFormat(vSumCostCust, 9) & _
                    ImpreFormat(vSumImpuesto, 9) & ImpreFormat(vSumNeto, 9) & gPrnSaltoLinea
                vRTFImp = vRTFImp & gPrnSaltoLinea
                vRTFImp = vRTFImp & ImpreFormat("RESUMEN", 8, 8) & ImpreFormat(vTotPrestamo, 19) & _
                    ImpreFormat(vTotInteres, 9) & ImpreFormat(vTotCostTasa, 9) & _
                    ImpreFormat(vTotCostCust, 9) & ImpreFormat(vTotImpuesto, 9) & _
                    ImpreFormat(vTotNeto, 9) & gPrnSaltoLinea
            Else
                Print #ArcSal, ""
                Print #ArcSal, ImpreFormat("Resumen", 8, 8) & ImpreFormat(vPlazo, 7, 0) & _
                    ImpreFormat(vSumPrestamo, 12) & ImpreFormat(vSumInteres, 9) & _
                    ImpreFormat(vSumCostTasa, 9) & ImpreFormat(vSumCostCust, 9) & _
                    ImpreFormat(vSumImpuesto, 9) & ImpreFormat(vSumNeto, 9)
                Print #ArcSal, ""
                Print #ArcSal, ImpreFormat("RESUMEN", 8, 8) & ImpreFormat(vTotPrestamo, 19) & _
                    ImpreFormat(vTotInteres, 9) & ImpreFormat(vTotCostTasa, 9) & _
                    ImpreFormat(vTotCostCust, 9) & ImpreFormat(vTotImpuesto, 9) & _
                    ImpreFormat(vTotNeto, 9)
                ImpreEnd
            End If
        End With
        RegTransDiariaPrend.Close
        Set RegTransDiariaPrend = Nothing
        prgList.Visible = False
        prgList.Value = 0
    End If
    Me.Caption = vNameForm
    MousePointer = 0
End Sub

'Préstamo Rescatados
Private Sub ImprimePresResc(vConexion As ADODB.Connection)
    Dim vPlazo As Integer
    Dim vIndice As Integer
    Dim vLineas As Integer
    Dim vNeto As Double
    Dim vSaldoCap As Double
    Dim vSumPrestamo As Double
    Dim vSumSaldoCap As Double
    Dim vSumInteres As Double
    Dim vSumCostCust As Double
    Dim vSumCostPreR As Double
    Dim vSumImpuesto As Double
    Dim vSumNeto As Double
    Dim vTotPrestamo As Double
    Dim vTotSaldoCap As Double
    Dim vTotInteres As Double
    Dim vTotCostCust As Double
    Dim vTotCostPreR As Double
    Dim vTotImpuesto As Double
    Dim vTotNeto As Double
    MousePointer = 11
    MuestraImpresion = True
    'vRTFImp = ""
    'gdFecSis = "27/11/1999"
    'Estado Diferido - 2  ; Falta limitar por fecha
    'SSQL = "SELECT cCodCta, nplazo, nPrestamo, nmontopag, nInteres, nCostoCusto, nCostoPreRe, nImpuesto " & _
        " FROM ConsdiarioPrend WHERE cestado = '2' AND cflag IS NULL " & _
        " AND DATEDIFF(dd,dFecha,'" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 "
        
    sSql = "SELECT tp.cCodCta, tp.nplazo, cp.nPrestamo, tp.nmontoTran, tp.nIntVenc, tp.nCostCustVenc, tp.nCostPrepRem, tp.nImpuesto, cp.cAgeBoveda " & _
        " FROM TransPrenda TP INNER JOIN CredPrenda CP ON tp.ccodcta = cp.ccodcta WHERE tp.cCodTran IN ('" & gsCanNorPrestamo & "','" & gsCanMorPrestamo & "','" & _
        gsCanNorEnOtAgPrestamo & "','" & gsCanMorEnOtAgPrestamo & "','" & gsCanNorEnOtCjPrestamo & "','" & gsCanMorEnOtCjPrestamo & "') AND tp.cflag IS NULL " & _
        " AND DATEDIFF(dd,tp.dFecha,'" & Format(txtFecha, "mm/dd/yyyy") & "') = 0 "
        
        If pbListGene = False Then      'Verifica el listado (General o no)
            sSql = sSql & " AND tp.ccodusu = '" & gsCodUser & "' "
        ElseIf cboUsuario <> "<Consolidado>" Then
            'SSQL = SSQL & " AND tp.ccodusu = '" & Trim(cboUsuario.Text) & "' "
            ' *** Cambio para la Operaciones con Interconexion
            sSql = sSql & " AND (tp.cCodusu = '" & Trim(cboUsuario.Text) & "'  or tp.cCodUsuRem = '" & Trim(cboUsuario.Text) & "' ) "
        End If
        If Len(Trim(vBoveda)) > 0 Then
            sSql = sSql & " AND CP.cAgeBoveda in " & vBoveda & " "
        End If
        sSql = sSql & " ORDER BY tp.nPlazo, tp.cCodCta"
    RegTransDiariaPrend.Open sSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
    If (RegTransDiariaPrend.BOF Or RegTransDiariaPrend.EOF) Then
        RegTransDiariaPrend.Close
        Set RegTransDiariaPrend = Nothing
        MsgBox " No existen Préstamos Rescatados ", vbInformation, " Aviso "
        MuestraImpresion = False
        MousePointer = 0
        Exit Sub
    Else
        vPage = vPage + 1: vCont = 0
        prgList.Min = 0
        prgList.Max = RegTransDiariaPrend.RecordCount
        If optImpresion(0).Value = True Then
            If prgList.Max > pPrevioMax Then
                RegTransDiariaPrend.Close
                Set RegTransDiariaPrend = Nothing
                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
                MousePointer = 0
                Exit Sub
            Else
                pPrevioMax = pPrevioMax - prgList.Max
            End If
            Cabecera "PresResc", vPage
        Else
            ImpreBegin False, pHojaFiMax
            vRTFImp = ""
            Cabecera "PresResc", vPage
            Print #ArcSal, ImpreCarEsp(vRTFImp);
            vRTFImp = ""
        End If
        prgList.Visible = True
        vIndice = 1
        vLineas = 7
        vTotPrestamo = Format(0, "#0.00"): vTotInteres = Format(0, "#0.00")
        vTotCostCust = Format(0, "#0.00"): vTotCostPreR = Format(0, "#0.00")
        vTotImpuesto = Format(0, "#0.00"): vTotNeto = Format(0, "#.00")
        vTotSaldoCap = Format(0, "#.00")
        vSumPrestamo = Format(0, "#0.00"): vSumInteres = Format(0, "#0.00")
        vSumCostCust = Format(0, "#0.00"): vSumCostPreR = Format(0, "#0.00")
        vSumImpuesto = Format(0, "#0.00"): vSumNeto = Format(0, "#.00")
        vSumSaldoCap = Format(0, "#.00")
        With RegTransDiariaPrend
            vPlazo = !nPlazo
            .MoveFirst
            Do While Not .EOF
                If Val(!nPlazo) <> Val(vPlazo) Then
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoLinea
                        vRTFImp = vRTFImp & ImpreFormat("Resumen", 8, 9) & ImpreFormat(vPlazo, 4, 0) & _
                            ImpreFormat(vSumPrestamo, 12) & ImpreFormat(vSumSaldoCap, 9) & _
                            ImpreFormat(vSumInteres, 9) & ImpreFormat(vSumImpuesto, 8) & _
                            ImpreFormat(vSumCostCust, 8) & ImpreFormat(vSumCostPreR, 8) & _
                            ImpreFormat(vSumNeto, 9) & gPrnSaltoLinea
                        vRTFImp = vRTFImp & gPrnSaltoLinea
                    Else
                        Print #ArcSal, ""
                        Print #ArcSal, ImpreFormat("Resumen", 8, 9) & ImpreFormat(vPlazo, 4, 0) & _
                            ImpreFormat(vSumPrestamo, 12) & ImpreFormat(vSumSaldoCap, 9) & _
                            ImpreFormat(vSumInteres, 9) & ImpreFormat(vSumImpuesto, 8) & _
                            ImpreFormat(vSumCostCust, 8) & ImpreFormat(vSumCostPreR, 8) & _
                            ImpreFormat(vSumNeto, 9)
                        Print #ArcSal, ""
                    End If
                    vLineas = vLineas + 3
                    
                    vTotPrestamo = vTotPrestamo + vSumPrestamo: vTotInteres = vTotInteres + vSumInteres
                    vTotCostCust = vTotCostCust + vSumCostCust: vTotCostPreR = vTotCostPreR + vSumCostPreR
                    vTotImpuesto = vTotImpuesto + vSumImpuesto: vTotNeto = vTotNeto + vSumNeto
                    vTotSaldoCap = vTotSaldoCap + vSumSaldoCap
                    vSumPrestamo = Format(0, "#0.00"): vSumInteres = Format(0, "#0.00")
                    vSumCostCust = Format(0, "#0.00"): vSumCostPreR = Format(0, "#0.00")
                    vSumImpuesto = Format(0, "#0.00"): vSumNeto = Format(0, "#0.00")
                    vSumSaldoCap = Format(0, "#0.00")
                End If
                vSaldoCap = !nMontoTran - (Abs(!nIntVenc) + Abs(!nImpuesto) + Abs(!nCostCustVenc) + Abs(!nCostPrepRem))
                vNeto = !nMontoTran ' Abs(!nPrestamo) + Abs(!nIntVenc) + Abs(!nImpuesto) + Abs(!nCostCustVenc) + Abs(!nCostPrepRem)
                If optImpresion(0).Value = True Then
                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 4, 0) & ImpreFormat(!nPlazo, 4, 0) & Space(1) & _
                        FormatoContratro(!cCodCta) & ImpreFormat(!nPrestamo, 9) & ImpreFormat(vSaldoCap, 9) & _
                        ImpreFormat(!nIntVenc, 9) & ImpreFormat(!nImpuesto, 8) & _
                        ImpreFormat(!nCostCustVenc, 8) & ImpreFormat(!nCostPrepRem, 8) & _
                        ImpreFormat(vNeto, 9) & Space(3) & Mid(!cAgeBoveda, 4, 2) & gPrnSaltoLinea
                Else
                    Print #ArcSal, ImpreFormat(vIndice, 4, 0) & ImpreFormat(!nPlazo, 4, 0) & Space(1) & _
                        FormatoContratro(!cCodCta) & ImpreFormat(!nPrestamo, 9) & ImpreFormat(vSaldoCap, 9) & _
                        ImpreFormat(!nIntVenc, 9) & ImpreFormat(!nImpuesto, 8) & _
                        ImpreFormat(!nCostCustVenc, 8) & ImpreFormat(!nCostPrepRem, 8) & _
                        ImpreFormat(vNeto, 9) & Space(3) & Mid(!cAgeBoveda, 4, 2)
                End If
                vSumPrestamo = vSumPrestamo + Abs(!nPrestamo): vSumInteres = vSumInteres + Abs(!nIntVenc)
                vSumCostCust = vSumCostCust + Abs(!nCostCustVenc): vSumCostPreR = vSumCostPreR + IIf(IsNull(!nCostPrepRem) = True, 0, !nCostPrepRem)  'Abs(!nCostPrepRem)
                vSumImpuesto = vSumImpuesto + Abs(!nImpuesto): vSumNeto = vSumNeto + vNeto
                vSumSaldoCap = vSumSaldoCap + vSaldoCap
                vPlazo = !nPlazo
                vIndice = vIndice + 1: vLineas = vLineas + 1
                If vLineas > pLineasMax Then
                    vPage = vPage + 1
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoPagina
                        Cabecera "PresResc", vPage
                    Else
                        ImpreNewPage
                        vRTFImp = ""
                        Cabecera "PresResc", vPage
                        Print #ArcSal, ImpreCarEsp(vRTFImp);
                        vRTFImp = ""
                    End If
                    vLineas = 7
                End If
                vCont = vCont + 1
                prgList.Value = vCont
                Me.Caption = "Registro Nro.: " & vCont
                .MoveNext
            Loop
            vTotPrestamo = vTotPrestamo + vSumPrestamo: vTotInteres = vTotInteres + vSumInteres
            vTotCostCust = vTotCostCust + vSumCostCust: vTotCostPreR = vTotCostPreR + vSumCostPreR
            vTotImpuesto = vTotImpuesto + vSumImpuesto: vTotNeto = vTotNeto + vSumNeto
            vTotSaldoCap = vTotSaldoCap + vSumSaldoCap
            If optImpresion(0).Value = True Then
                vRTFImp = vRTFImp & gPrnSaltoLinea
                vRTFImp = vRTFImp & ImpreFormat("Resumen", 8, 8) & ImpreFormat(vPlazo, 4, 0) & _
                    ImpreFormat(vSumPrestamo, 12) & ImpreFormat(vSumSaldoCap, 9) & _
                    ImpreFormat(vSumInteres, 9) & ImpreFormat(vSumImpuesto, 8) & _
                    ImpreFormat(vSumCostCust, 8) & ImpreFormat(vSumCostPreR, 8) & _
                    ImpreFormat(vSumNeto, 9) & gPrnSaltoLinea
                vRTFImp = vRTFImp & gPrnSaltoLinea
                vRTFImp = vRTFImp & ImpreFormat("RESUMEN", 8, 8) & ImpreFormat(vTotPrestamo, 16) & _
                    ImpreFormat(vTotSaldoCap, 9) & ImpreFormat(vTotInteres, 9) & _
                    ImpreFormat(vTotImpuesto, 8) & ImpreFormat(vTotCostCust, 8) & _
                    ImpreFormat(vTotCostPreR, 8) & ImpreFormat(vTotNeto, 9) & gPrnSaltoLinea
            Else
                Print #ArcSal, ""
                Print #ArcSal, ImpreFormat("Resumen", 8, 8) & ImpreFormat(vPlazo, 4, 0) & _
                    ImpreFormat(vSumPrestamo, 12) & ImpreFormat(vSumSaldoCap, 9) & _
                    ImpreFormat(vSumInteres, 9) & ImpreFormat(vSumImpuesto, 8) & _
                    ImpreFormat(vSumCostCust, 8) & ImpreFormat(vSumCostPreR, 8) & _
                    ImpreFormat(vSumNeto, 9)
                Print #ArcSal,
                Print #ArcSal, ImpreFormat("RESUMEN", 8, 8) & ImpreFormat(vTotPrestamo, 16) & _
                    ImpreFormat(vTotSaldoCap, 9) & ImpreFormat(vTotInteres, 9) & _
                    ImpreFormat(vTotImpuesto, 8) & ImpreFormat(vTotCostCust, 8) & _
                    ImpreFormat(vTotCostPreR, 8) & ImpreFormat(vTotNeto, 9)
                ImpreEnd
            End If
        End With
        RegTransDiariaPrend.Close
        Set RegTransDiariaPrend = Nothing
        prgList.Visible = False
        prgList.Value = 0
    End If
    Me.Caption = vNameForm
    MousePointer = 0
End Sub

'Créditos Renovados
Private Sub ImprimeCredReno(vConexion As ADODB.Connection)
    Dim vPlazo As Integer
    Dim vIndice As Integer
    Dim vLineas As Integer
    Dim vNeto As Double
    Dim vInteres As Double
    Dim vSumPrestamo As Double
    Dim vSumMontAmor As Double
    Dim vSumPresRedu As Double
    Dim vSumInteres As Double
    Dim vSumCostCust As Double
    Dim vSumCostPreR As Double
    Dim vSumImpuesto As Double
    Dim vSumNeto As Double
    Dim vTotPrestamo As Double
    Dim vTotMontAmor As Double
    Dim vTotPresRedu As Double
    Dim vTotInteres As Double
    Dim vTotCostCust As Double
    Dim vTotCostPreR As Double
    Dim vTotImpuesto As Double
    Dim vTotNeto As Double
    MousePointer = 11
    MuestraImpresion = True
    'vRTFImp = ""
    'Estado Renovado - 7  ; Falta limitar por fecha
    'SSQL = "SELECT cCodCta, nplazo, nNumRenov, nPrestamo, nMontoPag, nSaldoCap, nInteres, nInteresVenc, nCostoCusto, nCostoPreRe, nImpuesto " & _
        " FROM ConsdiarioPrend WHERE cestado = '7' AND cflag IS NULL " & _
        " AND DATEDIFF(dd,dFecha,'" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 "
        
    sSql = "SELECT T.cCodCta, T.nplazo, nRenoDupli ,(T.nCapital + T.nSaldoCap) nPrestamo, T.nCapital, T.nSaldoCap, T.nInteres, T.nIntVenc, T.nCostCust, T.nCostPrepRem, T.nImpuesto " & _
      " FROM TransPrenda T INNER JOIN CredPrenda CP ON T.cCodCta = CP.cCodCta WHERE T.cCodTran IN ('" & gsRenPrestamo & "','" & gsRenMorPrestamo & "','" & gsRenEnOtAg & "','" & gsRenMorEnOtAg & "','" & _
      gsRenEnOtCj & "','" & gsRenMorEnOtCj & "') AND cflag IS NULL " & _
      " AND DATEDIFF(dd,T.dFecha,'" & Format(txtFecha, "mm/dd/yyyy") & "') = 0 "
        
        If pbListGene = False Then      'Verifica el listado (General o no)
            sSql = sSql & " AND T.ccodusu = '" & gsCodUser & "' "
        ElseIf cboUsuario <> "<Consolidado>" Then
            'SSQL = SSQL & " AND ccodusu = '" & Trim(cboUsuario.Text) & "' "
            ' *** Cambio para la Operaciones con Interconexion
            sSql = sSql & " AND (T.ccodusu = '" & Trim(cboUsuario.Text) & "'  or cCodUsuRem = '" & Trim(cboUsuario.Text) & "' ) "
        End If
        If Len(Trim(vBoveda)) > 0 Then
            sSql = sSql & " AND CP.cAgeBoveda in " & vBoveda & " "
        End If
        sSql = sSql & " ORDER BY T.nPlazo, T.cCodCta"
    RegTransDiariaPrend.Open sSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
    If (RegTransDiariaPrend.BOF Or RegTransDiariaPrend.EOF) Then
        RegTransDiariaPrend.Close
        Set RegTransDiariaPrend = Nothing
        MsgBox " No existen Créditos Renovados ", vbInformation, " Aviso "
        MuestraImpresion = False
        MousePointer = 0
        Exit Sub
    Else
        vPage = vPage + 1: vCont = 0
        prgList.Min = 0
        prgList.Max = RegTransDiariaPrend.RecordCount
        If optImpresion(0).Value = True Then
            If prgList.Max > pPrevioMax Then
                RegTransDiariaPrend.Close
                Set RegTransDiariaPrend = Nothing
                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
                MousePointer = 0
                Exit Sub
            Else
                pPrevioMax = pPrevioMax - prgList.Max
            End If
            Cabecera "CredReno", vPage
        Else
            ImpreBegin False, pHojaFiMax
            vRTFImp = ""
            Cabecera "CredReno", vPage
            Print #ArcSal, ImpreCarEsp(vRTFImp);
            vRTFImp = ""
        End If
        prgList.Visible = True
        vIndice = 1
        vLineas = 7
        vTotPrestamo = Format(0, "#0.00"): vTotInteres = Format(0, "#0.00")
        vTotMontAmor = Format(0, "#0.00"): vTotPresRedu = Format(0, "#0.00")
        vTotCostCust = Format(0, "#0.00"): vTotCostPreR = Format(0, "#0.00")
        vTotImpuesto = Format(0, "#0.00"): vTotNeto = Format(0, "#.00")
        vSumPrestamo = Format(0, "#0.00"): vSumInteres = Format(0, "#0.00")
        vSumMontAmor = Format(0, "#0.00"): vSumPresRedu = Format(0, "#0.00")
        vSumCostCust = Format(0, "#0.00"): vSumCostPreR = Format(0, "#0.00")
        vSumImpuesto = Format(0, "#0.00"): vSumNeto = Format(0, "#.00")
        With RegTransDiariaPrend
            vPlazo = !nPlazo
            .MoveFirst
            Do While Not .EOF
                If Val(!nPlazo) <> Val(vPlazo) Then
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoLinea
                        vRTFImp = vRTFImp & ImpreFormat("Resumen", 8, 7) & ImpreFormat(vPlazo, 9, 0) & _
                            ImpreFormat(vSumPrestamo, 11) & ImpreFormat(vSumMontAmor, 7) & _
                            ImpreFormat(vSumPresRedu, 8) & ImpreFormat(vSumInteres, 7) & _
                            ImpreFormat(vSumImpuesto, 6) & ImpreFormat(vSumCostCust, 6) & _
                            ImpreFormat(vSumCostPreR, 6) & ImpreFormat(vSumNeto, 8) & gPrnSaltoLinea
                        vRTFImp = vRTFImp & gPrnSaltoLinea
                    Else
                        Print #ArcSal, ""
                        Print #ArcSal, ImpreFormat("Resumen", 8, 7) & ImpreFormat(vPlazo, 9, 0) & _
                            ImpreFormat(vSumPrestamo, 11) & ImpreFormat(vSumMontAmor, 7) & _
                            ImpreFormat(vSumPresRedu, 8) & ImpreFormat(vSumInteres, 7) & _
                            ImpreFormat(vSumImpuesto, 6) & ImpreFormat(vSumCostCust, 6) & _
                            ImpreFormat(vSumCostPreR, 6) & ImpreFormat(vSumNeto, 8)
                        Print #ArcSal, ""
                    End If
                    vLineas = vLineas + 3
                    
                    vTotPrestamo = vTotPrestamo + vSumPrestamo: vTotInteres = vTotInteres + vSumInteres
                    vTotMontAmor = vTotMontAmor + vSumMontAmor: vTotPresRedu = vTotPresRedu + vSumPresRedu
                    vTotCostCust = vTotCostCust + vSumCostCust: vTotCostPreR = vTotCostPreR + vSumCostPreR
                    vTotImpuesto = vTotImpuesto + vSumImpuesto: vTotNeto = vTotNeto + vSumNeto
                    vSumPrestamo = Format(0, "#0.00"): vSumInteres = Format(0, "#0.00")
                    vSumMontAmor = Format(0, "#0.00"): vSumPresRedu = Format(0, "#0.00")
                    vSumCostCust = Format(0, "#0.00"): vSumCostPreR = Format(0, "#0.00")
                    vSumImpuesto = Format(0, "#0.00"): vSumNeto = Format(0, "#0.00")
                End If
                vInteres = Abs(!nInteres) + Abs(!nIntVenc)
                vNeto = Abs(!nCapital) + Abs(vInteres) + Abs(!nImpuesto) + Abs(!nCostCust) + Abs(!nCostPrepRem)
                If optImpresion(0).Value = True Then
                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 4, 0) & ImpreFormat(!nPlazo, 3, 0) & Space(1) & _
                        FormatoContratro(!cCodCta) & ImpreFormat(!nRenoDupli, 4, 0) & ImpreFormat(!nPrestamo, 8) & _
                        ImpreFormat(Round(!nCapital, 2), 7) & ImpreFormat(!nSaldoCap, 8) & _
                        ImpreFormat(Round(vInteres, 2), 7) & ImpreFormat(!nImpuesto, 6) & _
                        ImpreFormat(Round(!nCostCust, 2), 6) & ImpreFormat(Round(!nCostPrepRem, 2), 6) & _
                        ImpreFormat(vNeto, 8) & gPrnSaltoLinea
                Else
                    Print #ArcSal, ImpreFormat(vIndice, 4, 0) & ImpreFormat(!nPlazo, 3, 0) & Space(1) & _
                        FormatoContratro(!cCodCta) & ImpreFormat(!nRenoDupli, 4, 0) & ImpreFormat(!nPrestamo, 8) & _
                        ImpreFormat(Round(!nCapital, 2), 7) & ImpreFormat(!nSaldoCap, 8) & _
                        ImpreFormat(Round(vInteres, 2), 7) & ImpreFormat(!nImpuesto, 6) & _
                        ImpreFormat(!nCostCust, 6) & ImpreFormat(!nCostPrepRem, 6) & _
                        ImpreFormat(vNeto, 8)
                End If
                vSumPrestamo = vSumPrestamo + Abs(!nPrestamo): vSumInteres = vSumInteres + Round(vInteres, 2)
                vSumMontAmor = vSumMontAmor + Abs(Round(!nCapital, 2)): vSumPresRedu = vSumPresRedu + IIf(IsNull(!nSaldoCap) = True, 0, !nSaldoCap) 'Abs(!nsaldocap)
                vSumCostCust = vSumCostCust + Abs(Round(!nCostCust, 2)): vSumCostPreR = vSumCostPreR + IIf(IsNull(!nCostPrepRem) = True, 0, Round(!nCostPrepRem, 2)) 'Abs(!nCostPrepRem)
                vSumImpuesto = vSumImpuesto + Abs(!nImpuesto): vSumNeto = vSumNeto + Abs(vNeto)
                vPlazo = !nPlazo
                vIndice = vIndice + 1: vLineas = vLineas + 1
                If vLineas > pLineasMax Then
                    vPage = vPage + 1
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoPagina
                        Cabecera "CredReno", vPage
                    Else
                        ImpreNewPage
                        vRTFImp = ""
                        Cabecera "CredReno", vPage
                        Print #ArcSal, ImpreCarEsp(vRTFImp);
                        vRTFImp = ""
                    End If
                    vLineas = 7
                End If
                vCont = vCont + 1
                prgList.Value = vCont
                Me.Caption = "Registro Nro.: " & vCont
                .MoveNext
            Loop
            vTotPrestamo = vTotPrestamo + vSumPrestamo: vTotInteres = vTotInteres + vSumInteres
            vTotMontAmor = vTotMontAmor + vSumMontAmor: vTotPresRedu = vTotPresRedu + vSumPresRedu
            vTotCostCust = vTotCostCust + vSumCostCust: vTotCostPreR = vTotCostPreR + vSumCostPreR
            vTotImpuesto = vTotImpuesto + vSumImpuesto: vTotNeto = vTotNeto + vSumNeto
            
            If optImpresion(0).Value = True Then
                vRTFImp = vRTFImp & gPrnSaltoLinea
                vRTFImp = vRTFImp & ImpreFormat("Resumen", 8, 7) & ImpreFormat(vPlazo, 9, 0) & _
                    ImpreFormat(vSumPrestamo, 11) & ImpreFormat(vSumMontAmor, 7) & _
                    ImpreFormat(vSumPresRedu, 8) & ImpreFormat(vSumInteres, 7) & _
                    ImpreFormat(vSumImpuesto, 6) & ImpreFormat(vSumCostCust, 6) & _
                    ImpreFormat(vSumCostPreR, 6) & ImpreFormat(vSumNeto, 8) & gPrnSaltoLinea
               
                vRTFImp = vRTFImp & gPrnSaltoLinea
                vRTFImp = vRTFImp & ImpreFormat("RESUMEN", 8, 7) & ImpreFormat(vTotPrestamo, 20) & _
                    ImpreFormat(vTotMontAmor, 7) & ImpreFormat(vTotPresRedu, 8) & _
                    ImpreFormat(vTotInteres, 7) & ImpreFormat(vTotImpuesto, 6) & _
                    ImpreFormat(vTotCostCust, 6) & ImpreFormat(vTotCostPreR, 6) & _
                    ImpreFormat(vTotNeto, 8) & gPrnSaltoLinea
            Else
                Print #ArcSal, ""
                Print #ArcSal, ImpreFormat("Resumen", 8, 7) & ImpreFormat(vPlazo, 9, 0) & _
                    ImpreFormat(vSumPrestamo, 11) & ImpreFormat(vSumMontAmor, 7) & _
                    ImpreFormat(vSumPresRedu, 8) & ImpreFormat(vSumInteres, 7) & _
                    ImpreFormat(vSumImpuesto, 6) & ImpreFormat(vSumCostCust, 6) & _
                    ImpreFormat(vSumCostPreR, 6) & ImpreFormat(vSumNeto, 8)
               
                Print #ArcSal, ""
                Print #ArcSal, ImpreFormat("RESUMEN", 8, 7) & ImpreFormat(vTotPrestamo, 20) & _
                    ImpreFormat(vTotMontAmor, 7) & ImpreFormat(vTotPresRedu, 8) & _
                    ImpreFormat(vTotInteres, 7) & ImpreFormat(vTotImpuesto, 6) & _
                    ImpreFormat(vTotCostCust, 6) & ImpreFormat(vTotCostPreR, 6) & _
                    ImpreFormat(vTotNeto, 8)
                ImpreEnd
            End If
        End With
        RegTransDiariaPrend.Close
        Set RegTransDiariaPrend = Nothing
        prgList.Visible = False
        prgList.Value = 0
    End If
    Me.Caption = vNameForm
    MousePointer = 0
End Sub

'Prendas Nuevas
Private Sub ImprimePrenNuev(vConexion As ADODB.Connection)
    Dim vPlazo As Integer
    Dim vIndice As Integer
    Dim vLineas As Integer
    Dim vSumValoTasa As Currency 'suma de los valores de tasacion
    Dim vSumOroNeto As Currency  'suma del oro neto
    Dim vTotValoTasa As Currency
    Dim vTotOroNeto As Currency
    MousePointer = 11
    MuestraImpresion = True
    'vRTFImp = ""
    'Estado Desemboldado- 1  ; Falta limitar por fecha
    'SSQL = "SELECT cCodCta, nPlazo, nValTasac, nOroNeto " & _
        " FROM ConsdiarioPrend " & _
        " WHERE cestado = '1' AND cflag IS NULL " & _
        " AND DATEDIFF(dd,dFecha,'" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 "
        
    sSql = "SELECT tp.cCodCta, tp.nPlazo, cp.nValTasac, cp.nOroNeto " & _
        " FROM TransPrenda TP INNER JOIN CredPrenda CP ON tp.ccodcta = cp.ccodcta " & _
        " WHERE tp.cCodTran = '" & gsDesPrestamo & "' AND tp.cflag IS NULL " & _
        " AND DATEDIFF(dd,tp.dFecha,'" & Format(txtFecha, "mm/dd/yyyy") & "') = 0 "
        
        If pbListGene = False Then      'Verifica el listado (General o no)
            sSql = sSql & " AND tp.ccodusu = '" & gsCodUser & "' "
        ElseIf cboUsuario <> "<Consolidado>" Then
            sSql = sSql & " AND tp.ccodusu = '" & Trim(cboUsuario.Text) & "' "
        End If
        If Len(Trim(vBoveda)) > 0 Then
            sSql = sSql & " AND CP.cAgeBoveda in " & vBoveda & " "
        End If
        sSql = sSql & " ORDER BY tp.nPlazo, tp.cCodCta"
    RegTransDiariaPrend.Open sSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
    If (RegTransDiariaPrend.BOF Or RegTransDiariaPrend.EOF) Then
        RegTransDiariaPrend.Close
        Set RegTransDiariaPrend = Nothing
        MsgBox " No existen Prendas Nuevas ", vbInformation, " Aviso "
        MuestraImpresion = False
        MousePointer = 0
        Exit Sub
    Else
        vPage = vPage + 1: vCont = 0
        prgList.Min = 0
        prgList.Max = RegTransDiariaPrend.RecordCount
        If optImpresion(0).Value = True Then
            If prgList.Max > pPrevioMax Then
                RegTransDiariaPrend.Close
                Set RegTransDiariaPrend = Nothing
                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
                MousePointer = 0
                Exit Sub
            Else
                pPrevioMax = pPrevioMax - prgList.Max
            End If
            Cabecera "PrenNuev", vPage
        Else
            ImpreBegin False, pHojaFiMax
            vRTFImp = ""
            Cabecera "PrenNuev", vPage
            Print #ArcSal, ImpreCarEsp(vRTFImp);
            vRTFImp = ""
        End If
        prgList.Visible = True
        vIndice = 1
        vLineas = 7
        vTotValoTasa = Format(0, "#0.00"): vTotOroNeto = Format(0, "#0.00")
        vSumValoTasa = Format(0, "#0.00"): vSumOroNeto = Format(0, "#0.00")
        With RegTransDiariaPrend
            vPlazo = !nPlazo
            .MoveFirst
            Do While Not .EOF
                If Val(!nPlazo) <> Val(vPlazo) Then
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoLinea
                        vRTFImp = vRTFImp & ImpreFormat("Resumen", 8, 9) & ImpreFormat(vPlazo, 7, 0) & _
                            ImpreFormat(vSumValoTasa, 18) & ImpreFormat(vSumOroNeto, 13) & gPrnSaltoLinea
                        vRTFImp = vRTFImp & gPrnSaltoLinea
                    Else
                        Print #ArcSal, ""
                        Print #ArcSal, ImpreFormat("Resumen", 8, 9) & ImpreFormat(vPlazo, 7, 0) & _
                            ImpreFormat(vSumValoTasa, 18) & ImpreFormat(vSumOroNeto, 13)
                        Print #ArcSal, ""
                    End If
                    vLineas = vLineas + 3
                    
                    vTotValoTasa = vTotValoTasa + vSumValoTasa: vTotOroNeto = vTotOroNeto + vSumOroNeto
                    vSumValoTasa = Format(0, "#0.00"): vSumOroNeto = Format(0, "#0.00")
                End If
                If optImpresion(0).Value = True Then
                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 5, 0) & ImpreFormat(!nPlazo, 6, 0) & Space(2) & FormatoContratro(!cCodCta) & _
                        ImpreFormat(!nvaltasac, 14) & ImpreFormat(!noroneto, 13) & gPrnSaltoLinea
                Else
                    Print #ArcSal, ImpreFormat(vIndice, 5, 0) & ImpreFormat(!nPlazo, 6, 0) & Space(2) & FormatoContratro(!cCodCta) & _
                        ImpreFormat(!nvaltasac, 14) & ImpreFormat(!noroneto, 13)
                End If
                vSumValoTasa = vSumValoTasa + IIf(IsNull(!nvaltasac), 0, Abs(!nvaltasac)): vSumOroNeto = vSumOroNeto + IIf(IsNull(!noroneto), 0, Abs(!noroneto))
                vPlazo = !nPlazo
                vIndice = vIndice + 1: vLineas = vLineas + 1
                If vLineas > pLineasMax Then
                    vPage = vPage + 1
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoPagina
                        Cabecera "PrenNuev", vPage
                    Else
                        ImpreNewPage
                        vRTFImp = ""
                        Cabecera "PrenNuev", vPage
                        Print #ArcSal, ImpreCarEsp(vRTFImp);
                        vRTFImp = ""
                    End If
                    vLineas = 7
                End If
                vCont = vCont + 1
                prgList.Value = vCont
                Me.Caption = "Registro Nro.: " & vCont
                .MoveNext
            Loop
            vTotValoTasa = vTotValoTasa + vSumValoTasa: vTotOroNeto = vTotOroNeto + vSumOroNeto
            If optImpresion(0).Value = True Then
                vRTFImp = vRTFImp & gPrnSaltoLinea
                vRTFImp = vRTFImp & ImpreFormat("Resumen", 8, 9) & ImpreFormat(vPlazo, 7, 0) & _
                    ImpreFormat(vSumValoTasa, 18) & ImpreFormat(vSumOroNeto, 13) & gPrnSaltoLinea
                vRTFImp = vRTFImp & gPrnSaltoLinea
                vRTFImp = vRTFImp & ImpreFormat("RESUMEN", 8, 9) & ImpreFormat(vTotValoTasa, 25) & _
                    ImpreFormat(vTotOroNeto, 15) & gPrnSaltoLinea
            Else
                Print #ArcSal, ""
                Print #ArcSal, ImpreFormat("Resumen", 8, 9) & ImpreFormat(vPlazo, 7, 0) & _
                    ImpreFormat(vSumValoTasa, 18) & ImpreFormat(vSumOroNeto, 13)
                Print #ArcSal, ""
                Print #ArcSal, ImpreFormat("RESUMEN", 8, 9) & ImpreFormat(vTotValoTasa, 25) & _
                    ImpreFormat(vTotOroNeto, 15)
                ImpreEnd
            End If
        End With
        RegTransDiariaPrend.Close
        Set RegTransDiariaPrend = Nothing
        prgList.Visible = False
        prgList.Value = 0
    End If
    Me.Caption = vNameForm
    MousePointer = 0
End Sub

'Prendas Devueltas
Private Sub ImprimePrenDevu(vConexion As ADODB.Connection)
    Dim vPlazo As Integer
    Dim vIndice As Integer
    Dim vLineas As Integer
    Dim vSumValoTasa As Currency 'suma de los valores de tasacion
    Dim vSumOroNeto As Currency  'suma del oro neto
    Dim vTotValoTasa As Currency
    Dim vTotOroNeto As Currency
    MousePointer = 11
    MuestraImpresion = True
    'vRTFImp = ""
    'Estado Cancelada - 3  ; Falta limitar por fecha
    'y recogida Dentro del Plazo - cRescDife = 'D'
    'SSQL = "SELECT cCodCta, nPlazo, nValTasac, nOroNeto " & _
        " FROM ConsdiarioPrend WHERE cestado = '3' AND cflag IS NULL " & _
        " AND cRescDife = 'D' AND DATEDIFF(dd,dFecha,'" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 "
    
    sSql = "SELECT cCodCta, nPlazo, nValTasac, nOroNeto " & _
        " FROM TransPrenda WHERE cCodTran IN ('" & gsDevJoyas & "','" & gsDevJoyasEOA & "') AND cflag IS NULL " & _
        " AND DATEDIFF(dd,dFecha,'" & Format(txtFecha, "mm/dd/yyyy") & "') = 0 "
        
        If pbListGene = False Then      'Verifica el listado (General o no)
            sSql = sSql & " AND ccodusu = '" & gsCodUser & "' "
        ElseIf cboUsuario <> "<Consolidado>" Then
            sSql = sSql & " AND ccodusu = '" & Trim(cboUsuario.Text) & "' "
        End If
        sSql = sSql & " ORDER BY nPlazo, cCodCta"
    RegTransDiariaPrend.Open sSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
    If (RegTransDiariaPrend.BOF Or RegTransDiariaPrend.EOF) Then
        RegTransDiariaPrend.Close
        Set RegTransDiariaPrend = Nothing
        MsgBox " No existen Prendas Devueltas ", vbInformation, " Aviso "
        MuestraImpresion = False
        MousePointer = 0
        Exit Sub
    Else
        vPage = vPage + 1: vCont = 0
        prgList.Min = 0
        prgList.Max = RegTransDiariaPrend.RecordCount
        If optImpresion(0).Value = True Then
            If prgList.Max > pPrevioMax Then
                RegTransDiariaPrend.Close
                Set RegTransDiariaPrend = Nothing
                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
                MousePointer = 0
                Exit Sub
            Else
                pPrevioMax = pPrevioMax - prgList.Max
            End If
            Cabecera "PrenDevu", vPage
        Else
            ImpreBegin False, pHojaFiMax
            vRTFImp = ""
            Cabecera "PrenDevu", vPage
            Print #ArcSal, ImpreCarEsp(vRTFImp);
            vRTFImp = ""
        End If
        prgList.Visible = True
        vIndice = 1
        vLineas = 7
        vTotValoTasa = Format(0, "#0.00"): vTotOroNeto = Format(0, "#0.00")
        vSumValoTasa = Format(0, "#0.00"): vSumOroNeto = Format(0, "#0.00")
        With RegTransDiariaPrend
            vPlazo = !nPlazo
            .MoveFirst
            Do While Not .EOF
                If Val(!nPlazo) <> Val(vPlazo) Then
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoLinea
                        vRTFImp = vRTFImp & ImpreFormat("Resumen", 8, 8) & ImpreFormat(vPlazo, 7, 0) & _
                            ImpreFormat(vSumValoTasa, 18) & ImpreFormat(vSumOroNeto, 15) & gPrnSaltoLinea
                        vRTFImp = vRTFImp & gPrnSaltoLinea
                    Else
                        Print #ArcSal, ""
                        Print #ArcSal, ImpreFormat("Resumen", 8, 8) & ImpreFormat(vPlazo, 7, 0) & _
                            ImpreFormat(vSumValoTasa, 18) & ImpreFormat(vSumOroNeto, 15)
                        Print #ArcSal, ""
                    End If
                    vLineas = vLineas + 3
                    
                    vTotValoTasa = vTotValoTasa + vSumValoTasa: vTotOroNeto = vTotOroNeto + vSumOroNeto
                    vSumValoTasa = Format(0, "#0.00"): vSumOroNeto = Format(0, "#0.00")
                End If
                If optImpresion(0).Value = True Then
                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 6, 0) & ImpreFormat(!nPlazo, 7, 0) & _
                        ImpreFormat(!cCodCta, 12, 2) & ImpreFormat(Abs(!nvaltasac), 14) & _
                        ImpreFormat(Abs(!noroneto), 15) & gPrnSaltoLinea
                Else
                    Print #ArcSal, ImpreFormat(vIndice, 6, 0) & ImpreFormat(!nPlazo, 7, 0) & _
                        ImpreFormat(!cCodCta, 12, 2) & ImpreFormat(Abs(!nvaltasac), 14) & _
                        ImpreFormat(Abs(!noroneto), 15)
                End If
                vSumValoTasa = vSumValoTasa + Abs(!nvaltasac): vSumOroNeto = vSumOroNeto + Abs(!noroneto)
                
                vPlazo = !nPlazo
                vIndice = vIndice + 1: vLineas = vLineas + 1
                If vLineas > pLineasMax Then
                    vPage = vPage + 1
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoPagina
                        Cabecera "PrenDevu", vPage
                    Else
                        ImpreNewPage
                        vRTFImp = ""
                        Cabecera "PrenDevu", vPage
                        Print #ArcSal, ImpreCarEsp(vRTFImp);
                        vRTFImp = ""
                    End If
                    vLineas = 7
                End If
                vCont = vCont + 1
                prgList.Value = vCont
                Me.Caption = "Registro Nro.: " & vCont
                .MoveNext
            Loop
            vTotValoTasa = vTotValoTasa + vSumValoTasa: vTotOroNeto = vTotOroNeto + vSumOroNeto
            If optImpresion(0).Value = True Then
                vRTFImp = vRTFImp & gPrnSaltoLinea
                vRTFImp = vRTFImp & ImpreFormat("Resumen", 8, 8) & ImpreFormat(vPlazo, 7, 0) & _
                    ImpreFormat(vSumValoTasa, 18) & ImpreFormat(vSumOroNeto, 15) & gPrnSaltoLinea
                vRTFImp = vRTFImp & gPrnSaltoLinea
                vRTFImp = vRTFImp & ImpreFormat("RESUMEN", 8, 8) & ImpreFormat(vTotValoTasa, 25) & _
                    ImpreFormat(vTotOroNeto, 15) & gPrnSaltoLinea
            Else
                Print #ArcSal, ""
                Print #ArcSal, ImpreFormat("Resumen", 8, 8) & ImpreFormat(vPlazo, 7, 0) & _
                    ImpreFormat(vSumValoTasa, 18) & ImpreFormat(vSumOroNeto, 15)
                Print #ArcSal, ""
                Print #ArcSal, ImpreFormat("RESUMEN", 8, 8) & ImpreFormat(vTotValoTasa, 25) & _
                    ImpreFormat(vTotOroNeto, 15)
                ImpreEnd
            End If
        End With
        RegTransDiariaPrend.Close
        Set RegTransDiariaPrend = Nothing
        prgList.Visible = False
        prgList.Value = 0
    End If
    Me.Caption = vNameForm
    MousePointer = 0
End Sub

'Prendas Rescatadas en Condición de Diferidas
Private Sub ImprimePrenRescCondDife(vConexion As ADODB.Connection)
    Dim vPlazo As Integer
    Dim vIndice As Integer
    Dim vLineas As Integer
    Dim vSumValoTasa As Double 'suma de los valores de tasacion
    Dim vSumOroNeto As Double  'suma del oro neto
    Dim vTotValoTasa As Double
    Dim vTotOroNeto As Double
    MousePointer = 11
    MuestraImpresion = True
    'gdFecSis = "27/11/1999"
    'vRTFImp = ""
    'Estado Cancelada - 3  ; Falta limitar por fecha
    'y recogida en Dentro y Fuera del Plazo - cRescDife = 'F'
    'SSQL = "SELECT cCodCta, nPlazo, nValTasac, nOroNeto " & _
        " FROM ConsdiarioPrend WHERE cestado = '3' AND cflag IS NULL " & _
        " AND cRescDife IN('D','F') AND DATEDIFF(dd,dFecha,'" & Format(gdFecSis, "mm/dd/yyyy") & "') = 0 "
    
    sSql = "SELECT tp.cCodCta, tp.nPlazo, tp.nValTasac, tp.nOroNeto, cp.cAgeBoveda " & _
        " FROM TransPrenda tp INNER JOIN CredPrenda CP ON tp.ccodcta = cp.ccodcta WHERE tp.cCodTran IN ('" & gsDevJoyas & "','" & gsDevJoyasEOA & "') AND tp.cflag IS NULL " & _
        " AND DATEDIFF(dd,tp.dFecha,'" & Format(txtFecha, "mm/dd/yyyy") & "') = 0 "
        
        If pbListGene = False Then      'Verifica el listado (General o no)
            sSql = sSql & " AND tp.ccodusu = '" & gsCodUser & "' "
        ElseIf cboUsuario <> "<Consolidado>" Then
            ' *** Cambio para la Operaciones con Interconexion
            sSql = sSql & " AND (tp.cCodUsu = '" & Trim(cboUsuario.Text) & "'  or cCodUsuRem = '" & Trim(cboUsuario.Text) & "' ) "
        End If
        If Len(Trim(vBoveda)) > 0 Then
            sSql = sSql & " AND CP.cAgeBoveda in " & vBoveda & " "
        End If
        sSql = sSql & " ORDER BY tp.nPlazo, tp.cCodCta"
    RegTransDiariaPrend.Open sSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
    If (RegTransDiariaPrend.BOF Or RegTransDiariaPrend.EOF) Then
        RegTransDiariaPrend.Close
        Set RegTransDiariaPrend = Nothing
        MsgBox " No existen Prendas Rescatadas en Condición de Diferidas ", vbInformation, " Aviso "
        MuestraImpresion = False
        MousePointer = 0
        Exit Sub
    Else
        vPage = vPage + 1: vCont = 0
        prgList.Min = 0
        prgList.Max = RegTransDiariaPrend.RecordCount
        If optImpresion(0).Value = True Then
            If prgList.Max > pPrevioMax Then
                RegTransDiariaPrend.Close
                Set RegTransDiariaPrend = Nothing
                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
                MousePointer = 0
                Exit Sub
            Else
                pPrevioMax = pPrevioMax - prgList.Max
            End If
            Cabecera "PrReCoDi", vPage
        Else
            ImpreBegin False, pHojaFiMax
            vRTFImp = ""
            Cabecera "PrReCoDi", vPage
            Print #ArcSal, ImpreCarEsp(vRTFImp);
            vRTFImp = ""
        End If
        prgList.Visible = True
        vIndice = 1
        vLineas = 7
        vTotValoTasa = Format(0, "#0.00"): vTotOroNeto = Format(0, "#0.00")
        vSumValoTasa = Format(0, "#0.00"): vSumOroNeto = Format(0, "#0.00")
        With RegTransDiariaPrend
            vPlazo = !nPlazo
            .MoveFirst
            Do While Not .EOF
                If Val(!nPlazo) <> Val(vPlazo) Then
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoLinea
                        vRTFImp = vRTFImp & ImpreFormat("Resumen", 8, 8) & ImpreFormat(vPlazo, 7, 0) & _
                            ImpreFormat(vSumValoTasa, 18) & ImpreFormat(vSumOroNeto, 15) & gPrnSaltoLinea
                        vRTFImp = vRTFImp & gPrnSaltoLinea
                    Else
                        Print #ArcSal, ""
                        Print #ArcSal, ImpreFormat("Resumen", 8, 8) & ImpreFormat(vPlazo, 7, 0) & _
                            ImpreFormat(vSumValoTasa, 18) & ImpreFormat(vSumOroNeto, 15)
                        Print #ArcSal, ""
                    End If
                    vLineas = vLineas + 3
                    
                    vTotValoTasa = vTotValoTasa + vSumValoTasa: vTotOroNeto = vTotOroNeto + vSumOroNeto
                    vSumValoTasa = Format(0, "#0.00"): vSumOroNeto = Format(0, "#0.00")
                End If
                If optImpresion(0).Value = True Then
                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 5, 0) & ImpreFormat(!nPlazo, 5, 0) & Space(2) & _
                        FormatoContratro(!cCodCta) & ImpreFormat(Abs(!nvaltasac), 14) & _
                        ImpreFormat(Abs(!noroneto), 15) & Space(4) & Mid(!cAgeBoveda, 4, 2) & gPrnSaltoLinea
                Else
                    Print #ArcSal, ImpreFormat(vIndice, 5, 0) & ImpreFormat(!nPlazo, 5, 0) & Space(2) & _
                        FormatoContratro(!cCodCta) & ImpreFormat(Abs(!nvaltasac), 14) & _
                        ImpreFormat(Abs(!noroneto), 15) & Space(4) & Mid(!cAgeBoveda, 4, 2)
                End If

                vSumValoTasa = vSumValoTasa + Abs(!nvaltasac): vSumOroNeto = vSumOroNeto + IIf(IsNull(!noroneto) = True, Format(0, "#0.00"), Format(Abs(!noroneto), "#0.00"))
                
                vPlazo = !nPlazo
                vIndice = vIndice + 1: vLineas = vLineas + 1
                If vLineas > pLineasMax Then
                    vPage = vPage + 1
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoPagina
                        Cabecera "PrReCoDi", vPage
                    Else
                        ImpreNewPage
                        vRTFImp = ""
                        Cabecera "PrReCoDi", vPage
                        Print #ArcSal, ImpreCarEsp(vRTFImp);
                        vRTFImp = ""
                    End If
                    vLineas = 7
                End If
                vCont = vCont + 1
                prgList.Value = vCont
                Me.Caption = "Registro Nro.: " & vCont
                .MoveNext
            Loop
            vTotValoTasa = vTotValoTasa + vSumValoTasa: vTotOroNeto = vTotOroNeto + vSumOroNeto
            If optImpresion(0).Value = True Then
                vRTFImp = vRTFImp & gPrnSaltoLinea
                vRTFImp = vRTFImp & ImpreFormat("Resumen", 8, 8) & ImpreFormat(vPlazo, 7, 0) & _
                    ImpreFormat(vSumValoTasa, 18) & ImpreFormat(vSumOroNeto, 15) & gPrnSaltoLinea
                vRTFImp = vRTFImp & gPrnSaltoLinea
                vRTFImp = vRTFImp & ImpreFormat("RESUMEN", 8, 8) & ImpreFormat(vTotValoTasa, 25) & _
                    ImpreFormat(vTotOroNeto, 15) & gPrnSaltoLinea
            Else
                Print #ArcSal, ""
                Print #ArcSal, ImpreFormat("Resumen", 8, 8) & ImpreFormat(vPlazo, 7, 0) & _
                    ImpreFormat(vSumValoTasa, 18) & ImpreFormat(vSumOroNeto, 15)
                Print #ArcSal,
                Print #ArcSal, ImpreFormat("RESUMEN", 8, 8) & ImpreFormat(vTotValoTasa, 25) & _
                    ImpreFormat(vTotOroNeto, 15)
                ImpreEnd
            End If
        End With
        RegTransDiariaPrend.Close
        Set RegTransDiariaPrend = Nothing
        prgList.Visible = False
        prgList.Value = 0
    End If
    Me.Caption = vNameForm
    MousePointer = 0
End Sub

'Contratos que se encuentran en condición de diferidos
Private Sub ImprimeContCondDife(vConexion As ADODB.Connection)
    Dim RegTmp As New ADODB.Recordset
    Dim vIndice As Integer
    Dim vLineas As Integer
    Dim vSumValoTasa As Double 'suma de los valores de tasacion
    Dim vSumOroNeto As Double  'suma del oro neto
    Dim vDiasDife As Double
    Dim lsBloq As String * 5
    MousePointer = 11
    MuestraImpresion = True
    'vRTFImp = ""
    'Estado Rescatado
    sSql = "SELECT cCodCta, nPlazo, nValTasac, nOroNeto, dFecCance, cAgeBoveda " & _
        " FROM CredPrenda " & _
        " WHERE cestado = '2' " & _
        " AND DATEDIFF(dd,dfeccance,'" & Format(gdFecSis, "mm/dd/yyyy") & "') >= 0 "
    If Len(Trim(vBoveda)) > 0 Then
        sSql = sSql & " AND cAgeBoveda in " & vBoveda & " "
    End If
    sSql = sSql & " ORDER BY  cCodCta"
    RegTmp.Open sSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
    If (RegTmp.BOF Or RegTmp.EOF) Then
        RegTmp.Close
        Set RegTmp = Nothing
        MsgBox " No existen Contratos en Condición de Diferidos ", vbInformation, " Aviso "
        MuestraImpresion = False
        MousePointer = 0
        Exit Sub
    Else
        vPage = vPage + 1: vCont = 0
        prgList.Min = 0
        prgList.Max = RegTmp.RecordCount
        If optImpresion(0).Value = True Then
            If prgList.Max > pPrevioMax Then
                RegTmp.Close
                Set RegTmp = Nothing
                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
                MousePointer = 0
                Exit Sub
            Else
                pPrevioMax = pPrevioMax - prgList.Max
            End If
            Cabecera "ContCondDife", vPage
        Else
            ImpreBegin False, pHojaFiMax
            vRTFImp = ""
            Cabecera "ContCondDife", vPage
            Print #ArcSal, ImpreCarEsp(vRTFImp);
            vRTFImp = ""
        End If
        prgList.Visible = True
        vIndice = 1
        vLineas = 7
        vSumValoTasa = 0
        vSumOroNeto = 0
        With RegTmp
            .MoveFirst
            Do While Not .EOF
                vDiasDife = DateDiff("d", !dfeccance, gdFecSis) + 1
                lsBloq = IIf(IsCtaBlo(!cCodCta, vConexion), "BLOQ", "   ")
                If optImpresion(0).Value = True Then
                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 5, 0) & Space(2) & FormatoContratro(!cCodCta) & _
                        ImpreFormat(!nPlazo, 5, 0) & ImpreFormat(!nvaltasac, 15) & _
                        ImpreFormat(!noroneto, 15) & Space(4) & _
                        Format(!dfeccance, "dd/mm/yyyy") & ImpreFormat(vDiasDife, 10, 0) & Space(3) & lsBloq & Space(3) & Mid(!cAgeBoveda, 4, 2) & gPrnSaltoLinea
                Else
                    Print #ArcSal, ImpreFormat(vIndice, 5, 0) & Space(2) & FormatoContratro(!cCodCta) & _
                        ImpreFormat(!nPlazo, 5, 0) & ImpreFormat(!nvaltasac, 15) & _
                        ImpreFormat(!noroneto, 15) & Space(4) & _
                        Format(!dfeccance, "dd/mm/yyyy") & ImpreFormat(vDiasDife, 10, 0) & Space(3) & lsBloq & Space(3) & Mid(!cAgeBoveda, 4, 2)
                End If
                vSumValoTasa = vSumValoTasa + IIf(IsNull(!nvaltasac), 0, Abs(!nvaltasac))
                vSumOroNeto = vSumOroNeto + IIf(IsNull(!noroneto), 0, Abs(!noroneto))
                vIndice = vIndice + 1: vLineas = vLineas + 1
                If vLineas > pLineasMax Then
                    vPage = vPage + 1
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoPagina
                        Cabecera "ContCondDife", vPage
                    Else
                        ImpreNewPage
                        vRTFImp = ""
                        Cabecera "ContCondDife", vPage
                        Print #ArcSal, ImpreCarEsp(vRTFImp);
                        vRTFImp = ""
                    End If
                    vLineas = 7
                End If
                vCont = vCont + 1
                prgList.Value = vCont
                Me.Caption = "Registro Nro.: " & vCont
                .MoveNext
            Loop
            If optImpresion(0).Value = True Then
                vRTFImp = vRTFImp & gPrnSaltoLinea
                vRTFImp = vRTFImp & ImpreFormat("Resumen", 8, 9) & ImpreFormat(vSumValoTasa, 25) & _
                    ImpreFormat(vSumOroNeto, 15)
            Else
                Print #ArcSal, ""
                Print #ArcSal, ImpreFormat("Resumen", 8, 9) & ImpreFormat(vSumValoTasa, 25) & _
                    ImpreFormat(vSumOroNeto, 15)
                ImpreEnd
            End If
        End With
        RegTmp.Close
        Set RegTmp = Nothing
        prgList.Visible = False
        prgList.Value = 0
    End If
    Me.Caption = vNameForm
    MousePointer = 0
End Sub

'Listado de las Operaciones Extornadas durante el día
Private Sub ImprimeOperExto(vConexion As ADODB.Connection)
    Dim RegOpe As New ADODB.Recordset
    Dim vIndice As Integer
    Dim vLineas As Integer
    Dim vSumMonto As Double  'suma del oro neto
    MousePointer = 11
    MuestraImpresion = True
    'vRTFImp = ""
    'Estado Rescatado
    sSql = "SELECT cCodCta, cCodTran, dfecha, nMontoTran, cCodUsu, cCodUsuRem " & _
        " FROM TransPrenda " & _
        " WHERE cflag IS NOT NULL " & _
        " AND DATEDIFF(dd,dFecha,'" & Format(txtFecha, "mm/dd/yyyy") & "') = 0 "
        If pbListGene = False Then      'Verifica el listado (General o no)
            sSql = sSql & " AND ccodusu = '" & gsCodUser & "' "
        ElseIf cboUsuario <> "<Consolidado>" Then
            'SSQL = SSQL & " AND ccodusu = '" & Trim(cboUsuario.Text) & "' "
            ' *** Cambio para la Operaciones con Interconexion
            sSql = sSql & " AND (ccodusu = '" & Trim(cboUsuario.Text) & "'  or cCodUsuRem = '" & Trim(cboUsuario.Text) & "' ) "
            
        End If
        sSql = sSql & " ORDER BY  cCodCta, nnumtran"
    RegOpe.Open sSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
    If (RegOpe.BOF Or RegOpe.EOF) Then
        RegOpe.Close
        Set RegOpe = Nothing
        MsgBox " No existen Operaciones Extornadas en el día ", vbInformation, " Aviso "
        MuestraImpresion = False
        MousePointer = 0
        Exit Sub
    Else
        vPage = vPage + 1: vCont = 0
        prgList.Min = 0
        prgList.Max = RegOpe.RecordCount
        If optImpresion(0).Value = True Then
            If prgList.Max > pPrevioMax Then
                RegOpe.Close
                Set RegOpe = Nothing
                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
                MousePointer = 0
                Exit Sub
            Else
                pPrevioMax = pPrevioMax - prgList.Max
            End If
            Cabecera "OperExto", vPage
        Else
            ImpreBegin False, pHojaFiMax
            vRTFImp = ""
            Cabecera "OperExto", vPage
            Print #ArcSal, ImpreCarEsp(vRTFImp);
            vRTFImp = ""
        End If
        prgList.Visible = True
        vIndice = 1
        vLineas = 7
        vSumMonto = Format(0, "#0.00")
        With RegOpe
            .MoveFirst
            Do While Not .EOF
                'vOperacion = ContratoOperacion(!ccodtran)
                If optImpresion(0).Value = True Then
                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 5, 0) & Space(2) & FormatoContratro(!cCodCta) & _
                        Space(2) & Format(!dFecha, "dd/mm/yyyy hh:mm:ss") & ImpreFormat(ContratoOperacion(!ccodtran), 28, 2) & _
                        ImpreFormat(!nMontoTran, 11) & ImpreFormat(IIf(IsNull(!cCodUsu), !cCodusurem, !cCodUsu), 4, 8) & gPrnSaltoLinea
                Else
                    Print #ArcSal, ImpreFormat(vIndice, 5, 0) & Space(2) & FormatoContratro(!cCodCta) & _
                        Space(2) & Format(!dFecha, "dd/mm/yyyy hh:mm:ss") & ImpreFormat(ContratoOperacion(!ccodtran), 28, 2) & _
                        ImpreFormat(!nMontoTran, 11) & ImpreFormat(IIf(IsNull(!cCodUsu), !cCodusurem, !cCodUsu), 4, 8)
                End If
                vSumMonto = vSumMonto + !nMontoTran
                vIndice = vIndice + 1: vLineas = vLineas + 1
                If vLineas > pLineasMax Then
                    vPage = vPage + 1
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoPagina
                        Cabecera "OperExto", vPage
                    Else
                        ImpreNewPage
                        vRTFImp = ""
                        Cabecera "OperExto", vPage
                        Print #ArcSal, ImpreCarEsp(vRTFImp);
                        vRTFImp = ""
                    End If
                    vLineas = 7
                End If
                vCont = vCont + 1
                prgList.Value = vCont
                Me.Caption = "Registro Nro.: " & vCont
                .MoveNext
            Loop
            If optImpresion(0).Value = True Then
                vRTFImp = vRTFImp & gPrnSaltoLinea
                vRTFImp = vRTFImp & ImpreFormat("RESUMEN", 8, 8) & ImpreFormat(vSumMonto, 68) & gPrnSaltoLinea
            Else
                Print #ArcSal, ""
                Print #ArcSal, ImpreFormat("RESUMEN", 8, 8) & ImpreFormat(vSumMonto, 68)
                ImpreEnd
            End If
        End With
        RegOpe.Close
        Set RegOpe = Nothing
        prgList.Visible = False
        prgList.Value = 0
    End If
    Me.Caption = vNameForm
    MousePointer = 0
End Sub

'Pago de sobrantes
Private Sub ImprimePagoSobr(vConexion As ADODB.Connection)
    'Dim vNombre As String * 37
    Dim vIndice As Integer  'contador de Item
    Dim vLineas As Integer
    Dim vSumPagsob As Double
    MousePointer = 11
    MuestraImpresion = True
    'vRTFImp = ""
    sSql = "SELECT tp.cCodCta, tp.nMontoTran, tp.dfecha, tp.ccodusu " & _
        " FROM TransPrenda TP" & _
        " WHERE tp.ccodtran IN ('" & gsPagSobrante & "','" & gsPagSobDeOtAg & "') AND nmontotran > 0 " & _
        " AND DATEDIFF(dd, tp.dFecha,'" & Format(txtFecha, "mm/dd/yyyy") & "') = 0 "
        If pbListGene = False Then      'Verifica el listado (General o no)
            sSql = sSql & " AND tp.ccodusu = '" & gsCodUser & "' "
        ElseIf cboUsuario <> "<Consolidado>" Then
            'SSQL = SSQL & " AND tp.ccodusu = '" & Trim(cboUsuario.Text) & "' "
            ' *** Cambio para la Operaciones con Interconexion
            sSql = sSql & " AND (tp.ccodusu = '" & Trim(cboUsuario.Text) & "'  or tp.cCodUsuRem = '" & Trim(cboUsuario.Text) & "' ) "
        End If
        sSql = sSql & " ORDER BY tp.cCodCta"
    RegCredPrend.Open sSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
    If (RegCredPrend.BOF Or RegCredPrend.EOF) Then
        RegCredPrend.Close
        Set RegCredPrend = Nothing
        MsgBox " No existen Pago de Sobrantes ", vbInformation, " Aviso "
        MuestraImpresion = False
        MousePointer = 0
        Exit Sub
    Else
        vPage = vPage + 1: vCont = 0
        prgList.Min = 0
        prgList.Max = RegCredPrend.RecordCount
        If optImpresion(0).Value = True Then
            If prgList.Max > pPrevioMax Then
                RegCredPrend.Close
                Set RegCredPrend = Nothing
                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
                MousePointer = 0
                Exit Sub
            Else
                pPrevioMax = pPrevioMax - prgList.Max
            End If
            Cabecera "PagoSobr", vPage
        Else
            ImpreBegin False, pHojaFiMax
            vRTFImp = ""
            Cabecera "PagoSobr", vPage
            Print #ArcSal, ImpreCarEsp(vRTFImp);
            vRTFImp = ""
        End If
        prgList.Visible = True
        vIndice = 1
        vLineas = 7
        vSumPagsob = 0
        With RegCredPrend
            .MoveFirst
            Do While Not .EOF
                If optImpresion(0).Value = True Then
                    vRTFImp = vRTFImp & ImpreFormat(vIndice, 5, 0) & Space(2) & FormatoContratro(!cCodCta) & Space(1) & Format(!dFecha, "dd/mm/yyyy hh:mm:ss") & _
                        ImpreFormat(!nMontoTran, 10) & ImpreFormat(!cCodUsu, 4, 6) & gPrnSaltoLinea
                Else
                    Print #ArcSal, ImpreFormat(vIndice, 5, 0) & Space(2) & FormatoContratro(!cCodCta) & Space(1) & Format(!dFecha, "dd/mm/yyyy hh:mm:ss") & _
                        ImpreFormat(!nMontoTran, 10) & ImpreFormat(!cCodUsu, 4, 6)
                End If
                vSumPagsob = vSumPagsob + !nMontoTran
                vIndice = vIndice + 1
                vLineas = vLineas + 1
                If vLineas > pLineasMax Then
                    vPage = vPage + 1
                    If optImpresion(0).Value = True Then
                        vRTFImp = vRTFImp & gPrnSaltoPagina
                        Cabecera "PagoSobr", vPage
                    Else
                        ImpreNewPage
                        vRTFImp = ""
                        Cabecera "PagoSobr", vPage
                        Print #ArcSal, ImpreCarEsp(vRTFImp);
                        vRTFImp = ""
                    End If
                    vLineas = 7
                End If
                vCont = vCont + 1
                prgList.Value = vCont
                Me.Caption = "Registro Nro.: " & vCont
                .MoveNext
            Loop
            If optImpresion(0).Value = True Then
                vRTFImp = vRTFImp & gPrnSaltoLinea
                vRTFImp = vRTFImp & ImpreFormat("Total", 8, 32) & ImpreFormat(vSumPagsob, 12) & gPrnSaltoLinea
            Else
                Print #ArcSal, ""
                Print #ArcSal, ImpreFormat("Total", 8, 32) & ImpreFormat(vSumPagsob, 12)
                ImpreEnd
            End If
        End With
        RegCredPrend.Close
        Set RegCredPrend = Nothing
        prgList.Visible = False
        prgList.Value = 0
    End If
    Me.Caption = vNameForm
    MousePointer = 0
End Sub

'Procedimiento Externo que permitira indicar al formulario (al cargar) si las impresiones serán
'de nivel de usuario o supervisor
Public Sub ListadoGeneral(pbValor As Boolean)
    pbListGene = pbValor
End Sub

'Parametros para el formulario
Private Sub CargaParametros()
    dbCmact.CommandTimeout = 50
    pPrevioMax = 4000
    pLineasMax = 56
    pHojaFiMax = 66
End Sub

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub


' Valida fechas
Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFecha) Then Me.cmdAgencia.SetFocus
End If
End Sub

Private Sub txtFecha_Validate(Cancel As Boolean)
    If Not ValFecha(txtFecha) Then
        Cancel = True
    End If
End Sub

Private Function FormatoContratro(pContrato As String) As String
FormatoContratro = Mid(pContrato, 1, 2) & "-" & Mid(pContrato, 3, 4) & "-" & Mid(pContrato, 7, 5) & "-" & Mid(pContrato, 12, 1)
End Function
