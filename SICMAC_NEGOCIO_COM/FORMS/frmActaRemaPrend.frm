VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmColPRemateActa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio : Acta de Remate"
   ClientHeight    =   5205
   ClientLeft      =   1035
   ClientTop       =   705
   ClientWidth     =   6945
   Icon            =   "frmActaRemaPrend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdImprimir 
      Cancel          =   -1  'True
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   4635
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   4635
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Top             =   4635
      Width           =   975
   End
   Begin VB.Frame fraContenedor 
      Height          =   4320
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   90
      Width           =   6570
      Begin VB.Frame Frame1 
         Caption         =   "Remates "
         Height          =   2400
         Left            =   4680
         TabIndex        =   27
         Top             =   720
         Width           =   1365
         Begin VB.ListBox List1 
            Height          =   2010
            Left            =   105
            TabIndex        =   28
            Top             =   195
            Width           =   1155
         End
      End
      Begin VB.CommandButton cmdAgencia 
         Caption         =   "A&gencias..."
         Height          =   345
         Left            =   4440
         TabIndex        =   26
         Top             =   3555
         Width           =   1020
      End
      Begin VB.Frame fraContenedor 
         Caption         =   "Personas Responsables "
         Height          =   2595
         Index           =   2
         Left            =   135
         TabIndex        =   21
         Top             =   660
         Width           =   4170
         Begin VB.TextBox txtCodAud 
            Height          =   285
            Left            =   30
            TabIndex        =   36
            Top             =   2145
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.TextBox txtCodMar 
            Height          =   285
            Left            =   30
            TabIndex        =   35
            Top             =   1605
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.TextBox txtCodGer 
            Height          =   285
            Left            =   30
            TabIndex        =   34
            Top             =   1020
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.TextBox txtCodAdm 
            Height          =   285
            Left            =   30
            TabIndex        =   33
            Top             =   465
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.CommandButton cmdPersona 
            Caption         =   "..."
            Height          =   300
            Index           =   3
            Left            =   3735
            TabIndex        =   32
            Top             =   2145
            Width           =   315
         End
         Begin VB.CommandButton cmdPersona 
            Caption         =   "..."
            Height          =   300
            Index           =   2
            Left            =   3735
            TabIndex        =   31
            Top             =   1590
            Width           =   315
         End
         Begin VB.CommandButton cmdPersona 
            Caption         =   "..."
            Height          =   300
            Index           =   1
            Left            =   3735
            TabIndex        =   30
            Top             =   1005
            Width           =   315
         End
         Begin VB.CommandButton cmdPersona 
            Caption         =   "..."
            Height          =   300
            Index           =   0
            Left            =   3720
            TabIndex        =   29
            Top             =   450
            Width           =   315
         End
         Begin VB.TextBox txtNomGerente 
            Enabled         =   0   'False
            Height          =   330
            Left            =   150
            MaxLength       =   30
            TabIndex        =   5
            Top             =   990
            Width           =   3555
         End
         Begin VB.TextBox txtNomMartillero 
            Enabled         =   0   'False
            Height          =   315
            Left            =   150
            MaxLength       =   30
            TabIndex        =   6
            Top             =   1575
            Width           =   3555
         End
         Begin VB.TextBox txtNomAdministrador 
            Enabled         =   0   'False
            Height          =   315
            Left            =   150
            MaxLength       =   30
            TabIndex        =   4
            Top             =   435
            Width           =   3570
         End
         Begin VB.TextBox txtNomVeedor 
            Enabled         =   0   'False
            Height          =   315
            Left            =   150
            MaxLength       =   30
            TabIndex        =   7
            Top             =   2130
            Width           =   3555
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Gerente de Créditos :"
            Height          =   240
            Index           =   8
            Left            =   135
            TabIndex        =   25
            Top             =   780
            Width           =   1575
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Martillero Público :"
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   24
            Top             =   1365
            Width           =   1545
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Administrador(a) de Agencia :"
            Height          =   240
            Index           =   2
            Left            =   165
            TabIndex        =   23
            Top             =   225
            Width           =   2190
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Veedor de Auditoria :"
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   22
            Top             =   1920
            Width           =   1620
         End
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         TabIndex        =   3
         Top             =   210
         Width           =   1140
      End
      Begin VB.TextBox txtNumRemate 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   765
      End
      Begin VB.TextBox txtAgencia 
         Height          =   345
         Left            =   165
         MaxLength       =   30
         TabIndex        =   8
         Top             =   3570
         Width           =   4065
      End
      Begin MSMask.MaskEdBox txtFecRemate 
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Top             =   210
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
      Begin MSMask.MaskEdBox txtHorRemate 
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
         Left            =   3960
         TabIndex        =   2
         Top             =   195
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
      Begin VB.Label lblEtiqueta 
         Caption         =   "Número :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Hora :"
         Height          =   255
         Index           =   7
         Left            =   3480
         TabIndex        =   16
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Fecha :"
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   15
         Top             =   255
         Width           =   630
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Estado :"
         Height          =   255
         Index           =   6
         Left            =   4680
         TabIndex        =   14
         Top             =   225
         Width           =   645
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Agencia :"
         Height          =   240
         Index           =   4
         Left            =   210
         TabIndex        =   13
         Top             =   3330
         Width           =   615
      End
   End
   Begin RichTextLib.RichTextBox rtfImp 
      Height          =   315
      Left            =   8445
      TabIndex        =   18
      Top             =   4755
      Visible         =   0   'False
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmActaRemaPrend.frx":030A
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
      Height          =   330
      Left            =   8430
      TabIndex        =   19
      Top             =   4455
      Visible         =   0   'False
      Width           =   165
      _ExtentX        =   291
      _ExtentY        =   582
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmActaRemaPrend.frx":038A
   End
   Begin MSComctlLib.ProgressBar prgList 
      Height          =   330
      Left            =   240
      TabIndex        =   20
      Top             =   4665
      Visible         =   0   'False
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmColPRemateActa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* REMATE - ACTA DE REMATE
'Archivo:  frmColPRemateActa.frm
'LAYG   :  05/08/2001.
'Resumen:  Genera el acta de remate de Creditos Pignoraticios

Option Explicit
Dim bModifi As Boolean
Dim pPrevioMax As Double
Dim pLineasMax As Double
Dim pHojaFiMax As Integer
Dim pAgeRemSub As String * 2

Dim RegProcesar As New ADODB.Recordset
Dim RegJoyas As New ADODB.Recordset
Dim sSQL As String
Dim MuestraImpresion As Boolean
Dim vRTFImp As String
Dim vCont As Double
Dim vNomAge As String

Private Sub cmdAgencia_Click()
    frmSelectAgencias.Inicio Me
    frmSelectAgencias.Show 1
End Sub

'Cancela el proceso actual
Private Sub cmdCancelar_Click()
cmdImprimir.Enabled = False
Limpiar
If txtNomAdministrador.Enabled = True Then
    txtNomAdministrador.SetFocus
End If
End Sub

'Verifica que antes de imprimir se halla ingresado la información necesaria para
' realizar la impresión.
Private Sub cmdImprimir_Click()
On Error GoTo ControlError
If Len(Trim(txtNomAdministrador)) = 0 Then
    MsgBox " Ingrese nombre del Administrador ", vbInformation, " Aviso "
    txtNomAdministrador.SetFocus
ElseIf Len(Trim(txtNomGerente)) = 0 Then
    MsgBox " Ingrese nombre del Gerente ", vbInformation, " Aviso "
    txtNomGerente.SetFocus
ElseIf Len(Trim(txtNomMartillero)) = 0 Then
    MsgBox " Ingrese nombre del Martillero ", vbInformation, " Aviso "
    txtNomMartillero.SetFocus
ElseIf Len(txtNomVeedor) = 0 Then
    MsgBox " Ingrese nombre del Veedor ", vbInformation, " Aviso "
    txtNomVeedor.SetFocus
ElseIf Len(txtAgencia) = 0 Then
    MsgBox " Ingrese el Agencia ", vbInformation, " Aviso "
    txtAgencia.SetFocus
Else
    If bModifi Then
        
'        sSql = "UPDATE Remate SET cCodAdm = '" & txtCodAdm.Text & "', cCodGer = '" & txtCodGer.Text & "', " & _
'            " cCodMar = '" & txtCodMar.Text & "', cCodAud = '" & txtCodAud.Text & "', " & _
'            " cCodUsu = '" & gsCodUser & "', dFecModif = '" & Format(gdFecSis & " " & Time, "mm/dd/yyyy hh:mm:ss") & "' " & _
'            " WHERE cNroRemat = '" & txtNumRemate.Text & "'"
'        dbCmact.Execute sSql
        bModifi = False
    End If
    
    ImprimirActRemCab
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdPersona_Click(Index As Integer)
Dim loPers As UPersona
Dim lsPersCod As String
On Error GoTo ControlError

Set loPers = New UPersona
Set loPers = frmBuscaPersona.Inicio

If Not loPers Is Nothing Then
    lsPersCod = loPers.sPersCod
    bModifi = True
    If Index = 0 Then
        txtCodAdm.Text = loPers.sPersCod
        txtNomAdministrador.Text = PstaNombre(loPers.sPersNombre, False)
    ElseIf Index = 1 Then
        txtCodGer.Text = loPers.sPersCod
        txtNomGerente.Text = PstaNombre(loPers.sPersNombre, False)
    ElseIf Index = 2 Then
        txtCodMar.Text = loPers.sPersCod
        txtNomMartillero.Text = PstaNombre(loPers.sPersNombre, False)
    Else
        txtCodAud.Text = loPers.sPersCod
        txtNomVeedor.Text = PstaNombre(loPers.sPersNombre, False)
    End If
End If
Exit Sub
ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Cerrar el formulario
Private Sub cmdSalir_Click()
Unload frmSelectAgencias
Unload Me
End Sub

'Inicializa los campos para ser utilizados
Public Sub Limpiar()
txtCodAdm = ""
txtCodGer = ""
txtCodMar = ""
txtCodAud = ""
txtNomGerente = ""
txtNomMartillero = ""
txtNomAdministrador = ""
txtNomVeedor = ""
txtAgencia = ""
End Sub

'Inicializa el formulario y se busca el último contrato finalizado
Private Sub Form_Load()

    CargaParametros
    bModifi = False
    CargaListaRemates
    
End Sub

Private Sub CargaListaRemates()

Dim loValida As dColPFunciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String

On Error GoTo dError
'Leer información de Remate
    lsSQL = "Select cNroProceso, nRGEstado From ColocPigRecGar Where cTpoProceso ='R' ORDER BY cNroProceso DESC"
    
    Set loValida = New dColPFunciones
        Set lrs = loValida.dObtieneRecordSet(lsSQL)
    Set loValida = Nothing
    If lrs.BOF And lrs.EOF Then
    Else
        Do While Not lrs.EOF
            List1.AddItem lrs!cNroProceso
            lrs.MoveNext
        Loop
        lrs.Close
        List1.ListIndex = 0
    End If
Set lrs = Nothing
Exit Sub

dError:
    Err.Raise Err.Number, "Obtiene Datos Contrato en <<dObtieneDatosContrato>>", Err.Description
End Sub

Private Sub List1_Click()

Dim loValida As dColPFunciones
Dim lrs As ADODB.Recordset
Dim lsSQL As String
    
    'Leer información del Proceso
    lsSQL = " SELECT RG.cNroProceso, RG.dProceso, RG.nRGEstado, " & _
           " Adm = (Select Comit.cPersCod From ColocPigRGComite Comit Where Comit.cTpoProceso ='R' And Comit.cNroProceso ='" & List1.Text & "' And Comit.cRelacion = 'AD' ), " & _
           " Ger = (Select Comit.cPersCod From ColocPigRGComite Comit Where Comit.cTpoProceso ='R' And Comit.cNroProceso ='" & List1.Text & "' And Comit.cRelacion = 'GE' ), " & _
           " Mar = (Select Comit.cPersCod From ColocPigRGComite Comit Where Comit.cTpoProceso ='R' And Comit.cNroProceso ='" & List1.Text & "' And Comit.cRelacion = 'MA' ), " & _
           " Aud = (Select Comit.cPersCod From ColocPigRGComite Comit Where Comit.cTpoProceso ='R' And Comit.cNroProceso ='" & List1.Text & "' And Comit.cRelacion = 'AU' ) " & _
           " From ColocPigRecGar RG WHERE RG.cTpoProceso ='R' And RG.cNroProceso = '" & List1.Text & "' "

    Set loValida = New dColPFunciones
        Set lrs = loValida.dObtieneRecordSet(lsSQL)
    Set loValida = Nothing
    If lrs.BOF And lrs.EOF Then
    Else
        Limpiar
        With lrs
            txtNumRemate = !cNroProceso
            txtFecRemate = Format(!dProceso, "dd/mm/yyyy")
            txtHorRemate = Format(!dProceso, "hh:mm")
            'txtEstado = Switch(!cEstado = "N", "NO INICIADO", !cEstado = "I", "INICIADO", !cEstado = "F", "FINALIZADO")
            If Not IsNull(!Adm) Then
                txtCodAdm = !Adm
                lsSQL = "Select PersNombre From Persona Where cPersCod ='" & !Adm & "' "
                txtNomAdministrador = PstaNombre(fgDevuelveDatosQuery("Select cPersNombre as Campo FROM " & gsCentralPers & "Persona WHERE cPersCod = '" & !Adm & "'"), False)
            End If
            If Not IsNull(!Ger) Then
                txtCodGer = !Ger
                txtNomGerente = PstaNombre(fgDevuelveDatosQuery("Select cPersNombre as Campo FROM " & gsCentralPers & "Persona WHERE cPersCod = '" & !Ger & "'"), False)
            End If
            If Not IsNull(!Mar) Then
                txtCodMar = !Mar
                txtNomMartillero = PstaNombre(fgDevuelveDatosQuery("Select cPersNombre as Campo FROM " & gsCentralPers & "Persona WHERE cPersCod = '" & !Mar & "'"), False)
            End If
            If Not IsNull(!Aud) Then
                txtCodAud = !Aud
                txtNomVeedor = PstaNombre(fgDevuelveDatosQuery("Select cPersNombre as Campo FROM " & gsCentralPers & "Persona WHERE cPersCod = '" & !Aud & "'"), False)
            End If
        End With
    End If
    Set lrs = Nothing
End Sub

'Valida la información ingresada en el campo txtAgencia
Private Sub txtAgencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdImprimir.Enabled = True
    cmdImprimir.SetFocus
End If
End Sub

'Valida txtConNoRem
Private Sub txtConNoRem_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 And Len(Trim(txtConNoRem.Text)) >= 8 Then
'    lstContratos.AddItem txtConNoRem
'    sSql = "INSERT INTO RemSubRemota(cNroRemSub, cTipo, cCodCta, cestado)" & _
'        " VALUES ('" & txtNumRemate & "','R','" & txtConNoRem & "','N')"
'    dbCmact.Execute sSql
'    txtConNoRem.Text = ""
'End If
End Sub

'Valida la información ingresada en el campo txtNomAdministrador
Private Sub txtNomAdministrador_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNomGerente.SetFocus
End If
End Sub
'Valida la información ingresada en el campo txtNomGerente
Private Sub txtNomGerente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNomMartillero.SetFocus
End If
End Sub
'Valida la información ingresada en el campo txtNomMartillero
Private Sub txtNomMartillero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNomVeedor.SetFocus
End If
End Sub
'Valida la información ingresada en el campo txtNomVeedor
Private Sub txtNomVeedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtAgencia.SetFocus
End If
End Sub

' Permite elegir si la presentaciión se realiza en la pantalla del Previo
' o va directo a la impresora
Private Sub ImprimirActRemCab()
'    rtfCartas.FileName = App.Path & cPlantillaActaRema
'    ImprimirActRemInter
'    If MuestraImpresion And optImpresion(0).Value = True Then
'        frmPrevio.Previo rtfImp, " Acta de Remate ", True, 66
'    End If
End Sub

'Proceso de generación de la impresión para el ACTA DE REMATE
Public Sub ImprimirActRemInter()
'    Dim vIndice As Integer  'contador de Item
'    Dim vLineas As Integer
'    Dim vPage As Integer
'    Dim vCuenta As Integer
'    Dim vPrecio As Currency
'    Dim v14 As Currency, v16 As Currency, v18 As Currency, v21 As Currency
'    Dim x As Integer
'    Dim vConexion As ADODB.Connection
'    Dim pPaso As String
'    'Dim pBandActa  As Boolean
'    'pBandActa = False
'    MuestraImpresion = True
'    vRTFImp = ""
'    MousePointer = 11
'    vRTFImp = rtfCartas.Text
'    vRTFImp = Replace(vRTFImp, "<<NROREMATE>>", txtNumRemate, , , vbTextCompare)
'    vRTFImp = Replace(vRTFImp, "<<FECHAC>>", Format(txtFecRemate, "dd/mm/yyyy"), , , vbTextCompare)
'    vRTFImp = Replace(vRTFImp, "<<FECHAL>>", Format(txtFecRemate, "dddd,d mmmm yyyy"), , , vbTextCompare)
'    vRTFImp = Replace(vRTFImp, "<<MARTILLERO>>", txtNomMartillero, , , vbTextCompare)
'    vRTFImp = Replace(vRTFImp, "<<GERENTE>>", txtNomGerente, , , vbTextCompare)
'    vRTFImp = Replace(vRTFImp, "<<ADMINISTRADOR>>", txtNomAdministrador, , , vbTextCompare)
'    vRTFImp = Replace(vRTFImp, "<<VEEDOR>>", txtNomVeedor, , , vbTextCompare)
'    vRTFImp = Replace(vRTFImp, "<<HORA>>", txtHorRemate, , , vbTextCompare)
'    vRTFImp = Replace(vRTFImp, "<<AGENCIA>>", TxtAgencia, , , vbTextCompare)
'    vRTFImp = vRTFImp & gPrnSaltoLinea
'    If Not optImpresion(0).Value = True Then
'        ImpreBegin True, 66
'        Print #ArcSal, ImpreCarEsp(vRTFImp);
'        vRTFImp = ""
'    End If
'    'CONTRATOS  VENDIDOS - AGENCIA ?
'    For x = 1 To frmPigAgencias.List1.ListCount
'        If frmPigAgencias.List1.Selected(x - 1) = True Then
'            pPaso = "0"
'            vNomAge = ""
'            If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
'                vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAge & "'")
'                Set vConexion = dbCmact
'                pPaso = "1"
'            Else
'                If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
'                    vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
'                    Set vConexion = dbCmactN
'                    pPaso = "2"
'                End If
'                'CierraConeccion
'            End If
'            If pPaso <> "0" Then
'                sSql = "SELECT cp.ccodcta, cp.npiezas, cp.mdesclote, " & _
'                    " p.cnompers, p.ctidoci, p.cnudoci, p.ctidotr, p.cnudotr, " & _
'                    " dr.nPreBaseVta, dr.ncomision, dr.nsobrante " & _
'                    " FROM credprenda CP inner join detremate DR on cp.ccodcta = dr.ccodcta " & _
'                    " left join " & gcCentralPers & "Persona P on dr.ccodpers = p.ccodpers " & _
'                    " WHERE dr.cestado IN ('V','P') AND dr.cNroRemat = '" & Trim(txtNumRemate) & "' " & _
'                    " ORDER BY cp.ccodcta "
'                RegProcesar.Open sSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
'                If RegProcesar.BOF And RegProcesar.EOF Then
'                    RegProcesar.Close
'                    Set RegProcesar = Nothing
'                    'pBandActa = True
'                    MsgBox " No existen contratos Vendidos en la " & vNomAge, vbInformation, " Aviso "
'                Else
'                    vRTFImp = vRTFImp & "REMATE DE PRENDA DE LA " & vNomAge & gPrnSaltoLinea
'                    vRTFImp = vRTFImp & String(100, "-") & gPrnSaltoLinea
'                    vRTFImp = vRTFImp & "ITEM    CONTRATO    PZ        DESCRIPCION          14Kl.    16Kl.    18Kl.    21Kl.        PRECIO" & gPrnSaltoLinea
'                    vRTFImp = vRTFImp & String(100, "-") & gPrnSaltoLinea
'                    prgList.Min = 0: vCont = 0
'                    prgList.Max = RegProcesar.RecordCount
'                    If optImpresion(0).Value = True Then
'                        If prgList.Max > pPrevioMax Then
'                            RegProcesar.Close
'                            Set RegProcesar = Nothing
'                            MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
'                                " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
'                            MuestraImpresion = False
'                            MousePointer = 0
'                            Exit Sub
'                        End If
'                    Else
'                        Print #ArcSal, ImpreCarEsp(vRTFImp);
'                        vRTFImp = ""
'                    End If
'                    prgList.Visible = True
'                    vPage = 1
'                    vIndice = 1
'                    vLineas = 4 + 18
'                    vCuenta = 0
'                    v14 = 0: v16 = 0: v18 = 0: v21 = 0
'                    With RegProcesar
'                        Do While Not .EOF
'                            sSql = "SELECT * FROM Joyas WHERE ccodcta = '" & !cCodCta & "'"
'                            RegJoyas.Open sSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
'                            If (RegJoyas.BOF Or RegJoyas.EOF) Then
'                                MsgBox " No existe la descripción de sus Joyas " & !cCodCta, vbExclamation, " Error del Sistema "
'                                RegJoyas.Close
'                                Set RegJoyas = Nothing
'                                RegProcesar.Close
'                                Set RegProcesar = Nothing
'                                Exit Sub
'                            Else
'                                Do While Not RegJoyas.EOF
'                                    If RegJoyas!ckilataje = "14" Then
'                                        v14 = RegJoyas!nPesoOro: v14 = Round(v14, 2)
'                                    End If
'                                    If RegJoyas!ckilataje = "16" Then
'                                        v16 = RegJoyas!nPesoOro: v16 = Round(v16, 2)
'                                    End If
'                                    If RegJoyas!ckilataje = "18" Then
'                                        v18 = RegJoyas!nPesoOro: v18 = Round(v18, 2)
'                                    End If
'                                    If RegJoyas!ckilataje = "21" Then
'                                        v21 = RegJoyas!nPesoOro: v21 = Round(v21, 2)
'                                    End If
'                                    RegJoyas.MoveNext
'                                Loop
'                                RegJoyas.Close
'                                Set RegJoyas = Nothing
'                            End If
'                            vPrecio = !nprebasevta
'
'                            If optImpresion(0).Value = True Then
'                                vRTFImp = vRTFImp & ImpreFormat(vIndice, 5, 0) & ImpreFormat(!cCodCta, 13, 1) & ImpreFormat(!npiezas, 3, 0) & _
'                                    ImpreFormat(UCase(QuiebreTexto(!mDescLote, 1)), 24, 1) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & _
'                                    ImpreFormat(v18, 6) & ImpreFormat(v21, 6) & ImpreFormat(vPrecio, 12) & gPrnSaltoLinea
'                            Else
'                                Print #ArcSal, ImpreFormat(vIndice, 5, 0) & ImpreFormat(!cCodCta, 13, 1) & ImpreFormat(!npiezas, 3, 0) & _
'                                    ImpreFormat(UCase(QuiebreTexto(!mDescLote, 1)), 24, 1) & ImpreFormat(v14, 6) & ImpreFormat(v16, 6) & _
'                                    ImpreFormat(v18, 6) & ImpreFormat(v21, 6) & ImpreFormat(vPrecio, 12)
'                            End If
'                            vLineas = vLineas + 1
'                            Do While vCuenta < 15
'                                vCuenta = vCuenta + 1
'                                If Len(QuiebreTexto(!mDescLote, vCuenta + 1)) > 0 Then
'                                    If optImpresion(0).Value = True Then
'                                        vRTFImp = vRTFImp & Space(22) & UCase(QuiebreTexto(!mDescLote, vCuenta + 1)) & gPrnSaltoLinea
'                                    Else
'                                        Print #ArcSal, Space(22) & ImpreCarEsp(UCase(QuiebreTexto(!mDescLote, vCuenta + 1)))
'                                    End If
'                                    vLineas = vLineas + 1
'                                End If
'                            Loop
'                            If optImpresion(0).Value = True Then
'                                vRTFImp = vRTFImp & Space(23) & PstaNombre("" & !cNomPers, False) & gPrnSaltoLinea
'                                vRTFImp = vRTFImp & Space(23) & TipoDoCi(!cTidoci & "") & " " & Trim(!cNudoci) & "  " & TipoDoTr(!cTidotr & "") & " " & Trim(!cNudoTr) & gPrnSaltoLinea
'                                vRTFImp = vRTFImp & gPrnSaltoLinea
'                            Else
'                                Print #ArcSal, Space(23) & PstaNombre("" & !cNomPers, False)
'                                Print #ArcSal, Space(23) & TipoDoCi(!cTidoci & "") & " " & Trim(!cNudoci) & "  " & TipoDoTr(!cTidotr & "") & " " & Trim(!cNudoTr)
'                                Print #ArcSal, ""
'                            End If
'                            vLineas = vLineas + 3
'
'                            vCuenta = 0
'                            v14 = 0: v16 = 0: v18 = 0: v21 = 0
'                            vIndice = vIndice + 1
'                            If vLineas >= 55 Then
'                                vPage = vPage + 1
'                                If optImpresion(0).Value = True Then
'                                    vRTFImp = vRTFImp & gPrnSaltoPagina
'                                    vRTFImp = vRTFImp & "REMATE DE PRENDA DE LA " & vNomAge & gPrnSaltoLinea
'                                    vRTFImp = vRTFImp & String(100, "-") & gPrnSaltoLinea
'                                    vRTFImp = vRTFImp & "ITEM    CONTRATO    PZ        DESCRIPCION          14Kl.    16Kl.    18Kl.    21Kl.        PRECIO" & gPrnSaltoLinea
'                                    vRTFImp = vRTFImp & String(100, "-") & gPrnSaltoLinea
'                                Else
'                                    ImpreNewPage
'                                    Print #ArcSal, "REMATE DE PRENDA DE LA " & vNomAge
'                                    Print #ArcSal, String(100, "-")
'                                    Print #ArcSal, "ITEM    CONTRATO    PZ        DESCRIPCION          14Kl.    16Kl.    18Kl.    21Kl.        PRECIO"
'                                    Print #ArcSal, String(100, "-")
'                                End If
'                                vLineas = 4
'                            End If
'                            vCont = vCont + 1
'                            prgList.Value = vCont
'                            .MoveNext
'                        Loop
'                    End With
'                    prgList.Value = 0
'                    RegProcesar.Close
'                    Set RegProcesar = Nothing
'                    If optImpresion(0).Value = True Then
'                        vRTFImp = vRTFImp & gPrnSaltoPagina
'                    Else
'                        ImpreNewPage
'                    End If
'                End If
'            End If 'DE PPASO
'            If pPaso = "2" Then CierraConeccion
'        End If
'    Next x
'
'    'CONTRATOS  NO VENDIDOS
'    For x = 1 To frmPigAgencias.List1.ListCount
'        If frmPigAgencias.List1.Selected(x - 1) = True Then
'            pPaso = "0"
'            vNomAge = ""
'            If Right(Trim(gsCodAge), 2) = Mid(frmPigAgencias.List1.List(x - 1), 1, 2) Then
'                vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAge & "'")
'                Set vConexion = dbCmact
'                pPaso = "1"
'            Else
'                If AbreConeccion(Mid(frmPigAgencias.List1.List(x - 1), 1, 2) & "XXXXXXXXXX") Then
'                    vNomAge = FuncGnral("SELECT cNomTab as Campo FROM " & gcCentralCom & "TablaCod WHERE substring(ccodtab,1,2) = '47' AND cvalor = '" & gsCodAgeN & "'")
'                    Set vConexion = dbCmactN
'                    pPaso = "2"
'                End If
'                'CierraConeccion
'            End If
'            If pPaso <> "0" Then
'                sSql = "SELECT ccodcta " & _
'                    " FROM detremate " & _
'                    " WHERE cestado = 'N' AND detremate.cNroRemat = '" & Trim(txtNumRemate) & "' " & _
'                    " ORDER BY ccodcta "
'                RegProcesar.Open sSql, vConexion, adOpenStatic, adLockOptimistic, adCmdText
'                If RegProcesar.BOF And RegProcesar.EOF Then
'                    RegProcesar.Close
'                    Set RegProcesar = Nothing
'                    MsgBox " No existen contratos No Vendidos en la " & vNomAge, vbInformation, " Aviso "
'                Else
'                    'vRTFImp = vRTFImp & gPrnSaltoPagina
'                    vRTFImp = vRTFImp & gPrnSaltoLinea
'                    vRTFImp = vRTFImp & " NO HABIENDO POSTORES QUEDARON PARA EL PROXIMO REMATE O ADJUDICACION LOS " & gPrnSaltoLinea
'                    vRTFImp = vRTFImp & " SIGUIENTES LOTES DE LA " & vNomAge & gPrnSaltoLinea
'                    vRTFImp = vRTFImp & gPrnSaltoLinea
'                    If Not optImpresion(0).Value = True Then
'                        Print #ArcSal, vRTFImp
'                    End If
'                    vIndice = 1: vLineas = 0
'                    prgList.Min = 0: vCont = 0
'                    prgList.Max = RegProcesar.RecordCount
'                    With RegProcesar
'                        Do While Not .EOF
'                            If optImpresion(0).Value = True Then
'                                vRTFImp = vRTFImp & ImpreFormat(vIndice, 5, 0) & ImpreFormat(!cCodCta, 13) & Space(5)
'                            Else
'                                Print #ArcSal, ImpreFormat(vIndice, 5, 0) & ImpreFormat(!cCodCta, 13) & Space(5);
'                            End If
'                            If (vIndice Mod 4) = 0 Then
'                                If optImpresion(0).Value = True Then
'                                    vRTFImp = vRTFImp & gPrnSaltoLinea
'                                Else
'                                    Print #ArcSal, ""
'                                End If
'                            End If
'                            vLineas = vLineas + 1
'                            If vLineas > 55 * 4 And (vIndice Mod 4) = 0 Then
'                                vPage = vPage + 1
'                                If optImpresion(0).Value = True Then
'                                    vRTFImp = vRTFImp & gPrnSaltoPagina
'                                    vRTFImp = vRTFImp & gPrnSaltoLinea & gPrnSaltoLinea
'                                Else
'                                    ImpreNewPage
'                                    Print #ArcSal, "": Print #ArcSal, ""
'                                End If
'                                vLineas = 4
'                            End If
'                            vIndice = vIndice + 1
'                            vCont = vCont + 1
'                            prgList.Value = vCont
'                            .MoveNext
'                        Loop
'                    End With
'                    prgList.Visible = False
'                    prgList.Value = 0
'                    RegProcesar.Close
'                    Set RegProcesar = Nothing
'                    If optImpresion(0).Value = True Then
'                        vRTFImp = vRTFImp & gPrnSaltoPagina
'                    Else
'                        ImpreNewPage
'                    End If
'                End If
'            End If 'DE PPASO
'            If pPaso = "2" Then CierraConeccion
'        End If
'    Next x
'
'    'Impresion de las firmas
'    If vLineas > 40 * 4 Then
'        vPage = vPage + 1
'        If optImpresion(0).Value = True Then
'            vRTFImp = vRTFImp & gPrnSaltoPagina
'            vRTFImp = vRTFImp & gPrnSaltoLinea & gPrnSaltoLinea
'        Else
'            ImpreNewPage
'            Print #ArcSal, "": Print #ArcSal, ""
'        End If
'        vLineas = 4
'    End If
'    If optImpresion(0).Value = True Then
'        vRTFImp = vRTFImp & gPrnSaltoLinea & gPrnSaltoLinea
'        vRTFImp = vRTFImp & " Siendo las " & Format(Time, "hh:mm:ss") & " se dió por concluido el acto del Remate " & gPrnSaltoLinea
'        vRTFImp = vRTFImp & " firmando en señal de conformidad." & gPrnSaltoLinea
'        vRTFImp = vRTFImp & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea
'        vRTFImp = vRTFImp & Space(10) & "__________________________ " & Space(15) & "__________________________ " & gPrnSaltoLinea
'        vRTFImp = vRTFImp & Space(14) & "Gerente de Créditos" & Space(22) & " Martillero Público" & gPrnSaltoLinea
'        vRTFImp = vRTFImp & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea & gPrnSaltoLinea
'        vRTFImp = vRTFImp & Space(10) & "__________________________ " & Space(15) & "__________________________ " & gPrnSaltoLinea
'        vRTFImp = vRTFImp & Space(11) & "Administrador de Agencia" & Space(20) & "Veedor de Auditoria" & gPrnSaltoLinea
'        rtfImp.Text = vRTFImp
'    Else
'        Print #ArcSal, "": Print #ArcSal, ""
'        Print #ArcSal, " Siendo las " & Format(Time, "hh:mm:ss") & " se dió por concluido el acto del Remate "
'        Print #ArcSal, " firmando en señal de conformidad."
'        Print #ArcSal, "": Print #ArcSal, "": Print #ArcSal, "": Print #ArcSal, "": Print #ArcSal, "": Print #ArcSal, ""
'        Print #ArcSal, Space(10) & "__________________________ " & Space(15) & "__________________________ "
'        Print #ArcSal, Space(14) & "Gerente de Créditos" & Space(22) & " Martillero Público"
'        Print #ArcSal, "": Print #ArcSal, "": Print #ArcSal, "": Print #ArcSal, "": Print #ArcSal, "": Print #ArcSal, ""
'        Print #ArcSal, Space(10) & "__________________________ " & Space(15) & "__________________________ "
'        Print #ArcSal, Space(11) & "Administrador de Agencia" & Space(20) & "Veedor de Auditoria"
'        ImpreEnd
'    End If
'    MousePointer = 0
End Sub

Private Sub CargaParametros()
    pPrevioMax = 2000
    pLineasMax = 56
    pHojaFiMax = 66
End Sub

