VERSION 5.00
Begin VB.Form frmCapCajaChica 
   Caption         =   "Desembolso Para Caja Chica"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6285
   Icon            =   "frmCapCajaChica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdDesembolsar 
      Caption         =   "&Desembolsar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtMontoDesembolsar 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Colaborador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txtNombreColaborador 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   5055
      End
      Begin VB.TextBox txtNombreArea 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   4335
      End
      Begin VB.TextBox txtNombreAgencia 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Width           =   4335
      End
      Begin VB.TextBox txtCodigoArea 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   710
      End
      Begin VB.TextBox txtCodigoAgencia 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1320
         Width           =   710
      End
      Begin SICMACT.TxtBuscar txtCodigoColaborador 
         Height          =   345
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   609
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Agencia:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Área:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Label lblMontoDesembolsar 
      Caption         =   "Monto a Desembolsar S/."
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   1815
   End
End
Attribute VB_Name = "frmCapCajaChica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'****************************************************************
'***Nombre      : frmCapCajaChica
'***Descripción : Formulario para Desembolsar las Aperturas y las
'                 Autorizaciones de Reembolso de Caja Chica.
'***Creación    : ELRO el 20120423, según OYP-RFC047-2012
'****************************************************************
Private fnMovNroCH As Long
Private fnProcNroCH As Integer

Private Sub limpiarCampos()
    txtNombreColaborador = ""
    txtCodigoArea = ""
    txtNombreArea = ""
    txtCodigoAgencia = ""
    txtNombreAgencia = ""
    txtMontoDesembolsar = "0.00"
End Sub

Private Sub imprimirBoleta(ByVal psBoleta As String, Optional ByVal sMensaje As String = "Boleta Operación")
Dim nFicSal As Integer
Do
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
    Print #nFicSal, psBoleta
    Close #nFicSal
Loop Until MsgBox("¿Desea Re-Imprimir " & sMensaje & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo
End Sub

Private Function validarDatosColaborador() As Boolean

validarDatosColaborador = False

If txtCodigoColaborador = gsCodPersUser Then
    MsgBox "No puedes realizar esta operación con prorpio usuario ", vbInformation, "Aviso"
    cmdDesembolsar.Enabled = False
    txtCodigoColaborador = ""
    txtCodigoColaborador.SetFocus
    Exit Function
End If

If txtCodigoColaborador = "" Then
    MsgBox "Falta ingresar el código de colaborador", vbInformation, "Aviso"
    cmdDesembolsar.Enabled = False
    txtCodigoColaborador = ""
    txtCodigoColaborador.SetFocus
    Exit Function
End If

validarDatosColaborador = True
End Function

Private Sub cargarDatosColaboradorCHParaDesembolsar()
Dim oNCOMCajaGeneral As COMNCajaGeneral.NCOMCajaGeneral
Set oNCOMCajaGeneral = New COMNCajaGeneral.NCOMCajaGeneral
Dim rsCHParaDesembolsar As ADODB.Recordset
Set rsCHParaDesembolsar = New ADODB.Recordset

If validarDatosColaborador = False Then
    Exit Sub
End If

Call limpiarCampos


Set rsCHParaDesembolsar = oNCOMCajaGeneral.obtenerAprobacionCajaChicaParaDesembolsar(txtCodigoColaborador.psCodigoPersona)

If Not rsCHParaDesembolsar.BOF And Not rsCHParaDesembolsar.EOF Then

    txtNombreColaborador = rsCHParaDesembolsar!cPersNombre
    txtCodigoArea = rsCHParaDesembolsar!cAreaCod
    txtNombreArea = rsCHParaDesembolsar!cAreaDescripcion
    txtCodigoAgencia = rsCHParaDesembolsar!cAgeCod
    txtNombreAgencia = rsCHParaDesembolsar!cAgeDescripcion
    txtMontoDesembolsar = Format(rsCHParaDesembolsar!nImporte, "#,##0.00")
    fnMovNroCH = rsCHParaDesembolsar!nMovNro
    fnProcNroCH = rsCHParaDesembolsar!nProcNro
    
    If cmdDesembolsar.Enabled = False Then
        cmdDesembolsar.Enabled = True
    End If
   
       
Else
    MsgBox "No tiene Caja Chica pendientes a desembolsar.", vbInformation, "Aviso"
    cmdDesembolsar.Enabled = False
    txtCodigoColaborador = ""
    txtCodigoColaborador.SetFocus
    
End If

Set oNCOMCajaGeneral = Nothing
Set rsCHParaDesembolsar = Nothing
End Sub

Private Sub cmdDesembolsar_Click()
    'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    'fin Comprobacion si es RFIII
    
    Dim oNCOMCaptaMovimiento As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set oNCOMCaptaMovimiento = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim oNCOMContFunciones As COMNContabilidad.NCOMContFunciones
    Set oNCOMContFunciones = New COMNContabilidad.NCOMContFunciones
    Dim lsMovNro As String, lsBoleta As String, lsDJ As String
    
    If MsgBox("¿Desea Grabar la Operación?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Set oNCOMCaptaMovimiento = Nothing
        Set oNCOMContFunciones = Nothing
        Exit Sub
    Else
        oNCOMCaptaMovimiento.IniciaImpresora gImpresora
        lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, _
                                                   gsCodAge, _
                                                   gsCodUser)
        
        
           Call oNCOMCaptaMovimiento.registrarDesembolsoParaCajaChica(CCur(txtMontoDesembolsar), _
                                                                      fnMovNroCH, _
                                                                      gsOpeCod, _
                                                                      lsMovNro, _
                                                                      "Desembolso para Caja Chica", _
                                                                      txtNombreColaborador, _
                                                                      gsNomAge, _
                                                                      Moneda.gMonedaNacional, _
                                                                      fnProcNroCH, _
                                                                      lsBoleta, _
                                                                      sLpt, _
                                                                      gbImpTMU)

        
        If Trim(lsBoleta) <> "" Then
            Call imprimirBoleta(lsBoleta)
            'INICIO JHCU ENCUESTA 16-10-2019
            Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
            'FIN
        Else
            MsgBox "No se realizo la operación", vbInformation, "Aviso"
        End If
    End If
    

           
    Set oNCOMCaptaMovimiento = Nothing
    Set oNCOMContFunciones = Nothing
    lsMovNro = ""
    lsBoleta = ""
    cmdsalir_Click
End Sub

Private Sub cmdsalir_Click()
fnMovNroCH = 0
fnProcNroCH = 0
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
fnMovNroCH = 0
fnProcNroCH = 0
End Sub

Private Sub txtCodigoColaborador_EmiteDatos()
    Call cargarDatosColaboradorCHParaDesembolsar
    If cmdDesembolsar.Enabled Then
        cmdDesembolsar.SetFocus
    End If
End Sub
