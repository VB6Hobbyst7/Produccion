VERSION 5.00
Begin VB.Form frmCredFichaSobreEndeudamiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha de Seguimiento Sobreendeudado"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   Icon            =   "frmCredFichaSobreEndeudamiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
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
      Left            =   9480
      TabIndex        =   35
      Top             =   8740
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Evaluación de Sobreendeudado (Admisión)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   25
      Top             =   6300
      Width           =   11775
      Begin VB.TextBox txtResultado 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   27
         Top             =   240
         Width           =   4695
      End
      Begin SICMACT.FlexEdit feAdmEvalSobr 
         Height          =   1815
         Left            =   75
         TabIndex        =   26
         Top             =   560
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   3201
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   1
         EncabezadosNombres=   "N°-Código-Evaluación-Plan de Mitigación-MantCodigo-nEval"
         EncabezadosAnchos=   "400-1300-3000-5500-1200-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "R-L-L-C-C-C"
         FormatosEdit    =   "3-1-0-1-1-1"
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         TipoBusPersona  =   2
      End
      Begin VB.Label Label11 
         Caption         =   "Resultado:"
         Height          =   255
         Left            =   75
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
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
      Left            =   8280
      TabIndex        =   22
      Top             =   8740
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
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
      Left            =   10680
      TabIndex        =   19
      Top             =   8740
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
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
      Left            =   0
      TabIndex        =   4
      Top             =   8740
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ficha de Sobreendeudado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5640
      Left            =   0
      TabIndex        =   3
      Top             =   620
      Width           =   11775
      Begin VB.TextBox txtResultadoNew 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1005
         TabIndex        =   33
         Top             =   3460
         Width           =   4695
      End
      Begin VB.TextBox txtCalfCamc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9720
         TabIndex        =   18
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtSalEndSistFin 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   5060
         TabIndex        =   17
         Top             =   1040
         Width           =   1455
      End
      Begin VB.TextBox txtSalCapConso 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   5060
         TabIndex        =   16
         Top             =   620
         Width           =   1455
      End
      Begin VB.TextBox txtMarClieAval 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9840
         TabIndex        =   15
         Top             =   1040
         Width           =   735
      End
      Begin VB.TextBox txtMonCmac 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   5060
         TabIndex        =   14
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtAna 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   5060
         TabIndex        =   13
         Top             =   200
         Width           =   840
      End
      Begin VB.TextBox txtAge 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9720
         TabIndex        =   12
         Top             =   200
         Width           =   1935
      End
      Begin SICMACT.FlexEdit feNewEvalSobre 
         Height          =   1815
         Left            =   120
         TabIndex        =   20
         Top             =   3760
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   3201
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   1
         EncabezadosNombres=   "N°-Código-Evaluación-Plan de Mitigación-nEval"
         EncabezadosAnchos=   "400-1250-2900-6850-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-3-X"
         ListaControles  =   "0-0-0-3-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "R-L-L-C-C"
         FormatosEdit    =   "3-1-0-0-0"
         CantEntero      =   50
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         MaxLength       =   50
         ColWidth0       =   405
         RowHeight0      =   300
         TipoBusPersona  =   2
      End
      Begin SICMACT.FlexEdit feDiaAtraso 
         Height          =   1215
         Left            =   5520
         TabIndex        =   21
         Top             =   1940
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   2143
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   1
         EncabezadosNombres=   "N°-Credito-Estado-Dias Atraso-CodEstado"
         EncabezadosAnchos=   "400-1900-2500-1200-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "R-L-L-C-C"
         FormatosEdit    =   "3-1-0-1-0"
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         TipoBusPersona  =   2
      End
      Begin SICMACT.FlexEdit feIfi 
         Height          =   1215
         Left            =   120
         TabIndex        =   29
         Top             =   1940
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   2143
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   1
         EncabezadosNombres=   "N°-Nombre-Monto"
         EncabezadosAnchos=   "400-3300-1200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X"
         ListaControles  =   "0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "R-L-R"
         FormatosEdit    =   "3-1-2"
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         TipoBusPersona  =   2
      End
      Begin VB.Label Label12 
         Caption         =   "Resultado:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   3460
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Instituciones Financieras"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1760
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Días de Atraso"
         Height          =   255
         Left            =   5520
         TabIndex        =   24
         Top             =   1760
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Evaluación de Sobreendeudado (Seguimiento)"
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
         Left            =   120
         TabIndex        =   23
         Top             =   3220
         Width           =   4095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Marca de confirmación de cliente aval (""Si"", Si es cliente aval, ""No"" Caso contrario) :"
         Height          =   495
         Left            =   6680
         TabIndex        =   11
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label9 
         Caption         =   "Monto de última provisión CMAC MAYNAS :"
         Height          =   255
         Left            =   1940
         TabIndex        =   10
         Top             =   1500
         Width           =   3135
      End
      Begin VB.Label Label8 
         Caption         =   "Calificación en CMAC :"
         Height          =   255
         Left            =   8040
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Saldo endeudamiento en el sistema financiero según el RCC (Reporte Consolidado de Créditos) sin considerar monto de CMAC MAYNAS :"
         Height          =   615
         Left            =   100
         TabIndex        =   8
         Top             =   960
         Width           =   4935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Saldo de capital consolidado a la fecha de reporte, consolidado al Tipo de Cambio Fijo vigente del momento :"
         Height          =   615
         Left            =   105
         TabIndex        =   7
         Top             =   500
         Width           =   4920
      End
      Begin VB.Label Label2 
         Caption         =   "Analista de Crédito Responsable :"
         Height          =   255
         Left            =   2620
         TabIndex        =   6
         Top             =   220
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Agencia :"
         Height          =   255
         Left            =   8960
         TabIndex        =   5
         Top             =   220
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.TextBox txtCodPersM 
         Enabled         =   0   'False
         Height          =   300
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdExam 
         Caption         =   "Examinar"
         Height          =   375
         Left            =   2400
         TabIndex        =   31
         Top             =   180
         Width           =   855
      End
      Begin VB.TextBox txtNombrePers 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   2
         Top             =   240
         Width           =   8295
      End
      Begin SICMACT.TxtBuscar txtCodPers 
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
         EnabledText     =   0   'False
      End
   End
End
Attribute VB_Name = "frmCredFichaSobreEndeudamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************'
'** Nombre      : frmCredFichaSobreEndeudamiento                       '
'** Descripción : Formulario de Evaluacion de Riesgo Sobreendeudamiento'
'** Referencia  : ERS054-2016                                          '
'** Creación    : JOEP, 20-29-2016 10:30:00 AM                           '
'**********************************************************************'

Option Explicit

Dim fcPersCodBus As String
Dim fsPersCod As String
Dim fsAgeCod As String
Dim fscCtaCodNew As String
Dim fscCtaCodEvalAdm As String
Dim fdFechaCierre As Date
Dim MatIfis As Variant
Dim MatDiasAtrs As Variant
Dim MatFichaNew As Variant
Dim MatFichaAdm As Variant
Dim rsTotalEvalSobreEndNuevo As ADODB.Recordset

Dim fnTpOrig As Integer
Dim dFechaReg As Date
Dim fnTipoReg As Integer
Dim fnTipoPermiso As Integer


Public Function Inicio(ByVal fnTipoRegMant As Integer) As Boolean

    fnTipoPermiso = 0
    fnTipoReg = fnTipoRegMant
    fnTpOrig = fnTipoRegMant
    fcPersCodBus = ""
    fsPersCod = ""
    fscCtaCodNew = ""
    fscCtaCodEvalAdm = ""
    
    Dim PermisoFichaSobrEnd As COMNCredito.NCOMCredito
    Set PermisoFichaSobrEnd = New COMNCredito.NCOMCredito
    
'(2: JefeNegocio, 1: JefeAgencia)
fnTipoPermiso = PermisoFichaSobrEnd.ObtieneTipoPermisoFichaSobreEnd(fnTipoRegMant, gsCodCargo) ' Obtener el tipo de Permiso, Segun Cargo

    If CargaControlesTipoPermiso(fnTipoPermiso) Then
    
        If fnTipoRegMant = 1 Then
            
            txtCodPersM.Visible = False
            cmdExam.Enabled = False
                    
        ElseIf fnTipoRegMant = 2 Then
            
            txtCodPers.Visible = False
            txtCodPersM.Visible = True
            cmdExam.Enabled = True
        
        End If
    Else
        Unload Me
        Exit Function
    End If
Me.Show 1

End Function

Private Function CargaControlesTipoPermiso(ByVal TipoPermiso As Integer) As Boolean
    '1: JefeAgencia o JefeNegocio->
    If TipoPermiso = 3 Then
        'Call HabilitaControles(False)
        CargaControlesTipoPermiso = True
    Else
        MsgBox "No tiene Permisos para este módulo", vbInformation, "Aviso"
        'Call HabilitaControles(False)
        CargaControlesTipoPermiso = False
    End If
End Function

Private Sub LimpiarFlex()
    LimpiaFlex feIfi
    LimpiaFlex feDiaAtraso
    LimpiaFlex feNewEvalSobre
    LimpiaFlex feAdmEvalSobr
End Sub
Private Sub LimpiarCabecera()
'Limpiar Cabecera
    'txtCodPersM.Text = ""
    'txtCodPers.Text = ""
    txtNombrePers.Text = ""
    txtAna.Text = ""
    txtAge.Text = ""
    txtSalCapConso.Text = ""
    txtCalfCamc.Text = ""
    txtSalEndSistFin.Text = ""
    txtMarClieAval.Text = ""
    txtMonCmac.Text = ""
    txtResultado.Text = ""
    txtResultadoNew.Text = ""
End Sub

Private Sub cmdCancelar_Click()
 Call LimpiarCabecera
 Call LimpiarFlex
 txtCodPersM.Text = ""
 txtCodPers.Text = ""

 
 If cmdImprimir.Enabled = True Then
    cmdGuardar.Enabled = True
 End If
 
    If fnTpOrig = 1 Then
        fnTipoReg = 1
    Else
        fnTipoReg = 2
    End If

 cmdImprimir.Enabled = False

End Sub

Private Sub cmdExam_Click()
    fcPersCodBus = frmCredFichaSobreLista.Inicio()
    If Len(fcPersCodBus) > 0 Then
    txtCodPersM.Text = fcPersCodBus
        Call LimpiarFlex
        Call ObtDatReg
    End If
End Sub

Private Sub Cmdguardar_Click()

Dim oCredFicha As COMNCredito.NCOMCredito
Dim oValidaReg As COMDCredito.DCOMCredito
Dim GrabarFicha As Boolean
Dim rsValidaRegistro As ADODB.Recordset
Dim dfecReg As Date
Dim i As Integer

Set oCredFicha = New COMNCredito.NCOMCredito
Set oValidaReg = New COMDCredito.DCOMCredito
    
If Valida Then
    
    If fnTipoReg = 2 Then
    Set rsValidaRegistro = oValidaReg.ValidaReg(fscCtaCodNew, txtCodPersM, dFechaReg)
        
        txtCodPers.Text = txtCodPersM
        dfecReg = Mid(dFechaReg, 1, 10)
        If rsValidaRegistro!NVECESREG = 2 Then
            MsgBox "Ud. ya no puede Guardar el Registro. Ya supero el maximo de Registro", vbInformation, "Aviso"
            Exit Sub
        End If
    Else
    dfecReg = gdFecSis
    End If
    
If (feIfi.TextMatrix(1, 1)) <> "" Then
    ReDim MatIfis(feIfi.rows - 1, 6)
                    For i = 1 To feIfi.rows - 1
                        MatIfis(i, 0) = feIfi.TextMatrix(i, 1)
                        MatIfis(i, 1) = feIfi.TextMatrix(i, 2)
                    Next i
End If

If (feDiaAtraso.TextMatrix(1, 1)) <> "" Then
    ReDim MatDiasAtrs(feDiaAtraso.rows - 1, 6)
                    For i = 1 To feDiaAtraso.rows - 1
                        MatDiasAtrs(i, 0) = feDiaAtraso.TextMatrix(i, 1)
                        MatDiasAtrs(i, 1) = feDiaAtraso.TextMatrix(i, 4)
                        MatDiasAtrs(i, 2) = feDiaAtraso.TextMatrix(i, 3)
                    Next i
End If

If (feNewEvalSobre.TextMatrix(1, 1)) <> "" Then
    ReDim MatFichaNew(feNewEvalSobre.rows - 1, 6)
                    For i = 1 To feNewEvalSobre.rows - 1
                        MatFichaNew(i, 0) = Right(feNewEvalSobre.TextMatrix(i, 1), 1)
                        MatFichaNew(i, 1) = feNewEvalSobre.TextMatrix(i, 4)
                        MatFichaNew(i, 2) = feNewEvalSobre.TextMatrix(i, 3)
                    Next i
Else
                ReDim MatFichaNew(0)
End If

If (feAdmEvalSobr.TextMatrix(1, 1)) <> "" Then
    ReDim MatFichaAdm(feAdmEvalSobr.rows - 1, 6)
                    For i = 1 To feAdmEvalSobr.rows - 1
                        MatFichaAdm(i, 0) = Right(feAdmEvalSobr.TextMatrix(i, 1), 1)
                        MatFichaAdm(i, 1) = feAdmEvalSobr.TextMatrix(i, 5)
                        MatFichaAdm(i, 2) = feAdmEvalSobr.TextMatrix(i, 3)
                        MatFichaAdm(i, 3) = feAdmEvalSobr.TextMatrix(i, 4)
                    Next i
Else
                ReDim MatFichaAdm(0)
End If
            
       
    If MsgBox("Los Datos serán Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
          
        GrabarFicha = oCredFicha.GrabarFichaSobreEnd(fnTipoReg, fscCtaCodNew, txtCodPers.Text, txtAna.Text, fsAgeCod, txtSalCapConso.Text, _
                                                    IIf(Left(txtCalfCamc.Text, 1) = "", 6, Left(txtCalfCamc.Text, 1)), txtSalEndSistFin.Text, IIf((txtMarClieAval.Text) = "Si", 1, 0), txtMonCmac.Text, _
                                                    fscCtaCodNew, fscCtaCodEvalAdm, MatIfis, MatDiasAtrs, MatFichaNew, MatFichaAdm, dfecReg, fdFechaCierre)
   
        If GrabarFicha Then
           
           If fnTipoReg = 1 Then
               MsgBox "Los Datos se Grabaron Correctamente", vbInformation, "Aviso"
               
               cmdGuardar.Enabled = False
               cmdImprimir.Enabled = True
           Else
               MsgBox "Los Datos se Actualizaron Correctamente", vbInformation, "Aviso"
               
               cmdGuardar.Enabled = False
               cmdImprimir.Enabled = True
           End If
           
        Else
           MsgBox "Hubo error al grabar la informacion", vbError, "Error"
           
        End If
End If
End Sub

Private Sub feNewEvalSobre_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim rsDatosCodigos As ADODB.Recordset
            
Set oCredito = New COMDCredito.DCOMCredito
If feNewEvalSobre.TextMatrix(feNewEvalSobre.row, 4) <> "" Then
    If feNewEvalSobre.TextMatrix(feNewEvalSobre.row, 4) <> 0 Then
        If feNewEvalSobre.Col = 3 Then
            Set rsDatosCodigos = oCredito.MostrarDatosPlanMitig(6)
                feNewEvalSobre.CargaCombo rsDatosCodigos
        End If
    End If
End If

Set oCredito = Nothing
RSClose rsDatosCodigos
End Sub

Private Sub feNewEvalSobre_DblClick()

If feNewEvalSobre.TextMatrix(feNewEvalSobre.row, 4) <> "" Then
    If feNewEvalSobre.TextMatrix(feNewEvalSobre.row, 4) = 0 Then
        feNewEvalSobre.ListaControles = "X-X-X-X-X"
    Else
        feNewEvalSobre.ListaControles = "X-X-X-3-X"
    End If
End If

'If feNewEvalSobre.Col = 3 Then
'    feNewEvalSobre.MaxLength = "200"
'End If

End Sub
'
Private Sub feNewEvalSobre_EnterCell()

If feNewEvalSobre.TextMatrix(feNewEvalSobre.row, 4) <> "" Then
    If feNewEvalSobre.TextMatrix(feNewEvalSobre.row, 4) = 0 Then
        feNewEvalSobre.ListaControles = "X-X-X-X-X"
    Else
        feNewEvalSobre.ListaControles = "X-X-X-3-X"
    End If
End If

'If feNewEvalSobre.Col = 3 Then
 '   feNewEvalSobre.MaxLength = "200"
 ' End If
End Sub
'
Private Sub feNewEvalSobre_KeyPress(KeyAscii As Integer)
  If feNewEvalSobre.Col = 3 Then
    feNewEvalSobre.MaxLength = "200"
  End If
End Sub
'
Private Sub feNewEvalSobre_OnCellChange(pnRow As Long, pnCol As Long)
 If feNewEvalSobre.Col = 3 Then
    feNewEvalSobre.MaxLength = "200"
  End If
End Sub

Private Sub feNewEvalSobre_RowColChange() 'PresionarEnter
    If feNewEvalSobre.Col = 3 Then
        feNewEvalSobre.AvanceCeldas = Vertical
        If feNewEvalSobre.TextMatrix(feNewEvalSobre.row, 4) <> "" Then 'LUCV20180226
            If feNewEvalSobre.TextMatrix(feNewEvalSobre.row, 4) = 0 Then
                feNewEvalSobre.ListaControles = "X-X-X-X-X"
            Else
                feNewEvalSobre.ListaControles = "X-X-X-3-X"
            End If
        End If
    Else
        feNewEvalSobre.AvanceCeldas = Horizontal
    End If
End Sub


Private Sub Form_Load()
    cmdImprimir.Enabled = False
    CentraForm Me
End Sub

Private Sub ObtDatReg()
Dim oObtDatos As COMDCredito.DCOMCredito
Dim rsValReg As ADODB.Recordset

Set oObtDatos = New COMDCredito.DCOMCredito
'Verifica si Trae datos
        
        If fnTipoReg = 1 Then
            fcPersCodBus = fsPersCod
            Set rsValReg = oObtDatos.ValReg(Trim(txtCodPers.Text))
        ElseIf fnTipoReg = 2 Then
            Set rsValReg = oObtDatos.ValReg(Trim(txtCodPersM.Text))
        End If
        
        If rsValReg.RecordCount = 0 Then
            fnTipoReg = 1
        Else
            dFechaReg = rsValReg!dFechaReg
                If fnTipoReg = 2 Then
                    Call Mantenimiento(Trim(txtCodPersM.Text))
                ElseIf fnTipoReg = 1 Then
                    Call Mantenimiento(Trim(txtCodPers.Text))
                End If
            Exit Sub
        End If
End Sub



Private Sub txtCodPers_EmiteDatos()

    Dim i As Integer
    Dim j As Integer
    
    Dim oDPersonaS As COMDPersona.DCOMPersonas
    Dim sPersCod As String
    Dim lnFila As Integer
    Dim oRs As ADODB.Recordset
    Dim nContador0 As Integer
    Dim nContador1 As Integer
    Dim nContador2 As Integer
        nContador0 = 0
        nContador1 = 0
        nContador2 = 0
    
    Dim oObtenerSobreEnd As COMDCredito.DCOMCredito
    Dim rsObtenerObtenerCabSobreEnd As ADODB.Recordset
    Dim rsObtenerIfis As ADODB.Recordset
    Dim rsObtenerDiasAtraso As ADODB.Recordset
    Dim rsObtenerEvalAdm As ADODB.Recordset
    Dim rsObtenerEvalSobreEndAdm As ADODB.Recordset
    Dim rsObtenerEvalSobreEndNuevo As ADODB.Recordset
           
    Dim nValorRep As Integer
    Dim cValRep As String
    Dim cCodFinal As String
    Call LimpiarCabecera
    Call LimpiarFlex
    
    
    If Trim(txtCodPers.Text) = "" Then Exit Sub
              
       sPersCod = Trim(txtCodPers.Text)
       
       fsPersCod = sPersCod
       
       Set oDPersonaS = New COMDPersona.DCOMPersonas
       Set oRs = oDPersonaS.BuscaCliente(sPersCod, BusquedaCodigo)
       Set oDPersonaS = Nothing
                                   
       If Not oRs.EOF And Not oRs.BOF Then
        txtNombrePers.Text = oRs!cPersNombre
       End If
               
       Call ObtDatReg
       If fnTipoReg = 2 Then
        Exit Sub
       End If
       Set oObtenerSobreEnd = New COMDCredito.DCOMCredito
       
       Set rsObtenerObtenerCabSobreEnd = oObtenerSobreEnd.ObtenerCabeceraSobreEnd(fsPersCod)
       Set rsObtenerIfis = oObtenerSobreEnd.ObtenerIfis(fsPersCod)
       Set rsObtenerDiasAtraso = oObtenerSobreEnd.ObtenerDiasAtraso(fsPersCod)
       
       Set rsObtenerEvalSobreEndAdm = oObtenerSobreEnd.ObtenerEvalSobreEndAdm(fsPersCod)
       Set rsObtenerEvalSobreEndNuevo = oObtenerSobreEnd.ObtenerEvalSobreEndNuevo(fsPersCod)
       
       'Obtener Cabecera
       If Not (rsObtenerObtenerCabSobreEnd.EOF And rsObtenerObtenerCabSobreEnd.BOF) Then
                        
         txtAna.Text = rsObtenerObtenerCabSobreEnd!cCodAnalista
         txtAge.Text = rsObtenerObtenerCabSobreEnd!cAgeDescripcion
         txtSalCapConso.Text = Format(rsObtenerObtenerCabSobreEnd!nSaldoCap, "#,##0.00")
         txtCalfCamc.Text = rsObtenerObtenerCabSobreEnd!cCalifSistF
         txtSalEndSistFin.Text = Format(rsObtenerObtenerCabSobreEnd!nSaldEndSf, "#,##0.00")
         txtMarClieAval.Text = IIf((rsObtenerObtenerCabSobreEnd!nPersGarant) = "1", "Si", "No")
         txtMonCmac.Text = Format(rsObtenerObtenerCabSobreEnd!nProvision, "#,##0.00")
         
         fscCtaCodNew = rsObtenerObtenerCabSobreEnd!cCtaCod
         fsAgeCod = rsObtenerObtenerCabSobreEnd!cAgeCod
         fdFechaCierre = rsObtenerObtenerCabSobreEnd!dFecha
         
       Else
        
         MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
         Exit Sub
       End If
       
       'Obtener IFIS
       Call LimpiaFlex(feIfi)
         If Not (rsObtenerIfis.EOF And rsObtenerIfis.BOF) Then
            feIfi.Clear
            feIfi.FormaCabecera
            Call LimpiaFlex(feIfi)
                Do While Not rsObtenerIfis.EOF
                    feIfi.AdicionaFila
                    lnFila = feIfi.row
                    feIfi.TextMatrix(lnFila, 1) = rsObtenerIfis!Nombre
                    feIfi.TextMatrix(lnFila, 2) = Format(rsObtenerIfis!Saldo, "#,##0.00")
                    rsObtenerIfis.MoveNext
                Loop
            rsObtenerIfis.Close
            Set rsObtenerIfis = Nothing
         End If
         
         'Obtener Dias de Atraso
         Call LimpiaFlex(feDiaAtraso)
         If Not (rsObtenerDiasAtraso.EOF And rsObtenerDiasAtraso.BOF) Then
            feDiaAtraso.Clear
            feDiaAtraso.FormaCabecera
            Call LimpiaFlex(feDiaAtraso)
                Do While Not rsObtenerDiasAtraso.EOF
                    feDiaAtraso.AdicionaFila
                    lnFila = feDiaAtraso.row
                    feDiaAtraso.TextMatrix(lnFila, 1) = rsObtenerDiasAtraso!cCtaCod
                    feDiaAtraso.TextMatrix(lnFila, 2) = rsObtenerDiasAtraso!cConsDescripcion
                    feDiaAtraso.TextMatrix(lnFila, 3) = rsObtenerDiasAtraso!nDiasAtraso
                    feDiaAtraso.TextMatrix(lnFila, 4) = rsObtenerDiasAtraso!nPrdEstado
                    rsObtenerDiasAtraso.MoveNext
                Loop
            rsObtenerDiasAtraso.Close
            Set rsObtenerDiasAtraso = Nothing
         End If
         
         'Obtener Evaluacion de SobreEndeudamiento Nueva
         Call LimpiaFlex(feNewEvalSobre)
            If Not (rsObtenerEvalSobreEndNuevo.EOF And rsObtenerEvalSobreEndNuevo.BOF) Then
                nContador0 = 0
                nContador1 = 0
                nContador2 = 0
               feNewEvalSobre.Clear
               feNewEvalSobre.FormaCabecera
               Call LimpiaFlex(feNewEvalSobre)
                   Do While Not rsObtenerEvalSobreEndNuevo.EOF
                       feNewEvalSobre.AdicionaFila
                       lnFila = feNewEvalSobre.row
                       feNewEvalSobre.TextMatrix(lnFila, 1) = rsObtenerEvalSobreEndNuevo!cCodigo
                       feNewEvalSobre.TextMatrix(lnFila, 2) = rsObtenerEvalSobreEndNuevo!cResultado
                       feNewEvalSobre.TextMatrix(lnFila, 4) = rsObtenerEvalSobreEndNuevo!nResultado
                       
                        'If rsObtenerEvalSobreEndNuevo!nResultado = 0 Then
                            'nContador0 = nContador0 + 1
                        'ElseIf rsObtenerEvalSobreEndNuevo!nResultado = 1 Then
                            'nContador1 = nContador1 + 1
                        'ElseIf rsObtenerEvalSobreEndNuevo!nResultado = 2 Then
                            'nContador2 = nContador2 + 1
                            cCodFinal = rsObtenerEvalSobreEndNuevo!nCodFinal
                    'End If
                       
                       rsObtenerEvalSobreEndNuevo.MoveNext
                   Loop
                   
                   'If nContador0 = 5 Then
                    'txtResultadoNew.Text = "NO APLICA"
                'ElseIf nContador2 >= 2 Then
                    'txtResultadoNew.Text = "SOBREENDEUDADO"
                'Else
                    'txtResultadoNew.Text = "POTENCIAL SOBREENDEUDADO"
                'End If
                   
                   txtResultadoNew.Text = cCodFinal
                   
               rsObtenerEvalSobreEndNuevo.Close
               Set rsObtenerEvalSobreEndNuevo = Nothing
            End If
cCodFinal = ""

        'Obtener Evaluacion de SobreEndeudamiento Admision
         Call LimpiaFlex(feAdmEvalSobr)
         txtResultado.Text = ""
         If Not (rsObtenerEvalSobreEndAdm.EOF And rsObtenerEvalSobreEndAdm.BOF) Then
            nContador0 = 0
            nContador1 = 0
            nContador2 = 0
            feAdmEvalSobr.Clear
            feAdmEvalSobr.FormaCabecera
            Call LimpiaFlex(feAdmEvalSobr)
                Do While Not rsObtenerEvalSobreEndAdm.EOF
                    feAdmEvalSobr.AdicionaFila
                    lnFila = feAdmEvalSobr.row
                    feAdmEvalSobr.TextMatrix(lnFila, 1) = rsObtenerEvalSobreEndAdm!cCodigo
                    feAdmEvalSobr.TextMatrix(lnFila, 2) = rsObtenerEvalSobreEndAdm!cResultado
                    feAdmEvalSobr.TextMatrix(lnFila, 3) = rsObtenerEvalSobreEndAdm!cPlanmitigacion
                    feAdmEvalSobr.TextMatrix(lnFila, 5) = rsObtenerEvalSobreEndAdm!nResultado
                    
                    fscCtaCodEvalAdm = rsObtenerEvalSobreEndAdm!cCtaCod
                    
                    cCodFinal = rsObtenerEvalSobreEndAdm!nCodFinal
                    
                    'If rsObtenerEvalSobreEndAdm!nResultado = 0 Then
                        'nContador0 = nContador0 + 1
                    'ElseIf rsObtenerEvalSobreEndAdm!nResultado = 1 Then
                        'nContador1 = nContador1 + 1
                    'ElseIf rsObtenerEvalSobreEndAdm!nResultado = 2 Then
                        'nContador2 = nContador2 + 1
                    'End If
                
                    rsObtenerEvalSobreEndAdm.MoveNext
                Loop
                                                
                'If nContador0 = 5 Then
                    'txtResultado.Text = "No Aplica"
                'ElseIf nContador2 >= 2 Then
                    txtResultado.Text = cCodFinal
                'Else
                    'txtResultado.Text = "POTENCIAL SOBREENDEUDADO"
                'End If
            
            rsObtenerEvalSobreEndAdm.Close
            Set rsObtenerEvalSobreEndAdm = Nothing
            
    'Comparacion de Codigos si se repite
        For i = 1 To feNewEvalSobre.rows - 1
            If feNewEvalSobre.TextMatrix(i, 1) <> "" Then
                If Right(feNewEvalSobre.TextMatrix(i, 1), 1) = "4" Then
                    cValRep = feNewEvalSobre.TextMatrix(i, 1)
                    nValorRep = feNewEvalSobre.TextMatrix(i, 4)
                End If
            End If
        Next i
        
        For i = 1 To feAdmEvalSobr.rows - 1
            If feNewEvalSobre.TextMatrix(1, 1) <> "" Then
                If feAdmEvalSobr.TextMatrix(i, 1) <> "" Then
                    If Right(feAdmEvalSobr.TextMatrix(i, 1), 1) = "4" Then
                        If feAdmEvalSobr.TextMatrix(i, 1) = cValRep And feAdmEvalSobr.TextMatrix(i, 5) = nValorRep Then
                            feAdmEvalSobr.TextMatrix(i, 4) = "Si"
                        Else
                            feAdmEvalSobr.TextMatrix(i, 4) = "No"
                        End If
                    Else
                        feAdmEvalSobr.TextMatrix(i, 4) = "No"
                    End If
                End If
            End If
        Next i
    End If
        
    RSClose oRs
    'RSClose rsObtenerSobreEnd
End Sub

Private Sub cmdImprimir_Click()
    Dim oAgencia As COMDCredito.DCOMCredito
    Dim oEval As COMDCredito.DCOMCredito
    Dim rsAgencia As ADODB.Recordset
    Dim rsCabFichaSobEnd As ADODB.Recordset
    Dim rsFichaSobEndIfis As ADODB.Recordset
    Dim rsFichaSobEndDiasAtr As ADODB.Recordset
    Dim rsFichaSobEndEvalNew As ADODB.Recordset
    Dim rsFichaSobEndEvalAdm As ADODB.Recordset
    
    Dim nContS As Integer
    Dim nCodFinal As String
    Dim A As Integer
    Dim i As Integer
    Dim oDoc  As cPDF
    
    Dim K As Integer
                   
    Set oDoc = New cPDF
    Set oAgencia = New COMDCredito.DCOMCredito
    Set oEval = New COMDCredito.DCOMCredito
    Set rsAgencia = oAgencia.RecuperaAgencia(gsCodAge)
    Set rsCabFichaSobEnd = oEval.ObtenerCabFichaSobEnd(IIf(fnTpOrig = 1, Trim(txtCodPers.Text), Trim(txtCodPersM.Text)), Format(gdFecSis, "yyyyMMdd"))
    Set rsFichaSobEndDiasAtr = oEval.ObtenerFichaSobEndDiasAtrs(IIf(fnTpOrig = 1, Trim(txtCodPers.Text), Trim(txtCodPersM.Text)), Format(gdFecSis, "yyyyMMdd"))
    Set rsFichaSobEndEvalNew = oEval.PDFObtenerFichaSobEndNew(IIf(fnTpOrig = 1, Trim(txtCodPers.Text), Trim(txtCodPersM.Text)), Format(gdFecSis, "yyyyMMdd"))
    Set rsFichaSobEndEvalAdm = oEval.PDFObtenerFichaSobEndAdm(IIf(fnTpOrig = 1, Trim(txtCodPers.Text), Trim(txtCodPersM.Text)), Format(gdFecSis, "yyyyMMdd"))
    
'Creación del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Ficha de Evaluacion de Riesgo Nº " & fsPersCod
    oDoc.Title = "Evaluacion de Riesgo Nº " & IIf(fnTpOrig = 1, Trim(txtCodPers.Text), Trim(txtCodPersM.Text))
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\FichaEvalRiesgoSobreEnd" & fsPersCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
'Contenido
    
    oDoc.Fonts.Add "F1", "Courier New", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Courier New", TrueType, Bold, WinAnsiEncoding
    
    oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"
    
If Not (rsCabFichaSobEnd.BOF Or rsCabFichaSobEnd.EOF) Then

    'Tamaño de hoja A4
    oDoc.NewPage A4_Vertical

    '---------- cabecera ---------------
    oDoc.WImage 60, 45, 60, 113, "Logo"
    oDoc.WTextBox 50, 55, 20, 100, rsAgencia!cAgeDescripcion, "F2", 8, hCenter

    oDoc.WTextBox 40, 30, 45, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F2", 7.5, hRight

    oDoc.WTextBox 100, 100, 10, 400, "FICHA DE EVALUACION DE RIESGO SOBREENDEUDAMIENTO", "F2", 12, hCenter
    oDoc.WTextBox 120, 50, 35, 490, "USUARIO: " & Trim(gsCodUser), "F2", 7.5, hLeft
    oDoc.WTextBox 130, 50, 35, 490, "FECHA DE CIERRE: " & Format(rsCabFichaSobEnd!dFechaCierre, "dd/mm/yyyy") & "", "F2", 7.5, hLeft
    
    A = 0
    A = 40
    
    'oDoc.WTextBox 110 + A, 50, 300, 500, "", "F1", 12, hCenter, vTop, vbBlack, 0.6, vbBlack
'Cabecera
    oDoc.WTextBox 110 + A, 50, 20, 270, "", "F1", 7.5, hCenter, vMiddle, vbWhite, 1, vbBlack
    oDoc.WTextBox 110 + A, 55, 10, 200, "DESCRIPCION", "F2", 10, hCenter
    oDoc.WTextBox 110 + A, 320, 20, 230, "", "F1", 7.5, hLeft, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 110 + A, 320, 10, 230, "VALOR", "F2", 10, hCenter
'Cabecera
    
'Contenido
    oDoc.WTextBox 130 + A, 50, 20, 270, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 135 + A, 55, 10, 200, "AGENCIA: ", "F2", 7.5, hLeft
    oDoc.WTextBox 130 + A, 320, 20, 230, "", "F1", 7.5, hLeft, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 135 + A, 330, 10, 230, rsCabFichaSobEnd!cAgeDescripcion, "F2", 7.5, hLeft
    
    oDoc.WTextBox 150 + A, 50, 20, 270, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 155 + A, 55, 10, 200, "ANALISTA DE CREDITO RESPONSABLE:", "F2", 7.5, hLeft
    oDoc.WTextBox 150 + A, 320, 20, 230, "", "F1", 7.5, hLeft, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 155 + A, 330, 10, 230, rsCabFichaSobEnd!cAnaCod, "F2", 7.5, hLeft
    
    oDoc.WTextBox 170 + A, 50, 20, 270, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 175 + A, 55, 10, 200, "CODIGO DE CLIENTE: ", "F2", 7.5, hLeft
    oDoc.WTextBox 170 + A, 320, 20, 230, "", "F1", 7.5, hLeft, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 175 + A, 330, 10, 230, rsCabFichaSobEnd!cPersCod, "F2", 7.5, hLeft
    
    oDoc.WTextBox 190 + A, 50, 20, 270, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 195 + A, 55, 10, 200, "NOMBRE Y APELLIDO DEL CLIENTE: ", "F2", 7.5, hLeft
    oDoc.WTextBox 190 + A, 320, 20, 230, "", "F1", 7.5, hLeft, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 195 + A, 330, 10, 230, rsCabFichaSobEnd!cPersNombre, "F2", 7.5, hLeft
    
    oDoc.WTextBox 210 + A, 50, 35, 270, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 215 + A, 55, 10, 250, "SALDO DE CAPITAL CONSOLIDADO A LA FECHA DE REPORTE, CONSOLIDADO AL TIPO DE CAMBIO FIJO VIGENTE DEL MOMENTO: ", "F2", 7.5, hLeft
    oDoc.WTextBox 210 + A, 320, 35, 230, "", "F1", 7.5, hLeft, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 215 + A, 330, 10, 230, Format(rsCabFichaSobEnd!nSaldoCap, "#,##0.00"), "F2", 7.5, hLeft
    
    oDoc.WTextBox 245 + A, 50, 35, 270, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 250 + A, 55, 10, 250, "SALDO ENDEUDAMIENTO EN EL SISTEMA FINANCIERO SEGUN EL ULTIMO RCC (REPORTE CONSOLIDADO DE CREDITO) SIN CONSIDERAR MONTO DE CMAC MAYNAS: ", "F2", 7.5, hLeft
    oDoc.WTextBox 245 + A, 320, 35, 230, "", "F1", 7.5, hLeft, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 250 + A, 330, 10, 230, Format(rsCabFichaSobEnd!nSaldoEnd, "#,##0.00"), "F2", 7.5, hLeft
    
    oDoc.WTextBox 280 + A, 50, 20, 500, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 285 + A, 55, 10, 200, "DIAS ATRASO A LA FECHA DE REPORTE: ", "F2", 7.5, hLeft
    
    oDoc.WTextBox 300 + A, 50, 20, 150, "", "F1", 7.5, hLeft, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 303 + A, 60, 20, 100, "Creditos", "F2", 7.5, hLeft
    oDoc.WTextBox 300 + A, 200, 20, 120, "", "F1", 7.5, hLeft, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 303 + A, 210, 20, 100, "Estado", "F2", 7.5, hLeft
    oDoc.WTextBox 300 + A, 320, 20, 230, "", "F1", 7.5, hLeft, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 303 + A, 330, 20, 230, "Dias de Atraso", "F2", 7.5, hLeft
    
    For i = 1 To (rsFichaSobEndDiasAtr.RecordCount)
    oDoc.WTextBox 320 + A, 50, 20, 150, "", "F1", 7.5, hLeft, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 325 + A, 55, 10, 230, rsFichaSobEndDiasAtr!cCtaCod, "F2", 7.5, hLeft
    oDoc.WTextBox 320 + A, 200, 20, 120, "", "F1", 7.5, hLeft, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 325 + A, 205, 10, 230, rsFichaSobEndDiasAtr!cConsDescripcion, "F2", 7.5, hLeft
    oDoc.WTextBox 320 + A, 320, 20, 230, "", "F1", 7.5, hLeft, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 325 + A, 330, 10, 230, rsFichaSobEndDiasAtr!nDiasAtraso, "F2", 7.5, hLeft
    A = A + 20
    rsFichaSobEndDiasAtr.MoveNext
    Next i

    oDoc.WTextBox 320 + A, 50, 20, 270, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 325 + A, 55, 10, 200, "CALIFICACION EN CMAC MAYNAS: ", "F2", 7.5, hLeft
    oDoc.WTextBox 320 + A, 320, 20, 230, "", "F1", 7.5, hLeft, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 325 + A, 330, 10, 230, rsCabFichaSobEnd!nCalifSf, "F2", 7.5, hLeft
    
    oDoc.WTextBox 340 + A, 50, 20, 270, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 345 + A, 55, 10, 500, "MONTO DE ULTIMA PROVISION CMAC MAYNAS: ", "F2", 7.5, hLeft
    oDoc.WTextBox 340 + A, 320, 20, 230, "", "F1", 7.5, hLeft, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 345 + A, 330, 10, 230, Format(rsCabFichaSobEnd!nMontoProv, "#,##0.00"), "F2", 7.5, hLeft
    
    oDoc.WTextBox 360 + A, 50, 20, 270, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 361 + A, 55, 10, 250, "MARCA DE CONFIRMACION DE CLIENTE AVAL ('Si', SI ES CLIENTE AVAL, 'No' CASO CONTRARIO): ", "F2", 7.5, hLeft
    oDoc.WTextBox 360 + A, 320, 20, 230, "", "F1", 7.5, hLeft, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 365 + A, 330, 10, 230, IIf(rsCabFichaSobEnd!nMarcaGart = 0, "Si", "No"), "F2", 7.5, hLeft
          
    oDoc.WTextBox 380 + A, 50, 20, 500, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 385 + A, 65, 10, 200, "Evaluacion (Seguimiento) ", "F2", 7.5, hLeft
          
     oDoc.WTextBox 400 + A, 50, 15, 80, "", "F1", 3, hCenter, vTop, vbBlack, 1, vbBlack
     oDoc.WTextBox 405 + A, 65, 10, 200, "CODIGOS ", "F2", 7.5, hLeft
     oDoc.WTextBox 400 + A, 130, 15, 190, "", "F1", 3, hCenter, vTop, vbBlack, 1, vbBlack
     oDoc.WTextBox 405 + A, 190, 10, 200, "EVALUACION ", "F2", 7.5, hLeft
     oDoc.WTextBox 400 + A, 320, 15, 230, "", "F1", 3, hCenter, vTop, vbBlack, 1, vbBlack
     oDoc.WTextBox 405 + A, 400, 10, 200, "PLAN DE MITIGACION ", "F2", 7.5, hLeft
       
     If Not (rsFichaSobEndEvalNew.BOF Or rsFichaSobEndEvalNew.EOF) Then
            For i = 1 To (rsFichaSobEndEvalNew.RecordCount)
                
                        oDoc.WTextBox 415 + A, 50, 28, 80, "", "F1", 5, hCenter, vTop, vbBlack, 1, vbBlack
                        oDoc.WTextBox 420 + A, 55, 10, 200, rsFichaSobEndEvalNew!cCodigo, "F2", 7.5, hLeft
                        oDoc.WTextBox 415 + A, 130, 28, 190, "", "F1", 5, hCenter, vTop, vbBlack, 1, vbBlack
                        oDoc.WTextBox 420 + A, 140, 20, 190, rsFichaSobEndEvalNew!cEval, "F2", 7.5, hLeft
                        oDoc.WTextBox 415 + A, 320, 28, 230, "", "F1", 5, hCenter, vTop, vbBlack, 1, vbBlack
                        oDoc.WTextBox 415 + A, 322, 10, 230, rsFichaSobEndEvalNew!cPlanmitigacion, "F2", 7.5, hLeft '425330
                        A = A + 28
                        'If rsFichaSobEndEvalNew!nEval = 2 Then
                            'nContS = nContS + 1
                            nCodFinal = rsFichaSobEndEvalNew!nCodFinal
                        'End If
                        
                        rsFichaSobEndEvalNew.MoveNext
                
            Next i
                    
                    'If nContS >= 2 Then
                        oDoc.WTextBox 415 + A, 50, 20, 500, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                        oDoc.WTextBox 420 + A, 65, 200, 700, "RESULTADO FINAL :" & nCodFinal, "F2", 7.5, hLeft
                   ' Else
                        'oDoc.WTextBox 420 + A, 50, 20, 500, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                        'oDoc.WTextBox 425 + A, 65, 200, 500, "RESULTADO FINAL : POTENCIAL SOBREENDEUDADO", "F2", 7.5, hLeft
                    'End If
                
    End If
    
    nContS = 0
    nCodFinal = ""
    
    oDoc.WTextBox 440 + A, 50, 20, 500, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 445 + A, 65, 10, 200, "Evaluacion (Admision) ", "F2", 7.5, hLeft
    
    oDoc.WTextBox 460 + A, 50, 20, 80, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 465 + A, 65, 10, 200, "CODIGOS ", "F2", 7.5, hLeft
    oDoc.WTextBox 460 + A, 130, 20, 190, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 465 + A, 190, 10, 200, "EVALUACION ", "F2", 7.5, hLeft
    oDoc.WTextBox 460 + A, 320, 20, 230, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
    oDoc.WTextBox 465 + A, 400, 10, 200, "PLAN DE MITIGACION ", "F2", 7.5, hLeft
    
    If Not (rsFichaSobEndEvalAdm.BOF Or rsFichaSobEndEvalAdm.EOF) Then
            For i = 1 To (rsFichaSobEndEvalAdm.RecordCount)
           
                oDoc.WTextBox 480 + A, 50, 20, 80, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox 481 + A, 55, 10, 200, rsFichaSobEndEvalAdm!cCodigo, "F2", 7.5, hLeft
                oDoc.WTextBox 480 + A, 130, 20, 190, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox 481 + A, 140, 10, 190, rsFichaSobEndEvalAdm!cEval, "F2", 7.5, hLeft
                oDoc.WTextBox 480 + A, 320, 20, 230, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox 481 + A, 322, 10, 230, rsFichaSobEndEvalAdm!cPlanmitigacion, "F2", 7.5, hLeft
                A = A + 20
                
                'If rsFichaSobEndEvalAdm!nEvaluacion = 2 Then
                    'nContS = nContS + 1
                    nCodFinal = rsFichaSobEndEvalAdm!nCodFinal
                'End If
                
                rsFichaSobEndEvalAdm.MoveNext
           
            Next i
            
            'If nContS >= 2 Then
                oDoc.WTextBox 480 + A, 50, 20, 500, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                oDoc.WTextBox 485 + A, 65, 200, 700, "RESULTADO FINAL :" & nCodFinal, "F2", 7.5, hLeft
            'Else
                'oDoc.WTextBox 480 + A, 50, 20, 500, "", "F1", 7.5, hCenter, vTop, vbBlack, 1, vbBlack
                'oDoc.WTextBox 485 + A, 65, 200, 500, "RESULTADO FINAL : POTENCIAL SOBREENDEUDADO", "F2", 7.5, hLeft
            'End If
            
    End If
'Contenido
    
    oDoc.PDFClose
    oDoc.Show
Else
    MsgBox "Los Datos de la Ficha de Sobreendeudado se mostrara después de GRABAR ", vbInformation, "Aviso"
End If

End Sub

Private Sub Mantenimiento(ByVal pcPersCod As String)
    Dim lnFila As Integer
    Dim i As Integer
    Dim fnConEval As Integer
    Dim FinalCod As String
    
    Dim oDCOMMantFicha As COMDCredito.DCOMCredito
    
    Dim rsCabFichaSobEnd As ADODB.Recordset
    Dim rsFichaSobEndIfis As ADODB.Recordset
    Dim rsFichaSobEndDiasAtrs As ADODB.Recordset
    Dim rsFichaSobEndNew As ADODB.Recordset
    Dim rsFichaSobEndAdm As ADODB.Recordset
        
    Call LimpiarFlex

    Set oDCOMMantFicha = New COMDCredito.DCOMCredito
    Set rsCabFichaSobEnd = oDCOMMantFicha.ObtenerCabFichaSobEnd(pcPersCod, Format(dFechaReg, "yyyyMMdd"))
    Set rsFichaSobEndIfis = oDCOMMantFicha.ObtenerFichaSobEndIfis(pcPersCod, Format(dFechaReg, "yyyyMMdd"))
    Set rsFichaSobEndDiasAtrs = oDCOMMantFicha.ObtenerFichaSobEndDiasAtrs(pcPersCod, Format(dFechaReg, "yyyyMMdd"))
    Set rsFichaSobEndNew = oDCOMMantFicha.ObtenerFichaSobEndNew(pcPersCod, Format(dFechaReg, "yyyyMMdd"))
    Set rsFichaSobEndAdm = oDCOMMantFicha.ObtenerFichaSobEndAdm(pcPersCod, Format(dFechaReg, "yyyyMMdd"))
    
    If Not (rsCabFichaSobEnd.EOF And rsCabFichaSobEnd.BOF) Then
        fscCtaCodNew = rsCabFichaSobEnd!cCtaCodEval
        fsAgeCod = rsCabFichaSobEnd!cAgeCod
        txtCodPersM.Text = rsCabFichaSobEnd!cPersCod
        txtNombrePers.Text = rsCabFichaSobEnd!cPersNombre
        txtAna.Text = rsCabFichaSobEnd!cAnaCod
        txtAge.Text = rsCabFichaSobEnd!cAgeDescripcion
        txtSalCapConso.Text = Format(rsCabFichaSobEnd!nSaldoCap, "#,##0.00")
        txtCalfCamc.Text = rsCabFichaSobEnd!nCalifSf
        txtSalEndSistFin.Text = Format(rsCabFichaSobEnd!nSaldoEnd, "#,##0.00")
        txtMarClieAval.Text = IIf((rsCabFichaSobEnd!nMarcaGart) = "1", "Si", "No")
        txtMonCmac.Text = Format(rsCabFichaSobEnd!nMontoProv, "#,##0.00")
        fdFechaCierre = rsCabFichaSobEnd!dFechaCierre
    End If
            
    If Not (rsFichaSobEndIfis.EOF And rsFichaSobEndIfis.BOF) Then
        'feIfi
        For i = 1 To rsFichaSobEndIfis.RecordCount
            feIfi.AdicionaFila
            lnFila = feIfi.row
            
            feIfi.TextMatrix(lnFila, 1) = rsFichaSobEndIfis!cIfis
            feIfi.TextMatrix(lnFila, 2) = Format(rsFichaSobEndIfis!nMonto, "#,##0.00")
            
            rsFichaSobEndIfis.MoveNext
            
        Next i
        rsFichaSobEndIfis.Close
        Set rsFichaSobEndIfis = Nothing
    End If
    
    If Not (rsFichaSobEndDiasAtrs.EOF And rsFichaSobEndDiasAtrs.BOF) Then
        'feDiaAtraso
        For i = 1 To rsFichaSobEndDiasAtrs.RecordCount
            feDiaAtraso.AdicionaFila
            lnFila = feDiaAtraso.row
            
            feDiaAtraso.TextMatrix(lnFila, 1) = rsFichaSobEndDiasAtrs!cCtaCod
            feDiaAtraso.TextMatrix(lnFila, 2) = rsFichaSobEndDiasAtrs!cConsDescripcion
            feDiaAtraso.TextMatrix(lnFila, 3) = rsFichaSobEndDiasAtrs!nDiasAtraso
            feDiaAtraso.TextMatrix(lnFila, 4) = rsFichaSobEndDiasAtrs!cCredEstado
            
            rsFichaSobEndDiasAtrs.MoveNext
            
        Next i
        rsFichaSobEndDiasAtrs.Close
        Set rsFichaSobEndDiasAtrs = Nothing
    End If

    If Not (rsFichaSobEndNew.EOF And rsFichaSobEndNew.BOF) Then
        'feNewEvalSobre
        fnConEval = 0
        For i = 1 To rsFichaSobEndNew.RecordCount
            feNewEvalSobre.AdicionaFila
            lnFila = feNewEvalSobre.row
            
            feNewEvalSobre.TextMatrix(lnFila, 1) = rsFichaSobEndNew!cCodigo
            feNewEvalSobre.TextMatrix(lnFila, 2) = rsFichaSobEndNew!cEval
            feNewEvalSobre.TextMatrix(lnFila, 3) = rsFichaSobEndNew!cPlanmitigacion
            feNewEvalSobre.TextMatrix(lnFila, 4) = rsFichaSobEndNew!nEval
            
            FinalCod = rsFichaSobEndNew!nCodFinal
             'If rsFichaSobEndNew!nEval = 2 Then
                'fnConEval = fnConEval + 1
            'End If
            
            'If fnConEval >= 2 Then
                'txtResultadoNew.Text = "SOBREENDEUDADO"
            'Else
                'txtResultadoNew.Text = "POTENCIAL SOBREENDEUDADO"
            'End If
            
            rsFichaSobEndNew.MoveNext
            
        Next i
        
        txtResultadoNew.Text = FinalCod
        
        rsFichaSobEndNew.Close
        Set rsFichaSobEndNew = Nothing
    End If

FinalCod = ""

    If Not (rsFichaSobEndAdm.EOF And rsFichaSobEndAdm.BOF) Then
        'feAdmEvalSobr
        fnConEval = 0
        For i = 1 To rsFichaSobEndAdm.RecordCount
            feAdmEvalSobr.AdicionaFila
            lnFila = feAdmEvalSobr.row
            fscCtaCodEvalAdm = rsFichaSobEndAdm!cCtaCod
            feAdmEvalSobr.TextMatrix(lnFila, 1) = rsFichaSobEndAdm!cCodigo
            feAdmEvalSobr.TextMatrix(lnFila, 2) = rsFichaSobEndAdm!cEval
            feAdmEvalSobr.TextMatrix(lnFila, 3) = rsFichaSobEndAdm!cPlanmitigacion
            feAdmEvalSobr.TextMatrix(lnFila, 4) = rsFichaSobEndAdm!cMantCodigo
            feAdmEvalSobr.TextMatrix(lnFila, 5) = rsFichaSobEndAdm!nEvaluacion
            
            FinalCod = rsFichaSobEndAdm!nCodFinal
            
            'If rsFichaSobEndAdm!nEvaluacion = 2 Then
                'fnConEval = fnConEval + 1
            'End If
            
            'If fnConEval >= 2 Then
                'txtResultado.Text = "SOBREENDEUDADO"
            'Else
                'txtResultado.Text = "POTENCIAL SOBREENDEUDADO"
            'End If
            
            rsFichaSobEndAdm.MoveNext
            
        Next i
        rsFichaSobEndAdm.Close
        Set rsFichaSobEndAdm = Nothing
        fnConEval = 0
    End If
End Sub

Private Sub cmdSalir_Click()
    Call LimpiarCabecera
    Call LimpiarFlex
    Unload Me
End Sub

Public Function Valida() As Boolean
Dim i As Integer
Valida = True

If fnTipoReg = 1 Then
    'Persona Registro
    If txtCodPers.Text = "" Then
        MsgBox "Seleccione al Cliente", vbInformation, "Aviso"
        txtCodPers.SetFocus
        Valida = False
        Exit Function
    End If
Else
    'Persona Mantenimiento
    If txtCodPersM.Text = "" Then
        MsgBox "Seleccione al Cliente", vbInformation, "Aviso"
        cmdExam.SetFocus
        Valida = False
        Exit Function
    End If
End If

If (feNewEvalSobre.rows - 1) > 1 Then
    For i = 1 To (feNewEvalSobre.rows - 1)
        If CInt(feNewEvalSobre.TextMatrix(i, 4)) <> 0 Then
            If feNewEvalSobre.TextMatrix(i, 3) = "" Then
                MsgBox "Registre el Plan de Mitigación del " & feNewEvalSobre.TextMatrix(i, 1) & "", vbInformation, "Aviso"
                feNewEvalSobre.SetFocus
                Valida = False
                Exit Function
            End If
        End If
    Next i
End If

If (feAdmEvalSobr.TextMatrix(1, 1) = "" And feNewEvalSobre.TextMatrix(1, 1) = "") Then
    MsgBox "El Cliente no Tuvo Evaluación de Sobreendeudamiento.", vbInformation, "Aviso"
    Valida = False
    Exit Function
End If

'If (feAdmEvalSobr.Rows - 1) <= 1 Then
 '   MsgBox "El Cliente No tuvo Evaluacion de Admision", vbInformation, "Aviso"
  '  Valida = False
   ' Exit Function
'End If

'If (feNewEvalSobre.Rows - 1) <= 1 Then
'    MsgBox "El Cliente No tiene Evaluacion Nueva", vbInformation, "Aviso"
 '   Valida = False
  '  Exit Function
'End If
End Function

Private Sub feNewEvalSobre_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String

If pnCol = 3 Then
    If CInt(feNewEvalSobre.TextMatrix(pnRow, 4)) = 0 Then
        Cancel = False
        MsgBox "No se permite ingresar Plan de Mitigación a códigos que no corresponde.", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If

End If
End Sub
