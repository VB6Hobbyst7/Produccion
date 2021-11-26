VERSION 5.00
Begin VB.Form frmOpeComisionOtros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comisión: "
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   Icon            =   "frmOpeComisionOtros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   2880
      TabIndex        =   5
      Top             =   2040
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   4200
      TabIndex        =   4
      Top             =   2040
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin SICMACT.TxtBuscar TxtBCodPers 
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1860
         _extentx        =   3281
         _extenty        =   582
         appearance      =   1
         appearance      =   1
         font            =   "frmOpeComisionOtros.frx":030A
         appearance      =   1
         tipobusqueda    =   3
         stitulo         =   ""
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   9
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label lblDOI 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   8
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "D.O.I. : "
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1245
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente: "
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   780
         Width           =   615
      End
   End
   Begin VB.Label Label12 
      Caption         =   "Monto S/.:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2115
      Width           =   840
   End
   Begin VB.Label lblComision 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "frmOpeComisionOtros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmOpeComisionOtros
'** Descripción : Formulario para registrar la comision para otros conceptos creado segun TI-ERS029-2013
'** Creación : JUEZ, 20130411 09:00:00 AM
'**********************************************************************************************

Option Explicit
Dim R As ADODB.Recordset
Dim fsOpeCod As Long
Dim fsConceptoCod As Integer
Dim fsGlosa As String
Dim fsTitVoucher As String

Public Sub Inicia(ByVal psOpecod As Long, ByVal pnConcepto As Integer, ByVal psTitulo As String, ByVal psGlosa As String, ByVal psTitVoucher As String)
    Dim loParam As COMDColocPig.DCOMColPCalculos
    fsOpeCod = psOpecod
    fsConceptoCod = pnConcepto
    fsGlosa = psGlosa
    fsTitVoucher = psTitVoucher
    
    Me.Caption = Me.Caption & psTitulo
    
    Set loParam = New COMDColocPig.DCOMColPCalculos
    lblComision.Caption = Format(loParam.dObtieneColocParametro(fsConceptoCod), "#,##0.00")
    Set loParam = Nothing
    Me.Show 1
End Sub

Private Sub CmdAceptar_Click()
    If MsgBox("Desea Grabar la Información?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Dim oCredMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set oCredMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsCont As COMNContabilidad.NCOMContFunciones
    Set clsCont = New COMNContabilidad.NCOMContFunciones
    Dim lsMov As String
    Dim lsBoleta As String

    lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    gnMovNro = 0
    gnMovNro = oCredMov.OtrasOperaciones(lsMov, fsOpeCod, CDbl(lblComision.Caption), "", fsGlosa & " cliente: " & lblCliente.Caption, gMonedaNacional, TxtBCodPers.Text, , , , , , , gnMovNro)
    If gnMovNro <> 0 Then
    
        If fsOpeCod = gComisionEvalPolEndosada Then
            Dim oCred As COMDCredito.DCOMCredActBD
            Set oCred = New COMDCredito.DCOMCredActBD
            Call oCred.dInsertComision(gnMovNro, TxtBCodPers.Text, CDbl(lblComision.Caption))
            Set oCred = Nothing
        End If
        Dim oBol As COMNCredito.NCOMCredDoc
        Set oBol = New COMNCredito.NCOMCredDoc
            lsBoleta = oBol.ImprimeBoletaComision(fsTitVoucher, Left("Total pago comision", 36), "", str(CDbl(lblComision.Caption)), lblCliente.Caption, lblDOI.Caption, "________" & gMonedaNacional, False, "", "", , gdFecSis, gsNomAge, gsCodUser, sLpt, , gbImpTMU)
        Set oBol = Nothing
        
        Do
           If Trim(lsBoleta) <> "" Then
                lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                    Print #nFicSal, oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                    Print #nFicSal, ""
                Close #nFicSal
          End If
            
        Loop While MsgBox("Desea Re Imprimir ?", vbQuestion + vbYesNo, "Aviso") = vbYes
        Set oBol = Nothing
        Limpiar
    Else
        MsgBox "Hubo un error en el registro", vbInformation, "Aviso"
    End If
End Sub

Private Sub Limpiar()
    lblCliente.Caption = ""
    lblDOI.Caption = ""
    TxtBCodPers.Text = ""
    cmdAceptar.Enabled = False
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub TxtBCodPers_EmiteDatos()
    If TxtBCodPers.Text <> "" Then
        Dim oCred As COMDCredito.DCOMCredito
        Set oCred = New COMDCredito.DCOMCredito
        If fsOpeCod = gComisionEvalPolEndosada Then
            If oCred.ExisteCreditoTitular(TxtBCodPers.Text, True, True, , True) Then 'WIOR 20130829 MOSTRAR CREDITOS VIGENTES
                CargaDatos
            Else
                MsgBox "Sólo se puede realizar pago por este concepto a nombre de un titular de crédito en estado aprobado", vbInformation, "Aviso"
                Limpiar
            End If
        ElseIf fsOpeCod = gComisionDupTasacion Then
            If oCred.ExisteCreditoTitular(TxtBCodPers.Text, , True, True) Then
                CargaDatos
            Else
                MsgBox "Sólo se puede realizar pago por este concepto a nombre de un titular de crédito", vbInformation, "Aviso"
                Limpiar
            End If
        Else
            CargaDatos
        End If
    Else
        Limpiar
    End If
    Set oCred = Nothing
End Sub

Private Sub CargaDatos()
    Dim oCred As COMDCredito.DCOMCredito
    Set oCred = New COMDCredito.DCOMCredito
    Set R = oCred.RecuperaDatosComision(TxtBCodPers.Text, 2)
    Set oCred = Nothing
    lblCliente.Caption = R!cPersNombre
    lblDOI.Caption = R!cPersIDnro
    Set R = Nothing
    cmdAceptar.Enabled = True
End Sub
