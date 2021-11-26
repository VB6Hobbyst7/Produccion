VERSION 5.00
Begin VB.Form frmCFExtornoRenovacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extorno"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11010
   Icon            =   "frmCFExtornoRenovacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame11 
      Caption         =   "Cartas Fianzas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10575
      Begin VB.TextBox txtGlosa 
         Height          =   615
         Left            =   120
         MaxLength       =   700
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2400
         Width           =   7695
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   390
         Left            =   9240
         TabIndex        =   3
         ToolTipText     =   "Ir al Menu Principal"
         Top             =   2400
         Width           =   1185
      End
      Begin VB.CommandButton cmdExtornar 
         Caption         =   "&Extornar"
         Enabled         =   0   'False
         Height          =   390
         Left            =   7920
         TabIndex        =   2
         ToolTipText     =   "Grabar Datos de Aprobacion de Credito"
         Top             =   2400
         Width           =   1185
      End
      Begin SICMACT.FlexEdit feCartaFianza 
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10320
         _ExtentX        =   17992
         _ExtentY        =   3201
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Nº Carta Fianza-Afianzado-Nueva Emision-Periodo-Moneda-Monto-cMovAct-nMovNro-Usuario"
         EncabezadosAnchos=   "400-1800-2500-1400-1200-1200-1500-0-0-800"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-5-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Glosa"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmCFExtornoRenovacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************************
'*  APLICACION : Carta Fianza
'*  ARCHIVO : frmCFExtornoRenovacion
'*  CREACION: 03/12/2013     - WIOR
'*******************************************************************************************
'*  RESUMEN: EXTORNO DE APROBACION DE RENOVACION DE CARTA FIANZA Y EXTORNO DE LA RENOVACION.
'*******************************************************************************************

Option Explicit
Dim fnTpoExtorno As Integer

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Public Sub Inicio(ByVal pbTpoExt As Integer)

Select Case pbTpoExt
    Case 1: Me.Caption = "Extorno de Autorizaciones de Renovación"
    Case 2: Me.Caption = "Extorno de Renovación de Carta Fianza"
End Select

fnTpoExtorno = pbTpoExt
Call CargaDatos
Me.Show 1
End Sub

Private Sub CargaDatos()
Dim oCF As COMDCartaFianza.DCOMCartaFianza
Dim rsCF As ADODB.Recordset
Dim i As Long

Set oCF = New COMDCartaFianza.DCOMCartaFianza

Select Case fnTpoExtorno
    Case 1: Set rsCF = oCF.ObternerAutRenovacion()
    Case 2: Set rsCF = oCF.ObternerRenovacionesHechas()
End Select


Call LimpiaFlex(feCartaFianza)
If rsCF.RecordCount > 0 Then
    For i = 0 To rsCF.RecordCount - 1
        feCartaFianza.AdicionaFila
        feCartaFianza.TextMatrix(i + 1, 0) = i + 1
        feCartaFianza.TextMatrix(i + 1, 1) = rsCF!cCtaCod
        feCartaFianza.TextMatrix(i + 1, 2) = rsCF!Afianzado
        feCartaFianza.TextMatrix(i + 1, 3) = Format(rsCF!NEmision, "dd/mm/yyyy")
        feCartaFianza.TextMatrix(i + 1, 4) = rsCF!nPeriodo
        feCartaFianza.TextMatrix(i + 1, 5) = rsCF!Moneda
        feCartaFianza.TextMatrix(i + 1, 6) = Format(rsCF!nMonto, "#0,000.00")
        feCartaFianza.TextMatrix(i + 1, 7) = rsCF!cMovAut
        feCartaFianza.TextMatrix(i + 1, 8) = rsCF!nMovNro
        feCartaFianza.TextMatrix(i + 1, 9) = Right(Trim(rsCF!cMovAut), 4)
        rsCF.MoveNext
    Next i
Else
    MsgBox "No hay datos.", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdExtornar_Click()
If ValidaDatos Then
    If fnTpoExtorno = 1 Then
        Call ExtornoAutRenovacion
    Else
        Call ExtornoRenovacionRealizada
    End If
End If
End Sub

Private Sub feCartaFianza_Click()
    cmdExtornar.Enabled = True
End Sub

Private Function ValidaDatos() As Boolean
If Trim(txtGlosa.Text) = "" Then
    MsgBox "Ingrese la Glosa.", vbInformation, "Aviso"
    ValidaDatos = False
    Exit Function
End If

ValidaDatos = True
End Function

Private Sub ExtornoAutRenovacion()
Dim cMovAut As String
Dim oCF As COMDCartaFianza.DCOMCartaFianza
Dim rsCF As ADODB.Recordset

Set oCF = New COMDCartaFianza.DCOMCartaFianza
Set rsCF = oCF.ObternerAutRenovacion(Trim(feCartaFianza.TextMatrix(feCartaFianza.row, 1)), 0)

cMovAut = Trim(feCartaFianza.TextMatrix(feCartaFianza.row, 7))

If Not (rsCF.EOF And rsCF.BOF) Then
    If rsCF.RecordCount > 0 Then
        If MsgBox("Estas seguro de Extornar la Autorizacion de la Renovación de la Carta Fianza Nº " & _
            Trim(feCartaFianza.TextMatrix(feCartaFianza.row, 1)) & _
            Chr(10) & "Autorizado Por " & Right(cMovAut, 4) & " el día " & _
            Mid(cMovAut, 7, 2) & "/" & Mid(cMovAut, 5, 2) & "/" & Mid(cMovAut, 1, 4) & _
            " a las " & Mid(cMovAut, 9, 2) & ":" & Mid(cMovAut, 11, 2) & ":" & Mid(cMovAut, 13, 2) & " ?", _
            vbInformation + vbYesNo, "Aviso") = vbYes Then
        
        
            Call oCF.InsActAutRenovacionCF(2, Trim(feCartaFianza.TextMatrix(feCartaFianza.row, 1)), , , , , , , , cMovAut, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), , Trim(txtGlosa.Text), 4)

        
            MsgBox "Datos Extornados Satisfactoriamente"
            Call LimpiarDatos
        End If
    Else
        MsgBox "No se puede extornar la Autorización de la Carta Fianza"
        Call LimpiarDatos
    End If
Else
    MsgBox "No se puede extornar la Autorización de la Carta Fianza"
    Call LimpiarDatos
End If
End Sub

Private Sub LimpiarDatos()
txtGlosa.Text = ""
cmdExtornar.Enabled = False
Call CargaDatos
End Sub



Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
 KeyAscii = fgIntfMayusculas(KeyAscii)
    If KeyAscii = 13 Then
        If cmdExtornar.Enabled Then
            cmdExtornar.SetFocus
        End If
    End If
End Sub
Private Sub ExtornoRenovacionRealizada()
Dim nRenova As Integer
Dim sCtaCod As String

Dim oCF As COMDCartaFianza.DCOMCartaFianza
Dim rsCF As ADODB.Recordset
Dim oCFExtorno As COMDCredito.DCOMCredActBD

Set oCFExtorno = New COMDCredito.DCOMCredActBD
Set oCF = New COMDCartaFianza.DCOMCartaFianza

nRenova = Trim(feCartaFianza.TextMatrix(feCartaFianza.row, 8))
sCtaCod = Trim(feCartaFianza.TextMatrix(feCartaFianza.row, 1))
Set rsCF = oCF.OperacionesCFRestaura(2, sCtaCod, nRenova)

If Not (rsCF.EOF And rsCF.BOF) Then
    If rsCF.RecordCount > 0 Then
        If MsgBox("Estas seguro de Extornar la Renovación Nº " & (nRenova + 1) & " de la Carta Fianza Nº " & sCtaCod & " ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            
            Call oCFExtorno.dUpdateProducto(sCtaCod, , CDbl(rsCF!nMonto), gColocEstVigNorm, CDate(rsCF!dVigCol), -2, False)
            Call oCFExtorno.dUpdateColocaciones(sCtaCod, , CDate(rsCF!dVencimiento), CDbl(rsCF!nMonto), , , CDate(rsCF!dVigCol), False, False, "")
            Call oCFExtorno.dUpdateColocCartaFianza(sCtaCod, , , CDate(rsCF!dEmision), CDate(rsCF!dVencimiento), , , , nRenova, CDate(rsCF!dEmision), True)
            
            Call oCFExtorno.dDeleteColocCFEstado(sCtaCod, gColocEstRenovada)
            
            If nRenova > 0 Then
                Call oCFExtorno.dInsertColocCFEstado(sCtaCod, CDate(rsCF!dPrdEstado), gColocEstRenovada, CDbl(rsCF!nMonto), CDate(rsCF!dVencimiento), "Renovacion de Carta Fianza", CInt(rsCF!nMotivoRechazo), False, True)
            End If
    
            Call oCF.RegistroExtornoRenovacionCF(sCtaCod, nRenova, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), Trim(txtGlosa.Text))
        
            Call oCF.OperacionesCFRestaura(3, sCtaCod, nRenova)

        
            MsgBox "Datos Extornados Satisfactoriamente"
            Call LimpiarDatos
        End If
    Else
        MsgBox "No se puede Extornar la Renovación de la Carta Fianza"
        Call LimpiarDatos
    End If
Else
    MsgBox "No se puede Extornar la Renovación de la Carta Fianza"
    Call LimpiarDatos
End If
End Sub

