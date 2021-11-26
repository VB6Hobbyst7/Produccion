VERSION 5.00
Begin VB.Form frmCredAmpliadoNew 
   Caption         =   "Ampliación de Créditos"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8040
   Icon            =   "frmCredAmpliadoNew.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   6720
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin SICMACT.FlexEdit feAmpliacion 
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   7815
      _extentx        =   13785
      _extenty        =   5953
      cols0           =   7
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   "#--Chk-Crédito-Tipo Producto-Moneda-Monto"
      encabezadosanchos=   "400-0-500-2000-2300-900-1300"
      font            =   "frmCredAmpliadoNew.frx":030A
      font            =   "frmCredAmpliadoNew.frx":0336
      font            =   "frmCredAmpliadoNew.frx":0362
      font            =   "frmCredAmpliadoNew.frx":038E
      font            =   "frmCredAmpliadoNew.frx":03BA
      fontfixed       =   "frmCredAmpliadoNew.frx":03E6
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      columnasaeditar =   "X-X-2-X-X-X-X"
      listacontroles  =   "0-0-4-0-0-0-0"
      encabezadosalineacion=   "C-C-C-C-L-L-R"
      formatosedit    =   "0-0-0-0-0-0-0"
      textarray0      =   "#"
      lbeditarflex    =   -1
      lbbuscaduplicadotext=   -1
      colwidth0       =   405
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.TextBox txtNomCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   5055
      End
      Begin VB.TextBox txtCodCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmCredAmpliadoNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmCredAmpliadoNew
'***     Descripcion:       Permite seleccionar los creditos a cancelar
'***     Creado por:        FRHU
'***     Fecha-Tiempo:         24/04/2014 01:00:00 PM
'*****************************************************************************************
Option Explicit
Private rs_Ampliado As ADODB.Recordset
Private MatCalend As Variant
Public nIdCampana As Integer
Private Type TColumna
    sCheck As String
    sCredito As String
    sTipoProducto As String
    sMoneda As String
    nMonto As Double
End Type
Dim MatDatos() As TColumna
Dim MontoTotal As Double

Public Sub Inicio(ByVal psPersCod As String, ByVal psMoneda As String, ByVal psNombre As String, ByRef nMontoTotal As Double, ByVal prsAmpliado As ADODB.Recordset, Optional ByVal cCtaCodPreSol As String = "", Optional ByVal cCodProd As String = "") 'JOEP20190919 ERS042 CP-2018 Agrego c
    Dim oRs As ADODB.Recordset
    Dim fila As Integer
    Dim m As Integer
    Dim i As Integer
    Dim objAmpliado As New COMDCredito.DCOMAmpliacion
    Dim xfila As Integer
    
    Me.txtCodCliente.Text = psPersCod
    Me.txtNomCliente.Text = psNombre
    Set oRs = objAmpliado.ObtenerCreditoXPersona(psPersCod, psMoneda, cCodProd)
    fila = 1
    xfila = 1
    Do While Not oRs.EOF
        Call CargarDatos(oRs!cCtaCod, fila)
        oRs.MoveNext
    Loop
    For i = 1 To fila - 1
        Me.feAmpliacion.AdicionaFila
        Me.feAmpliacion.TextMatrix(i, 1) = "0"
        If Not prsAmpliado Is Nothing Then
            prsAmpliado.MoveFirst
            Do While Not prsAmpliado.EOF
                If MatDatos(i).sCredito = prsAmpliado(0) Then
                    Me.feAmpliacion.TextMatrix(i, 2) = "1"
                End If
                prsAmpliado.MoveNext
            Loop
        End If
        
        'agregado por vapi SEGÙN ERS TI-ERS001-2017
        If cCtaCodPreSol <> "" And cCtaCodPreSol = MatDatos(i).sCredito Then
            Me.feAmpliacion.TextMatrix(i, 2) = "1"
        End If
        'fin agregado por vapi
        
        Me.feAmpliacion.TextMatrix(i, 3) = MatDatos(i).sCredito
        Me.feAmpliacion.TextMatrix(i, 4) = MatDatos(i).sTipoProducto
        Me.feAmpliacion.TextMatrix(i, 5) = MatDatos(i).sMoneda
        Me.feAmpliacion.TextMatrix(i, 6) = Format(MatDatos(i).nMonto, "#,###,##0.00")
    Next i
    Me.Show 1
    nMontoTotal = MontoTotal
End Sub

Private Sub CargarDatos(ByVal psCtaCod As String, ByRef pnFila As Integer)
    Dim rs As ADODB.Recordset
    Dim oNegCredito As COMNCredito.NCOMCredito
    Dim oAmpliado As COMDCredito.DCOMAmpliacion
    Dim nInteresFecha As Double
    Dim nMontoFecha As Double
    
    Dim Item As ListItem
    
    Set oNegCredito = New COMNCredito.NCOMCredito
    MatCalend = oNegCredito.RecuperaMatrizCalendarioPendiente(psCtaCod)
    If Not IsArray(MatCalend) Then
        'MsgBox "La Cuenta no tiene Calendario pendiente", vbInformation, "Mensaje"
        Set oNegCredito = Nothing
        Exit Sub
    End If
    If UBound(MatCalend) = 0 Then
        'MsgBox "La Cuenta no tiene Calendario pendiente", vbInformation, "Mensaje"
        Set oNegCredito = Nothing
        Exit Sub
    End If
    'MAVM 14092010 Se incluyo el Int Grac
    'nInteresFecha = oNegCredito.MatrizInteresGastosAFecha(Me.ActXCodCta1.NroCuenta, MatCalend, gdFecSis, True)
    nInteresFecha = oNegCredito.MatrizInteresGastosAFecha(psCtaCod, MatCalend, gdFecSis, True) + oNegCredito.MatrizInteresGraAFecha(psCtaCod, MatCalend, gdFecSis)
    nMontoFecha = oNegCredito.MatrizCapitalAFecha(psCtaCod, MatCalend)
    
    Set oNegCredito = Nothing
    
    nMontoFecha = nInteresFecha + nMontoFecha
    
    Set oAmpliado = New COMDCredito.DCOMAmpliacion
    Set rs = oAmpliado.ListaDatosAmpliacion(psCtaCod)
    Set oAmpliado = Nothing
    
    ReDim Preserve MatDatos(pnFila)
    If Not rs.BOF And Not rs.EOF Then
         MatDatos(pnFila).sCredito = psCtaCod
         MatDatos(pnFila).sTipoProducto = rs!TipoProducto
         MatDatos(pnFila).sMoneda = IIf(IsNull(rs!Moneda), "", rs!Moneda)
         MatDatos(pnFila).nMonto = Format(nMontoFecha, "#0.00")
         
        nIdCampana = rs!idCampana
        pnFila = pnFila + 1
    End If
    Set rs = Nothing
End Sub

Private Sub cmdAceptar_Click()
    Dim i As Integer
    MontoTotal = 0
    If validatos Then
        If Not CredAmpCumpleCondiciones Then Exit Sub 'FRHU 20160615 ERS002-2016
        If MsgBox("Desea establecer como un credito ampliado", vbInformation + vbYesNo) = vbYes Then
            Call InicializarRecord
           
           For i = 1 To Me.feAmpliacion.rows - 1
                If Me.feAmpliacion.TextMatrix(i, 2) = "." Then
                    With rs_Ampliado
                        rs_Ampliado.AddNew
                        rs_Ampliado(0) = Me.feAmpliacion.TextMatrix(i, 3)
                        rs_Ampliado(1) = Mid(Me.feAmpliacion.TextMatrix(i, 4), 1, 30)
                        rs_Ampliado(2) = gdFecSis
                        rs_Ampliado(3) = Me.feAmpliacion.TextMatrix(i, 6)
                        rs_Ampliado(4) = Me.feAmpliacion.TextMatrix(i, 5)
                        
                        .Update
                    End With
                    MontoTotal = MontoTotal + CDbl(Me.feAmpliacion.TextMatrix(i, 6))
                End If
            Next i
            Set frmCredSolicitud.rsAmpliado = Nothing
            Set frmCredSolicitud.rsAmpliado = rs_Ampliado
            Set rs_Ampliado = Nothing
            Unload Me
        End If
    Else
        MsgBox "Debe seleccionar al menos un crédito para continuar con el procedimiento", vbInformation
        Exit Sub
    End If
End Sub
Sub InicializarRecord()
    Set rs_Ampliado = New ADODB.Recordset
    With rs_Ampliado.Fields
        .Append "cCtaCod", adVarChar, 18
        .Append "TipoProducto", adVarChar, 30
        .Append "dFecha", adDate
        .Append "nMonto", adDouble
        .Append "Moneda", adVarChar, 10
    End With
    rs_Ampliado.Open
End Sub
Private Function validatos() As Boolean
    Dim i As Integer
    validatos = False
    For i = 1 To Me.feAmpliacion.rows - 1
         If Me.feAmpliacion.TextMatrix(i, 2) = "." Then
            validatos = True
         End If
    Next i
End Function

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'FRHU 20160615 ERS002-2016
Private Function CredAmpCumpleCondiciones() As Boolean
    Dim objCred As New COMDCredito.DCOMCredito
    Dim rsCred As New ADODB.Recordset
    Dim lnPase As Integer
    Dim lsTpoDesc As String
    Dim i As Integer
    Dim lsCtaCodAmp As String
    
On Error GoTo ErrorCredAmpCumpleCondiciones
    CredAmpCumpleCondiciones = False
    For i = 1 To Me.feAmpliacion.rows - 1
        If feAmpliacion.TextMatrix(i, 2) = "." Then
            lsCtaCodAmp = feAmpliacion.TextMatrix(i, 3)
            Set rsCred = objCred.ValidarCondicionesCredAmpliados(lsCtaCodAmp, Trim(Me.txtCodCliente.Text), gdFecSis, gsCodUser, gsCodAge)
            If Not (rsCred.BOF And rsCred.EOF) Then
                lnPase = rsCred!nPase
                lsTpoDesc = rsCred!cTpoDesc
                If lnPase = 0 Then
                    MsgBox lsTpoDesc, vbInformation, "AVISO"
                    Exit Function
                End If
            End If
        End If
    Next i
    CredAmpCumpleCondiciones = True
    
    Exit Function
ErrorCredAmpCumpleCondiciones:
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
'FIN FRHU 20160615
