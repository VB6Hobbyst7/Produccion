VERSION 5.00
Begin VB.Form frmCredSolicitud_SelecFtes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar Fuentes de Ingreso para el Credito"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8685
   Icon            =   "frmCredSolicitud_SelecFtes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin SICMACT.FlexEdit FeFuentes 
      Height          =   2055
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   8175
      _extentx        =   14420
      _extenty        =   3625
      cols0           =   6
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   "-cNumFuente-Sel-Fuente-Tipo-FechEval"
      encabezadosanchos=   "400-0-400-4520-600-1200"
      font            =   "frmCredSolicitud_SelecFtes.frx":030A
      font            =   "frmCredSolicitud_SelecFtes.frx":0336
      font            =   "frmCredSolicitud_SelecFtes.frx":0362
      font            =   "frmCredSolicitud_SelecFtes.frx":038E
      font            =   "frmCredSolicitud_SelecFtes.frx":03BA
      fontfixed       =   "frmCredSolicitud_SelecFtes.frx":03E6
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      columnasaeditar =   "X-X-2-X-X-X"
      listacontroles  =   "0-0-4-0-0-0"
      encabezadosalineacion=   "C-C-C-C-C-C"
      formatosedit    =   "0-0-0-0-0-0"
      lbeditarflex    =   -1
      lbpuntero       =   -1
      lbordenacol     =   -1
      colwidth0       =   405
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5145
      TabIndex        =   1
      Top             =   2610
      Width           =   1230
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3930
      TabIndex        =   0
      Top             =   2610
      Width           =   1230
   End
End
Attribute VB_Name = "frmCredSolicitud_SelecFtes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sPersCod As String
Public MatFuentes As Variant
Public MatFuentesF As Variant


Public Sub Inicio(ByVal psPersCod As String)
    sPersCod = psPersCod
    Call CargaFuentesIngreso(psPersCod)
    Me.Show 1
End Sub

Private Sub CargaFuentesIngreso(ByVal psPersCod As String)
Dim i As Integer
Dim MatFte As Variant
Dim oPersona As UPersona_Cli

    On Error GoTo ErrorCargaFuentesIngreso
    
    Set oPersona = New UPersona_Cli 'COMDPersona.DCOMPersona
    Call oPersona.RecuperaPersona_Solicitud(psPersCod, gdFecSis) 'oPersona.RecuperaPersona(psPersCod)
    
    MatFte = oPersona.FiltraFuentesIngresoPorRazonSocial
    
    FeFuentes.Clear
    FeFuentes.Rows = 2
    FeFuentes.FormaCabecera
    
    If IsArray(MatFte) Then
        For i = 0 To UBound(MatFte) - 1
            'cmbFuentes.AddItem MatFte(i, 2) & Space(100 - Len(MatFte(i, 2))) & MatFte(i, 6) & Space(50 - Len(MatFte(i, 6))) & MatFte(i, 8)
            'MatFuentes(i) = MatFte(i, 1)
            With FeFuentes
                .AdicionaFila
                .TextMatrix(i + 1, 2) = "0"
                .TextMatrix(i + 1, 1) = MatFte(i, 8)
                .TextMatrix(i + 1, 3) = MatFte(i, 2)
                .TextMatrix(i + 1, 4) = MatFte(i, 1)
                .TextMatrix(i + 1, 5) = MatFte(i, 4)
            End With
        Next i
    End If
    
    Exit Sub

ErrorCargaFuentesIngreso:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmdAceptar_Click()
'MatFuentes
Dim i As Integer
ReDim MatFuentes(0)
ReDim MatFuentesF(3, 0)
Dim retorno As Integer
retorno = 0
With FeFuentes
    For i = 1 To .Rows - 1
        If .TextMatrix(i, 2) = "." Then
            ReDim Preserve MatFuentes(UBound(MatFuentes) + 1)
            ReDim Preserve MatFuentesF(3, 1)
            MatFuentes(UBound(MatFuentes) - 1) = i - 1 '.TextMatrix(i, 1)
            '**ALPA****18/04/2008*********************
            MatFuentesF(1, 1) = FeFuentes.TextMatrix(FeFuentes.Row, 1)
            If Right(FeFuentes.TextMatrix(FeFuentes.Row, 5), 10) = "1900/01/01" Then
                ReDim MatFuentes(0)
                ReDim MatFuentesF(3, 1)
                retorno = 1
                Exit For
            End If
            MatFuentesF(2, 1) = Right(FeFuentes.TextMatrix(FeFuentes.Row, 5), 10)
            MatFuentesF(3, 1) = FeFuentes.TextMatrix(FeFuentes.Row, 4)
            retorno = 2
            '**End*************************************
        End If
    Next
End With
'**ALPA****18/04/2008*********************
If retorno = 1 Then
    MsgBox "Fecha no esta Vigente", vbInformation, "Aviso"
    CmdAceptar.SetFocus
End If
If retorno = 0 Then
    MsgBox "Selecciona Fuente de Ingreso", vbInformation, "Aviso"
    CmdAceptar.SetFocus
End If
If retorno = 2 Then
Unload Me
End If
'*****end***********************************
End Sub

Private Sub cmdCancelar_Click()
    Dim sMaTem() As String
    ReDim Preserve sMaTem(3, 1)
    MatFuentesF = sMaTem
    MatFuentesF(3, 1) = ""
    Unload Me
End Sub
'**ALPA****18/04/2008*********************
'Private Sub FeFuentes_RowColChange()
'If FeFuentes.Col = 5 Then
'    Dim oCrDoc As New ComdCredito.DCOMCredDoc
'    FeFuentes.CargaCombo oCrDoc.ObtenerFechaProNumFuente(FeFuentes.TextMatrix(FeFuentes.Row, 1), gdFecSis)
'    Set oCrDoc = Nothing
'FeFuentes.TextMatrix(FeFuentes.Row, 5) = Right(FeFuentes.TextMatrix(FeFuentes.Row, 5), 10)
''*****end***********************************
'End If
'End Sub

Private Sub Form_Load()
    CentraForm Me
End Sub
