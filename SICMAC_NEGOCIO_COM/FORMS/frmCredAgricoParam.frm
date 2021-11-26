VERSION 5.00
Begin VB.Form frmCredAgricoParam 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Tipos de Créditos Agropecuarios"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8010
   Icon            =   "frmCredAgricoParam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   1170
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   3600
      Width           =   1170
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Quitar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   3600
      Width           =   1170
   End
   Begin SICMACT.FlexEdit feTpoAgrico 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5953
      Cols0           =   5
      HighLight       =   1
      EncabezadosNombres=   "Nº-Tipo-Subtipo-Minimo-CodTipo"
      EncabezadosAnchos=   "500-1800-4200-1000-0"
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
      ColumnasAEditar =   "X-1-2-3-X"
      ListaControles  =   "0-3-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-C-C"
      FormatosEdit    =   "0-0-0-0-0"
      TextArray0      =   "Nº"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      TipoBusqueda    =   3
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   495
      RowHeight0      =   300
   End
End
Attribute VB_Name = "frmCredAgricoParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private i As Integer
Private fsCodSubProd As String
Private fsCodSubProdActivar As String
Private fnOpeRealiza As Integer

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdNuevo_Click()
If ValidaDatos Then
  feTpoAgrico.AdicionaFila
    feTpoAgrico.SetFocus
    SendKeys "{Enter}"
End If
End Sub

Private Sub cmdQuitar_Click()
If MsgBox("Estas seguro de quitar este registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    Dim oParam As COMDCredito.DCOMParametro
    Set oParam = New COMDCredito.DCOMParametro
    
    If Trim(feTpoAgrico.TextMatrix(feTpoAgrico.Row, 4)) <> "" Then
        Call oParam.ActualizaParametrosAgro(CLng(Trim(Right(feTpoAgrico.TextMatrix(feTpoAgrico.Row, 1), 10))), Trim(feTpoAgrico.TextMatrix(feTpoAgrico.Row, 4)), Trim(feTpoAgrico.TextMatrix(feTpoAgrico.Row, 2)), 0, CLng(feTpoAgrico.TextMatrix(feTpoAgrico.Row, 3)))
    End If
    
    feTpoAgrico.EliminaFila feTpoAgrico.Row
    
    CargaGrid
End If
End Sub

Private Sub feTpoAgrico_OnCellChange(pnRow As Long, pnCol As Long)
If ValidaDatos Then
    Dim oParam As COMDCredito.DCOMParametro
    Set oParam = New COMDCredito.DCOMParametro

    If fnOpeRealiza = 1 Then
        Call oParam.InsertaParametrosAgro(CLng(Trim(Right(feTpoAgrico.TextMatrix(feTpoAgrico.Row, 1), 10))), Trim(feTpoAgrico.TextMatrix(feTpoAgrico.Row, 2)), , CLng(feTpoAgrico.TextMatrix(feTpoAgrico.Row, 3)))
    ElseIf fnOpeRealiza = 2 Then
        Call oParam.ActualizaParametrosAgro(CLng(Trim(Right(feTpoAgrico.TextMatrix(feTpoAgrico.Row, 1), 10))), fsCodSubProdActivar, Trim(feTpoAgrico.TextMatrix(feTpoAgrico.Row, 2)), , CLng(feTpoAgrico.TextMatrix(feTpoAgrico.Row, 3)))
    End If
    CargaGrid
End If
End Sub

Private Sub Form_Load()
   CargaControles
End Sub
Private Sub CargaControles()
Dim oConst As COMDConstantes.DCOMConstantes
Set oConst = New COMDConstantes.DCOMConstantes

feTpoAgrico.CargaCombo oConst.RecuperaConstantes(7067)
CargaGrid
End Sub

Private Sub CargaGrid()
Dim oParam As COMDCredito.DCOMParametro
Dim rsParam As ADODB.Recordset

Set oParam = New COMDCredito.DCOMParametro

Set rsParam = oParam.ObtenerParametrosAgro
LimpiaFlex feTpoAgrico
If Not (rsParam.EOF And rsParam.BOF) Then
    For i = 1 To rsParam.RecordCount
        feTpoAgrico.AdicionaFila
        feTpoAgrico.TextMatrix(i, 1) = rsParam!TipoDesc & Space(75) & rsParam!nTipo
        feTpoAgrico.TextMatrix(i, 2) = rsParam!cSubTipo
        feTpoAgrico.TextMatrix(i, 3) = rsParam!nMin
        feTpoAgrico.TextMatrix(i, 4) = rsParam!nSubTipo
        rsParam.MoveNext
    Next i
End If
End Sub

Private Function ValidaDatos() As Boolean
Dim oParam As COMDCredito.DCOMParametro
Dim rsParam As ADODB.Recordset

Dim lnTipo As Long
Dim lsSubProd As String

Dim lnCantSubProd As Double

Set oParam = New COMDCredito.DCOMParametro
fnOpeRealiza = 1
fsCodSubProdActivar = ""

If feTpoAgrico.TextMatrix(1, 1) <> "" Then
    For i = 1 To feTpoAgrico.Rows - 1
        If feTpoAgrico.TextMatrix(i, 1) = "" Then
            MsgBox "Seleccione el Tipo  en la Fila " & i, vbInformation, "Aviso"
            ValidaDatos = False
            Exit Function
        End If

        If feTpoAgrico.TextMatrix(i, 2) = "" Then
            MsgBox "Ingrese la Descripción del Sub Tipo en la Fila " & i, vbInformation, "Aviso"
            ValidaDatos = False
            Exit Function
        End If
    Next i
    
    
    
    lnTipo = CLng(Trim(Right(feTpoAgrico.TextMatrix(feTpoAgrico.Row, 1), 10)))
    lsSubProd = Trim(feTpoAgrico.TextMatrix(feTpoAgrico.Row, 2))
    fsCodSubProd = Trim(feTpoAgrico.TextMatrix(feTpoAgrico.Row, 4))
    If lnTipo = 1 Then
        If feTpoAgrico.TextMatrix(feTpoAgrico.Row, 3) = "" Then
            feTpoAgrico.TextMatrix(feTpoAgrico.Row, 3) = "0"
        End If
    Else
        If feTpoAgrico.TextMatrix(feTpoAgrico.Row, 3) = "" Or feTpoAgrico.TextMatrix(feTpoAgrico.Row, 3) = "0" Then
            MsgBox "Favor de Ingresar Valor Mayor a 0 en la columna Minimo para Tipo Pecuario", vbInformation, "Aviso"
            ValidaDatos = False
            Exit Function
        End If
    End If
    
    Set rsParam = oParam.ObtenerParametrosAgro(lnTipo, lsSubProd, "0,1")
    
    If fsCodSubProd <> "" Then
        If Not (rsParam.EOF And rsParam.BOF) Then
            lnCantSubProd = oParam.ExisteCredAgrico(fsCodSubProd)
            If lnCantSubProd > 0 Then
                If MsgBox("Existe" & IIf(lnCantSubProd > 1, "n ", " ") & lnCantSubProd & " Crédito" & IIf(lnCantSubProd > 1, "s", "") & " con este sub tipo, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                    fnOpeRealiza = 2
                    fsCodSubProdActivar = fsCodSubProd
                Else
                    feTpoAgrico.TextMatrix(feTpoAgrico.Row, 1) = rsParam!TipoDesc & Space(75) & rsParam!nTipo
                    feTpoAgrico.TextMatrix(feTpoAgrico.Row, 2) = rsParam!cSubTipo
                    ValidaDatos = False
                    Exit Function
                End If
            Else
                fnOpeRealiza = 2
                fsCodSubProdActivar = fsCodSubProd
            End If
        Else
            Set rsParam = Nothing
            Set rsParam = oParam.ObtenerParametrosAgro(lnTipo, , , fsCodSubProd)
            
            If Not (rsParam.EOF And rsParam.BOF) Then
            lnCantSubProd = oParam.ExisteCredAgrico(fsCodSubProd)
                If lnCantSubProd > 0 Then
                    If MsgBox("Existe " & IIf(lnCantSubProd > 1, "n", "") & lnCantSubProd & " Crédito" & IIf(lnCantSubProd > 1, "s", "") & " con este sub tipo, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                        fnOpeRealiza = 2
                        fsCodSubProdActivar = fsCodSubProd
                    Else
                        feTpoAgrico.TextMatrix(feTpoAgrico.Row, 1) = rsParam!TipoDesc & Space(75) & rsParam!nTipo
                        feTpoAgrico.TextMatrix(feTpoAgrico.Row, 2) = rsParam!cSubTipo
                        ValidaDatos = False
                        Exit Function
                    End If
                Else
                    fnOpeRealiza = 2
                    fsCodSubProdActivar = fsCodSubProd
                End If
            End If
            
        End If
    Else
        If Not (rsParam.EOF And rsParam.BOF) Then
            If CInt(rsParam!nEstado) = 1 Then
                MsgBox "Registro se encuentra duplicado", vbInformation, "Aviso"
                feTpoAgrico.TextMatrix(feTpoAgrico.Row, 2) = ""
                ValidaDatos = False
                Exit Function
            Else
                If MsgBox("Registro se encuentra duplicado con SubTipo desactivado, Desea volver activarlo?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                    fnOpeRealiza = 2
                    fsCodSubProdActivar = Trim(rsParam!nSubTipo)
                Else
                    feTpoAgrico.TextMatrix(feTpoAgrico.Row, 2) = ""
                    ValidaDatos = False
                    Exit Function
                End If
            End If
        End If
    End If
End If
ValidaDatos = True
End Function
 
