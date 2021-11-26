VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMdiMain 
   BackColor       =   &H8000000C&
   ClientHeight    =   8955
   ClientLeft      =   1620
   ClientTop       =   2595
   ClientWidth     =   17745
   Icon            =   "frmMdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMdiMain.frx":030A
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pbxFondo 
      Align           =   1  'Align Top
      Height          =   8055
      Left            =   0
      Picture         =   "frmMdiMain.frx":C9F4
      ScaleHeight     =   7995
      ScaleWidth      =   17685
      TabIndex        =   2
      Top             =   660
      Width           =   17745
      Begin VB.Image Image1 
         Height          =   13500
         Left            =   0
         Picture         =   "frmMdiMain.frx":5DC70
         Top             =   0
         Width           =   21600
      End
   End
   Begin Sicmact.Usuario Usuario1 
      Left            =   0
      Top             =   960
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   480
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilsMain 
      Left            =   1080
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":84B29
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":85B7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":86BCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdiMain.frx":87C1F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17745
      _ExtentX        =   31300
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ilsMain"
      HotImageList    =   "ilsMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Operaciones"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "OpeCajero"
                  Text            =   "Operacione de Cajero"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "OpeCaja"
                  Text            =   "Operacione de Caja"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "OpeLog"
                  Text            =   "Operaciones de Logística"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Clientes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Calculadora"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mantenimiento Permiso"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar staMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8700
      Width           =   17745
      _ExtentX        =   31300
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   2999
            MinWidth        =   2999
            TextSave        =   "10:33 AM"
         EndProperty
      EndProperty
   End
   Begin VB.Menu M0100000000 
      Caption         =   "&Archivo"
      Index           =   0
      Begin VB.Menu M0101000000 
         Caption         =   "Configurar &Impresora"
         Index           =   0
      End
      Begin VB.Menu M0101000000 
         Caption         =   "Configurar Caracteres de Impresión"
         Index           =   1
      End
      Begin VB.Menu M0101000000 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu M0101000000 
         Caption         =   "&Salir"
         Index           =   3
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu M0200000000 
      Caption         =   "A&dministración"
      Index           =   0
      Begin VB.Menu M0201000000 
         Caption         =   "&Tipo de Cambio"
         Index           =   0
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Cobertura de Seguro por Agencia"
         Index           =   1
      End
   End
   Begin VB.Menu M0300000000 
      Caption         =   "Con&tabilidad"
      Index           =   0
      Begin VB.Menu M0301000000 
         Caption         =   "&Definiciones"
         Index           =   0
         Begin VB.Menu M0301010000 
            Caption         =   "Parámetros de Viáticos"
            Index           =   0
            Begin VB.Menu M0301010100 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0301010100 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Tabla de Documentos"
            Index           =   1
            Begin VB.Menu M0301010200 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0301010200 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Tabla de Impuestos"
            Index           =   2
            Begin VB.Menu M0301010300 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0301010300 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Tabla de Operaciones"
            Index           =   3
            Begin VB.Menu M0301010400 
               Caption         =   "Clasificación de Operaciones"
               Index           =   0
               Begin VB.Menu M0301010401 
                  Caption         =   "&Mantenimiento"
                  Index           =   0
               End
               Begin VB.Menu M0301010401 
                  Caption         =   "&Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu M0301010400 
               Caption         =   "Definición de Operaciones"
               Index           =   1
               Begin VB.Menu M0301010402 
                  Caption         =   "&Mantenimiento"
                  Index           =   0
               End
               Begin VB.Menu M0301010402 
                  Caption         =   "&Consulta"
                  Index           =   1
               End
            End
         End
         Begin VB.Menu M0301010000 
            Caption         =   "&Configuración de Reportes"
            Index           =   4
            Begin VB.Menu M0301010500 
               Caption         =   "Reporte de Cuentas Contables por Columnas"
               Index           =   0
            End
            Begin VB.Menu M0301010500 
               Caption         =   "Reportes en Base a Fórmulas"
               Index           =   1
            End
            Begin VB.Menu M0301010500 
               Caption         =   "Reportes de Gastos x Niveles"
               Index           =   2
            End
            Begin VB.Menu M0301010500 
               Caption         =   "Proyecciones Anuales de Gastos"
               Index           =   3
            End
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Pendientes &Históricas"
            Index           =   5
            Begin VB.Menu M0301010600 
               Caption         =   "Moneda &Nacional"
               Index           =   1
            End
            Begin VB.Menu M0301010600 
               Caption         =   "Moneda &Extranjera"
               Index           =   2
            End
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Indice &VAC"
            Index           =   6
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Tabla Valores del sistema"
            Index           =   7
            Begin VB.Menu M0301010700 
               Caption         =   "Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0301010700 
               Caption         =   "Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Tabla de porcentajes de gastos por agencias"
            Index           =   8
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Tabla de porcentajes de Seguros Patrimoniales por agencias"
            Index           =   9
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Parámetros de Posición Cambiaria"
            Index           =   10
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Parámetros de Límites y Alertas Tempranas"
            Index           =   11
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Estructura Contable"
         Index           =   1
         Begin VB.Menu M0301020000 
            Caption         =   "&Cuentas Contables"
            Index           =   0
            Begin VB.Menu M0301020100 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0301020100 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Cuentas Base de SBS"
            Index           =   1
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Objetos"
            Index           =   2
            Begin VB.Menu M0301020200 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0301020200 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Generación con Variables"
            Index           =   4
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Manteniento de Cuentas por Agencias (AyB)"
            Index           =   5
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Grupo Operaciones Adeudados"
            Index           =   6
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Operaciones"
         Index           =   2
         Begin VB.Menu M0301030000 
            Caption         =   "&Saldos Iniciales para Puesta en Producción"
            Index           =   1
         End
         Begin VB.Menu M0301030000 
            Caption         =   "&Registro de Asientos Manuales"
            Index           =   2
         End
         Begin VB.Menu M0301030000 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu M0301030000 
            Caption         =   "&Consultas"
            Index           =   4
            Begin VB.Menu M0301034000 
               Caption         =   "por &Asientos Contables"
               Index           =   1
            End
            Begin VB.Menu M0301034000 
               Caption         =   "por &Movimientos"
               Index           =   2
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "M&odificación de Asientos Contables"
            Index           =   5
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Modi&ficación de Datos no Contables"
            Index           =   6
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Extorno con Generación de Asientos"
            Index           =   7
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Extorno con Eliminación de Asientos "
            Index           =   8
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Asientos &Eliminados"
            Index           =   9
         End
         Begin VB.Menu M0301030000 
            Caption         =   "-"
            Index           =   10
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Transferencia de Cuentas en Saldos"
            Index           =   11
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Transferencia de Cuentas en Asientos"
            Index           =   12
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Transferencia de Asiento de Agencias"
            Index           =   13
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Pólizas de Seguros Patrimoniales"
            Index           =   14
            Begin VB.Menu M0301035000 
               Caption         =   "Registro"
               Index           =   1
            End
            Begin VB.Menu M0301035000 
               Caption         =   "Porcentajes"
               Index           =   2
            End
            Begin VB.Menu M0301035000 
               Caption         =   "Distribución de gastos"
               Index           =   3
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Fondo Seguro de Depósito"
            Index           =   15
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Intangibles"
            Index           =   16
            Begin VB.Menu M0301036000 
               Caption         =   "Registro de Intangibles"
               Index           =   1
            End
            Begin VB.Menu M0301036000 
               Caption         =   "Amortizaciòn de Intangibles"
               Index           =   2
            End
            Begin VB.Menu M0301036000 
               Caption         =   "Extorno de Amortizaciones de Intangibles"
               Index           =   3
            End
            Begin VB.Menu M0301036000 
               Caption         =   "Baja de Intangibles"
               Index           =   4
            End
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Provisiones"
         Index           =   3
      End
      Begin VB.Menu M0301000000 
         Caption         =   "A&justes"
         Index           =   4
         Begin VB.Menu M0301040000 
            Caption         =   "Ajuste de Tipo de Cambio"
            Index           =   0
         End
         Begin VB.Menu M0301040000 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Ajuste por &Inflacion"
            Index           =   2
            Begin VB.Menu M0301020300 
               Caption         =   "Indice de Precios al por Mayor"
               Index           =   0
               Begin VB.Menu M0301010403 
                  Caption         =   "&Mantenimiento"
                  Index           =   0
               End
               Begin VB.Menu M0301010403 
                  Caption         =   "&Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu M0301020300 
               Caption         =   "&Generación de Ajuste"
               Index           =   1
            End
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Ajuste por Inflación en Base a Detalle &Histórico"
            Index           =   3
            Begin VB.Menu M0301020400 
               Caption         =   "&Detalle Histórico"
               Index           =   0
            End
            Begin VB.Menu M0301020400 
               Caption         =   "&Generación de Ajuste"
               Index           =   1
            End
         End
         Begin VB.Menu M0301040000 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu M0301040000 
            Caption         =   "&Otros Ajustes"
            Index           =   5
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Asientos Fractal"
            Index           =   6
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Análisis  de &Pendientes"
         Index           =   5
         Begin VB.Menu M0301140000 
            Caption         =   "&Definición de Cuentas Pendientes"
            Index           =   0
         End
         Begin VB.Menu M0301140000 
            Caption         =   "&Operaciones con Pendientes"
            Index           =   1
         End
         Begin VB.Menu M0301140000 
            Caption         =   "&Asignación de Referencias de Asientos"
            Index           =   2
         End
         Begin VB.Menu M0301140000 
            Caption         =   "&Reporte de Pendientes"
            Index           =   4
         End
         Begin VB.Menu M0301140000 
            Caption         =   "Reporte de Pendientes Detalle"
            Index           =   5
         End
         Begin VB.Menu M0301140000 
            Caption         =   "Consulta de Operaciones Tramite Negocio"
            Index           =   6
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Balances"
         Index           =   7
         Begin VB.Menu M0301050000 
            Caption         =   "&Proceso de &Validación"
            Index           =   0
         End
         Begin VB.Menu M0301050000 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Balance &Histórico"
            Index           =   2
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Balance &Ajustado"
            Index           =   3
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Balance General Forma A y B"
            Index           =   4
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Balance &Sectorial"
            Index           =   5
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Procesos"
         Index           =   8
         Begin VB.Menu M0301060000 
            Caption         =   "&Actualización de Saldos"
            Index           =   0
         End
         Begin VB.Menu M0301060000 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu M0301060000 
            Caption         =   "Cierre Contable &Diario"
            Index           =   2
         End
         Begin VB.Menu M0301060000 
            Caption         =   "Cierre Contable &Mensual"
            Index           =   3
         End
         Begin VB.Menu M0301060000 
            Caption         =   "Cierre Contable &Anual"
            Index           =   4
         End
         Begin VB.Menu M0301060000 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu M0301060000 
            Caption         =   "Asignación de IGV"
            Index           =   6
         End
         Begin VB.Menu M0301060000 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu M0301060000 
            Caption         =   "Extorno Cierre"
            Index           =   8
         End
         Begin VB.Menu M0301060000 
            Caption         =   "Comisión Banco de la Nación"
            Index           =   9
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Reportes"
         Index           =   9
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Instituciones"
         Index           =   10
         Begin VB.Menu M0301070000 
            Caption         =   "&SUNAT"
            Index           =   0
            Begin VB.Menu M0301020500 
               Caption         =   "COA &Transferencia de Comprobantes"
               Index           =   0
            End
            Begin VB.Menu M0301020500 
               Caption         =   "Contribuyentes no &hábidos"
               Index           =   1
            End
            Begin VB.Menu M0301020500 
               Caption         =   "Renta de Cuarta Categoria"
               Index           =   2
            End
            Begin VB.Menu M0301020500 
               Caption         =   "Programa de Libros Electrónicos"
               Index           =   3
            End
            Begin VB.Menu M0301020500 
               Caption         =   "COA Generacion Archivo PVS"
               Index           =   4
            End
         End
         Begin VB.Menu M0301070000 
            Caption         =   "S&BS"
            Index           =   1
            Begin VB.Menu M0301020600 
               Caption         =   "BCIENT para Sucave"
               Index           =   0
            End
         End
         Begin VB.Menu M0301070000 
            Caption         =   "&FONCODES"
            Index           =   2
            Begin VB.Menu M0301020700 
               Caption         =   "Abono/Cargo por Operaciones"
               Index           =   0
            End
         End
         Begin VB.Menu M0301070000 
            Caption         =   "&BCR"
            Index           =   3
            Begin VB.Menu M0301020900 
               Caption         =   "Balance Sectorial"
               Index           =   0
            End
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Ane&xos SBS"
         Index           =   12
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Re&portes SBS"
         Index           =   13
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Otras Operaciones"
         Index           =   14
         Begin VB.Menu M0301090000 
            Caption         =   "Entidades con Operaciones Reciprocas"
            Index           =   0
         End
         Begin VB.Menu M0301090000 
            Caption         =   "Calculo resumido de ventas de Cred. Pignoraticios"
            Index           =   1
         End
         Begin VB.Menu M0301090000 
            Caption         =   "Resumen de Gastos Judiciales por Agencia"
            Index           =   2
         End
         Begin VB.Menu M0301090000 
            Caption         =   "Pagos Adelantados"
            Index           =   3
         End
         Begin VB.Menu M0301090000 
            Caption         =   "Detracciones"
            Index           =   4
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Contingencias"
         Index           =   16
         Begin VB.Menu M0301100000 
            Caption         =   "Registro"
            Index           =   0
            Begin VB.Menu M0301020800 
               Caption         =   "Activo Contingente"
               Enabled         =   0   'False
               Index           =   1
               Visible         =   0   'False
            End
            Begin VB.Menu M0301020800 
               Caption         =   "Pasivo Contingente"
               Index           =   2
            End
         End
         Begin VB.Menu M0301100000 
            Caption         =   "Extorno"
            Enabled         =   0   'False
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu M0301100000 
            Caption         =   "Mant. Tipo de Monto"
            Index           =   2
         End
         Begin VB.Menu M0301100000 
            Caption         =   "Consulta"
            Index           =   3
         End
         Begin VB.Menu M0301100000 
            Caption         =   "Seguimiento"
            Enabled         =   0   'False
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu M0301100000 
            Caption         =   "Extorno Liberacion"
            Enabled         =   0   'False
            Index           =   5
            Visible         =   0   'False
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Análisis de Gastos"
         Index           =   18
         Begin VB.Menu M0301200000 
            Caption         =   "Plantilla de Gastos por Movilidad"
            Index           =   0
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "-"
         Index           =   19
      End
      Begin VB.Menu M0301000000 
         Caption         =   "ONP/AFP"
         Index           =   20
         Begin VB.Menu M0301300000 
            Caption         =   "Porcentajes de Comisiones y Seguro AFP"
            Index           =   0
         End
         Begin VB.Menu M0301300000 
            Caption         =   "Datos ONP/AFP de Personas"
            Index           =   1
         End
      End
   End
   Begin VB.Menu M0400000000 
      Caption         =   "&Finanzas"
      Index           =   0
      Begin VB.Menu M0401000000 
         Caption         =   "&Movimientos"
         Index           =   0
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Reportes"
         Index           =   2
      End
      Begin VB.Menu M0401000000 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Rece&pción de Cheques"
         Index           =   6
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Consolidación de &Estadísticas"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Consolidación de &Billetaje"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu M0401000000 
         Caption         =   "-"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Cierre Diario de Caja"
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu M0401000000 
         Caption         =   "-"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Control de Paquetes de Adeudados"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Mantenimiento de linea"
         Index           =   13
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Reporte FONDEMI"
         Index           =   14
      End
      Begin VB.Menu M0401000000 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Gestión Financiera"
         Index           =   16
         Begin VB.Menu M0401100000 
            Caption         =   "Registro Flujo Caja Proyectado"
            Index           =   0
         End
         Begin VB.Menu M0401100000 
            Caption         =   "Proyecciones de Ahorros por Agencia"
            Index           =   1
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Reporte FONCODES"
         Index           =   17
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Registro de Operaciones"
         Index           =   18
         Begin VB.Menu M0401200000 
            Caption         =   "Cheques"
            Index           =   0
            Begin VB.Menu M0401201000 
               Caption         =   "Registro Talonario"
               Index           =   0
            End
            Begin VB.Menu M0401201000 
               Caption         =   "Mantenimiento de Cheques"
               Index           =   1
            End
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Encaje Legal Diario"
         Index           =   19
         Begin VB.Menu M0401300000 
            Caption         =   "Generar Formato"
            Index           =   0
         End
         Begin VB.Menu M0401300000 
            Caption         =   "Calendario de Proyecciones"
            Index           =   1
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Lineas de Crédito"
         Index           =   20
         Begin VB.Menu M0401400000 
            Caption         =   "Listar líneas de crédito"
            Index           =   0
         End
         Begin VB.Menu M0401400000 
            Caption         =   "Priorizar líneas de crédito"
            Index           =   1
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Reporte RESPONSABILITY"
         Index           =   21
      End
   End
   Begin VB.Menu M0600000000 
      Caption         =   "Herra&mientas"
      Index           =   0
      Begin VB.Menu M0601000000 
         Caption         =   "Spooler de Impresión"
         Index           =   1
      End
      Begin VB.Menu M0601000000 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Movimientos Diarios"
         Index           =   11
      End
   End
   Begin VB.Menu M0700000000 
      Caption         =   "Perso&nas"
      Index           =   0
      Begin VB.Menu M0701000000 
         Caption         =   "&Personas"
         Index           =   0
         Begin VB.Menu M0701010000 
            Caption         =   "&Consulta"
            Index           =   2
         End
      End
      Begin VB.Menu M0701000000 
         Caption         =   "&Instituciones Financieras"
         Index           =   1
         Begin VB.Menu M0701020000 
            Caption         =   "Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M0701020000 
            Caption         =   "Consulta"
            Index           =   1
         End
      End
      Begin VB.Menu M0701000000 
         Caption         =   "Grupos Economicos"
         Index           =   2
         Begin VB.Menu M0701030000 
            Caption         =   "Mantenimiento"
            Index           =   0
         End
      End
   End
   Begin VB.Menu M0800000000 
      Caption         =   "A&yuda"
      Index           =   0
      Begin VB.Menu M0801000000 
         Caption         =   "&Contenido"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu M0801000000 
         Caption         =   "&Indice"
         Index           =   1
      End
      Begin VB.Menu M0801000000 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu M0801000000 
         Caption         =   "&Soporte Técnico"
         Index           =   3
      End
      Begin VB.Menu M0801000000 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu M0801000000 
         Caption         =   "&Acerca del Sistema..."
         Index           =   5
      End
   End
   Begin VB.Menu M1000000000 
      Caption         =   "&Seguridad"
      Index           =   0
      Begin VB.Menu M1001000000 
         Caption         =   "&Permisos"
         Index           =   0
      End
   End
   Begin VB.Menu M1100000000 
      Caption         =   "&Banco Pagador"
      Index           =   0
      Begin VB.Menu M1101000000 
         Caption         =   "Confirmación de Abonos"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objPista As COMManejador.Pista 'ARLO20170511

Private Sub M0101000000_Click(Index As Integer)
Select Case Index
    Case 0
        frmImpresora.Show 1
    Case 1
        frmCaracImpresion.Show 1
    Case 3
        If MsgBox("Esta Seguro que Desea Salir del Sistema ?", vbInformation + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
            Call SalirSICMACMAdmnstrativo 'AGREGADO POR ARLO
            End
        End If
End Select
End Sub

Private Sub M0201000000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmMantTipoCambio.Show 1
        Case 1
            frmCoberturaSeguroAgencia.Show 1
        Case 5
            frmCierreContDia.inicio True
       
    End Select
End Sub

'***Agregado por ELRO el 20111026, según Acta 277-2011/TI-D
Private Sub M0202000000_Click(Index As Integer)
    Select Case Index
        Case 0: frmCoberturaSeguroAgencia.Show 1
    End Select
End Sub
'***Fin Agregado por ELRO**********************************

Private Sub M0301000000_Click(Index As Integer)
    Select Case Index
        Case 3
            frmOperaciones.inicio Left(gContProvisionProveeMN, 2) & "_" & Mid(gContProvisionProveeMN, 4, 1), True
        Case 9
            frmReportes.inicio "76"
        Case 12
            frmReportes.inicio "77", True
        Case 13
            frmReportes.inicio "78", True

    End Select
End Sub

Private Sub M0301010000_Click(Index As Integer)
Select Case Index
    Case 6: frmMntIndiceVac.Show 1, Me
    'ALPA 20091109**********************************************
    Case 8: frmAgenciaPorcentajeGastos.Show 1, Me
    '***********************************************************
    '*** PEAC 20100907
    Case 9: frmAgenciaPorcentajeSeguPatri.Show 1, Me
    Case 10: frmLimitePosCam.Show 1 'MIOL 20130722, SEGUN RQ13387 - ERS088-2013
    Case 11: frmAlertasTempranas.Show 1 'VAPA ERS002-2017
End Select
End Sub

Private Sub M0301010100_Click(Index As Integer)
    Select Case Index
       Case 0: frmMntParaViaticos.inicio False
       Case 1: frmMntParaViaticos.inicio True
    End Select
End Sub

Private Sub M0301010200_Click(Index As Integer)
    Select Case Index
       Case 0:  frmMntDocumento.inicio False
       Case 1:  frmMntDocumento.inicio True
    End Select
End Sub

Private Sub M0301010300_Click(Index As Integer)
    Select Case Index
       Case 0: frmMntImpuestos.inicio False
       Case 1: frmMntImpuestos.inicio True
    End Select
End Sub

Private Sub M0301010401_Click(Index As Integer)
    Select Case Index
        Case 0:  frmMntOperacionGru.inicio False
        Case 1:  frmMntOperacionGru.inicio True
    End Select
End Sub

Private Sub M0301010402_Click(Index As Integer)
    Select Case Index
       Case 0:  frmMntOperacion.inicio False
       Case 1:  frmMntOperacion.inicio True
    End Select
End Sub

Private Sub M0301010403_Click(Index As Integer)
    Select Case Index
        Case 0:  frmMntIPM.inicio False
        Case 1:  frmMntIPM.inicio True
    End Select
End Sub

Private Sub M0301010500_Click(Index As Integer)
    Select Case Index
       Case 0: frmMntRepColumnas.Show 0, Me
       Case 1: frmMntRepFormula.Show 0, Me
       Case 2: frmConfigRepGastoxNiveles.Show 1 'MIOL 20130528, SEGUN RQ13286
       Case 3: frmRegProyAnual.Show 1 'MIOL 20130528, SEGUN RQ13287
    End Select
End Sub

Private Sub M0301010600_Click(Index As Integer)
    gsOpeCod = CodigoOperacion(gOpePendMntHistoric, Index)
   frmAnalisisCtaHisto.Show 1
End Sub

Private Sub M0301010700_Click(Index As Integer)
    Select Case Index
       Case 0: frmConstSistema.inicio False
       Case 1: frmConstSistema.inicio True
    End Select
End Sub

Private Sub M0301020000_Click(Index As Integer)
    Select Case Index
        Case 1: frmMntCtasContBase.inicio False, "CtaContBase"
        Case 4: frmContabManVar.Show 1
        Case 5: frmContCtasxAge.Show 1
        'ALPA 20110829***************
        Case 6: frmGrupoOperacionesContable.Show 1
        '*****************************
    End Select
End Sub

Private Sub M0301020100_Click(Index As Integer)
    Select Case Index
       Case 0: frmMntCtasContab.Show 0, Me
       Case 1: frmMntCtasContabN.inicio , True
    End Select
End Sub

Private Sub M0301020200_Click(Index As Integer)
    Select Case Index
       Case 0: frmMntObjetos.inicio False
       Case 1: frmMntObjetos.inicio True
    End Select
End Sub

Private Sub M0301020300_Click(Index As Integer)
    Select Case Index
        Case 1:  gsOpeCod = gContAjusteInflaIngre
            frmAjusteInflacion.Show 0, Me
            
    End Select
End Sub

Private Sub M0301020400_Click(Index As Integer)
    Select Case Index
        Case 0:  frmMntAjusteInfla.inicio False
        Case 1:  gsOpeCod = gContAjusteInflaHisto
            frmAjusteInflaHisto.Show 1
    End Select
End Sub

Private Sub M0301020500_Click(Index As Integer)
Select Case Index
   Case 0:     frmTransferenciaCoa.Show 0, Me
   Case 1:     frmMntContribuyeNoHabido.Show 0, Me
   Case 2:      frmRentaCuart.Show 0, Me
   Case 3: frmPlanillasElectronicas.Show 1 'EJVG20130325
   Case 4:      frmTransferenciaPVS.Show 1 '*** PEAC 20130612
End Select
End Sub

Private Sub M0301020600_Click(Index As Integer)
    frmSucaveBCient.Show 0, Me
End Sub

Private Sub M0301020700_Click(Index As Integer)
Select Case Index
    Case 0
        frmFoncodesNotas.Show 1
        'frmReporteFoncodes.Show 1
End Select
End Sub

'** Juez 20120614 *****************************************
Private Sub M0301020800_Click(Index As Integer)
    Select Case Index
       'Case 1: frmContingenciaReg.RegistrarContingencia gActivoContingente
       Case 2: frmContingenciaReg.RegistrarContingencia gPasivoContingente
    End Select
End Sub
'** End Juez **********************************************
'PASI20170727 ********************************************
Private Sub M0301020900_Click(Index As Integer)
    Select Case Index
        Case 0: frmFTPBSI_Reportxt.Show 0, Me
    End Select
End Sub
'PASI END*************************************************

Private Sub M0301030000_Click(Index As Integer)
    Select Case Index
       Case 1:  frmMntSaldosIni.inicio False
       Case 2:  frmOperaciones.inicio "70[12]1", True
       Case 3:
       Case 4:
       Case 5:  frmAsientoModificaCont.inicio False, False
       Case 6:  frmAsientoModificaCont.inicio False, False, False, 1, True
       Case 7:  frmAsientoModificaCont.inicio False, True, False
       Case 8:  frmAsientoModificaCont.inicio False, True, True
       Case 9:  frmAsientoModificaCont.inicio True, False, , 2
       Case 10:
       Case 11:  frmMntSaldosIni.inicio True
       Case 12: frmMntSaldosMov.Show 1
       Case 13: frmContabMigracion.Show 1
       Case 14:
       Case 15: frmFondoSeguroDeposito.Show 1 '*** PEAC 20101005
       
    End Select
End Sub

Private Sub M0301034000_Click(Index As Integer)
Select Case Index
    Case 1:  frmAsientoModificaCont.inicio True, False
    Case 2:  frmMovimientoConsulta.Show 1
End Select
End Sub

Private Sub M0301035000_Click(Index As Integer)
    Select Case Index
        Case 1:  frmPolizaSeguroPatri.Show 1
        Case 2:  frmPolizasListaTasasOK.Show 1
        Case 3:  frmPolizaSeguPatriDistri.Show 1
    End Select
End Sub

Private Sub M0301040000_Click(Index As Integer)
    Select Case Index
       Case 0:  gsOpeCod = gContAjusteTipoCambio
                frmAjusteTipCambio.Show 0, Me
        Case 5: frmOperaciones.inicio Left(gContAjReclasiCartera, 2) & "_" & Mid(gContAjReclasiCartera, 4, 1), True
        Case 6: frmAsientosFractal.Show 1 'ALPA 20120403
    End Select

End Sub

Private Sub M0301050000_Click(Index As Integer)
    Select Case Index
       Case 0: frmBalanceValida.Show , Me
       Case 2: frmBalanceHisto.inicio 1
       Case 3: frmBalanceHisto.inicio 2
       Case 4: frmBalanceGeneral.Show , Me
       Case 5: frmBalanceSec.Show , Me
       'Case 7: frmExtornoCierre.Show , Me
    End Select
End Sub

Private Sub M0301060000_Click(Index As Integer)
    Select Case Index
        Case 0: frmCierreContDia.inicio True
        Case 2: frmCierreContDia.inicio False
        Case 3: frmCierreContMes.Show 1
        Case 4: gsOpeCod = gContCierreAnual
            frmCierreContAnual.Show 1
        Case 6:
            frmIGVReversion.Show 1
        Case 8:
            frmExtornoCierre.Show 1
        Case 9: frmAsntoComisionBcoNac.Show 1 '*** PEAC 20120806

    End Select
End Sub

Private Sub M0301080000_Click(Index As Integer)
    If Index = 0 Then
        frmRepConsolAho.Show 1
    ElseIf Index = 1 Then
        frmRepRiesgos.Show 1
    End If
    
End Sub

'*** PEAC 20110602
Private Sub M0301090000_Click(Index As Integer)
Select Case Index
    Case 0
        frmEntidadesOpeRecipro.Show 1
    Case 1
        frmResumenCredPignoRecup.Show 1
    Case 2
        frmResumenGastosJudiciales.Show 1
'***Agregado por ELRO el 20111215, según Acta N°323-2011/TI-D
    Case 3
        frmPagosAdelantados.Show 1
'***Fin Agregado por ELRO************************************
'***Agregado por ELRO el 20120521, según Acta N°328-2011/TI-D
    Case 4
        frmMntConceptoDetracion.Show 1
'***Fin Agregado por ELRO************************************

End Select
End Sub

'** Juez 20120614 *****************************
Private Sub M0301100000_Click(Index As Integer)
    Select Case Index
       'Case 1: frmContingenciaCons.Extorno 1
       Case 2: frmContingTipoMonto.Show 1
       Case 3: frmContingenciaCons.Consulta 1
       'Case 4: frmContingSeguimiento.Show 1
       'Case 5: frmContingLiberarExt.Show 1
    End Select
End Sub
'** End Juez **********************************

Private Sub M0301140000_Click(Index As Integer)
Select Case Index
    Case 0
        frmMntCtaContPend.Show 1
    Case 1
        frmOperaciones.inicio "74"
    Case 2
        frmAsignaReferencia.Show 1
    Case 4
        frmRepAnaCtas.Show 1
    Case 5
        frmRepAnaCtasDH.Show 1
    Case 6
        frmMuestraOpeTramNeg.Show 1
End Select
End Sub
'***Agregado por ELRO el 20130128, según OYP-RFC126-2012
Private Sub M0301200000_Click(Index As Integer)
frmReporteGastosPorMovilidad.Show 1
End Sub
'***Fin Agregado por ELRO el 20130128*******************
'EJVG20140805 ***
Private Sub M0301300000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmProveedorTasasAFPSistemaPension.Show 1
        Case 1
            frmProveedorRegSistemaPensionLista.Show 1
    End Select
End Sub
'END EJVG *******
Private Sub M0401000000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmOperaciones.inicio "4[0-5]"
        Case 1
            frmOperaciones.inicio "9[0-5]%"
        Case 2
            frmReportes.inicio "4[6-9]"
'        Case 4
'            frmMantOpCaja.Show 1
        Case 6
            frmTransCheque.Show 1
        Case 7
            frmConsEstadisticas.Show 1
        Case 8
            frmConsBilletaje.Show 1
        Case 10
            frmACGCierreDiario.Show 1
        Case 12
            frmMntCredSaldosAdeudo.Show 1
        'ALPA 20111028***********
        Case 13
            frmMntLineasCredito.Show 1
        Case 14
            frmCarteraFondemi.Show 1
        '************************
        'MIOL 20121105, SEGUN RFC103- 2012
        Case 17
            frmReporteEstadFoncodes.Show 1
        'END MIOL ************************
        'PASIERS0872014*******************************
        Case 21
            frmRepResponsability.Show 1
        'END PASI************************************
        
    End Select
    
End Sub
'EJVG20121022 ***
Private Sub M0401100000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmFlujoCajaProyectado.Show 1
        Case 1
            frmAhorrosProyecccionesPorAgencia.Show 1 'FRHU20140118 RQ13825
    End Select
End Sub

Private Sub M0401201000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmChequeTalonario.Show 1
        Case 1
            frmChequeMantenimiento.Show 1
    End Select
End Sub

Private Sub M0401300000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmParaEncDiario.Show 1
        Case 1 '****Agregado por PASI TI-ERS0012014 20140322
            frmEncDiarioCalenProy.Show 1
    End Select
End Sub
'ALPA 20150215************************************
Private Sub M0401400000_Click(Index As Integer)
    Select Case Index
        Case 0
            FrmCredLineaCreditoLista.Show 1
        Case 1
            FrmCredLineaCreditoPriorizar.Show 1
    End Select
End Sub
'************************************************
'END EJVG *******
Private Sub M0601000000_Click(Index As Integer)
    Select Case Index
        Case 1
            frmSpooler.Show 1, Me
'        Case 9
'            frmMntBienesNavidad.Show 1
'        Case 10
'           ' FrmVerAsiento.Show vbModal
        Case 11
            frmAACGValidaMovs.Show vbModal
    End Select
End Sub

Private Sub M0601010000_Click(Index As Integer)

End Sub

Private Sub M0701030000_Click(Index As Integer)
    If Index = 0 Then
        frmPersEcoGruRel1.Show 1
    End If
End Sub

Private Sub M0801000000_Click(Index As Integer)
    Select Case Index
        Case 0
            With cDialog
                  .HelpFile = App.path & "\Ayuda\" & "sicmact.hlp"
                  .HelpCommand = &HB Or cdlHelpSetContents
                  .ShowHelp
            End With
        Case 1
            With cDialog
                .HelpFile = App.path & "\Ayuda\" & "sicmact.hlp"
                .HelpCommand = cdlHelpKey
                .ShowHelp
                
            End With
    End Select
End Sub

Private Sub M0701010000_Click(Index As Integer)
    'Persona
    Select Case Index
        'Case 0 'Registro               RECO20140311 ERS160-2013 COMENTADO
            'frmPersona.Registrar
        'Case 1 'mantenimiento          RECO20140311 ERS160-2013 COMENTADO
            'frmPersona.Mantenimeinto
        Case 2 'Consulta
            frmPersona.Consultar
    End Select
        
End Sub

Private Sub M0701020000_Click(Index As Integer)
    
    'Instituciones Financieras
    Select Case Index
        Case 0
            frmMntInstFinanc.InicioActualizar
        Case 1
            frmMntInstFinanc.InicioConsulta
    End Select
End Sub
Private Sub M0901000000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmAnxEncajeBCR.Show 1
        Case 1
            'frmArchivos.Show 1
            frmCajaGenRemCheques.Show 1
        Case 2
            
        Case 3
            'frmCGBancosEjemplo.Show 1
            frmGeneraDataProyFinan.Show 1
        Case 4
            
            
    End Select
End Sub

Private Sub M1001000000_Click(Index As Integer)
    Select Case Index
        Case 0
            'frmMantPermisos.Show 1
            frmMantPermisosNuevo.Show 1 'WIOR 20130201
        Case 1
            
    End Select
End Sub

'AMDO 20131210 Banco Pagador ******
Private Sub M1101000000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmBancoPagadorConfirmacionAbonos.Show 1
    End Select
End Sub
'END AMDO *************************

Private Sub MDIForm_Load()
    Me.Usuario1.inicio gsCodUser
    staMain.Panels(1).Text = "Sistema Listo"
    CargaVarSistema True
    
    '->***** LUCV20190323, Según RO-1000373
    'Quita el borde de los dos controles
    Image1.BorderStyle = 0
    pbxFondo.BorderStyle = 0
    '->***** Fin LUCV20190323
        
    '**Modificado por DAOR 20100406, Control de versión ********************************
    'frmMdiMain.Caption = "SICMAC - MODULO FINANCIERO " & Space(20) & UCase(gsCodUser) & Space(7) & gsServerName & "\" & gsDBName & Space(5) & Format(gdFecSis, "dd/mm/yyyy")
    frmMdiMain.Caption = "SICMACM FINANCIERO " & space(20) & UCase(gsCodUser) & space(7) & gsServerName & "\" & gsDBName & space(5) & Format(gdFecSis, "dd/mm/yyyy")
    frmMdiMain.Caption = frmMdiMain.Caption & space(5) & " - Versión: " & Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & "." & Format(App.Revision, "#0") & "-" & gsFechaVersion 'Cambiar la fecha cada vez que se compila
    '********************************************************************

    frmMdiMain.staMain.Panels(2).Text = Format(gdFecSis, "dddd - dd - mmmm - yyyy") & space(3) & Format(Time, "hh:mm AMPM")
    GetTipCambio gdFecSis
    gsCodAge = Usuario1.CodAgeAct
    gsCodArea = Usuario1.cAreaCodAct
    CentraForm Me
    frmAdeuAlertaVencimientoPago.Show 1 'ALPA 20150910
    CargaMensajes 'WIOR 20130826
    'If Len(Usuario1.CodAgeAct) = 2 Then
    '    gsCodAge = gsCodCMAC & gsCodAge
    'End If
 '********* Temporalmente ***********
'    Dim R As New ADODB.Recordset
'    Dim oConec As DConecta
'    Dim sSql As String
'    Dim Y As Integer
'
'    gsRutaIcono = "\cm.ico"
'    'Timer1.Enabled = False
'    'NroRegOpe = 1628
'    'gdFecSis = Date
'    'Vusuario = "ARCV"
'    'gsCodAge = "01"
'    'gsCodCMAC = "108"
'
'    'Habilita Permiso para Operaciones y Reportes
'    Set oConec = New DConecta
'    oConec.AbreConexion
'    sSql = "Select cOpeCod,cOpeDesc,cOpeVisible,nOpeNiv,cOpeGruCod,cUltimaActualizacion from OpeTpo WHERE cOpeVisible ='1' Order by cOpeCod"
'    Set R = oConec.CargaRecordSet(sSql)
'    Y = 0
'    Do While Not R.EOF
'        Y = Y + 1
'        MatOperac(Y - 1, 0) = R!cOpeCod
'        MatOperac(Y - 1, 1) = R!cOpeDesc
'        MatOperac(Y - 1, 2) = IIf(IsNull(R!cOpeGruCod), "", R!cOpeGruCod)
'        MatOperac(Y - 1, 3) = R!cOpeVisible
'        MatOperac(Y - 1, 4) = R!nOpeNiv
'        R.MoveNext
'    Loop
'    NroRegOpe = Y
'    oConec.CierraConexion
'
 '************************************
 
 '->***** LUCV20190323, Según RO-1000373
    If VerificaGrupoMantenimientoUsuarios Then
        tlbMain.Buttons.Item(4).Visible = True
        tlbMain.Buttons.Item(4).Enabled = True
    Else
        tlbMain.Buttons.Item(4).Visible = False
        tlbMain.Buttons.Item(4).Enabled = False
    End If
 '<-***** Fin LUCV20190323
 
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'If MsgBox("Esta Seguro que Desea Salir del Sistema ?", vbInformation + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then 'COMENTADO POR ARLO20170511
    If MsgBox("Esta Seguro que Desea Salir del Sistema ?", vbInformation + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
        'Cancel = 1 'COMENTADO POR ARLO20170511
        Call SalirSICMACMAdmnstrativo 'AGREGADO POR ARLO 20170511
    End If
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        'frmSeleOpe.Show 1
    Case 2
            
    Case 3
    '->***** LUCV20190323, Según RO-1000373
        SCalculadora
    Case 4
        frmMantPermisosNuevo.Show 1
    '<-***** Fin LUCV20190323
    
End Select
Exit Sub
ErrorBoton:
End Sub
'WIOR 20130826 *************************************************************
Private Sub CargaMensajes()
Dim oSeg As UAcceso
Dim rsSeg As ADODB.Recordset

Set oSeg = New UAcceso
Set rsSeg = oSeg.ObtenerMensajeSeguridad

If Not (rsSeg.EOF And rsSeg.BOF) Then
    frmSegMensajeMostrar.inicio (Trim(rsSeg!cMensaje))
End If

Set oSeg = Nothing
Set rsSeg = Nothing
End Sub
'WIOR FIN *******************************************************************

'PASI 20140318************************************
Private Sub M0301036000_Click(Index As Integer)
    Select Case Index
        Case 1: frmIntangibleActivar.Show 1
        Case 2: frmIntangibleAmortizacion.Show 1
        Case 3: frmIntangibleExtorno.Show 1
        Case 4: frmIntangibleBaja.Show 1
    End Select
End Sub
'PASI FIN*****************************************

'ARLO 20170511
Sub SalirSICMACMAdmnstrativo()
    Set objPista = New COMManejador.Pista 'LogPistaIngresoSistema
    Call objPista.InsertarPista(gIngresarSalirSistema, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, TiposAccionesPistas.gSalirSistema, "Salida del Sicmac Financiero " & " Versión: " & Format(App.Major, "#0") & "." & Format(App.Minor, "#0") & "." & Format(App.Revision, "#0") & "-" & gsFechaVersion)
    Set objPista = Nothing
End Sub

'->***** LUCV20190323, Según RO-1000373
Public Sub SCalculadora()
    Dim Valor, I
    Valor = Shell("calc.exe", 1)  ' Ejecuta la Calculadora.
End Sub

Private Function VerificaGrupoMantenimientoUsuarios() As Boolean
    Dim oCons As NConstSistemas
    Set oCons = New NConstSistemas
    Dim sGrupoAutorizado As String
    Dim nGrupoTmp1 As String
    Dim nGrupoTmp2 As String
    Dim I As Integer
    Dim j As Integer
            
    sGrupoAutorizado = oCons.LeeConstSistema(519)
    VerificaGrupoMantenimientoUsuarios = False
    For I = 1 To Len(sGrupoAutorizado)
        If Not Mid(sGrupoAutorizado, I, 1) = "," Then
            nGrupoTmp1 = nGrupoTmp1 & Mid(sGrupoAutorizado, I, 1)
        Else
            For j = 1 To Len(gsGrupoUsu)
                If Not Mid(gsGrupoUsu, j, 1) = "," Then
                    nGrupoTmp2 = nGrupoTmp2 & Mid(gsGrupoUsu, j, 1)
                Else
                    If nGrupoTmp1 = nGrupoTmp2 Then
                        VerificaGrupoMantenimientoUsuarios = True
                        Exit Function
                    End If
                    nGrupoTmp2 = ""
                End If
            Next
            nGrupoTmp1 = ""
        End If
    Next
End Function
  
Private Sub MDIForm_Resize()
    'Posiciona el PictureBox que es el contenedor al ancho y alto que tenga el form mdi
    On Error Resume Next
    Dim ImageWidth As Single
    Dim ImageHeight As Single

    If WindowState = vbMaximized Then
        pbxFondo.Visible = False
        pbxFondo.AutoRedraw = True
        pbxFondo.Height = Me.ScaleHeight
        pbxFondo.BorderStyle = 0
    
        ImageWidth = pbxFondo.Picture.Width * 0.566893424036281
        ImageHeight = pbxFondo.Picture.Height * 0.566893424036281
    
        pbxFondo.PaintPicture pbxFondo.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, ImageWidth, ImageHeight
        Set Me.Picture = pbxFondo.Image
        pbxFondo.Refresh
    Else
        pbxFondo.BorderStyle = 0
        pbxFondo.Visible = False
        pbxFondo.AutoRedraw = True
        pbxFondo.PaintPicture pbxFondo.Picture, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, ImageWidth, ImageHeight
        pbxFondo.Move 0, 20, Me.Width, Me.Height
        Set Me.Picture = pbxFondo.Image
        pbxFondo.Refresh
    End If
End Sub
Private Sub pbxFondo_Resize()
    Dim Pos_x As Single
    Dim Pos_y As Single
  
    Pos_x = (pbxFondo.Width - Image1.Width) / 2
    Pos_y = (pbxFondo.Height - Image1.Height) / 2
    
    'Posiciona el control Image en el centro del Picturebox
    Image1.Move Pos_x, Pos_y - 680
End Sub
'<-***** Fin LUCV20190323

