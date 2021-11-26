VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDISicmact 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema del Negocio"
   ClientHeight    =   7305
   ClientLeft      =   1560
   ClientTop       =   2190
   ClientWidth     =   11400
   Icon            =   "mdisicmact2.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1217
      ButtonWidth     =   1773
      ButtonHeight    =   1058
      _Version        =   327682
      BorderStyle     =   1
      Begin VB.TextBox txtEstado3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   240
         Left            =   6255
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Pendientes"
         Top             =   225
         Width           =   2835
      End
      Begin VB.TextBox txtEstado2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Rechazadas"
         Top             =   225
         Width           =   2835
      End
      Begin VB.CommandButton cmdVer 
         Caption         =   "Consultar OP"
         Height          =   375
         Left            =   9105
         TabIndex        =   3
         Top             =   165
         Width           =   1275
      End
      Begin VB.TextBox txtEstado1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   240
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Aprobadas"
         Top             =   225
         Width           =   2835
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   2625
      Top             =   1815
   End
   Begin SICMACT.Usuario Usuario 
      Left            =   1530
      Top             =   3465
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin MSComctlLib.StatusBar SBBarra 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   7080
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   397
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11994
            MinWidth        =   11994
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "19/04/2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "04:17 PM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "NÚM"
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
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu M0101000000 
         Caption         =   "&Salir"
         Index           =   2
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu M0200000000 
      Caption         =   "&Captaciones"
      Index           =   0
      Begin VB.Menu M0201000000 
         Caption         =   "&Parámetros"
         Index           =   0
         Begin VB.Menu M0201010000 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M0201010000 
            Caption         =   "&Consulta"
            Index           =   1
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Tasas de Interés"
         Index           =   1
         Begin VB.Menu M0201020000 
            Caption         =   "&Ahorros"
            Index           =   0
            Begin VB.Menu M0201020100 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201020100 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0201020000 
            Caption         =   "&Plazo Fijo"
            Index           =   1
            Begin VB.Menu M0201020200 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201020200 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0201020000 
            Caption         =   "&CTS"
            Index           =   2
            Begin VB.Menu M0201020300 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201020300 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Mantenimiento"
         Index           =   2
         Begin VB.Menu M0201030000 
            Caption         =   "&Ahorros"
            Index           =   0
         End
         Begin VB.Menu M0201030000 
            Caption         =   "&Plazo Fijo"
            Index           =   1
         End
         Begin VB.Menu M0201030000 
            Caption         =   "&CTS"
            Index           =   2
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Bloqueo/Desbloqueo"
         Index           =   3
         Begin VB.Menu M0201040000 
            Caption         =   "&Ahorros"
            Index           =   0
         End
         Begin VB.Menu M0201040000 
            Caption         =   "&Plazo Fijo"
            Index           =   1
         End
         Begin VB.Menu M0201040000 
            Caption         =   "&CTS"
            Index           =   2
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Simulación Plazo Fijo"
         Index           =   4
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Tarjeta Magnética"
         Index           =   5
         Begin VB.Menu M0201050000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M0201050000 
            Caption         =   "Re&lación"
            Index           =   1
         End
         Begin VB.Menu M0201050000 
            Caption         =   "&Bloqueo/Desbloqueo"
            Index           =   2
         End
         Begin VB.Menu M0201050000 
            Caption         =   "&Cambio de Clave"
            Index           =   3
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Bene&ficiarios"
         Index           =   6
         Begin VB.Menu M0201060000 
            Caption         =   "&Mantenimiento"
            Index           =   0
         End
         Begin VB.Menu M0201060000 
            Caption         =   "&Consulta "
            Index           =   1
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Orden Pago"
         Index           =   7
         Begin VB.Menu M0201070000 
            Caption         =   "&Generación y Emisión"
            Index           =   0
            Begin VB.Menu M0201070100 
               Caption         =   "&Solicitud de Emisión"
               Index           =   0
            End
            Begin VB.Menu M0201070100 
               Caption         =   "&Consolidación y Envío"
               Index           =   1
            End
            Begin VB.Menu M0201070100 
               Caption         =   "&Recepción"
               Index           =   2
            End
            Begin VB.Menu M0201070100 
               Caption         =   "&Entrega"
               Index           =   3
            End
         End
         Begin VB.Menu M0201070000 
            Caption         =   "&Certificación"
            Index           =   1
         End
         Begin VB.Menu M0201070000 
            Caption         =   "&Anulación"
            Index           =   2
         End
         Begin VB.Menu M0201070000 
            Caption         =   "Con&sulta"
            Index           =   3
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Ser&vicios"
         Index           =   8
         Begin VB.Menu M0201080000 
            Caption         =   "&Universidad Nacional de Trujillo"
            Index           =   2
            Begin VB.Menu M0201080300 
               Caption         =   "&Parámetros"
               Index           =   0
            End
            Begin VB.Menu M0201080300 
               Caption         =   "&Reportes"
               Index           =   1
            End
            Begin VB.Menu M0201080300 
               Caption         =   "&Generación DBF"
               Index           =   2
            End
         End
         Begin VB.Menu M0201080000 
            Caption         =   "&Instituciones Convenio"
            Index           =   3
            Begin VB.Menu M0201080400 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0201080400 
               Caption         =   "&Cuentas"
               Index           =   1
            End
            Begin VB.Menu M0201080400 
               Caption         =   "&Plan de Pagos"
               Index           =   2
            End
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "&Reportes"
         Index           =   9
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Generar ACL Ahorros"
         Index           =   10
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Generar ACL Colocaciones"
         Index           =   11
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Parámetros por Personería"
         Index           =   12
         Begin VB.Menu M0201090000 
            Caption         =   "&Ahorros"
            Index           =   0
         End
         Begin VB.Menu M0201090000 
            Caption         =   "&Plazo Fijo"
            Index           =   1
         End
         Begin VB.Menu M0201090000 
            Caption         =   "&CTS"
            Index           =   2
         End
      End
      Begin VB.Menu M0201000000 
         Caption         =   "Autorización"
         Index           =   13
         Begin VB.Menu M0201100000 
            Caption         =   "Rangos"
            Index           =   0
         End
         Begin VB.Menu M0201100000 
            Caption         =   "Aprobacion / Rechazo"
            Index           =   1
         End
      End
   End
   Begin VB.Menu M0300000000 
      Caption         =   "Colocacio&nes"
      Index           =   0
      Begin VB.Menu M0301000000 
         Caption         =   "Cartas &Fianza"
         Index           =   0
         Begin VB.Menu M0301010000 
            Caption         =   "&Solicitud Carta Fianza"
            Index           =   0
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Gra&var Garantia"
            Index           =   1
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Sugerencia de &Analista"
            Index           =   2
         End
         Begin VB.Menu M0301010000 
            Caption         =   "&Resolver Carta Fianza"
            Index           =   3
            Begin VB.Menu M0301010100 
               Caption         =   "&Aprobar Carta Fianza"
               Index           =   0
            End
            Begin VB.Menu M0301010100 
               Caption         =   "&Rechazar Carta Fianza"
               Index           =   1
            End
            Begin VB.Menu M0301010100 
               Caption         =   "Retirar Carta Fianza Aprobada"
               Index           =   2
            End
            Begin VB.Menu M0301010100 
               Caption         =   "&Devolver Carta Fianza Emitida"
               Index           =   3
            End
         End
         Begin VB.Menu M0301010000 
            Caption         =   "&Emitir Carta Fianza"
            Index           =   4
         End
         Begin VB.Menu M0301010000 
            Caption         =   "&Honrar Carta Fianza"
            Index           =   5
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Consultas"
            Index           =   6
            Begin VB.Menu M0301010200 
               Caption         =   "Historial de Carta Fianza"
               Index           =   0
            End
         End
         Begin VB.Menu M0301010000 
            Caption         =   "Reportes"
            Index           =   7
            Begin VB.Menu M0301010300 
               Caption         =   "Reportes de Carta Fianza"
               Index           =   0
            End
         End
         Begin VB.Menu M0301010000 
            Caption         =   "&Manteminiento Tarifario"
            Index           =   8
         End
         Begin VB.Menu M0301010000 
            Caption         =   "&Relacionar con Credito"
            Index           =   9
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Creditos"
         Index           =   1
         Begin VB.Menu M0301020000 
            Caption         =   "&Definiciones"
            Index           =   0
            Begin VB.Menu M0301020100 
               Caption         =   "&Parametros de Control"
               Index           =   0
               Begin VB.Menu M0301020101 
                  Caption         =   "&Mantenimiento"
                  Index           =   0
               End
               Begin VB.Menu M0301020101 
                  Caption         =   "&Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu M0301020100 
               Caption         =   "&Lineas de Credito"
               Index           =   1
               Begin VB.Menu M0301020102 
                  Caption         =   "&Registro"
                  Index           =   0
               End
               Begin VB.Menu M0301020102 
                  Caption         =   "&Mantenimiento"
                  Index           =   1
               End
               Begin VB.Menu M0301020102 
                  Caption         =   "&Consulta"
                  Index           =   2
               End
            End
            Begin VB.Menu M0301020100 
               Caption         =   "&Niveles de Aprobacion"
               Index           =   2
               Begin VB.Menu M0301020103 
                  Caption         =   "&Mantenimiento"
                  Index           =   0
               End
               Begin VB.Menu M0301020103 
                  Caption         =   "&Consulta"
                  Index           =   1
               End
            End
            Begin VB.Menu M0301020100 
               Caption         =   "&Gastos"
               Index           =   3
               Begin VB.Menu M0301020104 
                  Caption         =   "&Mantenimiento"
                  Index           =   0
               End
               Begin VB.Menu M0301020104 
                  Caption         =   "&Consulta"
                  Index           =   1
               End
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Solicitud Credito"
            Index           =   1
            Begin VB.Menu M0301020200 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M0301020200 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Relaciones de Credito"
            Index           =   2
            Begin VB.Menu M0301020300 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0301020300 
               Caption         =   "Reasignacion de &Cartera en Lote"
               Index           =   1
            End
            Begin VB.Menu M0301020300 
               Caption         =   "Con&sulta"
               Index           =   2
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Garantias"
            Index           =   3
            Begin VB.Menu M0301020400 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M0301020400 
               Caption         =   "&Mantenimiento"
               Index           =   1
            End
            Begin VB.Menu M0301020400 
               Caption         =   "&Consulta"
               Index           =   2
            End
            Begin VB.Menu M0301020400 
               Caption         =   "Gravamen"
               Index           =   3
            End
            Begin VB.Menu M0301020400 
               Caption         =   "&Liberar o Bloquear Garantia"
               Index           =   4
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Sugerencia"
            Index           =   4
            Begin VB.Menu M0301020500 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M0301020500 
               Caption         =   "&Consulta"
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Resolver Creditos"
            Index           =   5
            Begin VB.Menu M0301020600 
               Caption         =   "&Aprobacion"
               Index           =   0
            End
            Begin VB.Menu M0301020600 
               Caption         =   "&Rechazo"
               Index           =   1
            End
            Begin VB.Menu M0301020600 
               Caption         =   "A&nulacion Creditos Aprobados"
               Index           =   2
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Reprogramacion Credito"
            Index           =   6
            Begin VB.Menu M0301020700 
               Caption         =   "Repr&ogramacion"
               Index           =   0
            End
            Begin VB.Menu M0301020700 
               Caption         =   "Reprogramacion en &Lote"
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Refinanciacion"
            Index           =   7
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Actualizacion de &Metodos de Liquidacion"
            Index           =   8
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Perdonar Mora"
            Index           =   9
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Gastos"
            Index           =   10
            Begin VB.Menu M0301020800 
               Caption         =   "&Administracion de Gastos en &Lote"
               Index           =   0
            End
            Begin VB.Menu M0301020800 
               Caption         =   "Mantenimiento de Penalidad de Cancelacion"
               Index           =   1
            End
            Begin VB.Menu M0301020800 
               Caption         =   "Administracion de Gastos"
               Index           =   2
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Reasignar &Institucion"
            Index           =   11
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Transferencia a Recuperaciones"
            Index           =   12
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Analista"
            Index           =   13
            Begin VB.Menu M0301020900 
               Caption         =   "&Nota de Analista"
               Index           =   0
            End
            Begin VB.Menu M0301020900 
               Caption         =   "&Metas de Analista"
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Si&mulaciones"
            Index           =   14
            Begin VB.Menu M0301021000 
               Caption         =   "Calendario de &Pagos"
               Index           =   0
            End
            Begin VB.Menu M0301021000 
               Caption         =   "Calendario de &Desembolsos Parciales"
               Index           =   1
            End
            Begin VB.Menu M0301021000 
               Caption         =   "Calendario de Cuota &Libre"
               Index           =   2
            End
            Begin VB.Menu M0301021000 
               Caption         =   "Simulador de Pagos"
               Index           =   3
            End
            Begin VB.Menu M0301021000 
               Caption         =   "Simulador de Nro de Cuotas"
               Index           =   4
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Cons&ultas"
            Index           =   15
            Begin VB.Menu M0301021100 
               Caption         =   "&Historial de Credito"
               Index           =   0
            End
            Begin VB.Menu M0301021100 
               Caption         =   "Historial de &Calendario"
               Index           =   1
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Reportes"
            Index           =   16
            Begin VB.Menu M0301021200 
               Caption         =   "&Duplicados"
               Index           =   0
            End
            Begin VB.Menu M0301021200 
               Caption         =   "Listados de Creditos"
               Index           =   1
            End
            Begin VB.Menu M0301021200 
               Caption         =   "Creditos Vinculados Titulares"
               Index           =   2
            End
            Begin VB.Menu M0301021200 
               Caption         =   "Creditos Vinculados Titulares y Garantes"
               Index           =   3
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Registro de Dacion de Pago"
            Index           =   17
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Cargo Automatico"
            Index           =   18
            Begin VB.Menu M0301021500 
               Caption         =   "Asignar Cargo Automatico"
               Index           =   1
            End
            Begin VB.Menu M0301021500 
               Caption         =   "Manteminiento"
               Index           =   2
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Modificar Codigos Modulares"
            Index           =   19
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Asignar Cuota Comodin"
            Index           =   20
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Administrar PrePagos Hipotecario"
            Index           =   21
         End
         Begin VB.Menu M0301020000 
            Caption         =   "CrediPago"
            Index           =   22
            Begin VB.Menu M0301021300 
               Caption         =   "Registro CrediPago"
               Index           =   0
            End
            Begin VB.Menu M0301021300 
               Caption         =   "Archivo Cobranza"
               Index           =   1
            End
            Begin VB.Menu M0301021300 
               Caption         =   "Archivo Resultado"
               Index           =   2
            End
         End
         Begin VB.Menu M0301020000 
            Caption         =   "Valorizar Cheque"
            Index           =   23
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Calendario de Desembolsos"
            Index           =   24
         End
         Begin VB.Menu M0301020000 
            Caption         =   "&Sustitución de Deudor"
            Index           =   25
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Pignoraticio"
         Index           =   2
         Begin VB.Menu M0301030000 
            Caption         =   "&Contrato"
            Index           =   0
            Begin VB.Menu M0301030100 
               Caption         =   "&Registro"
               Index           =   0
            End
            Begin VB.Menu M0301030100 
               Caption         =   "Mantenimiento Descripcion"
               Index           =   1
            End
            Begin VB.Menu M0301030100 
               Caption         =   "Anulación"
               Index           =   2
            End
            Begin VB.Menu M0301030100 
               Caption         =   "&Bloqueo"
               Index           =   3
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Rescate Joyas"
            Index           =   1
            Begin VB.Menu M0301030200 
               Caption         =   "Devolución de Joyas"
               Index           =   0
            End
            Begin VB.Menu M0301030200 
               Caption         =   "Devolución de Joyas No Desembolsadas"
               Index           =   1
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Remate"
            Index           =   2
            Begin VB.Menu M0301030300 
               Caption         =   "Preparacion Remate"
               Index           =   0
            End
            Begin VB.Menu M0301030300 
               Caption         =   "Remate"
               Index           =   1
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Adjudicacion"
            Index           =   3
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Subasta"
            Index           =   4
            Begin VB.Menu M0301030400 
               Caption         =   "Preparacion Subasta"
               Index           =   0
            End
            Begin VB.Menu M0301030400 
               Caption         =   "Subasta"
               Index           =   1
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "Chafaloneo"
            Index           =   5
         End
         Begin VB.Menu M0301030000 
            Caption         =   "&Consultas"
            Index           =   6
            Begin VB.Menu M0301030500 
               Caption         =   "Historial de Contrato"
               Index           =   0
            End
            Begin VB.Menu M0301030500 
               Caption         =   "Contrato de Persona"
               Index           =   1
            End
         End
         Begin VB.Menu M0301030000 
            Caption         =   "&Reportes"
            Index           =   7
            Begin VB.Menu M0301030600 
               Caption         =   "Movimientos Diario"
               Index           =   0
            End
            Begin VB.Menu M0301030600 
               Caption         =   "Listados Generales"
               Index           =   1
            End
            Begin VB.Menu M0301030600 
               Caption         =   "Estadisticas"
               Index           =   2
            End
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Pignoraticio"
         Index           =   3
         Begin VB.Menu M0301040000 
            Caption         =   "Generalidades"
            Index           =   0
            Begin VB.Menu M0301040100 
               Caption         =   "Tarifario"
               Index           =   0
            End
            Begin VB.Menu M0301040100 
               Caption         =   "Calificación Cliente Interno"
               Index           =   1
            End
         End
         Begin VB.Menu M0301040000 
            Caption         =   "&Contratos"
            Index           =   1
            Begin VB.Menu M0301040200 
               Caption         =   "Registro"
               Index           =   0
            End
            Begin VB.Menu M0301040200 
               Caption         =   "Mantenimiento de Descripción"
               Index           =   1
            End
            Begin VB.Menu M0301040200 
               Caption         =   "Anulación"
               Index           =   2
            End
            Begin VB.Menu M0301040200 
               Caption         =   "Bloqueo"
               Index           =   3
            End
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Remesa"
            Index           =   2
            Begin VB.Menu M0301040300 
               Caption         =   "Proyección"
               Index           =   0
            End
            Begin VB.Menu M0301040300 
               Caption         =   "Despacho"
               Index           =   1
            End
            Begin VB.Menu M0301040300 
               Caption         =   "Recepción"
               Index           =   2
            End
         End
         Begin VB.Menu M0301040000 
            Caption         =   "&Remate"
            Index           =   3
            Begin VB.Menu M0301040400 
               Caption         =   "Registro Remate"
               Index           =   0
            End
            Begin VB.Menu M0301040400 
               Caption         =   "Proceso Remate"
               Index           =   1
            End
            Begin VB.Menu M0301040400 
               Caption         =   "Verificador y Bloqueo de Piezas"
               Index           =   2
            End
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Selección de Joyas para Venta/Fundición"
            Index           =   4
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Consulta de Contratos"
            Index           =   5
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Proceso de Fundicion"
            Index           =   6
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Reporte"
            Index           =   7
         End
         Begin VB.Menu M0301040000 
            Caption         =   "Evaluacion Mensual de Clientes"
            Index           =   8
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "&Recuperaciones"
         Index           =   4
         Begin VB.Menu M0301050000 
            Caption         =   "Proceso &Judicial"
            Index           =   1
            Begin VB.Menu M0301050100 
               Caption         =   "&Registro de Expedientes"
               Index           =   0
            End
            Begin VB.Menu M0301050100 
               Caption         =   "&Actuaciones Procesales"
               Index           =   1
            End
         End
         Begin VB.Menu M0301050000 
            Caption         =   "&Gastos de Recuperacion"
            Index           =   2
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Actualizacion Metodo de &Liquidacion"
            Index           =   3
         End
         Begin VB.Menu M0301050000 
            Caption         =   "&Relaciones de Crédito"
            Index           =   4
            Begin VB.Menu M0301050200 
               Caption         =   "&Mantenimiento"
               Index           =   0
            End
            Begin VB.Menu M0301050200 
               Caption         =   "&Consulta"
               Index           =   1
            End
            Begin VB.Menu M0301050200 
               Caption         =   "Comision &Abogado"
               Index           =   2
            End
         End
         Begin VB.Menu M0301050000 
            Caption         =   "&Cancelar Credito"
            Index           =   6
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Cas&tigar Credito"
            Index           =   7
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Consulta"
            Index           =   8
            Begin VB.Menu M0301050300 
               Caption         =   "Historial en Recuperaciones"
               Index           =   0
            End
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Reportes"
            Index           =   9
            Begin VB.Menu M0301050400 
               Caption         =   "Reportes Mensuales"
               Index           =   0
            End
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Levantamiento Garantia"
            Index           =   10
         End
         Begin VB.Menu M0301050000 
            Caption         =   "Extorno Levantamiento Garantia"
            Index           =   11
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Calificación de Colocaciones"
         Index           =   5
         Begin VB.Menu M0301060000 
            Caption         =   "Evaluación de la Cartera"
            Index           =   0
         End
         Begin VB.Menu M0301060000 
            Caption         =   "Calificación de Colocaciones"
            Index           =   1
            Begin VB.Menu M0301060100 
               Caption         =   "Preparar Data"
               Index           =   0
            End
            Begin VB.Menu M0301060100 
               Caption         =   "Calificacion Sistema"
               Index           =   1
            End
            Begin VB.Menu M0301060100 
               Caption         =   "Mantenimiento"
               Index           =   2
            End
         End
         Begin VB.Menu M0301060000 
            Caption         =   "Reportes"
            Index           =   2
         End
         Begin VB.Menu M0301060000 
            Caption         =   "Revision de la Calificacion"
            Index           =   3
         End
         Begin VB.Menu M0301060000 
            Caption         =   "Consulta de Calificacion"
            Index           =   4
         End
      End
      Begin VB.Menu M0301000000 
         Caption         =   "Informe RCD"
         Index           =   6
         Begin VB.Menu M0301070000 
            Caption         =   "Parametros RCD"
            Index           =   0
         End
         Begin VB.Menu M0301070000 
            Caption         =   "Preparacion de Data"
            Index           =   1
            Begin VB.Menu M0301070100 
               Caption         =   "Datos Maestro RCD"
               Index           =   0
            End
            Begin VB.Menu M0301070100 
               Caption         =   "Persona desde MaestroRCD"
               Index           =   1
            End
         End
         Begin VB.Menu M0301070000 
            Caption         =   "Generar Datos"
            Index           =   2
            Begin VB.Menu M0301070200 
               Caption         =   "Informe RCD"
               Index           =   0
            End
            Begin VB.Menu M0301070200 
               Caption         =   "Informe IBM"
               Index           =   1
            End
            Begin VB.Menu M0301070200 
               Caption         =   "Verificar Datos"
               Index           =   2
            End
         End
         Begin VB.Menu M0301070000 
            Caption         =   "Reportes"
            Index           =   3
         End
         Begin VB.Menu M0301070000 
            Caption         =   "Validacion de Personas"
            Index           =   4
            Begin VB.Menu M0301070300 
               Caption         =   "RCDMaestro Persona"
               Index           =   0
            End
         End
      End
   End
   Begin VB.Menu M0400000000 
      Caption         =   "&Operaciones"
      Index           =   0
      Begin VB.Menu M0401000000 
         Caption         =   "Tipo de Ca&mbio"
         Index           =   0
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Operaciones"
         Index           =   2
         Shortcut        =   {F2}
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Operaciones CMACs &Recepción"
         Index           =   3
         Shortcut        =   {F3}
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Operaciones CMACs &Llamada"
         Index           =   4
         Shortcut        =   {F4}
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Estadisticas de Compra Venta de ME"
         Index           =   6
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Tarjeta &Magnetica"
         Index           =   7
         Begin VB.Menu M0401020000 
            Caption         =   "&Registro de Tarjeta"
            Index           =   0
         End
         Begin VB.Menu M0401020000 
            Caption         =   "Mantenimiento de Tarjetas Magneticas"
            Index           =   1
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Cierres"
         Index           =   8
         Begin VB.Menu M0401030000 
            Caption         =   "Cierre de &Día"
            Index           =   0
         End
         Begin VB.Menu M0401030000 
            Caption         =   "Cierre de &Mes"
            Index           =   1
         End
         Begin VB.Menu M0401030000 
            Caption         =   "Cierre Diario Batch"
            Index           =   2
         End
         Begin VB.Menu M0401030000 
            Caption         =   "Cierre Mes Batch"
            Index           =   3
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "E&xtornos "
         Index           =   9
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Asiento Contable"
         Index           =   10
         Begin VB.Menu M0401050000 
            Caption         =   "Asiento Contable &Dia"
            Index           =   0
         End
         Begin VB.Menu M0401050000 
            Caption         =   "Asiento Contable &Anterior"
            Index           =   1
         End
         Begin VB.Menu M0401050000 
            Caption         =   "Mantenimiento &Asiento Contable"
            Index           =   2
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "&Reportes"
         Index           =   11
         Begin VB.Menu M0401060000 
            Caption         =   "Resumen de Ingresos y Egresos Consolidado"
            Index           =   0
         End
         Begin VB.Menu M0401060000 
            Caption         =   "Detalle de &Operaciones"
            Index           =   1
         End
         Begin VB.Menu M0401060000 
            Caption         =   "Detalle de &Habilitación/Devolución Cajero"
            Index           =   2
         End
         Begin VB.Menu M0401060000 
            Caption         =   "Protocolo por &Usuario"
            Index           =   3
         End
         Begin VB.Menu M0401060000 
            Caption         =   "&Total Operaciones de Usuario"
            Index           =   4
         End
         Begin VB.Menu M0401060000 
            Caption         =   "Reportes de Cajamatic"
            Index           =   5
         End
      End
      Begin VB.Menu M0401000000 
         Caption         =   "Aviso Operaciones Pendientes"
         Index           =   12
         Visible         =   0   'False
         Begin VB.Menu M0401000001 
            Caption         =   "Automático"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu M0401000001 
            Caption         =   "Mostrar"
            Index           =   1
         End
      End
   End
   Begin VB.Menu M0500000000 
      Caption         =   "&Servicios"
      Index           =   0
      Begin VB.Menu M0501000000 
         Caption         =   "U.N.T"
         Index           =   0
         Begin VB.Menu M0501010000 
            Caption         =   "&Registrar Alumnos Nuevos UNT"
            Index           =   0
         End
         Begin VB.Menu M0501010000 
            Caption         =   "Generación Archivo &UNT"
            Index           =   1
         End
         Begin VB.Menu M0501010000 
            Caption         =   "&Actualiza Orden PAgo UNT"
            Index           =   2
         End
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Telefonica"
         Index           =   1
         Begin VB.Menu M0501020000 
            Caption         =   "&Generación Archivo Telefonica"
            Index           =   0
         End
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Consolidacion de Ser&vicios"
         Index           =   2
         Begin VB.Menu M0501030000 
            Caption         =   "&Hidrandina"
            Index           =   0
         End
         Begin VB.Menu M0501030000 
            Caption         =   "&Sedalib"
            Index           =   1
         End
         Begin VB.Menu M0501030000 
            Caption         =   "&Edelnor"
            Index           =   2
         End
      End
      Begin VB.Menu M0501000000 
         Caption         =   "Mantenimiento de Parámetros"
         Index           =   3
         Begin VB.Menu M0501040000 
            Caption         =   "&Hidrandina"
            Index           =   0
         End
         Begin VB.Menu M0501040000 
            Caption         =   "&Sedalib"
            Index           =   1
         End
         Begin VB.Menu M0501040000 
            Caption         =   "&Fideicomiso"
            Index           =   2
         End
         Begin VB.Menu M0501040000 
            Caption         =   "&Edelnor"
            Index           =   3
         End
      End
      Begin VB.Menu M0501000000 
         Caption         =   "SAT"
         Index           =   4
         Begin VB.Menu M0501050000 
            Caption         =   "&Carga de Valores"
            Index           =   0
         End
         Begin VB.Menu M0501050000 
            Caption         =   "&Distribucion de Fondos"
            Index           =   1
         End
      End
   End
   Begin VB.Menu M0600000000 
      Caption         =   "&Sistema"
      Index           =   0
      Begin VB.Menu M0601000000 
         Caption         =   "&Parametros del Sistema"
         Index           =   0
      End
      Begin VB.Menu M0601000000 
         Caption         =   "&Variables del Sistema"
         Index           =   1
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento &Permisos"
         Index           =   2
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento &Zonas"
         Index           =   3
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento Agencias"
         Index           =   4
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento de Ctas Contables"
         Index           =   5
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Copia de &Seguridad"
         Index           =   6
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento Grupo &Operaciones"
         Index           =   7
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento Operaciones Captaciones"
         Index           =   8
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento de Codigo Postal"
         Index           =   9
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Parámetros de Cheques"
         Index           =   10
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento CIIU"
         Index           =   11
      End
      Begin VB.Menu M0601000000 
         Caption         =   "Mantenimiento de Feriados"
         Index           =   12
      End
   End
   Begin VB.Menu M0700000000 
      Caption         =   "&Personas"
      Index           =   0
      Begin VB.Menu M0701000000 
         Caption         =   "&Personas"
         Index           =   0
         Begin VB.Menu M0701010000 
            Caption         =   "&Registro"
            Index           =   0
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Manteniemiento"
            Index           =   1
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Consulta"
            Index           =   2
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Exoneradas del Lavado de Dinero"
            Index           =   3
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Roles de Persona"
            Index           =   4
         End
         Begin VB.Menu M0701010000 
            Caption         =   "&Comentarios a Persona"
            Index           =   5
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
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu M0701000000 
         Caption         =   "&Posicion de Cliente"
         Index           =   3
      End
      Begin VB.Menu M0701000000 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu M0701000000 
         Caption         =   "Central de Riesgos"
         Index           =   5
         Begin VB.Menu M0701030000 
            Caption         =   "&SuperIntendencia de Banca y Seguros"
            Index           =   0
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu M0701030000 
            Caption         =   "&Infocorp"
            Index           =   1
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu M0701030000 
            Caption         =   "&Camara de Comercio"
            Index           =   2
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu M0701030000 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu M0701030000 
            Caption         =   "Morosos"
            Index           =   4
            Shortcut        =   ^{F4}
         End
         Begin VB.Menu M0701030000 
            Caption         =   "Direll"
            Index           =   5
         End
      End
      Begin VB.Menu M0701000000 
         Caption         =   "Reportes"
         Index           =   6
      End
   End
   Begin VB.Menu M0800000000 
      Caption         =   "Herra&mientas"
      Index           =   0
      Begin VB.Menu M0801000000 
         Caption         =   "Editor de Textos"
         Index           =   0
      End
      Begin VB.Menu M0801000000 
         Caption         =   "Spooler de Impresión"
         Index           =   1
      End
      Begin VB.Menu M0801000000 
         Caption         =   "Configuración de Periféricos"
         Index           =   2
      End
      Begin VB.Menu M0801000000 
         Caption         =   "Cargar Logo &Penware"
         Index           =   3
      End
      Begin VB.Menu M0801000000 
         Caption         =   "Explorador de Archivos"
         Index           =   4
      End
   End
   Begin VB.Menu M0900000000 
      Caption         =   "A&yuda"
      Index           =   0
      Begin VB.Menu M0901000000 
         Caption         =   "&Contenido"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu M0901000000 
         Caption         =   "&Indice"
         Index           =   1
      End
      Begin VB.Menu M0901000000 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu M0901000000 
         Caption         =   "&Acerca del Sistema..."
         Index           =   3
      End
   End
End
Attribute VB_Name = "MDISicmact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Api para Ejecutar Internet Explorer
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpoperation As String, _
ByVal lpfile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowcmd As Long) As Long


Private Sub cmdVer_Click()
      FrmCapAutOpeEstados.Show
End Sub

Private Sub M0101000000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmImpresora.Show 1
        Case 2
            If MsgBox("Esta Seguro que Desea Salir del Sistema ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
                End
            End If
    End Select
End Sub

Private Sub M0201000000_Click(Index As Integer)
    Select Case Index
        Case 4
            frmCapSimulacionPF.Show 1
        Case 9
            frmCapReportes.Show 1
        Case 10
            frmAclAhorros.Show
        Case 11
            frmAclColocaciones.Show
        
            
    End Select
End Sub

Private Sub M0201010000_Click(Index As Integer)
    Select Case Index
        Case 0
             frmCapParametros.Inicia False
        Case 1
            frmCapParametros.Inicia True
    End Select
End Sub

Private Sub M0201100000_Click(Index As Integer)


    Select Case Index
        Case 0 'Rangos
          '  Call frmCapAutorizacionRango.Inicia
           frmCapAutOpe.Show
        Case 1 'Aprobacion / rechazo
           Call frmCapAutMovOpe.Show
          ' Form1.Show
    End Select

End Sub

Private Sub M0201020100_Click(Index As Integer)
    Select Case Index
        Case 0 'mantenimiento
            frmCapTasaInt.Inicia gCapAhorros, False
        Case 1 'Consulta
            frmCapTasaInt.Inicia gCapAhorros, True
    End Select
End Sub

Private Sub M0201020200_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimiento
            frmCapTasaInt.Inicia gCapPlazoFijo, False
        Case 1 'Consulta
            frmCapTasaInt.Inicia gCapPlazoFijo, True
    End Select
End Sub

Private Sub M0201020300_Click(Index As Integer)
    Select Case Index
        Case 0
            frmCapTasaInt.Inicia gCapCTS, True
        Case 1
            frmCapTasaInt.Inicia gCapCTS, False
    End Select
End Sub

Private Sub M0201030000_Click(Index As Integer)
    'Mantenimiento
    Select Case Index
        Case 0 'Ahorro
            frmCapMantenimiento.Inicia gCapAhorros
        Case 1 'Plazo fijo
            frmCapMantenimiento.Inicia gCapPlazoFijo
        Case 2 'Cts
            frmCapMantenimiento.Inicia gCapCTS
    End Select
End Sub

Private Sub M0201040000_Click(Index As Integer)
    'Bloqueos / Desbloqueos
    Select Case Index
        Case 0 'Ahorros
            frmCapBloqueoDesbloqueo.Inicia gCapAhorros
        Case 1 'Plazo Fijo
            frmCapBloqueoDesbloqueo.Inicia gCapPlazoFijo
        Case 2 'Cts
            frmCapBloqueoDesbloqueo.Inicia gCapCTS
    End Select
End Sub

Private Sub M0201050000_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro
            frmCapTarjetaRegistro.Inicia
        Case 1 'Relacion
            frmCapTarjetaRelacion.Inicia False
        Case 2 'Bloqueo
            frmCapTarjetaBlqDesBlq.Show 1
        Case 3 'Cambio de Clave
            frmCapTarjetaCambioClave.Show 1
    End Select
End Sub

Private Sub M0201060000_Click(Index As Integer)
    Select Case Index
        Case 0
            frmCapBeneficiario.Inicia False
        Case 1
            frmCapBeneficiario.Inicia True
    End Select
End Sub

Private Sub M0201070000_Click(Index As Integer)
    Select Case Index
        'Case 0 'generacion
            'frmCapOrdPagGenEmi.Show 1
        Case 1 'Certificacion
            frmCapOrdPagAnulCert.Inicia gAhoOPCertificacion
        Case 2 'Anulacion
            frmCapOrdPagAnulCert.Inicia gAhoOPAnulacion
        Case 3 'Consulta
            frmCapOrdPagConsulta.Show 1
    End Select
End Sub

Private Sub M0201070100_Click(Index As Integer)
Select Case Index
    Case 0
        frmCapOrdPagSolicitud.Show 1
    Case 1
        frmCapOrdPagProceso.Inicia gCapTalOrdPagEstSolicitado
    Case 2
        frmCapOrdPagProceso.Inicia gCapTalOrdPagEstEnviado
    Case 3
        frmCapOrdPagProceso.Inicia gCapTalOrdPagEstRecepcionado
End Select
End Sub



Private Sub M0201080400_Click(Index As Integer)
    Select Case Index
        Case 0 'mantenimiento
            frmCapConvenioMant.Show 1
        Case 1 'Cuentas
            frmCapServConvCuentas.Inicia
        Case 2 'Plan Pagos
            frmCapServConvPlanPag.Inicia
    End Select
End Sub

Private Sub M0201090000_Click(Index As Integer)
Select Case Index
    Case 0
        frmCapPersParam.Inicia gCapAhorros
    Case 1
        frmCapPersParam.Inicia gCapPlazoFijo
    Case 2
        frmCapPersParam.Inicia gCapCTS
End Select
End Sub

Private Sub M0301010000_Click(Index As Integer)
    Select Case Index
    Case 0 'Solicitud CartaFianza
        frmCFSolicitud.Show 1
    Case 1 'Gravar Garantias
        frmCredGarantCred.Inicio PorMenu, , 1
    Case 2 'Solicitud CartaFianza
        Call frmCFSugerencia.Inicia
    Case 4 'Emitir CartaFianza
        frmCFEmision.Show 1
    Case 5 ' Honrar CartaFianza
        FrmCFHonrar.Show 1
    'Case 8 ' Niveles de Aprobacion
        'frmCFNivelApr.Show 1 YA NO SE USA
    
    Case 8 'Matenimiento de Tarifario
        frmCFTarifario.Show 1
    Case 9  ' Relacionar con Credito
        FrmCFHonrarCredito.Show 1
    End Select
End Sub

Private Sub M0301010100_Click(Index As Integer)
    Select Case Index
    Case 0 'Aprobacion
         frmCFAprobacion.Show 1
    Case 1 ' 'Rechazo
        FrmCFRechazar.Show 1
    Case 2
        FrmCFRetirarApr.Show 1
    Case 3
        FrmCFDevolucion.Show 1
    End Select

End Sub

Private Sub M0301010200_Click(Index As Integer)
    Select Case Index
    Case 0 'Consultas
        frmCFHistorial.Show 1
        
    End Select
End Sub

Private Sub M0301010300_Click(Index As Integer)
    Select Case Index
    Case 0 'Reportes
        frmCFReporte.Show 1
    End Select

End Sub

Private Sub M0301020000_Click(Index As Integer)
    Select Case Index
    Case 7 'Refinanciacion de Credito
        Call frmCredSolicitud.RefinanciaCredito(Registrar)
    Case 8 ' 'Actualizacion con Metodos de Liquidacion
        frmCredMntMetLiquid.Show 1
    Case 9 'Perdonar Mora
        frmCredPerdonarMora.Show 1
    Case 11 'Reasignar Institucion
        frmCredReasigInst.Show 1
    Case 12 'Transferencia a Recuperaciones
        frmCredTransARecup.Show 1
    Case 17 'Registro de Dacion de Pago
        frmCredRegisDacion.Show 1
    Case 18
        'frmCredCargoAuto.Show 1
    Case 19
        frmCredCodModular.Show 1
    Case 20
        frmCredAsigCComodin.Show 1
    Case 21
        frmCredAdmPrepago.Show 1
    Case 23
        frmCredValorizaCheque.Show 1
    ' CMACICA_CSTS - 05/11/2003 -------------------------------------------------
    Case 24
        frmCredCalendarioDesemb.Show 1
    Case 25
        Call frmCredSolicitud.SustitucionCredito(Registrar)
    ' --------------------------------------------------------------------------
    
    End Select
End Sub

Private Sub M0301020101_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimeinto de Parametros
            frmCredMantParametros.InicioActualizar
        Case 1 'Consulta de Parametros
            frmCredMantParametros.InicioCosultar
    End Select
End Sub

Private Sub M0301020102_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro de Lineas de Credito
            frmCredLineaCredito.Registrar
        Case 1 'Mantenimiento de Lineas de Credito
            frmCredLineaCredito.Actualizar
        Case 2 ' Consulta de lineas de Credito
            frmCredLineaCredito.Consultar
    End Select
End Sub

Private Sub M0301020103_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimeinto de Niveles de Aprobacion
            frmCredNivAprCred.Inicio MuestraNivelesActualizar
        Case 1 'Consulta de Niveles de Aprobacion
            frmCredNivAprCred.Inicio MuestraNivelesConsulta
    End Select
End Sub

Private Sub M0301020104_Click(Index As Integer)
    Select Case Index
        Case 0 'Mantenimeinto de Gastos
            frmCredMntGastos.Inicio InicioGastosActualizar
        Case 1
            frmCredMntGastos.Inicio InicioGastosConsultar
    End Select
End Sub

Private Sub M0301020200_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro de Solicitud
            frmCredSolicitud.Inicio Registrar
        Case 1 'Consulta de Solicitud
            frmCredSolicitud.Inicio Consulta
    End Select
End Sub

Private Sub M0301020300_Click(Index As Integer)
    Dim oCredRel As New UCredRelacion

    Select Case Index
        Case 0 'Mantenimiento de Relaciones de Credito
            frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioRegistroForm
            Set oCredRel = Nothing
        Case 1 'Reasignacion de Cartera en Lote
            frmCredReasigCartera.Show 1
        Case 2 'Consulta de Relaciones de Credito
            frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioConsultaForm
            Set oCredRel = Nothing
    End Select
End Sub

Private Sub M0301020400_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro de Garantia
            frmPersGarantias.Inicio RegistroGarantia
        Case 1 'Mantenimiento de Garantia
            frmPersGarantias.Inicio MantenimientoGarantia
        Case 2 'Consulta de Garantia
            frmPersGarantias.Inicio ConsultaGarant
        Case 3 'Gravament
            frmCredGarantCred.Inicio PorMenu
        Case 4 'Liberar Garantia
            frmCredLiberaGarantia.Show 1
    End Select
End Sub

Private Sub M0301020500_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro de Sugerencia
            frmCredSugerencia.Sugerencia lSugerTipoActRegistro
        Case 1 'Consulta de Sugerencia
            frmCredSugerencia.Sugerencia lSugerTipoActConsultar
    End Select
End Sub

Private Sub M0301020600_Click(Index As Integer)
    Select Case Index
        Case 0 'Aprobacion de Credito
            frmCredAprobacion.Show 1
        Case 1 'Rechazo de Credito
            frmCredRechazo.Rechazar
        Case 2 'Anulacion de Credito
            frmCredRechazo.Retirar
    End Select
End Sub

Private Sub M0301020700_Click(Index As Integer)
    Select Case Index
        Case 0 'Reprogramacion de Credito
            frmCredReprogCred.Show 1
        Case 1 'Reprogramacion en Lote
            frmCredReprogLote.Show 1
    End Select
End Sub

Private Sub M0301020800_Click(Index As Integer)
    Select Case Index
        Case 0 'Administracion de Gastos en Lote
            frmCredAsigGastosLote.Show 1
        Case 1 ' mantenimiento de Penalidad
            frmCredExonerarPen.Show 1
        Case 2
            frmCredAdmiGastos.Show 1
    End Select
End Sub

Private Sub M0301020900_Click(Index As Integer)
    Select Case Index
        Case 0 'Nota del Analista
            frmCredAsigNota.Show 1
        Case 1 'Meta del Analista
            frmCredMetasAnalista.Show 1
    End Select
End Sub

Private Sub M0301021000_Click(Index As Integer)
    Dim MatCalend As Variant
    Dim Matriz(0) As String
    
    Select Case Index
        Case 0 'Calendario de Pagos
            frmCredCalendPagos.Simulacion DesembolsoTotal
        Case 1 'Desembolsos Parciales
            frmCredCalendPagos.Simulacion DesembolsoParcial
        Case 2 'Cuota Libre
            MatCalend = frmCredCalendCuotaLibre.CalendarioLibre(True, gdFecSis, Matriz, 0#, 0, 0#)
        Case 3
            frmCredSimuladorPagos.Show 1
        Case 4
            frmCredSimNroCuotas.Show 1
    End Select
End Sub

Private Sub M0301021100_Click(Index As Integer)
    If Index = 0 Then
        frmCredConsulta.Show 1
    Else
        'frmCredHistCalendario.Show 1
        Call frmCredHistCalendario.Inicio
    End If
End Sub

Private Sub M0301021200_Click(Index As Integer)
    Select Case Index
        Case 0
            frmCredDupDoc.Show 1
        Case 1
            frmCredReportes.Inicia "Reportes de Créditos"
        Case 2
            frmCredVinculados.Ini True, "Créditos Vinculados - Titulares"
        Case 3
            frmCredVinculados.Ini False, "Créditos Vinculados - T y G Consolidado"
    End Select
End Sub



Private Sub M0301021300_Click(Index As Integer)
Select Case Index
        Case 0  'Registro para CrediPago
            frmCredCrediPago.Show 1
            
        Case 1
            frmCredCrediPagoArchivoCobranza.Show 1
        
        Case 2
            frmCredCrediPagoArchivoResultado.Show 1
End Select
End Sub

Private Sub M0301021500_Click(Index As Integer)
Select Case Index
    Case 1
        Dim oGen As DGeneral
        Dim lbCierreRealizado As Boolean
        
        Set oGen = New DGeneral
        lbCierreRealizado = oGen.GetCierreDiaRealizado(gdFecSis)
        Set oGen = Nothing
        
        If lbCierreRealizado Then
            MsgBox "El cierre ya fue ralizado, no puede ingresar a esta opción.", vbInformation, "Aviso"
            Exit Sub
        End If
        frmCredCargoAuto.Show 1
    Case 2
        FrmMantCargoAutomatico.Show 1
End Select
End Sub

Private Sub M0301030000_Click(Index As Integer)
    Select Case Index
        Case 3  'Adjudicacion
            frmColPAdjudicaLotes.Show 1
        Case 5 ' Chafaloneo
        
    End Select
End Sub

Private Sub M0301030100_Click(Index As Integer)
    Select Case Index
        Case 0 'Registro
            frmColPRegContratoDet.Show 1
        Case 1 'Mantenimiento
            frmColPMantPrestamoPig.Show 1
        Case 2 'Anulacion
            frmColPAnularPrestamoPig.Show 1
        Case 3 'Bloqueo
            frmColPBloqueo.Show 1
    End Select
End Sub

Private Sub M0301030200_Click(Index As Integer)
    Select Case Index
        Case 0 'Entrega Joyas
            frmColPRescateJoyas.Inicio gColPOpeDevJoyas, "Entrega de Joyas"
        Case 1 'Entrega Joyas No Desembolsadas
            frmColPRescateJoyas.Inicio gColPOpeDevJoyasNoDesemb, "Entrega Joyas No Desembolsadas"
    End Select
End Sub

Private Sub M0301030300_Click(Index As Integer)
    Select Case Index
        Case 0 'Preparacion de Remate
            frmColPRematePrepara.Show 1
        Case 1 'Remate
            frmColPRemateProceso.Show 1
    End Select
End Sub

Private Sub M0301030400_Click(Index As Integer)
    Select Case Index
        Case 0
            frmColPSubastaPrepara.Show 1
        Case 1
            frmColPSubastaProceso.Show 1
    End Select
End Sub

Private Sub M0301030500_Click(Index As Integer)
    
    Select Case Index
        Case 0
             frmColPMovs.Show 1
        Case 1
             frmColPContratosxCliente.Show 1
    End Select

End Sub

Private Sub M0301030600_Click(Index As Integer)
    Select Case Index
        Case 0
             frmColPRepo.Inicio 1
        Case 1
             frmColPRepo.Inicio 2
        Case 2
             frmColPRepo.Inicio 3
        Case 3
            
    End Select
End Sub

Private Sub M0301040000_Click(Index As Integer)
    Select Case Index
    Case 4
        frmPigSeleccionVentaFundicion.Show 1
    Case 5
        frmPigConsulta.Inicio
    Case 6
        frmPigFundicionJoya.Show 1
    Case 7
        FrmPigRepValores.Show 1
    Case 8
        frmPigEvaluacionMensualClientes.Show 1
    End Select
End Sub

Private Sub M0301040100_Click(Index As Integer)
    Select Case Index
    Case 0
        frmPigTarifario.Show 1
    Case 1
        FrmPigClasificaCli.Show 1
    End Select
End Sub

Private Sub M0301040200_Click(Index As Integer)
    Select Case Index
    Case 0
        frmPigRegContrato.Show 1
    Case 1
        frmPigMantContrato.Show 1
    Case 2
        frmPigAnularContrato.Show 1
    Case 3
        FrmPigBloqueo.Show 1
    End Select
End Sub

Private Sub M0301040300_Click(Index As Integer)
    Select Case Index
    Case 0
        frmPigProyeccionGuia.Show 1
    Case 1
        frmPigDespachoGuia.Show 1
    Case 2
        frmPigRecepcionValija.Show 1
    End Select
End Sub

Private Sub M0301040400_Click(Index As Integer)

    Select Case Index
    Case 0
        frmPigRegistroRemate.Show 1
    Case 1
        frmPigProcesoRemate.Show 1
    Case 2
        Dim oPigRemate As DPigContrato
        Dim rs As Recordset
        
        Set oPigRemate = New DPigContrato
        Set rs = oPigRemate.dObtieneDatosRemate(oPigRemate.dObtieneMaxRemate() - 1)
        If Not (rs.EOF And rs.BOF) Then
            If CStr(rs!cUbicacion) = Right(gsCodAge, 2) Then
                FrmPigVentaRemate.Show 1
            Else
                MsgBox "Usuario no se encuentra asignado en la Agencia de Remate", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
    End Select
End Sub

Private Sub M0301050000_Click(Index As Integer)
    Select Case Index
        Case 0 ' Ingreso a Recup de Otras Entidades
            frmColRecIngresoOtrasEnt.Show 1
        Case 2 ' Gastos en Recuperaciones
            frmColRecGastosRecuperaciones.Show 1
        Case 3 ' Metodo de Liquidacion
            frmColRecMetodoLiquid.Show 1
           
        Case 5 ' Pago
            frmColRecPagoCredRecup.Inicio gColRecOpePagJudSDEfe, "PAGO CREDITO EN RECUPERACIONES", gsCodCMAC, gsNomCmac, True
        Case 6 ' Cancelacion
            frmColRecCancelacion.Show 1
        Case 7 ' Castigo
            frmColRecCastigar.Show 1
        Case 8
        
        Case 10
            frmGarLevant.Show 1
        Case 11
            frmGarantExtorno.Show 1
    End Select
End Sub

Private Sub M0301050100_Click(Index As Integer)
    Select Case Index
        Case 0
            frmColRecExped.Show 1
        Case 1
            frmColRecActuacionesProc.Inicia "N"
    End Select
End Sub

Private Sub M0301050200_Click(Index As Integer)
 Dim oCredRel As New UCredRelacion
 Select Case Index
    Case 0
        frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioRegistroForm
    Case 1
        frmCredRelaCta.Inicio oCredRel, InicioMantenimiento, InicioConsultaForm
    Case 2
        frmColRecComision.Show 1
End Select
End Sub

Private Sub M0301050300_Click(Index As Integer)
Select Case Index
    Case 0
        frmColRecRConsulta.Inicia "Consulta de Pagos de Créditos Judiciales"
End Select
End Sub

Private Sub M0301050400_Click(Index As Integer)
Select Case Index
    Case 0
        frmColRecReporte.Inicia "Reportes de Recuperaciones"
End Select
End Sub

Private Sub M0301060000_Click(Index As Integer)
Select Case Index
    Case 0
     frmColocCalEvalCli.Inicio True
    Case 2
        FrmColocEvalRep.Show 0, MDISicmact
    Case 3
        frmColocCalEvalCli.Inicio False
    Case 4
        FrmColocEvalConsulta.Show 1
End Select
End Sub

Private Sub M0301060100_Click(Index As Integer)
Select Case Index
    Case 0 ' Preparar Data
         frmColocCalActualizaMaestroRCC.Show 1
    Case 1 ' Calificacion del Sistema
        frmColocCalSist.Show 1
    Case 2
        frmColocCalTabla.Show 1
        
End Select
End Sub

Private Sub M0301070000_Click(Index As Integer)
Select Case Index
    Case 0
        frmRCDParametro.Show 1
    Case 3
        frmRCDReporte.Show
        'frmColocCalEvalCli.Inicio True
End Select

End Sub

Private Sub M0301070100_Click(Index As Integer)
Select Case Index
    Case 0 ' Datos Maestro RCD
         frmRCDActualizaRCDMaestroPersona.Show 1
    Case 1 ' Persona desde Maestro RCD
        frmRCDActPersDeMaestroPersona.Show 1
End Select
End Sub

Private Sub M0301070101_Click(Index As Integer)
Select Case Index
    Case 0 ' Carga de Errores
        frmRCDErrorCargaTXT.Show 1
    Case 1 ' Correccion de Errores
        frmRCDErrorCorreccion.Show 1
End Select
End Sub

Private Sub M0301070200_Click(Index As Integer)
Select Case Index
    Case 0 'Informe RCD
        frmRCDGeneraDatosRCD.Show 1
    Case 1 ' Informe IBM
        frmRCDGeneraDatosIBM.Show 1
    Case 2
        frmRCDVericaDatos.Show 1
End Select
End Sub

Private Sub M0301070300_Click(Index As Integer)
Select Case Index
    Case 0 ' RCDMaestro Persona
        frmRCDMantMaestroPersona.Show 1
End Select
End Sub

Private Sub M0401000000_Click(Index As Integer)
If Index = 2 Or Index = 3 Or Index = 4 Or Index = 9 Then
    Dim clsTC As nTipoCambio
    Dim nTC As Double
    Set clsTC = New nTipoCambio
    nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
    Set clsTC = Nothing
    If nTC = 0 Then
        MsgBox "NO se ha registrado el TIPO DE CAMBIO. Debe registrarse para iniciar operaciones.", vbInformation, "Aviso"
        Exit Sub
    End If
End If

Dim sfiltro() As String
Dim lnFiltraTC As Integer
Dim lnFiltraMP As Integer
Dim oGen As DGeneral
Dim lbCierreRealizado As Boolean

Set oGen = New DGeneral
lnFiltraTC = CInt(oGen.LeeConstSistema(102))
lnFiltraMP = CInt(oGen.LeeConstSistema(103))
lbCierreRealizado = oGen.GetCierreDiaRealizado(gdFecSis)
Set oGen = Nothing

    Select Case Index
        Case 0
            frmMantTipoCambio.Show 1
        Case 2
            If lbCierreRealizado Then
                MsgBox "El cierre ya fue ralizado, no puede ingresar a esta opción.", vbInformation, "Aviso"
                Exit Sub
            End If
            ReDim sfiltro(5)
            If lnFiltraMP = 1 Then
                sfiltro(1) = "[1][01234][012]" 'Pigno Trujillo
            ElseIf lnFiltraMP = 2 Then
                sfiltro(1) = "[1][01345][012]"   'Pigno Lima
            ElseIf lnFiltraMP = 0 Then
                sfiltro(1) = "[1][012345][012]"   'Ambos
            End If
           
            sfiltro(2) = "[23][0-2][0123]"    'Captaciones
          
            If lnFiltraTC = 0 Then
                sfiltro(3) = "90002[0-3]"       'Compra Venta
            ElseIf lnFiltraTC = 1 Then
                sfiltro(3) = "90002[0-6]"
            End If
            sfiltro(4) = "9010[01][012356789]"    'Control de Efectivo Boveda y Cajero
            sfiltro(5) = "90003[0-5]"    'Operaciones con Cheques
            frmCajeroOperaciones.Inicia sfiltro, "Cajero - Operaciones"
        Case 3
            If lbCierreRealizado Then
                MsgBox "El cierre ya fue ralizado, no puede ingresar a esta opción.", vbInformation, "Aviso"
                Exit Sub
            End If
            
            ReDim sfiltro(3)
            sfiltro(1) = "260[0-3]" 'Operaciones de Captaciones
            sfiltro(2) = "126"      'Operaciones de Prendario
            sfiltro(3) = "106"      'Operaciones de Colocaciones
            frmCajeroOpeCMAC.Inicia sfiltro, "Cajero - Operaciones CMACs Recepción"
        Case 4
            If lbCierreRealizado Then
                MsgBox "El cierre ya fue ralizado, no puede ingresar a esta opción.", vbInformation, "Aviso"
                Exit Sub
            End If
            
            ReDim sfiltro(3)
            sfiltro(1) = "2605"     'Operaciones de Captaciones
            sfiltro(2) = "127"      'Operaciones de Prendario
            sfiltro(3) = "107"      'Operaciones de Colocaciones
            frmCajeroOpeCMAC.Inicia sfiltro, "Cajero - Operaciones CMACs Llamada"
        Case 9
            If lbCierreRealizado Then
                MsgBox "El cierre ya fue realizado, no puede ingresar a esta opción.", vbInformation, "Aviso"
                Exit Sub
            End If
        
            ReDim sfiltro(9)
            sfiltro(1) = "1[034]9[012]"         'Extornos de Colocaciones
            If lnFiltraMP = 1 Then
                sfiltro(2) = "129"          'Extornos de Prendario Trujillo
            ElseIf lnFiltraMP = 2 Then      'Extornos de Prendario Lima
                sfiltro(2) = "159"
            ElseIf lnFiltraMP = 0 Then      'Extornos de Prendario Lima
                sfiltro(2) = "1[25]9"
            End If
            sfiltro(3) = "2[3457]"      'Extornos de Captaciones
            sfiltro(4) = "3[569]"       'Extornos de Otras Operaciones
            If lnFiltraTC = 0 Then
                sfiltro(5) = "90900[0-3]"
            ElseIf lnFiltraTC = 1 Then
                sfiltro(5) = "90900[0-6]"
            End If
            sfiltro(6) = "90103[0-9]"   'Extornos de Operaciones de Boveda
            sfiltro(7) = "90102[1-9]"   'Extornos de Operaciones de Cajero
            sfiltro(8) = "90003[6-9]"   'Extornos de Operaciones con Cheque
            sfiltro(9) = "90004[4-6]"   'Extornos de Compra Venta - Tipo de Cambio Especial
            
            frmCajeroOperaciones.Inicia sfiltro, "Cajero - Extornos"
        Case 12
           'Poner timer para que muestre el form cada cierto tiempo
            
            
           ' AvisoOperacionesPendientes
            
    End Select
End Sub

Private Sub M0401010000_Click(Index As Integer)
Select Case Index
    Case 0
        'Case gOpeHabCajDevABoveMN, gOpeHabCajDevABoveME
        frmCajeroHab.Show 1
        'Case gOpeHabCajTransfEfectCajerosMN, gOpeHabCajTransfEfectCajerosME
    Case 1
        frmCajeroHab.Show 1
End Select
End Sub

Private Sub M0401000001_Click(Index As Integer)
Select Case Index
  Case 0
        If M0401000001(0).Checked Then
            Timer1.Enabled = False
            M0401000001(0).Checked = False
        ElseIf M0401000001(0).Checked = False Then
            Timer1.Enabled = True
            M0401000001(0).Checked = True
        End If
  Case 1
        FrmCapAutOpeEstados.Show
  End Select
End Sub

Private Sub M0401030000_Click(Index As Integer)
    Select Case Index
        Case 0
            Call frmCierreDiario.CierreDia
        Case 1
            Call frmCierreDiario.CierreMes
        Case 2
        
        Case 3
            
    End Select
            
End Sub

Private Sub M0401040000_Click(Index As Integer)
    Select Case Index
        Case 0 'Captaciones
            frmCapExtornos.Show 1
        
            
        Case 1 '    Extorno Credito
        
        Case 2 '    Pignoraticio
            'frmColPExtornoOpe.Show 1
        Case 3 '    recuperaciones
            
    End Select
End Sub

Private Sub M0401040100_Click(Index As Integer)
    Select Case Index
        Case 0
            frmCredExtornos.Show 1
        Case 1
            frmCredExtornos.Show 1
        Case 2
            frmCredExtPagoLote.Show 1
    End Select
End Sub

Private Sub M0401050000_Click(Index As Integer)
    Select Case Index
    Case 0
        frmAsientoDN.Inicio True
    Case 1
        frmAsientoDN.Inicio False
    Case 2
        frmCtaContMantenimiento.Show 1
    End Select
End Sub

Private Sub M0401060000_Click(Index As Integer)
Dim sCad As String
Dim oPrevio As clsPrevio
Select Case Index
    Case 0
        frmCajeroIngEgre.Inicia False, True
    Case 1
        Dim oRep As nCaptaReportes
        
        Set oRep = New nCaptaReportes
        sCad = oRep.ReporteTrasTotSM("DETALLE DE OPERACIONES", False, gsCodUser, Format$(gdFecSis, "yyyymmdd"), gsCodAge)
        Set oRep = Nothing
        
        Set oPrevio = New clsPrevio
        oPrevio.Show sCad, "DETALLE DE OPERACIONES", True
        Set oPrevio = Nothing
    Case 2
        Dim oHab As NCajeroImp
        
        Usuario.Inicio gsCodUser
        Set oHab = New NCajeroImp
        sCad = oHab.ReporteHabilitacionDevolucion(gsCodUser, Usuario.AreaCod, gsCodAge, gdFecSis, Usuario.UserNom, gsNomAge)
        
        Set oPrevio = New clsPrevio
        oPrevio.Show sCad, "DETALLE DE HABILITACIONES Y DEVOLUCIONES", True
        Set oPrevio = Nothing
    Case 3
        Dim oProt As nCaptaReportes
        Set oProt = New nCaptaReportes
        sCad = oProt.ProtocoloOperaciones("PROTOCOLO DE USUARIO SOLES", 0, 0, gsNomAge, gcEmpresa, gdFecSis, gMonedaNacional, gsCodUser, , Format(gdFecSis, gsFormatoFechaView), gsCodAge)
        sCad = sCad & oProt.ProtocoloOperaciones("PROTOCOLO DE USUARIO DOLARES", 0, 0, gsNomAge, gcEmpresa, gdFecSis, gMonedaExtranjera, gsCodUser, , Format(gdFecSis, gsFormatoFechaView), gsCodAge)
        
        Set oPrevio = New clsPrevio
        oPrevio.Show sCad, "PROTOCOLO DE USUARIO", True
        Set oPrevio = Nothing
    Case 4
        frmOperacionesNum.Show 1
End Select
End Sub

Private Sub M0501030000_Click(Index As Integer)
Select Case Index
    Case 0
        frmCapServGenReporte.Inicia gCapServHidrandina
    Case 1
        frmCapServGenReporte.Inicia gCapServSedalib
    Case 2
        frmCapServGenReporte.Inicia gCapServEdelnor
End Select
End Sub

Private Sub M0501040000_Click(Index As Integer)
Select Case Index
    Case 0
        frmCapServParametros.Inicia (gCapServHidrandina)
    Case 1
        frmCapServParametros.Inicia (gCapServSedalib)
    Case 2
        frmCapServParametros.Inicia (gCapServFideicomiso)
    Case 3
        frmCapServParametros.Inicia (gCapServEdelnor)
End Select
End Sub

Private Sub M0501050000_Click(Index As Integer)
Select Case Index
    Case 0
        FrmServSat.Show 1
    Case 1
       'frmServDisfondos.Inicia gCapServSATTInfraccion
End Select
End Sub

Private Sub M0601000000_Click(Index As Integer)
    Select Case Index
        Case 0 'Parametros
            
        Case 1 'Cosnsistema
             
        Case 2 'permisos
            frmMantPermisos.Show 1
        Case 3 'zonas
            frmMntZonas.Show 1
        Case 4 'Agencias
            frmMntAgencias.Show 1
        Case 5 'Ctas Contables
            frmCtaContMantenimiento.Show 1
        Case 6 'backUp
        Case 7
            frmCajeroGrupoOpe.Show 1
        Case 8
            frmCapMantOperacion.Show 1
        Case 9
            frmMantCodigoPostal.Show 1
        Case 10
            frmDocRecParam.Show 1
        Case 11
            frmMantCIIU.Show 1
        Case 12
            FrmMantFeriados.Show 1
    End Select
End Sub

Private Sub M0701000000_Click(Index As Integer)
    If Index = 3 Then
        frmPosicionCli.Show 1
    End If
    If Index = 6 Then  ' Reportes
        FrmPersReporte.Show 1
    End If
End Sub

Private Sub M0701010000_Click(Index As Integer)
    'Persona
    Select Case Index
        Case 0 'Registro
            frmPersona.Registrar
        Case 1 'mantenimiento
            frmPersona.Mantenimeinto
        Case 2 'Consulta
            frmPersona.Consultar
        Case 3 'Exoneradas del Lavado de Dinero
            frmPersLavDinero.Show 1
        Case 4 'Rol de Persona
            FrmPersonaRolMantenimiento.Show 1
        Case 5
            frmPersComentario.Show 1
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

Private Sub M0801000000_Click(Index As Integer)
    Select Case Index
        'Herramientas
        Case 1
            frmSpooler.Show 1
        Case 2
            frmSetupCOM.Show 1
    End Select
End Sub

Private Sub MDIForm_Click()
'  Form1.Show
End Sub

Private Sub MDIForm_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub


Private Sub AvisoOperacionesPendientes()

    
    
    Dim lsCadAux As String
    Dim lsSQL As String
    Dim coConex As DConecta
    Dim lsCola As String
    Dim lrst As ADODB.Recordset
    
    Set lrst = New ADODB.Recordset
    Set coConex = New DConecta
    coConex.AbreConexion

    
    Dim lrOperaciones As ADODB.Recordset
                        
                                                
     lsCola = " where (nAutEstado = " & gAhoEstAprobOpeAutorizado & _
        "  and (select count(nMovNro) from MovRef where MovRef.nMovNro = Mcao.nMovNro)=1 and  Mcao.cUserOri ='" & gsCodUser & "')" & _
        " or (nAutEstado = " & gAhoEstAprobOpeRechazado & " and    Mcao.cUserOri ='" & gsCodUser & "')  and cHabilitado ='S'"
        
'     lsCola = " where (nAutEstado = " & gAhoEstAprobOpeRechazado & " and    Mcao.cUserOri ='" & gsCodUser & "')"
'
'
'     lsCola = " where (nAutEstado = " & gAhoEstAprobOpeAutorizado & _
'        "  and (select count(nMovNro) from MovRef where MovRef.nMovNro = Mcao.nMovNro)=1 and  Mcao.cUserOri ='" & gsCodUser & "')"

                                    
    lsSQL = "Select cCtaCod,(select cPersNombre from Persona where Persona.cPerscod = Mcao.cPersCodCli)   as  cPersNombre," & _
    "(select cOpeDesc + replicate(' ',100) + OpeTpo.cOpeCod from OpeTpo where OpeTpo.cOpeCod = Mcao.cOpeCod  ) cOpeDes," & _
    "(select cOpeDesc from OpeTpo where OpeTpo.cOpeCod = Mcao.cOpeCodOri  ) cOpeDesOri," & _
    "(select case Mcao.nCodMon when 1 then 'S/.' + replicate(' ',100) + '1'   when 2 then '$.'+ replicate(' ',100) + '2' end) Monedas,Mcao.nMontoAprobado,  " & _
    "(select cConsDescripcion from constante where nConsCod = 9051  and  constante.nConsvalor = Mcao.nAutEstado) sAutEstadoDes," & _
    "Mcao.cAutObs, Mcao.cUltimaActualizacion UltAct, Mcao.nMovNro nMov , Mcao.cOpeCodOri OpeOri, Mcao.cUserOri UserOri " & _
    "From MovCapAutorizacionOpe  As Mcao" & lsCola
            
        
    Set lrst = coConex.CargaRecordSet(lsSQL)
    
    If lrst.EOF Then
        MsgBox "No Existen Operaciones Pendientes Autorizadas", vbInformation, "Aviso"
        Exit Sub
    End If
    
    'MsgBox lrst.RecordCount
    lsCadAux = ""
    
    lrst.MoveFirst
    Do While Not lrst.EOF
        'MsgBox lrst("cOpeDesOri")
        lsCadAux = lsCadAux + PstaNombre(Trim(lrst("cPersNombre")), True) + Space(10) + Trim(lrst("cOpeDesOri")) + Space(10) + Trim(lrst("sAutEstadoDes")) + Chr(13)
        lrst.MoveNext
    Loop
    MsgBox lsCadAux, vbInformation, "Operaciones Pendientes"


End Sub

Private Sub Timer1_Timer()
Dim sSql As String, rs As ADODB.Recordset
Dim oconecta As DConecta
   
   On Error GoTo MensaError
   
   
   sSql = "Select Sum(case when cautestado='A' then 1 else 0 end ) Aprobadas,Sum(case when cautestado='R' then 1 else 0 end ) Rechazadas,Sum(case when cautestado='P' then 1 else 0 end ) Pendientes "
   sSql = sSql & " From capautorizacionope "
   sSql = sSql & " where cautestado<>'E' and cuserori='" & Vusuario & "' and left(cultimaactualizacion,8)=convert(char(8),getdate(),112) "
   
   Set oconecta = New DConecta
   Set rs = New ADODB.Recordset
     oconecta.AbreConexion
     rs.CursorLocation = adUseClient
     Set rs = oconecta.CargaRecordSet(sSql)
     oconecta.CierraConexion
     Set oconecta = Nothing
     If rs.State = 1 Then
          If (rs!Aprobadas > 0 Or rs!Rechazadas > 0 Or rs!Pendientes > 0) Then
               Toolbar1.Visible = True
               txtEstado1.Text = "Aprobadas: " & CStr(rs!Aprobadas)
               txtEstado2.Text = "Rechazadas: " & CStr(rs!Rechazadas)
               txtEstado3.Text = "Pendientes: " & CStr(rs!Pendientes)
          Else
               Toolbar1.Visible = False
          End If
          If rs.State = 1 Then rs.Close
     End If
        Set rs = Nothing
        
'  Dim i As Long
'       i = 0
'       For i = 0 To Timer1.Interval
'         If i = Timer1.Interval Then
'            '  Unload FrmCapAutOpeEstados
'            '  FrmCapAutOpeEstados.Show 1
'         End If
'         If i = (Timer1.Interval / 2) Then
'             '  Unload FrmCapAutOpeEstados
'         End If
'       Next i
Exit Sub
MensaError:
     Call RaiseError(MyUnhandledError, "frmCapAutorizacionPrueba: CargaOperaciones  Method")
End Sub

