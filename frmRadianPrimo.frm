VERSION 5.00
Begin VB.Form frmRadianPrimo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   Caption         =   "Números Primos sobre Circunferencia"
   ClientHeight    =   11190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18960
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   6.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11190
   ScaleWidth      =   18960
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame9 
      BackColor       =   &H00808080&
      Caption         =   "Otros "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13320
      TabIndex        =   77
      Top             =   10200
      Width           =   5415
      Begin VB.CheckBox chkComprimido 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Comprimido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3720
         TabIndex        =   80
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdUlam 
         Caption         =   "ULAM"
         Height          =   495
         Left            =   240
         TabIndex        =   79
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkSoloPuntos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Solo Puntos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1680
         TabIndex        =   78
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00808080&
      Caption         =   "Efecto 3D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   13320
      TabIndex        =   68
      Top             =   8880
      Width           =   5415
      Begin VB.CheckBox chkVectorPosicion 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vector Posición"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   4080
         TabIndex        =   76
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtProfundidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   75
         Text            =   "20"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtCentoZ 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   74
         Text            =   "5500"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtAngulo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   73
         Text            =   "25"
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Altura"
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
         Left            =   1440
         TabIndex        =   71
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Distancia Órbitas"
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
         Left            =   2760
         TabIndex        =   70
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Posición"
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
         Left            =   4080
         TabIndex        =   69
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Ángulo"
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
         Left            =   120
         TabIndex        =   72
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CheckBox chkBaseDatos 
      Caption         =   "Base de Datos"
      Height          =   255
      Left            =   17280
      TabIndex        =   67
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00808080&
      Caption         =   "Ángulos "
      Height          =   1095
      Left            =   13320
      TabIndex        =   33
      Top             =   4080
      Width           =   5415
      Begin VB.TextBox txtAnguloAbsoluto 
         Alignment       =   1  'Right Justify
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
         Left            =   4080
         TabIndex        =   55
         Text            =   "  .351562"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtAnguloDeterminado 
         Alignment       =   1  'Right Justify
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
         Left            =   1440
         TabIndex        =   39
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdAnguloDeterminado 
         Caption         =   "Determinado"
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
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cboAngulo 
         Height          =   285
         Left            =   2760
         TabIndex        =   37
         Text            =   "0"
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdMas 
         Caption         =   "+"
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
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdMenos 
         Caption         =   "--"
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
         Left            =   720
         TabIndex        =   35
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtDelta 
         Alignment       =   1  'Right Justify
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
         Left            =   1440
         TabIndex        =   34
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Absoluto"
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
         Left            =   4080
         TabIndex        =   56
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00808080&
      Caption         =   "Puntos "
      Height          =   1095
      Left            =   13320
      TabIndex        =   30
      Top             =   1680
      Width           =   5415
      Begin VB.TextBox txtMin 
         Alignment       =   1  'Right Justify
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
         Left            =   4080
         TabIndex        =   47
         Text            =   "888"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtMax 
         Alignment       =   1  'Right Justify
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
         Left            =   4680
         TabIndex        =   46
         Text            =   "907"
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdRango 
         Caption         =   "Rango"
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
         Left            =   4080
         TabIndex        =   45
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtTamañoPunto 
         Alignment       =   1  'Right Justify
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
         Left            =   1440
         TabIndex        =   44
         Text            =   "25"
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrimoAnterior 
         Caption         =   "Primo Anterior"
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
         Left            =   2760
         TabIndex        =   41
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrimoSiguiente 
         Caption         =   "Primo Siguiente"
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
         Left            =   2760
         TabIndex        =   40
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdCompuestos 
         Caption         =   "Compuestos"
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
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrimos 
         Caption         =   "Primos"
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
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Tamaño Puntos"
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
         Left            =   1440
         TabIndex        =   43
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00808080&
      Caption         =   "Captura Imagen "
      Height          =   1095
      Left            =   13320
      TabIndex        =   24
      Top             =   6480
      Width           =   5415
      Begin VB.CommandButton cmdGuardarImagen 
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
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtRuta 
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
         Left            =   1440
         TabIndex        =   28
         Text            =   "C:\ImagenesPrimos\"
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtArchivo 
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
         Left            =   1440
         TabIndex        =   27
         Text            =   "NumerosPrimos.bmp"
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox txtAncho 
         Alignment       =   1  'Right Justify
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
         Left            =   720
         TabIndex        =   26
         Text            =   "9640"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtAlto 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   25
         Text            =   "9960"
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      Caption         =   "Animación "
      Height          =   1095
      Left            =   13320
      TabIndex        =   19
      Top             =   5280
      Width           =   5415
      Begin VB.CommandButton cmdGrabaSecuencia 
         Caption         =   "Grabar Secuencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4080
         TabIndex        =   66
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtIncremento 
         Alignment       =   1  'Right Justify
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
         Left            =   1440
         TabIndex        =   23
         Text            =   ".005"
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdAnimacion 
         Caption         =   "Mas"
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
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtAnimacion 
         Alignment       =   1  'Right Justify
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
         Left            =   2760
         TabIndex        =   21
         Text            =   "1200"
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdAnimacionMenos 
         Caption         =   "Menos"
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
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Pasos"
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
         Left            =   2760
         TabIndex        =   65
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Incremento"
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
         Left            =   1440
         TabIndex        =   64
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "Líneas "
      Height          =   1095
      Left            =   13320
      TabIndex        =   12
      Top             =   2880
      Width           =   5415
      Begin VB.CommandButton cmdPG 
         Caption         =   "Inter Gemelos"
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
         Left            =   2760
         TabIndex        =   61
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtPrimosGemelos 
         Alignment       =   1  'Right Justify
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
         Left            =   2760
         TabIndex        =   60
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtGap 
         Alignment       =   1  'Right Justify
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
         Left            =   4080
         TabIndex        =   18
         Text            =   "4"
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdUneGap 
         Caption         =   "GAP"
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
         Left            =   4080
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdUne 
         Caption         =   "Gemelos"
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
         Left            =   1440
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdRama 
         Caption         =   "Ramas"
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
         Left            =   1440
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdLineasCompuestos 
         Caption         =   "Compuestos"
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
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdLineas 
         Caption         =   "Primos"
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
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Información "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   13320
      TabIndex        =   8
      Top             =   7680
      Width           =   5415
      Begin VB.TextBox txtAviso 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   63
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdEstadisticaOrbita 
         Caption         =   "Estadística Órbitas"
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
         Left            =   4080
         TabIndex        =   59
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "C"
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
         Left            =   4080
         TabIndex        =   58
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtOrbitaCompuesto 
         Alignment       =   1  'Right Justify
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
         Left            =   4440
         TabIndex        =   57
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "P"
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
         Left            =   2760
         TabIndex        =   54
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdListaOrbitas 
         Caption         =   "Lista Órbitas"
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
         Left            =   2760
         TabIndex        =   53
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtOrbitaPrimo 
         Alignment       =   1  'Right Justify
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
         Left            =   3120
         TabIndex        =   52
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtNumeroOrbita 
         Alignment       =   1  'Right Justify
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
         Left            =   1440
         TabIndex        =   51
         Text            =   "2"
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdInformacionOrbita 
         Caption         =   "Información Órbita"
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
         Left            =   1440
         TabIndex        =   50
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtCantidadPrimos 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Cantidad Primos"
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
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Controles "
      Height          =   1455
      Left            =   13320
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton cmdEjes 
         Caption         =   "Ejes"
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
         Left            =   2760
         TabIndex        =   62
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkInverso 
         Caption         =   "Inverso"
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
         Left            =   2760
         TabIndex        =   49
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtInverso 
         Alignment       =   1  'Right Justify
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
         Left            =   4080
         TabIndex        =   48
         Text            =   "34"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEtiqueta 
         Caption         =   "Etiquetas"
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
         Left            =   4080
         TabIndex        =   42
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdColorFondo 
         Caption         =   "Color Fondo"
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
         Left            =   4080
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdOrbita 
         Caption         =   "Órbita"
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
         Left            =   2760
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdLejos 
         Caption         =   "Lejos"
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
         Left            =   1440
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdCerca 
         Caption         =   "Cerca"
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
         Left            =   1440
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrafica 
         Caption         =   "Mostrar"
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
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtN 
         Alignment       =   1  'Right Justify
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
         Left            =   1440
         TabIndex        =   3
         Text            =   "4096"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdDoble 
         Caption         =   "Doble"
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
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdMitad 
         Caption         =   "Mitad"
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
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmRadianPrimo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'* PROYECTO      : RADIAN PRIMOS
'* CONTENIDO     : CALCULAR NÚMEROS PRIMOS, MOSTRARLOS POR NIVELES EN CIRCUNFERENCIAS
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO / MIGUEL QUINTEIRO FERNANDEZ
'* INICIO        : 16 DE MAYO DE 2017
'* ACTUALIZACION : 16 DE MAYO DE 2017
'****************************************************************************************
Option Explicit

' Declaración de variables
Dim CentroX As Double
Dim CentroY As Double
Dim CentroZ As Double
Dim RadioEsfera As Double

' Declaración de variables
Dim miFactorCircular As Double
Dim miPi As Double
Dim X1 As Double
Dim Y1 As Double
Dim X2 As Double
Dim Y2 As Double

Dim Z1 As Double

Dim X3 As Double
Dim Y3 As Double
Dim X5 As Double
Dim Y5 As Double

Dim miRadio As Long
Dim min As Long
Dim miCuentaPrimos As Long
Dim miCuentaSuperior As Long
Dim miCuentaInferior As Long
Dim miMiniRadio As Long
Dim r As Integer
Dim miZoom As Long
Dim miTamañoPunto As Integer

Dim miEtiqueta As Boolean
Dim miCompuestos As Boolean
Dim miPrimos As Boolean
Dim miLineasP As Boolean
Dim miLineasC As Boolean
Dim miEjes As Boolean
Dim miUne As Boolean
Dim miUneGap As Boolean
Dim miOrbita As Boolean
Dim miPG As Boolean
Dim miRama As Boolean
Dim miRango As Boolean

Dim miOrbitaMaxima As Long
Dim miDelta As Double

Dim miColorFondo As Integer

' Declaración de arreglos
Dim miOrbitaP() As Long
Dim miOrbitaPG() As Boolean
Dim miOrbitaC() As Long

Dim miNumeros() As Long
Dim miX() As Long
Dim miY() As Long

' AL CARGAR EL FORMULARIO
Private Sub Form_Load()

' Comunicacion con la base de datos
  Set cn = New ADODB.Connection
  ' Conectar a la base de datos
  cn.Open _
      "Provider=sqloledb;" & _
                           "Data Source=LAPTOPMIGUEL\SQLEXPRESS;" & _
                           "Initial Catalog=NumerosPrimos;" & _
                           "Trusted_Connection=yes;"

  ' Inicialización de variable
  miPi = 3.1415926535
  miRadio = 4000
  miFactorCircular = 1.15
  min = 2
  miMiniRadio = 1
  miZoom = 20
  miOrbitaMaxima = 200
  miColorFondo = 15

  miDelta = 2.8125                    ' Para 128
  'miDelta = 0.5235987756             ' Reduccion
  'miDelta = 1.66016181584687E-03     ' Phi
  'miDelta = 2.718281828              ' e
  'miDelta = 14.13472514              ' Primer cero no trivial de la función Zeta de Riemann
  'miDelta = 29.9999995130823         ' A quince grados
  'miDelta = 57.2957795130823         ' Un Radian
  'miDelta = 69.1117795130823         ' Un Radian

  cboAngulo.AddItem "0.00000000000000"    ' Cero
  cboAngulo.AddItem "0.02197200000000"    ' Para 16384
  cboAngulo.AddItem "0.52359877560000"    ' Reduccion
  cboAngulo.AddItem "1.66016181584687"    ' Phi
  cboAngulo.AddItem "2.71828182800000"    ' Numero e
  cboAngulo.AddItem "14.1347251400000"    ' Primer cero no trivial de la función Zeta de Riemann
  cboAngulo.AddItem "29.9999995130823"    ' A quince grados
  cboAngulo.AddItem "57.2957795130823"    ' Un Radian
  cboAngulo.AddItem "137.500000000000"    ' Proporción aurea

  miEtiqueta = False
  miOrbita = True
  miCompuestos = True
  miPrimos = True
  miLineasP = False
  miLineasC = False
  miEjes = False
  miUne = False
  miUneGap = False
  miPG = False
  miRama = False
  miRango = False

  ReDim miOrbitaP(miOrbitaMaxima)
  ReDim miOrbitaPG(miOrbitaMaxima)
  ReDim miOrbitaC(miOrbitaMaxima)
End Sub

' AL DARLE DOBLE CLICK
Private Sub Form_DblClick()

' Dibuja Circulo
  min = InputBox("Ingrese el número N (Entre 1 y 30000)")
  Call Grafica
End Sub


'**************************************************************************************
' CONTROLES
'**************************************************************************************
' Pinta los puntos en la circunferencia
Private Sub cmdGrafica_Click()
  min = txtN
  Call Grafica
  'Call GraficaProyeccion
End Sub

' Reduce a la mitad la cantidad de puntos mostrados
Private Sub cmdMitad_Click()
  txtN.Text = Val(txtN.Text) / 2
  min = txtN
  Call Grafica
  'Call GraficaProyeccion
  Call Label2_Click
End Sub

' Duplica la cantidad de puntos mostrados
Private Sub cmdDoble_Click()
  txtN.Text = Val(txtN.Text) * 2
  min = txtN
  Call Grafica
  'Call GraficaProyeccion
  Call Label2_Click
End Sub

' Aleja la imagen de los puntos
Private Sub cmdLejos_Click()
  miZoom = miZoom * 1.1
  Call Grafica
  'Call GraficaProyeccion
End Sub

' Acerca la imagen de los puntos
Private Sub cmdCerca_Click()
  miZoom = miZoom / 1.1
  Call Grafica
  'Call GraficaProyeccion
End Sub

' Muestra las etiquetas de los números
Private Sub cmdOrbita_Click()
  If miOrbita = True Then
    miOrbita = False
  Else
    miOrbita = True
  End If
  Call Grafica
  'Call GraficaProyeccion
End Sub

' Muestra los ejes de coordenadas
Private Sub cmdEjes_Click()
  If miEjes = True Then
    miEjes = False
  Else
    miEjes = True
  End If
  Call Grafica
  'Call GraficaProyeccion
End Sub

' Color de fondo
Private Sub cmdColorFondo_Click()
  If miColorFondo < 15 Then
    miColorFondo = miColorFondo + 1
  Else
    miColorFondo = 0
  End If
  frmRadianPrimo.BackColor = QBColor(miColorFondo)
  frmRadianPrimo.Refresh
  DoEvents
End Sub

' Muestra las etiquetas de los números
Private Sub cmdEtiqueta_Click()
  If miEtiqueta = True Then
    miEtiqueta = False
  Else
    miEtiqueta = True
  End If
  Call Grafica
  'Call GraficaProyeccion
End Sub

'**************************************************************************************

'**************************************************************************************
' PUNTOS
'**************************************************************************************
' Muestra los números compuestos
Private Sub cmdPrimos_Click()
  If miPrimos = True Then
    miPrimos = False
  Else
    miPrimos = True
  End If
  Call Grafica
  'Call GraficaProyeccion
End Sub

' Muestra los números compuestos
Private Sub cmdCompuestos_Click()
  If miCompuestos = True Then
    miCompuestos = False
  Else
    miCompuestos = True
  End If
  Call Grafica
  'Call GraficaProyeccion
End Sub

' Busca primo anterior
Private Sub cmdPrimoAnterior_Click()
  Dim miDato As Long
  miDato = Val(txtN.Text)
  min = txtN
  If miDato > 2 Then
    miDato = miDato - 1
    While Not Primo(miDato)
      miDato = miDato - 1
    Wend
    Cls
    txtN.Text = miDato
    Call Grafica
    'Call GraficaProyeccion
  End If
  Call Label2_Click
  DoEvents
End Sub

' Busca primo siguiente
Private Sub cmdPrimoSiguiente_Click()
  Dim miDato As Long
  miDato = Val(txtN.Text)
  min = txtN
  If miDato > 2 Then
    miDato = miDato + 1
    While Not Primo(miDato)
      miDato = miDato + 1
    Wend
    Cls
    txtN.Text = miDato
    Call Grafica
    'Call GraficaProyeccion
  End If
  Call Label2_Click
  DoEvents
End Sub

' Muestra rango de valores
Private Sub cmdRango_Click()
  If miRango = True Then
    miRango = False
  Else
    miRango = True
  End If
  Call Grafica
  'Call GraficaProyeccion
End Sub

'**************************************************************************************

'**************************************************************************************
' LÍNEAS
'**************************************************************************************
' Muestra la línea de los primos
Private Sub cmdLineas_Click()
  If miLineasP = True Then
    miLineasP = False
  Else
    miLineasP = True
  End If
  Call Grafica
  'Call GraficaProyeccion
End Sub

' Muestra las líneas de los compuestos
Private Sub cmdLineasCompuestos_Click()
  If miLineasC = True Then
    miLineasC = False
  Else
    miLineasC = True
  End If
  Call Grafica
  'Call GraficaProyeccion
End Sub

' Muestra las líneas de los primos gemelos
Private Sub cmdUne_Click()
  If miUne = True Then
    miUne = False
  Else
    miUne = True
  End If
  Call Grafica
  'Call GraficaProyeccion
End Sub

'Muestra las líneas de ramas
Private Sub cmdRama_Click()
  If miRama = True Then
    miRama = False
  Else
    miRama = True
  End If
  Call Grafica
  'Call GraficaProyeccion
End Sub

' Muestra la línea de los primos gemelos internos
Private Sub cmdPG_Click()
  If miPG = True Then
    miPG = False
  Else
    miPG = True
  End If
  Call Grafica
  'Call GraficaProyeccion
End Sub

' Muestra las líneas de los primos por Gap
Private Sub cmdUneGap_Click()
  If miUneGap = True Then
    miUneGap = False
  Else
    miUneGap = True
  End If
  Call Grafica
  'Call GraficaProyeccion
End Sub

'**************************************************************************************

'**************************************************************************************
' ÁNGULOS
'**************************************************************************************
' Incrementa el Delta del desplazamiento
Private Sub cmdMas_Click()
'miDelta = CambiaComa(Val(txtDelta.Text))
  miDelta = miDelta + Val(txtIncremento.Text)
  'miDelta = miDelta + 0.01 + Val(txtIncremento.Text)
  txtDelta.Text = miDelta

  Call Grafica
  'Call GraficaProyeccion
End Sub

' Decrementa el Delta del desplazamiento
Private Sub cmdMenos_Click()
  miDelta = miDelta - Val(txtIncremento.Text)
  'miDelta = miDelta - 0.01
  txtDelta.Text = miDelta

  Call Grafica
  'Call GraficaProyeccion
End Sub

' Colocar un ángulo determinado
Private Sub cmdAnguloDeterminado_Click()
  miDelta = Val(txtAnguloDeterminado.Text)
  Call Grafica
  'Call GraficaProyeccion
End Sub

' Angulo Absoluto
Private Sub Label2_Click()
  miDelta = Val(txtAnguloAbsoluto.Text)
  Call Grafica
  'Call GraficaProyeccion
End Sub

'**************************************************************************************

'**************************************************************************************
' ANIMACIÓN
'**************************************************************************************
' Animación positiva
Private Sub cmdAnimacion_Click()
  Dim i As Long
  Dim r As Long
  'txtIncremento.Text = 0.0001
  'txtIncremento.Text = CambiaComa(txtIncremento.Text)
  For i = 1 To Val(txtAnimacion.Text)
    'miDelta = CambiaComa(Val(txtDelta.Text))
    miDelta = miDelta + Val(txtIncremento.Text)
    'miDelta = miDelta + 0.01 + Val(txtIncremento.Text)
    txtDelta.Text = miDelta

    Call Grafica
    'Call GraficaProyeccion

    For r = 1 To 200000
    Next r
    DoEvents
  Next i
End Sub

' Animación negativa
Private Sub cmdAnimacionMenos_Click()
  Dim i As Long
  Dim r As Long
  'txtIncremento.Text = 0.0001
  'txtIncremento.Text = CambiaComa(txtIncremento.Text)
  For i = 1 To Val(txtAnimacion.Text)
    miDelta = miDelta - Val(txtIncremento.Text)
    'miDelta = miDelta - 0.01
    txtDelta.Text = miDelta

    Call Grafica
    Call GraficaProyeccion

    For r = 1 To 200000
    Next r
    DoEvents
  Next i
End Sub

' Calcula la revolución según el incremento
Private Sub Label5_Click()
  txtAnimacion.Text = 360 / Val(txtIncremento.Text)
End Sub

Private Sub cmdGrabaSecuencia_Click()
  Dim i As Long
  Dim r As Long
  'txtIncremento.Text = 0.0001
  'txtIncremento.Text = CambiaComa(txtIncremento.Text)
  For i = 1 To Val(txtAnimacion.Text)
    'miDelta = CambiaComa(Val(txtDelta.Text))
    miDelta = miDelta + Val(txtIncremento.Text)
    'miDelta = miDelta + 0.01 + Val(txtIncremento.Text)
    txtDelta.Text = miDelta

    Call Grafica
    Call GraficaProyeccion

    For r = 1 To 200000
    Next r

    txtRuta.Text = "C:\ImagenesPrimos\Secuencia\"
    txtArchivo.Text = "Secuencia " & Trim(Str(i)) & ".bmp"

    Call cmdGuardarImagen_Click

    DoEvents
  Next i
  txtArchivo.Text = "NumerosPrimos.bmp"
  txtRuta.Text = "C:\ImagenesPrimos\"
End Sub

'**************************************************************************************

'**************************************************************************************
' CAPTURA IMAGEN
'**************************************************************************************
' GUARDAR EL FORMULARIO COMO IMAGEN
Private Sub cmdGuardarImagen_Click()
' Ajusta el tamaño del formulario
  Me.Height = Val(txtAlto.Text)   '9960
  Me.Width = Val(txtAncho.Text)   '9640
' Borra el portapapeles
  Clipboard.Clear
  DoEvents
  DoEvents
  ' Manda la pulsación de teclas para capturar la imagen de la pantalla
  On Error Resume Next
  Call keybd_event(&H2C, 1, 0, 0)
  DoEvents
  SavePicture Clipboard.GetData(vbCFBitmap), Trim(txtRuta.Text & txtArchivo.Text)
  DoEvents
  ' Ajusta el tamaño del formulario
  Me.Height = 9960
  Me.Width = 16050
End Sub

'**************************************************************************************

'**************************************************************************************
' INFORMACIÓN
'**************************************************************************************
' Muestra información de la orbita
Private Sub cmdInformacionOrbita_Click()
  Dim i As Long
  If Val(txtNumeroOrbita.Text) <> 0 Then
    txtOrbitaPrimo.Text = miOrbitaP(Val(txtNumeroOrbita.Text))
    txtOrbitaCompuesto.Text = miOrbitaC(Val(txtNumeroOrbita.Text))
    ' Marca la orbita con líneas
    Call Grafica
    If ((miRadio * (Val(txtNumeroOrbita.Text) / miZoom) * miFactorCircular) - 30) > 0 Then
      Circle (4750, 4750), (miRadio * (Val(txtNumeroOrbita.Text) / miZoom) * miFactorCircular) - 30, vbWhite
      Circle (4750, 4750), (miRadio * (Val(txtNumeroOrbita.Text) / miZoom) * miFactorCircular) + 30, vbWhite
    End If

    txtAviso.Text = miOrbitaPG(Val(txtNumeroOrbita.Text))
  End If
End Sub

' Información sobre las órbitas
Private Sub cmdListaOrbitas_Click()
  Dim miMensaje As String
  Dim o As Integer
  miMensaje = ""
  miMensaje = miMensaje + "****   Listado Órbitas   ****" + vbCrLf + vbCrLf
  miMensaje = miMensaje + "   #      P       C" + vbCrLf
  For o = 0 To miOrbitaMaxima
    miMensaje = miMensaje + Tabulado(Trim(Str(o)), 7) + _
                Tabulado(Trim(Str(miOrbitaP(o))), 7) + _
                Tabulado(Trim(Str(miOrbitaPG(o))), 7) + _
                Tabulado(Trim(Str(miOrbitaC(o))), 7) + vbCrLf
    If miOrbitaP(o) = 0 And miOrbitaC(o) = 0 Then
      o = miOrbitaMaxima + 1
    End If
  Next o
  MsgBox miMensaje, , "Información Orbitas"
End Sub

' Estadística de la orbita
Private Sub cmdEstadisticaOrbita_Click()
' Abre archivo para escritura
  Open "EstadisticaOrbita.txt" For Output As #1

  Dim i As Integer
  Dim acumP As Long
  Dim acumC As Long

  Dim relacion As Double
  Dim PorcentajePrimo As Double
  Dim PorcentajeCompuesto As Double

  Print #1, "Estadística para .--- "; min
  Print #1, ""
  Print #1, "    O.     P.     C.                 %P.     %C.     C/P."
  Print #1, "---------------------------------------------------------"

  acumP = 0
  acumC = 0
  For i = 0 To miOrbitaMaxima
    acumP = acumP + miOrbitaP(i)
    acumC = acumC + miOrbitaC(i)

    If (miOrbitaP(i) + miOrbitaC(i)) > 0 Then
      PorcentajePrimo = (miOrbitaP(i) * 100) / (miOrbitaP(i) + miOrbitaC(i))
      PorcentajeCompuesto = (miOrbitaC(i) * 100) / (miOrbitaP(i) + miOrbitaC(i))
    Else
      PorcentajePrimo = 0
      PorcentajeCompuesto = 0
    End If

    Print #1, Tabulado(Trim(Str(i)), 5); "  "; _
              Tabulado(Trim(Str(miOrbitaP(i))), 5); "  "; _
              Tabulado(Trim(Str(miOrbitaC(i))), 5); "  "; _
              "   ***    "; "  "; _
              Tabulado(Trim(Format(PorcentajePrimo, "##,##0.00")), 6); "  "; _
              Tabulado(Trim(Format(PorcentajeCompuesto, "##,##0.00")), 6); "  ";
    If miOrbitaP(i) > 0 Then
      relacion = miOrbitaC(i) / miOrbitaP(i)
      Print #1, Tabulado(Trim(Format(relacion, "##,##0.00")), 8); "   ";
      Print #1, Tabulado(Trim(Str(miOrbitaPG(i))), 5)
    Else
      Print #1, "--------"; "   ";
      Print #1, Tabulado(Trim(Str(miOrbitaPG(i))), 5)
    End If

    If (miOrbitaP(i) = 0) And (miOrbitaC(i) = 0) Then
      Print #1, "---------------------------------------------------------"
      Print #1, "T.  "; Tabulado(Trim(Str(acumP)), 8); " "; Tabulado(Trim(Str(acumC)), 8); "            ";

      PorcentajePrimo = (acumP * 100) / min
      PorcentajeCompuesto = (acumC * 100) / min
      Print #1, Tabulado(Trim(Format(PorcentajePrimo, "##,##0.00")), 6); "  "; _
                Tabulado(Trim(Format(PorcentajeCompuesto, "##,##0.00")), 6); "  ";
      If acumP <> 0 Then
        relacion = acumC / acumP
      End If
      Print #1, Tabulado(Trim(Format(relacion, "##,##0.00")), 8)

      i = miOrbitaMaxima + 1
    End If
  Next i
  ' Cierra archivo
  Close #1
End Sub

'**************************************************************************************



'**************************************************************************************
' OTROS
'**************************************************************************************
Private Sub cmdUlam_Click()
  Dim i As Long
  Dim m As Long
  Dim c As Long
  Dim posX As Long
  Dim posY As Long
  Dim posZ As Long
  Dim des As Long
  Dim Expande As Long

  Dim Derecha As Boolean
  Dim Arriba As Boolean
  Dim Izquierda As Boolean
  Dim Abajo As Boolean

  posX = 5800
  posY = 5800
  posZ = 1

  If chkComprimido.Value = 1 Then
    Expande = 150
    des = 75
  Else
    Expande = 30
    des = 350
  End If


  ' Borra la Pantalla
  Cls
  i = 1

  ' Muestra los Datos
  If chkSoloPuntos.Value = 1 Then
    Call PP2D(posX, posY, txtTamañoPunto, vbBlue)
    'Call PP3D(posX - 5000, posY + 1000, -posZ * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), vbBlue)    'vbWhite
  Else
    frmRadianPrimo.ForeColor = vbBlue
    PSet (posX, posY), vbBlue
    Print i
  End If

  Derecha = True
  Arriba = False
  Izquierda = False
  Abajo = False

  For c = 1 To Expande
    '***********************************************************************

    For m = 1 To c
      'posZ = posZ + 1
      i = i + 1
      ' Derecha
      If Derecha Then
        posX = posX + des
      End If
      ' Arriba
      If Arriba Then
        posY = posY - des
      End If
      ' Izquierda
      If Izquierda Then
        posX = posX - des
      End If
      ' Abajo
      If Abajo Then
        posY = posY + des
      End If
      If Primo(i) Then
        posZ = posZ + 10
        frmRadianPrimo.ForeColor = vbRed
      Else
        frmRadianPrimo.ForeColor = vbBlack
      End If
      ' Muestra los Datos
      If chkSoloPuntos.Value = 1 Then
        Call PP2D(posX, posY, txtTamañoPunto, frmRadianPrimo.ForeColor)
        'Call PP3D(posX - 5000, posY + 1000, -posZ * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), frmRadianPrimo.ForeColor)    'vbWhite
      Else
        PSet (posX, posY)
        Print i
      End If
    Next m
    'Cambia la direccion
    If Derecha Then
      Derecha = False
      Arriba = True
    ElseIf Arriba Then
      Arriba = False
      Izquierda = True
    ElseIf Izquierda Then
      Izquierda = False
      Abajo = True
    ElseIf Abajo Then
      Abajo = False
      Derecha = True
    Else
      Derecha = True
    End If
    '***********************************************************************


    '***********************************************************************
    For m = 1 To c
      'posZ = posZ + 1
      i = i + 1
      ' Derecha
      If Derecha Then
        posX = posX + des
      End If
      ' Arriba
      If Arriba Then
        posY = posY - des
      End If
      ' Izquierda
      If Izquierda Then
        posX = posX - des
      End If
      ' Abajo
      If Abajo Then
        posY = posY + des
      End If
      If Primo(i) Then
        posZ = posZ + 10
        frmRadianPrimo.ForeColor = vbRed
      Else
        frmRadianPrimo.ForeColor = vbBlack
      End If
      ' Muestra los Datos
      If chkSoloPuntos.Value = 1 Then
        Call PP2D(posX, posY, txtTamañoPunto, frmRadianPrimo.ForeColor)
        'Call PP3D(posX - 5000, posY + 1000, -posZ * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), frmRadianPrimo.ForeColor)    'vbWhite
      Else
        PSet (posX, posY)
        Print i
      End If
    Next m
    'Cambia la direccion
    If Derecha Then
      Derecha = False
      Arriba = True
    ElseIf Arriba Then
      Arriba = False
      Izquierda = True
    ElseIf Izquierda Then
      Izquierda = False
      Abajo = True
    ElseIf Abajo Then
      Abajo = False
      Derecha = True
    Else
      Derecha = True
    End If
    '***********************************************************************
  Next c
End Sub

'**************************************************************************************


'**************************************************************************************
' PROCEDIMIENTOS
'**************************************************************************************
' Dibuja los puntos de la imagen
Public Sub Grafica()
  Dim miOrbita3d As Double
  Dim miPXA As Double
  Dim miPYA As Double
  Dim miPXS As Double
  Dim miPYS As Double

  miPXA = 4750
  miPYA = 4750
  miPXS = 4750
  miPYS = 4750
  miOrbita3d = 0
  RadioEsfera = 200

  CentroZ = Val(txtCentoZ.Text)

  miTamañoPunto = Val(txtTamañoPunto.Text)

  miCuentaPrimos = 0

  If min <= 900000 Then
    ' Borra la pantalla
    Cls

    ' Marco
    'Line (100, 100)-(9500, 9500), , B

    ' Ejes de Coordenadas
    Line (4750, 0)-(4750, 9500)
    Line (0, 4750)-(9500, 4750)
    Line (0, 0)-(9500, 9500)
    Line (0, 9500)-(9500, 0)

    ' Borra el área de la circunferencia
    Dim r As Long
    For r = 1 To miRadio * miFactorCircular
      Circle (4750, 4750), r, frmRadianPrimo.BackColor
    Next r

    ' Círculo determinado
    'Circle (4750, 4750), 2512.44, vbBlack

    ' Ángulo Absoluto
    txtAnguloAbsoluto.Text = CambiaComa(360 / min)

    ' Circunferencia para números primos gemelos
    If miPG = True Then
      If Cos(miDelta * (miPi / 180)) > 0.0000001 Then
        Circle (4750, 4750), (4550 * Cos(miDelta * (miPi / 180))) * 0.945, QBColor(5)
      End If
      If Cos(miDelta * (miPi / 180)) < 0.0000001 Then
        If ((4550 * -1 * Cos(miDelta * (miPi / 180))) * 0.945) < 0 Then
          Circle (4750, 4750), (4550 * Cos(miDelta * (miPi / 180))) * 0.945, QBColor(4)
        Else
          Circle (4750, 4750), (4550 * -1 * Cos(miDelta * (miPi / 180))) * 0.945, QBColor(4)
        End If
      End If
    End If

    ' Inicializa los contadores de orbitas
    ReDim miOrbitaP(miOrbitaMaxima)
    ReDim miOrbitaPG(miOrbitaMaxima)
    ReDim miOrbitaC(miOrbitaMaxima)
    ReDim miNumeros(min)
    ReDim miX(min)
    ReDim miY(min)

    ' Recorre toda las circunferencia
    Dim i As Long
    For i = 1 To min

      miOrbita3d = miOrbita3d + 1

      If miOrbita = True Then
        ' Puntos iniciales
        If i = 1 Then
          miMiniRadio = 1
          'miMiniRadio = 0.5
        End If
        If i = 2 Then
          miMiniRadio = 2
          'miMiniRadio = 0.5
        End If
        If i = 3 Then
          miMiniRadio = 1
          'miMiniRadio = 1
        End If
      Else
        ' Puntos iniciales
        If i = 1 Then
          miMiniRadio = 50
        End If
        If i = 2 Then
          miMiniRadio = 50
        End If
        If i = 3 Then
          miMiniRadio = 50
        End If
      End If

      ' Cálculo de las coordenadas X, Y
      If chkInverso.Value = 1 Then
        If miOrbita = True Then
          miMiniRadio = (-1) * (miMiniRadio - Val(txtInverso.Text))
          X1 = 4750 + ((miRadio * (miMiniRadio / miZoom)) * Cos(i * miDelta * (miPi / 180)) * miFactorCircular)
          Y1 = 4750 + ((miRadio * (miMiniRadio / miZoom)) * -Sin(i * miDelta * (miPi / 180)) * miFactorCircular)
          miMiniRadio = (-1) * (miMiniRadio - Val(txtInverso.Text))
        Else
          miMiniRadio = (-1) * (miMiniRadio - Val(txtInverso.Text))
          X1 = 4750 + ((miRadio * (miMiniRadio / miZoom)) * -Cos(i * miDelta * (miPi / 180)) * miFactorCircular)
          Y1 = 4750 + ((miRadio * (miMiniRadio / miZoom)) * Sin(i * miDelta * (miPi / 180)) * miFactorCircular)
          miMiniRadio = (-1) * (miMiniRadio - Val(txtInverso.Text))
        End If

      Else
        X1 = 4750 + ((miRadio * (miMiniRadio / miZoom)) * Cos(i * miDelta * (miPi / 180)) * miFactorCircular)
        Y1 = 4750 + ((miRadio * (miMiniRadio / miZoom)) * -Sin(i * miDelta * (miPi / 180)) * miFactorCircular)
      End If

      ' Guarda el número y sus coordenadas
      miNumeros(i) = i
      'miX(i) = X1
      'miY(i) = Y1
      ''            miX(i) = EX
      ''            miY(i) = EY

      '****************************************************************************
      ' GUARDA LOS DATOS EN EL REGISTRO TEMPORAL DE LA BASE DE DATOS
      With rGraficaPrimos
        .Numero = i

        If Primo(i) Then
          .Primo = 1
        Else
          .Primo = 0
        End If

        .CX = X1
        .CY = Y1
        .Tamaño = miTamañoPunto

        If Primo(i) Then
          .Color = 12
        Else
          .Color = 0
        End If

        .PCX = 0
        .PCY = 0
      End With
      '****************************************************************************

      '****************************************************************************
      ' GRABAR EN LA BASE DE DATOS
      '****************************************************************************
      With rGraficaPrimos
        sql = "INSERT INTO GraficaPrimos (Numero,Primo,CX,CY,Tamaño,Color,PCX,PCY) VALUES ("
        sql = sql & Str(.Numero) & "," & .Primo & "," & Str(.CX) & "," & Str(.CY) & "," & Str(.Tamaño) & "," & Str(.Color) + "," & Str(.PCX) & "," & Str(.PCY) & ")"

        ' Insertar el registro
        If chkBaseDatos.Value = 1 Then
          cn.Execute sql
        End If
      End With
      '****************************************************************************

      ' Calcula si es primo
      If Primo(i) = True Then
        miCuentaPrimos = miCuentaPrimos + 1

        ' Calcula cantidad Superior e inferior
        If Y1 <= 4750 Then
          miCuentaSuperior = miCuentaSuperior + 1
        Else
          miCuentaInferior = miCuentaInferior + 1
        End If

        ' Control de Rango
        If miRango = True Then
          If i >= Val(txtMin.Text) And i <= Val(txtMax.Text) Then
            ' Solo muestra el rango

            ' Punto Primo
            If miPrimos = True Then
              Call PP3D(X1 - 5000, Y1 + 1000, -miOrbita3d * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), vbRed)    'vbWhite

              If miEtiqueta = True Then
                frmRadianPrimo.ForeColor = vbRed
                Print i
              End If
            Else

              Call PP3D(X1 - 5000, Y1 + 1000, -miOrbita3d * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), vbBlack)    'vbWhite

              If miEtiqueta = True Then
                frmRadianPrimo.ForeColor = vbBlack
                Print i
              End If

            End If

          End If
        Else
          ' Muestra todo

          '''''                    ' Muestra los primos
          '''''                    If miLineasP = True Then
          '''''                        ' Línea Prima
          '''''                        Line (EX, EY)-(CentroY + 6000, CentroZ + Val(txtCentoZ.Text)), vbYellow
          '''''                    End If

          ' Punto Primo
          If miPrimos = True Then
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++

            Call PP3D(X1 - 5000, Y1 + 1000, -miOrbita3d * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), vbRed)    'vbWhite

            '++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++

            If miEtiqueta = True Then
              frmRadianPrimo.ForeColor = vbRed
              Print i
            End If

            'miOrbita3d = 0

          End If
        End If

        miOrbita3d = 0

        ' Muestra los primos
        If miLineasP = True Then
          ' Línea Prima
          Line (EX, EY)-(CentroY + 6000, CentroZ + Val(txtCentoZ.Text)), vbYellow
        End If

        ' Muestra las líneas de las ramas
        If miRama = True Then
          ' Linea Rama
          miPXS = EX  'X1
          miPYS = EY  'Y1
          frmRadianPrimo.ForeColor = vbWhite

          'frmRadianPrimo.ForeColor = vbBlack
          If miPXA <> 4750 Then
            ' Control de Rango
            If miRango = True Then
              If i >= Val(txtMin.Text) And i <= Val(txtMax.Text) Then
                ' Solo muestra el rango
                Line (miPXA, miPYA)-(miPXS, miPYS)
              End If
            Else
              ' Muestra todo
              Line (miPXA, miPYA)-(miPXS, miPYS)
            End If
          End If
          miPXA = miPXS
          miPYA = miPYS
        End If

        ' Reinicia
        miPXA = 4750
        miPYA = 4750
        miPXS = 4750
        miPYS = 4750

        '*********************************************************************************
        ' Une Puntos
        If i <> 2 Then
          If miUne = True Then

            ' Caso especial 3 y 5
            If i = 3 Then
              X3 = EX  'X1
              Y3 = EY  'Y1
            End If
            If i = 5 Then
              X5 = EX  'X1
              Y5 = EY  'Y1
              ' Control de Rango
              If miRango = True Then
                If i >= Val(txtMin.Text) And i <= Val(txtMax.Text) Then
                  ' Solo muestra el rango
                  Line (X3, Y3)-(X5, Y5), vbGreen
                End If
              Else
                ' Muestra todo
                Line (X3, Y3)-(X5, Y5), vbGreen
              End If
            End If

            ' Almacena el actual
            If Primo(i + 2) Then
              X2 = EX  'X1
              Y2 = EY  'Y1
            End If

            ' Los une en el momento oportuno
            If (i - 2) > 0 Then
              If Primo(i - 2) Then
                ' Control de Rango
                If miRango = True Then
                  If i >= Val(txtMin.Text) And i <= Val(txtMax.Text) Then
                    ' Solo muestra el rango
                    'Line (X1, Y1)-(X2, Y2), vbGreen
                    Line (EX, EY)-(X2, Y2), vbGreen
                  End If
                Else
                  ' Muestra todo
                  'Line (X1, Y1)-(X2, Y2), vbGreen
                  Line (EX, EY)-(X2, Y2), vbGreen
                End If
              End If
            End If
          End If
        End If
        '*********************************************************************************

        ' Aumenta un primo en la orbita
        miOrbitaP(miMiniRadio) = miOrbitaP(miMiniRadio) + 1
        If Primo(i + 2) Then
          miOrbitaPG(miMiniRadio) = True
        End If

        ' Activa y desactiva las orbitas
        If miOrbita = True Then
          miMiniRadio = 1
        Else
          miMiniRadio = 50
        End If

      Else
        ' Control de Rango
        If miRango = True Then
          If i >= Val(txtMin.Text) And i <= Val(txtMax.Text) Then
            ' Solo muestra el rango
            ' Muestra los compuestos

            If miCompuestos = True Then
              Call PP3D(X1 - 5000, Y1 + 1000, -miOrbita3d * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), vbBlack)    'vbWhite

              If miEtiqueta = True Then
                frmRadianPrimo.ForeColor = vbBlack
                Print i
              End If
            End If
          End If
        Else
          ' Muestra todo
          ' Muestra los compuestos
          If miCompuestos = True Then
            Call PP3D(X1 - 5000, Y1 + 1000, -miOrbita3d * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), vbBlack)    'vbWhite

            If miEtiqueta = True Then
              frmRadianPrimo.ForeColor = vbBlack
              Print i
            End If
          End If
        End If

        'Si no es primo
        If miLineasC = True Then
          ' Control de Rango
          If miRango = True Then
            If i >= Val(txtMin.Text) And i <= Val(txtMax.Text) Then
              'Line (X1, Y1)-(4750, 4750), vbBlue
              Line (EX, EY)-(CentroY + 6000, CentroZ + Val(txtCentoZ.Text)), vbBlue
            End If
          Else
            ' Muestra todo
            'Line (X1, Y1)-(4750, 4750), vbBlue
            Line (EX, EY)-(CentroY + 6000, CentroZ + Val(txtCentoZ.Text)), vbBlue
          End If
        End If

        ' Muestra las líneas de las ramas
        If miRama = True Then
          ' Linea Rama
          miPXS = EX  'X1
          miPYS = EY  'Y1
          frmRadianPrimo.ForeColor = vbWhite
          'frmRadianPrimo.ForeColor = vbBlack
          If miPXA <> 4750 Then
            ' Control de Rango
            If miRango = True Then
              If i >= Val(txtMin.Text) And i <= Val(txtMax.Text) Then
                ' Solo muestra el rango
                Line (miPXA, miPYA)-(miPXS, miPYS)
              End If
            Else
              ' Muestra todo
              Line (miPXA, miPYA)-(miPXS, miPYS)
            End If
          End If
          miPXA = miPXS
          miPYA = miPYS
        End If

        ' Aumenta un compuesto en la orbita
        miOrbitaC(miMiniRadio) = miOrbitaC(miMiniRadio) + 1

        ' Activa y desactiva las orbitas
        If miOrbita = True Then
          miMiniRadio = miMiniRadio + 1
        End If

        ' Controla a la orbita maxima
        If miMiniRadio > miOrbitaMaxima Then
          miOrbitaMaxima = miOrbitaMaxima + 1
          ReDim Preserve miOrbitaP(miOrbitaMaxima)
          ReDim Preserve miOrbitaC(miOrbitaMaxima)
        End If
      End If

      ' Pinta recta en punto tangente a los primos gemelos
      If miPG = True Then
        If i <> 1 Then
          If Primo(i - 1) And Primo(i + 1) Then

            ' Control de Rango
            If miRango = True Then
              If i >= Val(txtMin.Text) And i <= Val(txtMax.Text) Then
                ' Solo muestra el rango
                'Line (X1, Y1)-(4750, 4750), QBColor(11)
                Line (EX, EY)-(CentroY + 6000, CentroZ + Val(txtCentoZ.Text)), QBColor(11)
              End If
            Else
              ' Muestra todo
              'Line (X1, Y1)-(4750, 4750), QBColor(11)
              Line (EX, EY)-(CentroY + 6000, CentroZ + Val(txtCentoZ.Text)), QBColor(11)
            End If

          End If
        End If
      End If

      miX(i) = EX
      miY(i) = EY

    Next i

    'Restablece color de fuentes y dibuja los ejes
    frmRadianPrimo.ForeColor = vbBlack
    If miEjes = True Then
      ' Ejes de Coordenadas
      Line (4750, 0)-(4750, 9500)
      Line (0, 4750)-(9500, 4750)
      'Line (0, 0)-(9500, 9500)
      'Line (0, 9500)-(9500, 0)
    End If

    ' Indica los Primos Gemelos
    If (miOrbitaP(2) - 1) > 0 Then
      txtPrimosGemelos.Text = miOrbitaP(2) - 1
    End If
  End If


  If miUneGap = True Then
    Call MostrarGap
  End If

  txtCantidadPrimos.Text = miCuentaPrimos
End Sub

' Seleccionar un ángulo determinado desde el combobox
Private Sub cboAngulo_Click()
  txtAnguloDeterminado.Text = cboAngulo.Text
  Call cmdAnguloDeterminado_Click
  DoEvents
End Sub

' DIBUJA UN CIRCULO
Public Sub DibujaCirculo(ByVal pX As Long, ByVal pY As Long, ByVal pRadio As Long, ByVal pColor As Integer)
  Circle (pX, pY), pRadio, QBColor(pColor)
End Sub

' Mostrar línea de Gap
Private Sub MostrarGap()
  Dim i As Long
  Dim miGap As Long

  miGap = Val(txtGap.Text)

  For i = 1 To (min - miGap)
    If Primo(i) Then
      If GapPrimos(i, miGap) Then
        Line (miX(i), miY(i))-(miX(i + miGap), miY(i + miGap)), QBColor(2)
      End If
    End If
  Next i
End Sub

' Grafica un punto en espacio vectorial de tres dimensiones 3D
Public Sub PP3D(ByVal pX As Double, ByVal pY As Double, ByVal pZ As Double, pTamaño As Integer, ByVal pColor As Long)
' Declaración de variables
  Dim ppX As Double
  Dim ppY As Double
  Dim ang As Double
  Dim Radio As Double
  ' Ajuste del Angulo
  ang = Val(txtAngulo.Text) * (miPi / 180)
  ' Coordenadas de Pantalla del Punto
  ppX = CentroY + (-pX * Cos(ang)) + (pY) + (0)
  ppY = CentroZ + (pX * Sin(ang)) + (0) + (-pZ)
  ' Mostrar el punto con control del Tamaño y Color
  If pTamaño <= 0 Then
    PSet (ppX, ppY), pColor
  Else
    For Radio = 1 To pTamaño
      Circle (ppX, ppY), Radio, pColor
    Next
  End If

  EX = ppX
  EY = ppY
  ' Muestra el vector posición
  'If pColor <> vbRed And pColor <> vbGreen And pColor <> vbBlue Then
  Call VectorPosicion(CentroY + 6000, CentroZ + Val(txtCentoZ.Text), ppX, ppY)
  'End If
End Sub

' Grafica un punto en espacio vectorial de tres dimensiones 3D
Public Sub PP2D(ByVal pX As Double, ByVal pY As Double, pTamaño As Integer, ByVal pColor As Long)
' Declaración de variables
  Dim Radio As Double
  ' Coordenadas de Pantalla del Punto
  ' Mostrar el punto con control del Tamaño y Color
  If pTamaño <= 0 Then
    PSet (pX, pY), pColor
  Else
    For Radio = 1 To pTamaño
      Circle (pX, pY), Radio, pColor
    Next
  End If
End Sub

' Muestra el vector posición
Public Sub VectorPosicion(ByVal PCX As Double, ByVal PCY As Double, ByVal pX As Double, ByVal pY As Double)
' Muestra el vector posición
  If chkVectorPosicion.Value = 1 Then
    Line (PCX, PCY)-(pX, pY), vbYellow
  End If
End Sub


'**************************************************************************************
' FUNCIONES
'**************************************************************************************
' FUNCION PARA CALCULAR SI EL NUMERO ES PRIMO
Public Function Primo(ByVal pN As Long) As Boolean
  Dim i As Long
  Primo = True
  If pN = 1 Then
    Primo = False
  Else
    For i = 2 To Sqr(pN)
      If (pN / i) = Int(pN / i) Then
        Primo = False
      End If
    Next i
  End If
End Function

' FUNCION PARA CALCULAR GAP ENTRE PRIMOS
Public Function GapPrimos(ByVal pN As Long, ByVal pG As Long) As Boolean
  Dim i As Long
  If Primo(pN) Then
    If Primo(pN + pG) Then
      GapPrimos = True
    End If

    For i = (pN + 1) To (pN + pG - 1)
      If Primo(i) Then
        GapPrimos = False
      End If
    Next
  Else
    GapPrimos = False
  End If
End Function

' Ajusta el punto por la coma
Private Function CambiaComa(ByRef n As Double) As String
  Dim i As Integer
  CambiaComa = ""
  For i = 1 To Len(n)
    If Mid(Str(n), i, 1) = "," Then
      CambiaComa = CambiaComa + "."
    Else
      CambiaComa = CambiaComa + Mid(Str(n), i, 1)
    End If
  Next i
End Function

' FUNCION PARA TABULAR
Public Function Tabulado(ByVal pT As String, ByVal pA As Integer) As String
  Dim i As Integer
  Dim miAncho As Integer
  miAncho = Len(Trim(pT))

  For i = 1 To (pA - miAncho)
    'pT = pT + " "
    pT = " " + pT
  Next i
  Tabulado = pT
End Function






' Dibuja los puntos de la imagen
Public Sub GraficaProyeccion()
  Dim miOrbita3d As Double
  Dim miPXA As Double
  Dim miPYA As Double
  Dim miPXS As Double
  Dim miPYS As Double
  Dim miProyX As Double
  Dim miProyY As Double
  Dim miProyZ As Double


  miPXA = 4750
  miPYA = 4750
  miPXS = 4750
  miPYS = 4750
  miOrbita3d = 0
  RadioEsfera = 200

  CentroZ = Val(txtCentoZ.Text)

  miTamañoPunto = Val(txtTamañoPunto.Text)

  miCuentaPrimos = 0

  If min <= 900000 Then
    ' Borra la pantalla
    Cls

    ' Marco
    'Line (100, 100)-(9500, 9500), , B

    ' Ejes de Coordenadas
    Line (4750, 0)-(4750, 9500)
    Line (0, 4750)-(9500, 4750)
    Line (0, 0)-(9500, 9500)
    Line (0, 9500)-(9500, 0)

    ' Borra el área de la circunferencia
    Dim r As Long
    For r = 1 To miRadio * miFactorCircular
      Circle (4750, 4750), r, frmRadianPrimo.BackColor
    Next r

    ' Círculo determinado
    'Circle (4750, 4750), 2512.44, vbBlack

    ' Ángulo Absoluto
    txtAnguloAbsoluto.Text = CambiaComa(360 / min)

    ' Circunferencia para números primos gemelos
    If miPG = True Then
      If Cos(miDelta * (miPi / 180)) > 0.0000001 Then
        Circle (4750, 4750), (4550 * Cos(miDelta * (miPi / 180))) * 0.945, QBColor(5)
      End If
      If Cos(miDelta * (miPi / 180)) < 0.0000001 Then
        If ((4550 * -1 * Cos(miDelta * (miPi / 180))) * 0.945) < 0 Then
          Circle (4750, 4750), (4550 * Cos(miDelta * (miPi / 180))) * 0.945, QBColor(4)
        Else
          Circle (4750, 4750), (4550 * -1 * Cos(miDelta * (miPi / 180))) * 0.945, QBColor(4)
        End If
      End If
    End If

    ' Inicializa los contadores de orbitas
    ReDim miOrbitaP(miOrbitaMaxima)
    ReDim miOrbitaPG(miOrbitaMaxima)
    ReDim miOrbitaC(miOrbitaMaxima)
    ReDim miNumeros(min)
    ReDim miX(min)
    ReDim miY(min)

    ' Recorre toda las circunferencia
    Dim i As Long
    For i = 1 To min

      miOrbita3d = miOrbita3d + 1

      If miOrbita = True Then
        ' Puntos iniciales
        If i = 1 Then
          miMiniRadio = 1
          'miMiniRadio = 0.5
        End If
        If i = 2 Then
          miMiniRadio = 2
          'miMiniRadio = 0.5
        End If
        If i = 3 Then
          miMiniRadio = 1
          'miMiniRadio = 1
        End If
      Else
        ' Puntos iniciales
        If i = 1 Then
          miMiniRadio = 50
        End If
        If i = 2 Then
          miMiniRadio = 50
        End If
        If i = 3 Then
          miMiniRadio = 50
        End If
      End If

      ' Cálculo de las coordenadas X, Y
      If chkInverso.Value = 1 Then
        If miOrbita = True Then
          miMiniRadio = (-1) * (miMiniRadio - Val(txtInverso.Text))
          X1 = 4750 + ((miRadio * (miMiniRadio / miZoom)) * Cos(i * miDelta * (miPi / 180)) * miFactorCircular)
          Y1 = 4750 + ((miRadio * (miMiniRadio / miZoom)) * -Sin(i * miDelta * (miPi / 180)) * miFactorCircular)
          miMiniRadio = (-1) * (miMiniRadio - Val(txtInverso.Text))
        Else
          miMiniRadio = (-1) * (miMiniRadio - Val(txtInverso.Text))
          X1 = 4750 + ((miRadio * (miMiniRadio / miZoom)) * -Cos(i * miDelta * (miPi / 180)) * miFactorCircular)
          Y1 = 4750 + ((miRadio * (miMiniRadio / miZoom)) * Sin(i * miDelta * (miPi / 180)) * miFactorCircular)
          miMiniRadio = (-1) * (miMiniRadio - Val(txtInverso.Text))
        End If

      Else
        X1 = 4750 + ((miRadio * (miMiniRadio / miZoom)) * Cos(i * miDelta * (miPi / 180)) * miFactorCircular)
        Y1 = 4750 + ((miRadio * (miMiniRadio / miZoom)) * -Sin(i * miDelta * (miPi / 180)) * miFactorCircular)
      End If


      ' Calcula las Coordenadas X,Y,Z de a Proyeccion
      miProyX = 100000 * proyX(EX, EY)
      miProyY = 100000 * proyY(EX, EY)
      miProyZ = 100000 * proyZ(EX, EY)
      'Call PP3D(X1 - 5000, Y1 + 1000, -miOrbita3d * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), vbRed)    'vbWhite
      If Primo(i) Then
        Call PP3D(miProyX, miProyY, miProyZ, Val(txtTamañoPunto.Text), vbRed)   'vbWhite
      Else
        Call PP3D(miProyX, miProyY, miProyZ, Val(txtTamañoPunto.Text), vbBlack)   'vbWhite
      End If

      ' Guarda el número y sus coordenadas
      miNumeros(i) = i
      'miX(i) = X1
      'miY(i) = Y1
      ''            miX(i) = EX
      ''            miY(i) = EY

      '****************************************************************************
      ' GUARDA LOS DATOS EN EL REGISTRO TEMPORAL DE LA BASE DE DATOS
      With rGraficaPrimos
        .Numero = i

        If Primo(i) Then
          .Primo = 1
        Else
          .Primo = 0
        End If

        .CX = X1
        .CY = Y1
        .Tamaño = miTamañoPunto

        If Primo(i) Then
          .Color = 12
        Else
          .Color = 0
        End If

        .PCX = 0
        .PCY = 0
      End With
      '****************************************************************************

      ''            '****************************************************************************
      ''            ' GRABAR EN LA BASE DE DATOS
      ''            '****************************************************************************
      ''            With rGraficaPrimos
      ''              sql = "INSERT INTO GraficaPrimos (Numero,Primo,CX,CY,Tamaño,Color,PCX,PCY) VALUES ("
      ''              sql = sql & Str(.Numero) & "," & .Primo & "," & Str(.CX) & "," & Str(.CY) & "," & Str(.Tamaño) & "," & Str(.Color) + "," & Str(.PCX) & "," & Str(.PCY) & ")"
      ''
      ''              ' Insertar el registro
      ''              If chkBaseDatos.Value = 1 Then
      ''                cn.Execute sql
      ''              End If
      ''            End With
      ''            '****************************************************************************


      ' Calcula si es primo
      If Primo(i) = True Then
        miCuentaPrimos = miCuentaPrimos + 1

        ' Calcula cantidad Superior e inferior
        If Y1 <= 4750 Then
          miCuentaSuperior = miCuentaSuperior + 1
        Else
          miCuentaInferior = miCuentaInferior + 1
        End If

        ' Control de Rango
        If miRango = True Then
          If i >= Val(txtMin.Text) And i <= Val(txtMax.Text) Then
            ' Solo muestra el rango

            ' Punto Primo
            If miPrimos = True Then
              Call PP3D(X1 - 5000, Y1 + 1000, -miOrbita3d * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), vbRed)    'vbWhite

              ' Calcula las Coordenadas X,Y,Z de a Proyeccion
              miProyX = proyX(EX, EY)
              miProyY = proyY(EX, EY)
              miProyZ = proyZ(EX, EY)
              'Call PP3D(X1 - 5000, Y1 + 1000, -miOrbita3d * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), vbRed)    'vbWhite
              Call PP3D(miProyX, miProyY, miProyZ, Val(txtTamañoPunto.Text), vbRed)   'vbWhite




              If miEtiqueta = True Then
                frmRadianPrimo.ForeColor = vbRed
                Print i
              End If
            Else

              Call PP3D(X1 - 5000, Y1 + 1000, -miOrbita3d * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), vbBlack)    'vbWhite

              ' Calcula las Coordenadas X,Y,Z de a Proyeccion
              miProyX = proyX(EX, EY)
              miProyY = proyY(EX, EY)
              miProyZ = proyZ(EX, EY)
              'Call PP3D(X1 - 5000, Y1 + 1000, -miOrbita3d * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), vbRed)    'vbWhite
              Call PP3D(miProyX, miProyY, miProyZ, Val(txtTamañoPunto.Text), vbRed)   'vbWhite


              If miEtiqueta = True Then
                frmRadianPrimo.ForeColor = vbBlack
                Print i
              End If

            End If

          End If
        Else
          ' Muestra todo

          '''''                    ' Muestra los primos
          '''''                    If miLineasP = True Then
          '''''                        ' Línea Prima
          '''''                        Line (EX, EY)-(CentroY + 6000, CentroZ + Val(txtCentoZ.Text)), vbYellow
          '''''                    End If

          ' Punto Primo
          If miPrimos = True Then
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++

            Call PP3D(X1 - 5000, Y1 + 1000, -miOrbita3d * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), vbRed)    'vbWhite

            '++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++


            ' Calcula las Coordenadas X,Y,Z de a Proyeccion
            miProyX = proyX(EX, EY)
            miProyY = proyY(EX, EY)
            miProyZ = proyZ(EX, EY)
            'Call PP3D(X1 - 5000, Y1 + 1000, -miOrbita3d * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), vbRed)    'vbWhite
            Call PP3D(miProyX, miProyY, miProyZ, Val(txtTamañoPunto.Text), vbRed)   'vbWhite





            If miEtiqueta = True Then
              frmRadianPrimo.ForeColor = vbRed
              Print i
            End If

            'miOrbita3d = 0

          End If
        End If

        miOrbita3d = 0

        ' Muestra los primos
        If miLineasP = True Then
          ' Línea Prima
          Line (EX, EY)-(CentroY + 6000, CentroZ + Val(txtCentoZ.Text)), vbYellow
        End If

        ' Muestra las líneas de las ramas
        If miRama = True Then
          ' Linea Rama
          miPXS = EX  'X1
          miPYS = EY  'Y1
          frmRadianPrimo.ForeColor = vbWhite

          'frmRadianPrimo.ForeColor = vbBlack
          If miPXA <> 4750 Then
            ' Control de Rango
            If miRango = True Then
              If i >= Val(txtMin.Text) And i <= Val(txtMax.Text) Then
                ' Solo muestra el rango
                Line (miPXA, miPYA)-(miPXS, miPYS)
              End If
            Else
              ' Muestra todo
              Line (miPXA, miPYA)-(miPXS, miPYS)
            End If
          End If
          miPXA = miPXS
          miPYA = miPYS
        End If

        ' Reinicia
        miPXA = 4750
        miPYA = 4750
        miPXS = 4750
        miPYS = 4750

        '*********************************************************************************
        ' Une Puntos
        If i <> 2 Then
          If miUne = True Then

            ' Caso especial 3 y 5
            If i = 3 Then
              X3 = EX  'X1
              Y3 = EY  'Y1
            End If
            If i = 5 Then
              X5 = EX  'X1
              Y5 = EY  'Y1
              ' Control de Rango
              If miRango = True Then
                If i >= Val(txtMin.Text) And i <= Val(txtMax.Text) Then
                  ' Solo muestra el rango
                  Line (X3, Y3)-(X5, Y5), vbGreen
                End If
              Else
                ' Muestra todo
                Line (X3, Y3)-(X5, Y5), vbGreen
              End If
            End If

            ' Almacena el actual
            If Primo(i + 2) Then
              X2 = EX  'X1
              Y2 = EY  'Y1
            End If

            ' Los une en el momento oportuno
            If (i - 2) > 0 Then
              If Primo(i - 2) Then
                ' Control de Rango
                If miRango = True Then
                  If i >= Val(txtMin.Text) And i <= Val(txtMax.Text) Then
                    ' Solo muestra el rango
                    'Line (X1, Y1)-(X2, Y2), vbGreen
                    Line (EX, EY)-(X2, Y2), vbGreen
                  End If
                Else
                  ' Muestra todo
                  'Line (X1, Y1)-(X2, Y2), vbGreen
                  Line (EX, EY)-(X2, Y2), vbGreen
                End If
              End If
            End If
          End If
        End If
        '*********************************************************************************

        ' Aumenta un primo en la orbita
        miOrbitaP(miMiniRadio) = miOrbitaP(miMiniRadio) + 1
        If Primo(i + 2) Then
          miOrbitaPG(miMiniRadio) = True
        End If

        ' Activa y desactiva las orbitas
        If miOrbita = True Then
          miMiniRadio = 1
        Else
          miMiniRadio = 50
        End If

      Else
        ' Control de Rango
        If miRango = True Then
          If i >= Val(txtMin.Text) And i <= Val(txtMax.Text) Then
            ' Solo muestra el rango
            ' Muestra los compuestos

            If miCompuestos = True Then
              Call PP3D(X1 - 5000, Y1 + 1000, -miOrbita3d * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), vbBlack)    'vbWhite

              ' Calcula las Coordenadas X,Y,Z de a Proyeccion
              miProyX = proyX(EX, EY)
              miProyY = proyY(EX, EY)
              miProyZ = proyZ(EX, EY)
              'Call PP3D(X1 - 5000, Y1 + 1000, -miOrbita3d * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), vbRed)    'vbWhite
              Call PP3D(miProyX, miProyY, miProyZ, Val(txtTamañoPunto.Text), vbBlack)   'vbWhite



              If miEtiqueta = True Then
                frmRadianPrimo.ForeColor = vbBlack
                Print i
              End If
            End If
          End If
        Else
          ' Muestra todo
          ' Muestra los compuestos
          If miCompuestos = True Then
            Call PP3D(X1 - 5000, Y1 + 1000, -miOrbita3d * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), vbBlack)    'vbWhite

            ' Calcula las Coordenadas X,Y,Z de a Proyeccion
            miProyX = proyX(EX, EY)
            miProyY = proyY(EX, EY)
            miProyZ = proyZ(EX, EY)
            'Call PP3D(X1 - 5000, Y1 + 1000, -miOrbita3d * Val(txtProfundidad.Text), Val(txtTamañoPunto.Text), vbRed)    'vbWhite
            Call PP3D(miProyX, miProyY, miProyZ, Val(txtTamañoPunto.Text), vbBlack)   'vbWhite




            If miEtiqueta = True Then
              frmRadianPrimo.ForeColor = vbBlack
              Print i
            End If
          End If
        End If

        'Si no es primo
        If miLineasC = True Then
          ' Control de Rango
          If miRango = True Then
            If i >= Val(txtMin.Text) And i <= Val(txtMax.Text) Then
              'Line (X1, Y1)-(4750, 4750), vbBlue
              Line (EX, EY)-(CentroY + 6000, CentroZ + Val(txtCentoZ.Text)), vbBlue
            End If
          Else
            ' Muestra todo
            'Line (X1, Y1)-(4750, 4750), vbBlue
            Line (EX, EY)-(CentroY + 6000, CentroZ + Val(txtCentoZ.Text)), vbBlue
          End If
        End If

        ' Muestra las líneas de las ramas
        If miRama = True Then
          ' Linea Rama
          miPXS = EX  'X1
          miPYS = EY  'Y1
          frmRadianPrimo.ForeColor = vbWhite
          'frmRadianPrimo.ForeColor = vbBlack
          If miPXA <> 4750 Then
            ' Control de Rango
            If miRango = True Then
              If i >= Val(txtMin.Text) And i <= Val(txtMax.Text) Then
                ' Solo muestra el rango
                Line (miPXA, miPYA)-(miPXS, miPYS)
              End If
            Else
              ' Muestra todo
              Line (miPXA, miPYA)-(miPXS, miPYS)
            End If
          End If
          miPXA = miPXS
          miPYA = miPYS
        End If

        ' Aumenta un compuesto en la orbita
        miOrbitaC(miMiniRadio) = miOrbitaC(miMiniRadio) + 1

        ' Activa y desactiva las orbitas
        If miOrbita = True Then
          miMiniRadio = miMiniRadio + 1
        End If

        ' Controla a la orbita maxima
        If miMiniRadio > miOrbitaMaxima Then
          miOrbitaMaxima = miOrbitaMaxima + 1
          ReDim Preserve miOrbitaP(miOrbitaMaxima)
          ReDim Preserve miOrbitaC(miOrbitaMaxima)
        End If
      End If

      ' Pinta recta en punto tangente a los primos gemelos
      If miPG = True Then
        If i <> 1 Then
          If Primo(i - 1) And Primo(i + 1) Then

            ' Control de Rango
            If miRango = True Then
              If i >= Val(txtMin.Text) And i <= Val(txtMax.Text) Then
                ' Solo muestra el rango
                'Line (X1, Y1)-(4750, 4750), QBColor(11)
                Line (EX, EY)-(CentroY + 6000, CentroZ + Val(txtCentoZ.Text)), QBColor(11)
              End If
            Else
              ' Muestra todo
              'Line (X1, Y1)-(4750, 4750), QBColor(11)
              Line (EX, EY)-(CentroY + 6000, CentroZ + Val(txtCentoZ.Text)), QBColor(11)
            End If

          End If
        End If
      End If

      miX(i) = EX
      miY(i) = EY

    Next i

    'Restablece color de fuentes y dibuja los ejes
    frmRadianPrimo.ForeColor = vbBlack
    If miEjes = True Then
      ' Ejes de Coordenadas
      Line (4750, 0)-(4750, 9500)
      Line (0, 4750)-(9500, 4750)
      'Line (0, 0)-(9500, 9500)
      'Line (0, 9500)-(9500, 0)
    End If

    ' Indica los Primos Gemelos
    If (miOrbitaP(2) - 1) > 0 Then
      txtPrimosGemelos.Text = miOrbitaP(2) - 1
    End If
  End If


  If miUneGap = True Then
    Call MostrarGap
  End If

  txtCantidadPrimos.Text = miCuentaPrimos
End Sub


