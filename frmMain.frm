VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "GENOMIN: Open a FASTA file  ..."
   ClientHeight    =   11730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16080
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11730
   ScaleWidth      =   16080
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "Process on buffer"
      Height          =   255
      Left            =   3120
      TabIndex        =   84
      Top             =   120
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Process on window"
      Height          =   255
      Left            =   3120
      TabIndex        =   83
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   9480
      TabIndex        =   81
      Text            =   "0"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   9480
      TabIndex        =   80
      Text            =   "0"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox SSR 
      Caption         =   "Save a short report"
      Height          =   255
      Left            =   240
      TabIndex        =   77
      Top             =   1050
      Width           =   1935
   End
   Begin VB.ComboBox DubleN 
      Height          =   315
      Left            =   9000
      TabIndex        =   70
      Text            =   "CG"
      Top             =   5160
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11160
      Top             =   240
   End
   Begin VB.CheckBox Erase_rec 
      Caption         =   "Erase chart before recalibration"
      Height          =   255
      Left            =   1680
      TabIndex        =   64
      Top             =   4680
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CheckBox LOB2 
      Caption         =   "Use lines or bars"
      Height          =   255
      Left            =   1680
      TabIndex        =   63
      Top             =   5385
      Width           =   1815
   End
   Begin VB.CheckBox LOB1 
      Caption         =   "Use lines or bars"
      Height          =   255
      Left            =   1680
      TabIndex        =   62
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CheckBox Rec_buff 
      Caption         =   "Use recalibration, the highest percentage becomes 100%"
      Height          =   255
      Left            =   1680
      TabIndex        =   61
      Top             =   4920
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.CheckBox UTN 
      Caption         =   "Use text notation (not recommended for small motifs)"
      Height          =   255
      Left            =   1680
      TabIndex        =   60
      Top             =   7880
      Width           =   4335
   End
   Begin VB.Frame Frame2 
      Caption         =   "GENOMIN status"
      Height          =   1455
      Left            =   120
      TabIndex        =   58
      Top             =   10200
      Width           =   15375
      Begin VB.TextBox GIN 
         Height          =   1095
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   59
         Text            =   "frmMain.frx":57250
         Top             =   240
         Width           =   13695
      End
      Begin VB.Label Clock_diff1 
         Caption         =   "Start: 00:00:00"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Clock_diff2 
         Caption         =   "Stop: 00:00:00"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Min_past 
         Caption         =   "Total: 0 Minute(s)"
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.CommandButton Save_GENOMIN 
      Caption         =   "Save results  ..."
      Height          =   375
      Left            =   240
      TabIndex        =   57
      Top             =   600
      Width           =   2775
   End
   Begin VB.CheckBox UG2 
      Caption         =   "Use grid"
      Height          =   255
      Left            =   1680
      TabIndex        =   53
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CheckBox UG1 
      Caption         =   "Use grid"
      Height          =   255
      Left            =   1680
      TabIndex        =   52
      Top             =   2745
      Width           =   1335
   End
   Begin VB.TextBox Rep 
      Height          =   285
      Left            =   10920
      TabIndex        =   28
      Text            =   "1"
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox motif_sequence 
      Height          =   285
      Left            =   8760
      TabIndex        =   24
      Text            =   "AAGCTT"
      Top             =   7800
      Width           =   2775
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00008000&
      Height          =   1455
      Left            =   600
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   725
      TabIndex        =   21
      Top             =   8160
      Width           =   10935
      Begin VB.Label Mes3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Example3 off"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   480
         Left            =   3480
         TabIndex        =   75
         Top             =   480
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   104
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   600
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   725
      TabIndex        =   15
      Top             =   5640
      Width           =   10935
      Begin VB.Label Mes2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Example2 off"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   480
         Left            =   3480
         TabIndex        =   74
         Top             =   480
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Line Lin6 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   0
         X2              =   728
         Y1              =   74
         Y2              =   74
      End
      Begin VB.Line Lin4 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   0
         X2              =   728
         Y1              =   26
         Y2              =   26
      End
      Begin VB.Line Lin5 
         BorderColor     =   &H00808080&
         BorderStyle     =   2  'Dash
         Visible         =   0   'False
         X1              =   0
         X2              =   728
         Y1              =   50
         Y2              =   50
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   104
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Info && settings"
      Height          =   9975
      Left            =   11760
      TabIndex        =   12
      Top             =   120
      Width           =   4215
      Begin VB.CheckBox Dynamic_buffer 
         Caption         =   "Use automatic setting optimization"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   4680
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.CommandButton Kill_processing 
         Caption         =   "Stop analyze"
         Height          =   495
         Left            =   2040
         TabIndex        =   66
         Top             =   9000
         Width           =   1575
      End
      Begin VB.CheckBox EraseG 
         Caption         =   "Erase graphics at next analysis "
         Height          =   255
         Left            =   360
         TabIndex        =   55
         Top             =   9600
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.CommandButton Start_Stop 
         Caption         =   "Start analyze"
         Enabled         =   0   'False
         Height          =   495
         Left            =   360
         TabIndex        =   54
         Top             =   9000
         Width           =   1575
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   360
         Max             =   6000
         Min             =   10
         TabIndex        =   43
         Top             =   4320
         Value           =   50
         Width           =   3615
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   360
         Max             =   10000
         Min             =   50
         TabIndex        =   42
         Top             =   3600
         Value           =   70
         Width           =   3615
      End
      Begin VB.TextBox Default_Window_Length 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "50"
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox Default_buff 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "70"
         Top             =   3240
         Width           =   855
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   7440
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   6360
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   5280
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00808080&
         BorderStyle     =   2  'Dash
         X1              =   240
         X2              =   3960
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label CHI 
         BackStyle       =   0  'Transparent
         Caption         =   "Chi-test: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   85
         Top             =   2040
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         BorderStyle     =   2  'Dash
         X1              =   240
         X2              =   3960
         Y1              =   1090
         Y2              =   1090
      End
      Begin VB.Label LCG 
         BackStyle       =   0  'Transparent
         Caption         =   "CSLCG: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   82
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label GCNT 
         BackStyle       =   0  'Transparent
         Caption         =   "Total (CG)n : 0"
         Height          =   255
         Left            =   240
         TabIndex        =   79
         Top             =   2520
         Width           =   3015
      End
      Begin VB.Label TotalCGcontent 
         BackStyle       =   0  'Transparent
         Caption         =   "Total CG content: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   76
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Motif_F 
         BackStyle       =   0  'Transparent
         Caption         =   "Total motifs found: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   2760
         Width           =   3015
      End
      Begin VB.Label Contigs 
         BackStyle       =   0  'Transparent
         Caption         =   "Total headers found: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Line Line8 
         X1              =   120
         X2              =   4080
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Label Label19 
         Caption         =   "Default window size (b):"
         Height          =   255
         Left            =   1320
         TabIndex        =   41
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label17 
         Caption         =   "Default buffer size (b):"
         Height          =   255
         Left            =   1440
         TabIndex        =   39
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Line Line7 
         X1              =   360
         X2              =   3600
         Y1              =   8400
         Y2              =   8400
      End
      Begin VB.Line Line6 
         X1              =   240
         X2              =   3960
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Genom_files 
         BackStyle       =   0  'Transparent
         Caption         =   "File path and size: ..."
         Height          =   495
         Left            =   480
         TabIndex        =   36
         Top             =   8520
         Width           =   3135
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":57265
         Height          =   855
         Left            =   480
         TabIndex        =   33
         Top             =   7440
         Width           =   3615
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":572FE
         Height          =   1095
         Left            =   480
         TabIndex        =   31
         Top             =   6360
         Width           =   3615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":573D4
         Height          =   975
         Left            =   480
         TabIndex        =   29
         Top             =   5280
         Width           =   3615
      End
      Begin VB.Label Relative_pos 
         Caption         =   "Relative chromosome position: 0b"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label up_limit 
         Caption         =   "Y - Content: 0 %"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label buffer_number 
         Caption         =   "X - Buffer number: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   3015
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   1680
      Width           =   9975
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   600
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   725
      TabIndex        =   3
      Top             =   3000
      Width           =   10935
      Begin VB.Label Mes1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Example1 off"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   480
         Left            =   3480
         TabIndex        =   73
         Top             =   480
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Line Lin3 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   0
         X2              =   728
         Y1              =   70
         Y2              =   70
      End
      Begin VB.Line Lin1 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   0
         X2              =   728
         Y1              =   22
         Y2              =   22
      End
      Begin VB.Line Lin2 
         BorderColor     =   &H00808080&
         BorderStyle     =   2  'Dash
         Visible         =   0   'False
         X1              =   0
         X2              =   728
         Y1              =   46
         Y2              =   46
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   104
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1320
      Width           =   9975
   End
   Begin VB.CommandButton cmdOpenFASTA 
      Caption         =   "Open a FASTA file  ..."
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label W_B 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sliding Windows/Buffer: 0"
      Height          =   255
      Left            =   4560
      TabIndex        =   78
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Caption         =   "("
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   71
      Top             =   5160
      Width           =   135
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "for Data Mining"
      Height          =   255
      Left            =   8880
      TabIndex        =   65
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "EXAMPLE 3"
      Height          =   255
      Left            =   600
      TabIndex        =   51
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "EXAMPLE 2"
      Height          =   255
      Left            =   600
      TabIndex        =   50
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "EXAMPLE 1"
      Height          =   255
      Left            =   600
      TabIndex        =   49
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "25 % -"
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label23 
      Caption         =   "75 % -"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label22 
      Caption         =   "25 % -"
      Height          =   255
      Left            =   120
      TabIndex        =   46
      Top             =   6680
      Width           =   495
   End
   Begin VB.Label Label20 
      Caption         =   "75 % -"
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   5940
      Width           =   495
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "V1.0"
      Height          =   255
      Left            =   4560
      TabIndex        =   37
      Top             =   985
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   ")n - Input n:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   35
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   8520
      X2              =   9960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   5880
      X2              =   4560
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label program_message 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GENOMIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   480
      Left            =   3240
      TabIndex        =   26
      Top             =   480
      Width           =   7935
   End
   Begin VB.Label Label9 
      Caption         =   "Input the motif sequence:"
      Height          =   255
      Left            =   6840
      TabIndex        =   25
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label Label18 
      Caption         =   "1b"
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Top             =   9720
      Width           =   375
   End
   Begin VB.Label max_nucleo2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0b"
      Height          =   255
      Left            =   9360
      TabIndex        =   22
      Top             =   9720
      Width           =   2175
   End
   Begin VB.Label Label13 
      Caption         =   "100 % -"
      Height          =   255
      Left            =   25
      TabIndex        =   20
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "0 % -"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "1b"
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label max_nucleo1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0b"
      Height          =   255
      Left            =   9480
      TabIndex        =   17
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "50 % -"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   6300
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "50 % -"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Clean_DNA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0b"
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Label max_nucleo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0b"
      Height          =   255
      Left            =   9600
      TabIndex        =   9
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "1b"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "100 % -"
      Height          =   255
      Left            =   25
      TabIndex        =   7
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "0 % -"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Window content:"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Buffer Stream:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ________________________________                          _____________________
' /  GENOMIN for Data Mining       \________________________/       v1.00         |
' |                                                                               |
' |            Name:  GENOMIN                                                     |
' |        Category:  open source software                                        |
' |          Author:  Paul A. Gagniuc                                             |
' |                                                                               |
' |    Date Created:  September 2008                                              |
' |       Tested On:  Windows Vista, Windows XP, Windows 7                        |
' |           Notes:  buffering methods for reading biological information        |
' |                                                                               |
' |                  _____________________________                                |
' |_________________/                             \_______________________________|
'

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function ShellExecute Lib "Shell32.dll" _
                Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long

Dim buff As Variant
Dim sFile As String
Dim Window As Variant
Dim sector As Variant
Dim stop_proc As Boolean
Dim motif_count As Variant
Dim start_cronos As Variant
Dim stopp_cronos As Variant
Dim Window_Length As Variant
Dim total_sequence As Variant
Dim total_cr_headers As Variant
Dim general_CG_buffer As String
Dim Global_CG_content As Variant
Dim Global_No_content As Variant
Dim Global_CG_nr_buff As Variant
Dim position_in_sequence As Variant
Dim EXEMPLE_1_old_y As Variant
Dim EXEMPLE_1_old_x As Variant
Dim EXEMPLE_2_old_y As Variant
Dim EXEMPLE_2_old_x As Variant


Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Mes1.Visible = False
    Else
        Mes1.Visible = True
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Mes2.Visible = False
    Else
        Mes2.Visible = True
    End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        Mes3.Visible = False
    Else
        Mes3.Visible = True
    End If
End Sub

Private Sub cmdOpenFASTA_Click()

    Dim CC As cCommonDialog
    Set CC = New cCommonDialog
    
    If CC.VBGetOpenFileName(sFile, , True, , , True, "FASTA files (*.fa, *.fasta, *.fas)|*.fa;*.fasta;*.fas|All files|*.*", , , "Open a FASTA file", "", frmMain.hwnd, 0) Then
        
        Kill_processing.Enabled = True
        Start_Stop.Enabled = False
        Call FASTA_FILE
        Start_Stop.Enabled = True
        Kill_processing.Enabled = False
        
    End If
    
End Sub


Function Stop_Time()

    Timer1.Enabled = False
    stopp_cronos = GetTickCount
    
    Min_past.Caption = "Total: " & Int(((stopp_cronos - start_cronos) / 1000) / 60) & " Minute(s)"
    
    Clock_diff2.Caption = "Stop :" & Time
    
    GIN.Text = GIN.Text & Clock_diff2.Caption & vbCrLf
    GIN.Text = GIN.Text & Min_past.Caption & vbCrLf
    GIN.Text = GIN.Text & "GENOMIN has finished processing the data !" & vbCrLf

End Function


Private Sub Dynamic_buffer_Click()
    If Dynamic_buffer.Value = 1 Then
    
        Default_buff.Enabled = False
        Default_Window_Length.Enabled = False
        HScroll1.Enabled = False
        HScroll2.Enabled = False
    Else
    
        Default_buff.Enabled = True
        Default_Window_Length.Enabled = True
        HScroll1.Enabled = True
        HScroll2.Enabled = True
    
    End If
End Sub

Private Sub Form_Terminate()
    stop_proc = True
    Call Stop_Time
End Sub

Private Sub Form_Unload(Cancel As Integer)
    stop_proc = True
    Call Stop_Time
End Sub

Private Sub Kill_processing_Click()
    stop_proc = True
    Start_Stop.Enabled = True
    Kill_processing.Enabled = False
    
    HScroll1.Enabled = True
    HScroll2.Enabled = True
    Option1.Enabled = True
    Option2.Enabled = True
    
    Call Stop_Time
    
    GIN.Text = GIN.Text & "GENOMIN has been stopped ..." & vbCrLf
End Sub


Private Sub Rep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        Exit Sub
    End If
    
    '48 to 57 r ascii of numbers 0 to 9
    For NUMK = 48 To 57
            
            'blocks all keys except numbers 0 to 9
            If KeyAscii = NUMK Then
            Exit Sub
            End If
            
    Next
    
    KeyAscii = 0
End Sub

Private Sub Save_GENOMIN_Click()
On Error Resume Next

    Dim CC As cCommonDialog
    Set CC = New cCommonDialog
    Dim CurEnvFile As String
    
    If CC.VBGetSaveFileName("GenominProject1", CurEnvFile, , "FASTA files (*.htm)|*.htm;|All files|*.*", , "", "bla bla", ".htm", frmMain.hwnd, 0) Then
        
    Save_GENOMIN.Enabled = False
    cmdOpenFASTA.Enabled = False
    Start_Stop.Enabled = False
    Kill_processing.Enabled = False
    
    tmp1 = CurEnvFile & "1.bmp"
    tmp2 = CurEnvFile & "2.bmp"
    tmp3 = CurEnvFile & "3.bmp"
        
    Kill tmp1
    Kill tmp2
    Kill tmp3
    Kill CurEnvFile
        
    SavePicture Picture1.Image, tmp1
    SavePicture Picture2.Image, tmp2
    SavePicture Picture3.Image, tmp3
    
    If SSR.Value = 1 Then GoTo 2
    
    nFileNum = FreeFile
    
    Open App.Path & "\tmp.txt" For Input As nFileNum
        lLineCount = 1
        ' Read the no. of lines from tmp.txt file
        Do While Not EOF(nFileNum)
            Line Input #nFileNum, sNextLine
            lLineCount = lLineCount + 1
        Loop
    Close nFileNum
    
    
    Open App.Path & "\tmp.txt" For Input As nFileNum
        liCount = 1
        ' Read the contents of tmp.txt file
        Do While Not EOF(nFileNum)
           Line Input #nFileNum, sNextLine
           sNextLine = sNextLine & vbCrLf
           sText = sText & sNextLine
           
            'update caption for the user.
            liCount = liCount + 1
            program_message.Caption = "Save results " & Int((100 / lLineCount) * liCount) & "%"
            
            If stop_proc = True Then
                GoTo 2
            End If
            
            DoEvents
        
        Loop
    Close nFileNum
    
2:
    
    tmp_text = sText
    tmp_text = Replace(tmp_text, Chr(13), "<br>")
    
    tmp_f = Replace(Genom_files.Caption, Chr(13), "<br>")
    html_header = html_header & tmp_f & "<br>"
    
    html_header = html_header & Clock_diff1.Caption & "<br>"
    html_header = html_header & Clock_diff2.Caption & "<br>"
    html_header = html_header & Min_past.Caption & "<br>"
    
    If Option1.Value = True Then
        html_header = html_header & "Processing on buffer<br>"
    Else
        html_header = html_header & "Processing on window<br>"
    End If
    
    html_header = html_header & Motif_F.Caption & "<br>"
    html_header = html_header & Contigs.Caption & "<br>"
    html_header = html_header & W_B.Caption & "<br>"
    html_header = html_header & TotalCGcontent.Caption & "<br>"
    html_header = html_header & "(CG)n - Input n:" & Rep.Text & "<br>"
    html_header = html_header & "Input motif sequence:" & motif_sequence.Text & "<br>"
    html_header = html_header & "Default buffer size (b):" & Default_buff.Text & "<br>"
    html_header = html_header & "Default window size (b):" & Default_Window_Length.Text & "<br>"
    If Rec_buff.Value = 1 Then rec_tmp = "yes" Else rec_tmp = "no"
    html_header = html_header & "Recalibrated:" & rec_tmp & "<br>"
    
    html_save = "<font size=8>GENOMIN V1.0</font><hr><br>" & html_header & "<br><img border=1 src='" & tmp1 & "'><br><img border=1 src='" & tmp2 & "'><br><img border=1 src='" & tmp3 & "'><br><hr>" & tmp_text
    
    file1 = FreeFile
    Open CurEnvFile For Append As file1
        Print #file1, html_save
    Close #file1
        
    OpenLink = ShellExecute(hwnd, "open", CurEnvFile, vbNull, vbNull, 1)
        
    End If
    
    program_message.Caption = "GENOMIN"
    
    Save_GENOMIN.Enabled = True
    cmdOpenFASTA.Enabled = True
    Start_Stop.Enabled = True
    Kill_processing.Enabled = True

End Sub


Private Function FileExists(ByVal sFileName As String) As Boolean
    Dim intReturn As Integer
    
        On Error GoTo FileExists_Error
        intReturn = GetAttr(sFileName)
        FileExists = True
        
    Exit Function
FileExists_Error:
        FileExists = False
End Function


Function FASTA_FILE()
    'On Error Resume Next
    
    start_cronos = GetTickCount 'Starts the count.
    Clock_diff1.Caption = "Start :" & Time
    Timer1.Enabled = True
    
    Window = ""
    
    GIN.Text = ""
    GIN.Text = GIN.Text & Clock_diff1.Caption & vbCrLf
    
    Global_CG_content = 0
    Global_No_content = 0
    Global_CG_nr_buff = 0
    total_cr_headers = 0
    GCNT.Caption = "No. (CG)" & Rep.Text & " = 0"
    DoEvents
    
    old_sir_buf = Empty
    
    stop_proc = False
    HScroll1.Enabled = False
    HScroll2.Enabled = False
    Option1.Enabled = False
    Option2.Enabled = False
    
    general_CG_buffer = Empty
    EXEMPLE_1_old_y = 0
    EXEMPLE_1_old_x = 0
    EXEMPLE_2_old_y = 0
    EXEMPLE_2_old_x = 0
    motif_count = 0
    position_in_sequence = 0

    If FileExists(App.Path & "\tmp.txt") = True Then Kill App.Path & "\tmp.txt"

    If EraseG.Value = 1 Then
        If Check1.Value = 1 Then Picture1.Cls
        If Check2.Value = 1 Then Picture2.Cls
        If Check3.Value = 1 Then Picture3.Cls
    End If
    
    Dim FileNum As Integer
    Dim alta_secventa As Boolean
    Dim dat As String
    Dim i As Long
    

    
    FileNum = FreeFile
    
    Open sFile For Binary As #FileNum
    
        lungime = LOF(FileNum)
        
        If Dynamic_buffer.Value = 1 Then '################################
        
        If lungime < 20000 Then
            HScroll1.Value = 70
            Default_buff.Text = 70
        End If
        
        If lungime < 100000 And lungime > 20000 Then
            HScroll1.Value = 300
            Default_buff.Text = 300
        End If
        
        If lungime < 1000000 And lungime > 100000 Then
            HScroll1.Value = 1000
            Default_buff.Text = 1000
        End If
        
        If lungime > 1000000 And lungime < 1000000 Then
            HScroll1.Value = 3000
            Default_buff.Text = 3000
        End If
        
        If lungime > 1000000 Then
            HScroll1.Value = 10000
            Default_buff.Text = 10000
        End If
        
        End If '##########################################################
        
        
        Genom_files.Caption = "File name: " & Split(sFile, "\")(UBound(Split(sFile, "\"))) & vbCrLf & "File size: " & lungime & " bytes"
        GIN.Text = GIN.Text & Genom_files.Caption & vbCrLf
        'we read the total number of characters from the
        'FASTA file(DNA sequence + total_sequence_header).
        total_sequence = FileLen(sFile)
        
        
        'extract all CRLF characters
        max_nucleo.Caption = total_sequence 'Int(total_sequence - (total_sequence / 70)) & "b"
        Text4.Text = total_sequence
        max_nucleo1.Caption = max_nucleo.Caption
        max_nucleo2.Caption = max_nucleo.Caption
        
        'we allocate memory space for "dat" variable.
        dat = String$(buff, vbNullChar)
            
        For i = buff To lungime Step buff
1:
            Get #FileNum, , dat
            Seek #FileNum, i + 1
            'we use Seek function for a sequential
            'data reading, regardless of file size.
            
            tmp_dat = tmp_dat & dat
                    
            ' a double buffer space is sufficient even if the ">" sign is at the
            ' end of the buffer space, the buffer being equal to or greater than
            ' the first sequence header.
                
            If InStr(dat, ">") Then
                i = i + buff             ' double buffering , buff * 2 ...
                alta_secventa = True     ' we mark the beginning of the other sequence within the same file !
                GoTo 1
            End If
                
            If stop_proc = True Then
                Close #FileNum
                Exit Function
            End If
                    
            Call procesare(tmp_dat, alta_secventa)
            tmp_dat = Empty
            alta_secventa = False
                
        Next i
        
    Close #FileNum
    
      
    If Rec_buff.Value = 1 Then
    If Erase_rec.Value = 1 And Check2.Value = 1 Then Picture2.Cls
    program_message.ForeColor = vbRed
    program_message.Caption = "Recalibrate, please wait ..."
    DoEvents
    
    If Check2.Value = 1 Then Call Recalibrate
    
    program_message.ForeColor = &H404040
    program_message.Caption = "GENOMIN"
    End If
        
    HScroll1.Enabled = True
    HScroll2.Enabled = True
    Option1.Enabled = True
    Option2.Enabled = True
    
    Start_Stop.Enabled = True
    
    Call Stop_Time
    
    MsgBox "Processing finished for " & Split(sFile, "\")(UBound(Split(sFile, "\"))) & " file !"
End Function



Function procesare(ByVal x As String, ByVal alta_seq As Boolean)
    
    x = LCase(x)
    ' if a new sequence begins in the same file, we
    ' exclude the header sequence from the buffer.
    
        If alta_seq = True Then
            
            tmp_1 = Split(x, ">")(0)
            tmp = Split(x, ">")(1)
            
               If InStr(tmp, Chr(10)) Then
                   tmp_2 = Split(tmp, Chr(10))(1)
                   
                   total_cr_headers = total_cr_headers + 1
                   
                   Call add_tmp_result("[" & total_cr_headers & "] Sequence: " & Split(tmp, Chr(10))(0) & vbCrLf)
                   GIN.Text = GIN.Text & "[" & total_cr_headers & "] Sequence: " & Split(tmp, Chr(10))(0) & vbCrLf
                   Contigs.Caption = "Total headers found: " & total_cr_headers
            
               Else
                   tmp_2 = ""
               End If
        
        
         x = tmp_1 & tmp_2
        
        End If
    
    
    
    x = Replace(x, vbCrLf, "")
    x = Replace(x, Chr(13), "")
    x = Replace(x, Chr(10), "")
    
    
    'Below is a list of other characters that can occur
    'inside NCBI FASTA files, which can be filtered according
    'to the rules imposed by the researcher.
    
    
    If InStr(x, "n") <> 0 Then
    par = Picture1.ScaleWidth / total_sequence
    y = 10
    XX = par * position_in_sequence
    
    Picture1.Line (XX, y)-(XX + sector, y), vbGray
    Picture1.Line (XX, y + 2)-(XX + sector, y + 2), vbGray
    End If
    
    
    x = Replace(x, "m", "")
    x = Replace(x, "s", "")
    x = Replace(x, "w", "")
    x = Replace(x, "b", "")
    x = Replace(x, "u", "")
    x = Replace(x, "d", "")
    x = Replace(x, "r", "")
    x = Replace(x, "h", "")
    x = Replace(x, "y", "")
    x = Replace(x, "v", "")
    x = Replace(x, "k", "")
    x = Replace(x, "n", "")
    x = Replace(x, "-", "")
    
    'M -> A C (amino)
    'S -> G C (strong)
    'W -> A T (weak)
    'B -> G T C
    'U -> uridine
    'D -> G A T
    'R -> G A (purine)
    'H -> A C T
    'Y -> T C (pyrimidine)
    'V -> G C A
    'K -> G T (keto)
    'N -> A G C T (any)
    '- -> gap of indeterminate length
    
    'Variable "x" contains a clear DNA sequence,
    'which can be processed by an external function.
    'To avoid disruptions in the DNA sequence, in the sliding
    'window case, we introduce the Buffer_Stream variable.
    
    'If we have many "N" characters, they will be removed
    'and the buffer will be empty. If variable x is empty,
    'processing is temporarily halted.
    
    If x <> "" Then
    
    Buffer_Stream = Window & x
    
    Text1.Text = Buffer_Stream
    
    'GIN.Text = GIN.Text & Buffer_Stream
    
    position_in_sequence = position_in_sequence + Len(x)
    Clean_DNA.Caption = "Clean sequence = " & position_in_sequence & "b"
    Text3.Text = position_in_sequence
    'From here we can do any experiments on the DNA sequence.
    
    If Option1.Value = True Then
        Call Buffer_Sector(Buffer_Stream)
    Else
        Call Slide_Window(Buffer_Stream)
    End If
    
    End If
    
    DoEvents
End Function

Function Buffer_Sector(ByVal x As String)

    'pentru a lua ultima bucata din buffer n-1.
    'Window = "|" & Mid(x, Len(x) - Len(motif_sequence.Text) + 2, Len(motif_sequence.Text)) & "|"
    
    Window = Mid(x, Len(x) - Len(motif_sequence.Text) + 2, Len(motif_sequence.Text)) '& "|"
    
    'GIN.Text = GIN.Text & x & vbCrLf
    
    If Check1.Value = 1 Then Call EXEMPLE_1(x)
    If Check2.Value = 1 Then Call EXEMPLE_2(x)
    If Check3.Value = 1 Then Call EXEMPLE_3(x)
    
    DoEvents
End Function



Function Slide_Window(ByVal x As String)

    For i = 1 To Len(x) - Window_Length
        Window = Mid(x, i, Window_Length)
        
        Text2.Text = Window
        Call Process_Window(Window)
    Next i

End Function

Function Process_Window(ByVal Window As String)
    
    If Check1.Value = 1 Then Call EXEMPLE_1(Window)
    If Check2.Value = 1 Then Call EXEMPLE_2(Window)
    If Check3.Value = 1 Then Call EXEMPLE_3(Window)
    
    DoEvents
End Function


Function EXEMPLE_1(ByRef Window As String)
    
    For i = 1 To Len(Window)
    
      nucleotida = Mid(Window, i, 1)
      If nucleotida = "a" Then A = A + 1
      If nucleotida = "t" Then T = T + 1
      If nucleotida = "g" Then G = G + 1
      If nucleotida = "c" Then C = C + 1
    
    Next i
    
    Total_CG = (100 / (C + G + T + A)) * (C + G)
    
    
    Global_CG_content = Global_CG_content + Total_CG
    Global_No_content = Global_No_content + 1
    tmp_gg = Global_CG_content / Global_No_content
    
    
    tmp_proc = Format$(tmp_gg, "##.###")
    TotalCGcontent.Caption = "Total CG content: " & tmp_proc & "%"
    
    Dim dif_len As String
    dif_len = (100 / total_sequence) * tmp_proc
    
    'on some computers, the values are "treated" differently ex. 0.546547648 is 0,546547648, so we need the line below
    dif_len = Replace(dif_len, ",", ".")
    
    LCG.Caption = "CSLCG: " & Split((dif_len), ".")(0) & "." & Mid(Split((dif_len), ".")(1), 1, 3)
    
    'tmp_chi = ((tmp_proc - 50) ^ 2 / 50) + (((100 - tmp_proc) - 50) ^ 2 / 50)
    
    'on some computers, the values are "treated" differently ex. 0.546547648 is 0,546547648, so we need the line below
    'tmp_chi = Replace(tmp_chi, ",", ".")
    
    'CHI.Caption = "Chi-test: " & Split((tmp_chi), ".")(0) & "." & Mid(Split((tmp_chi), ".")(1), 1, 3)
    
    par = Picture1.ScaleWidth / total_sequence
    y = Picture1.ScaleHeight / 100
    x = par * position_in_sequence
    
    
    If LOB1.Value = 1 Then
        Picture1.Line (x, 100)-(x, 100 - (Total_CG * y)), vbRed
    Else
        Picture1.Line (EXEMPLE_1_old_x, 100 - EXEMPLE_1_old_y)-(x, 100 - (Total_CG * y)), vbRed
        EXEMPLE_1_old_y = Total_CG * y
        EXEMPLE_1_old_x = x
    End If
    
    Line1.X1 = (par * position_in_sequence) + 1
    Line1.X2 = (par * position_in_sequence) + 1
    
    Call add_tmp_result("EXAMPLE 1 - No. buffer: [" & Int(x) & "] - > (G and C) percentage:" & Format$(Total_CG, "##.###") & "%" & vbCrLf)
    
    DoEvents
End Function

Function add_tmp_result(ByVal content As String)
    file1 = FreeFile
    Open App.Path & "\tmp.txt" For Append As file1
    Print #file1, content
    Close #file1
End Function



Function EXEMPLE_2(ByRef Window As String)
    Dim CG_nr() As String
    
    nucleo_test = LCase(DubleN.Text)
    For ye = 1 To Val(Rep.Text)
        rep_CG = rep_CG & LCase(nucleo_test)
    Next ye
    
    
    CG_nr = Split(Window, rep_CG)
    CG_nr_buff = UBound(CG_nr)
    
    op = CG_nr_buff * (2 * Val(Rep.Text))
    Total_CG = (100 / Len(Window)) * op
    
    par = Picture2.ScaleWidth / total_sequence
    y = Picture2.ScaleHeight / 100
    x = par * position_in_sequence
    
    general_CG_buffer = general_CG_buffer & "," & Total_CG & "#" & x
    
    If CG_nr_buff > 0 Then
        Global_CG_nr_buff = Global_CG_nr_buff + CG_nr_buff
        GCNT.Caption = "Total (CG)" & Rep.Text & " = " & Global_CG_nr_buff
    End If
    
    DoEvents
    
    If Rec_buff.Value = 0 Then
        Call add_tmp_result("EXAMPLE 2 - No. buffer: [" & Int(x) & "] - > (GC)<font size=2>n</font>, n=" & Rep.Text & " -> percentage:" & Int(Total_CG) & "%" & vbCrLf)
    End If
    
    
    If LOB2.Value = 1 Then
        Picture2.Line (x, 100)-(x, 100 - Total_CG), vbBlue
    Else
        Picture2.Line (EXEMPLE_2_old_x, 100 - EXEMPLE_2_old_y)-(x, 100 - (Total_CG * y)), vbBlue
        EXEMPLE_2_old_y = Total_CG * y
        EXEMPLE_2_old_x = x
    End If
    
    Line2.X1 = (par * position_in_sequence) + 1
    Line2.X2 = (par * position_in_sequence) + 1
    
    DoEvents
End Function

Function Recalibrate()
    On Error Resume Next
    
    Dim sir_buf()  As String
    sir_buf = Split(general_CG_buffer, ",")
    
    For T = 0 To UBound(sir_buf)
        If Val(Split(sir_buf(T), "#")(0)) > old_sir_buf Then
            old_sir_buf = Val(Split(sir_buf(T), "#")(0))
        End If
    Next T
    
    buffers_no = UBound(sir_buf)
    
    For T = 0 To buffers_no
    
        Total_CG = Int((100 / old_sir_buf) * Val(Split(sir_buf(T), "#")(0)))
        
        par = Picture2.ScaleWidth / total_sequence
        x = Val(Split(sir_buf(T), "#")(1))
        
        If LOB2.Value = 1 Then
            Picture2.Line (x, 100)-(x, 100 - Total_CG), vbBlue
        Else
            Picture2.Line (x_tmp, 100 - y_tmp)-(x, 100 - Total_CG), vbBlue
            x_tmp = x
            y_tmp = Total_CG
        End If
        
        Line2.X1 = (par * position_in_sequence) + 1
        Line2.X2 = (par * position_in_sequence) + 1
        
        
        Call add_tmp_result("EXAMPLE 2 - No. buffer: [" & Int(Split(sir_buf(T), "#")(1)) & "] - > (GC)<font size=2>n</font>, n=" & Rep.Text & " -> percentage:" & Int(Split(sir_buf(T), "#")(0)) & "%" & vbCrLf)
        
        If stop_proc = True Then Exit Function
        
        program_message.Caption = "Recalibrate, please wait " & Int((100 / buffers_no) * T) & "%"
        
        DoEvents
    Next T

End Function


Function EXEMPLE_3(ByRef Window As String)
    
    'Window = Replace(Window, LCase(motif_sequence.Text), "*")
    
    '#################################################
    
    Dim n_motifs() As String
    
    n_motifs = Split(LCase(Window), LCase(motif_sequence.Text))
    
    tmp_motif = UBound(n_motifs)
    If tmp_motif > 0 Then flag = 100 Else flag = 0
    
    motif_count = motif_count + tmp_motif
    Motif_F.Caption = "Total motifs found: " & motif_count
    
    'GIN.Text = GIN.Text & Window & " - " & tmp_motif & vbCrLf
    
    '#################################
    
    'motif = LCase(motif_sequence.Text)
    'oo = 1
    'flag = 0
    
    '1:
    'If pp = 0 Then oo = 1 Else oo = pp + Len(motif_sequence.Text) - 1 '1
    'pp = InStr(oo, Window, motif)
    'If pp <> 0 Then
    
    'flag = 100
    'motif_count = motif_count + 1
    'tmp_motif = tmp_motif + 1
    '
    
    'GoTo 1
    
    'End If
    
    '##########################################
    par = Picture1.ScaleWidth / total_sequence
    Line3.X1 = (par * position_in_sequence) + 1
    Line3.X2 = (par * position_in_sequence) + 1
    
    If flag = 100 Then
    
        Total_CG = flag
        y = Picture1.ScaleHeight / 100
        x = par * position_in_sequence
        
        Picture3.Line (x, 0)-(x, Total_CG * y), &H8000&
        
        If UTN.Value = 1 Then
            Picture3.CurrentX = x + 1
            Picture3.CurrentY = 20
            Picture3.Font.Size = 8
            Picture3.Print "M=" & tmp_motif
        End If
        
        Call add_tmp_result("EXAMPLE 3 - No. motifs found: [" & tmp_motif & "] - > Relative chromosome position:" & position_in_sequence & "b" & vbCrLf)
    
    End If
    
    DoEvents
End Function

Private Sub Form_Load()
    'for the rest of the file, the first line at the beginning
    'of each sequence - has a constant length.
    
    buff = HScroll1.Value
    Window_Length = HScroll2.Value
    W_B.Caption = "Sliding Windows/Buffer: " & HScroll1.Value - HScroll2.Value
    
    old_sir_buf = 0
    position_in_sequence = 0
    stop_proc = False
    
    DubleN.AddItem "AT"
    DubleN.AddItem "AC"
    DubleN.AddItem "AG"
    DubleN.AddItem "AA"
    
    DubleN.AddItem "CA"
    DubleN.AddItem "CT"
    DubleN.AddItem "CG"
    DubleN.AddItem "CC"
    
    DubleN.AddItem "GT"
    DubleN.AddItem "GA"
    DubleN.AddItem "GC"
    DubleN.AddItem "GG"
    
    DubleN.AddItem "TC"
    DubleN.AddItem "TG"
    DubleN.AddItem "TA"
    DubleN.AddItem "TT"
End Sub

Private Sub HScroll1_Change()
    If HScroll1.Value <= HScroll2.Value Then HScroll2.Value = HScroll1.Value - 1
    
    Default_buff.Text = HScroll1.Value
    buff = HScroll1.Value
    
    W_B.Caption = "Sliding Windows/Buffer: " & HScroll1.Value - HScroll2.Value
End Sub

Private Sub HScroll2_Change()
    If HScroll2.Value >= HScroll1.Value Then HScroll1.Value = HScroll2.Value + 1
    
    Default_Window_Length.Text = HScroll2.Value
    Window_Length = HScroll2.Value
    W_B.Caption = "Sliding Windows/Buffer: " & HScroll1.Value - HScroll2.Value
End Sub

Private Sub motif_sequence_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) And (KeyAscii <> 97) And (KeyAscii <> 99) And (KeyAscii <> 103) And (KeyAscii <> 116) And _
    (KeyAscii <> 65) And (KeyAscii <> 67) And (KeyAscii <> 71) And (KeyAscii <> 84) Then KeyAscii = 0
    
    'a = 97
    'c = 99
    'g = 103
    't = 116
    
    'A = 65
    'C = 67
    'G = 71
    'T = 84
    
    motif_sequence.MaxLength = HScroll1.Value
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    up_limit.Caption = "Y - Content:" & (Picture1.ScaleHeight - y) & " %"
    
    If total_sequence = Empty Then Exit Sub
    nr_total_buffere = total_sequence / HScroll1.Value
    sector = (Picture1.ScaleWidth / nr_total_buffere)
    buffer_number.Caption = "X - Buffer number:" & Int(x / sector) + 1
    
    Relative_pos.Caption = "Relative chromosome position:" & Int((total_sequence / Picture1.ScaleWidth) * x) & "b"
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    up_limit.Caption = "Y - Content:" & (Picture2.ScaleHeight - y) & " %"
    
    
    If total_sequence = Empty Then Exit Sub
    nr_total_buffere = total_sequence / HScroll1.Value
    sector = (Picture2.ScaleWidth / nr_total_buffere)
    buffer_number.Caption = "X - Buffer number:" & Int(x / sector) + 1
    
    Relative_pos.Caption = "Relative chromosome position:" & Int((total_sequence / Picture1.ScaleWidth) * x) & "b"
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    up_limit.Caption = "Y - Content:" & (Picture3.ScaleHeight - y) & " %"
    
    If total_sequence = Empty Then Exit Sub
    nr_total_buffere = total_sequence / HScroll1.Value
    sector = (Picture3.ScaleWidth / nr_total_buffere)
    buffer_number.Caption = "X - Buffer number:" & Int(x / sector) + 1
    
    Relative_pos.Caption = "Relative chromosome position:" & Int((total_sequence / Picture1.ScaleWidth) * x) & "b"
End Sub


Private Sub Rep_Change()

    If Option1.Value = True Then
    
        'on buffer (high processing speed)
        If Val(Rep.Text) >= HScroll1.Value Then
            MsgBox "CG repetitions > buffer size !"
            Rep.Text = HScroll1.Value - 1
        End If
    
    Else
    
        'on window (low processing speed, but may be needed for different experiments)
        If Val(Rep.Text) >= HScroll2.Value Then
            MsgBox "CG repetitions > window size !"
            Rep.Text = HScroll1.Value - 1
        End If
    
    End If

End Sub

Private Sub Start_Stop_Click()
    Start_Stop.Enabled = False
    Kill_processing.Enabled = True
    Call FASTA_FILE
End Sub

Private Sub Timer1_Timer()
    ' measuring the processing time
    Clock_diff2.Caption = "Stop: " & Time
    t_min = Int(((GetTickCount - start_cronos) / 1000) / 60)
    
    If t_min > 60 Then
        min_ramas = t_min Mod 60
        h_time = Int(t_min / 60)
        Min_past.Caption = "Total: " & h_time & "," & min_ramas & " Hour(s)"
    Else
        Min_past.Caption = "Total: " & t_min & " Minute(s)"
    End If
    
    DoEvents
End Sub

Private Sub UG1_Click()
    If UG1.Value = 1 Then
        Lin1.Visible = True
        Lin2.Visible = True
        Lin3.Visible = True
    Else
        Lin1.Visible = False
        Lin2.Visible = False
        Lin3.Visible = False
    End If
End Sub

Private Sub UG2_Click()
    If UG2.Value = 1 Then
        Lin4.Visible = True
        Lin5.Visible = True
        Lin6.Visible = True
    Else
        Lin4.Visible = False
        Lin5.Visible = False
        Lin6.Visible = False
    End If
End Sub
