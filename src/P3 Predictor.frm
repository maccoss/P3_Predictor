VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form P3Frm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P3 Predictor"
   ClientHeight    =   10605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10605
   ScaleWidth      =   17610
   Begin VB.Frame RTFrm 
      Caption         =   "RT Prediction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11040
      TabIndex        =   45
      Top             =   3960
      Width           =   3015
      Begin VB.TextBox RTWidthTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         TabIndex        =   51
         Text            =   "10"
         Top             =   200
         Width           =   615
      End
      Begin VB.TextBox RTIntTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   49
         Text            =   "-22.57"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox RTSlpTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   48
         Text            =   "1.5879"
         Top             =   220
         Width           =   615
      End
      Begin VB.Label RTWidthLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Width:"
         Height          =   255
         Left            =   1680
         TabIndex        =   50
         Top             =   240
         Width           =   495
      End
      Begin VB.Label RTIntLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Intercept:"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   600
         Width           =   735
      End
      Begin VB.Label RTSlpLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Slope:"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   615
      End
   End
   Begin RichTextLib.RichTextBox ProteinRTB 
      Height          =   3975
      Left            =   480
      TabIndex        =   44
      Top             =   600
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7011
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"P3 Predictor.frx":0000
   End
   Begin VB.ListBox SelectProductIonList 
      Appearance      =   0  'Flat
      Height          =   2055
      Left            =   14640
      Style           =   1  'Checkbox
      TabIndex        =   41
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton RemoveAllCmd 
      Caption         =   "Remove All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15960
      TabIndex        =   39
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton AddAllSelectedCmd 
      Caption         =   "Add Selected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15960
      TabIndex        =   38
      Top             =   5400
      Width           =   1455
   End
   Begin VB.ComboBox OrganismCombo 
      Height          =   315
      ItemData        =   "P3 Predictor.frx":0082
      Left            =   11280
      List            =   "P3 Predictor.frx":0084
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   2400
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog dbsaveresults 
      Left            =   16320
      Top             =   9960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      PrinterDefault  =   0   'False
   End
   Begin VB.Frame ProteinFeaFrm 
      Caption         =   "Protein Features"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   11040
      TabIndex        =   28
      Top             =   120
      Width           =   3135
      Begin VB.ComboBox EnzymeCombo 
         Height          =   315
         Left            =   240
         TabIndex        =   42
         Text            =   "EnzymeCombo"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox ExcludeFirstAAtxt 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1320
         TabIndex        =   34
         Text            =   "25"
         Top             =   680
         Width           =   615
      End
      Begin VB.CheckBox RaggedEndsChk 
         Caption         =   "Exclude Potential Ragged Ends"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.Label EnzymeLbl 
         Caption         =   "Enzyme:"
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
         Left            =   240
         TabIndex        =   43
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label ElimLbl2 
         Caption         =   "AAs"
         Height          =   255
         Left            =   2040
         TabIndex        =   35
         Top             =   720
         Width           =   495
      End
      Begin VB.Label ElimLbl 
         Caption         =   "Eliminate first "
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CommandButton SaveCSVCmd 
      Cancel          =   -1  'True
      Caption         =   "Output CSV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15960
      TabIndex        =   27
      Top             =   9240
      Width           =   1455
   End
   Begin VB.CommandButton RemoveCmd 
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15960
      TabIndex        =   26
      Top             =   6600
      Width           =   1455
   End
   Begin VB.ListBox SRMOutputList 
      Appearance      =   0  'Flat
      Height          =   4710
      Left            =   12600
      Sorted          =   -1  'True
      TabIndex        =   25
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Frame CEFrm 
      Caption         =   "Collision Energy Prediction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11040
      TabIndex        =   16
      Top             =   2880
      Width           =   3015
      Begin VB.TextBox CEIntTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Text            =   "3.314"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox CESlpTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Text            =   "0.034"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label CEIntLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Intercept:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label CESlpLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Slope:"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.Frame PeptideFeaFrm 
      Caption         =   "Peptide Features"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   7560
      TabIndex        =   9
      Top             =   120
      Width           =   3135
      Begin VB.CheckBox RPnKPContPepChk 
         Caption         =   "R-P or K-P"
         Height          =   255
         Left            =   600
         TabIndex        =   31
         Top             =   3720
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox GlycosContPepChk 
         Caption         =   "N-X-S/T motif"
         Height          =   255
         Left            =   600
         TabIndex        =   30
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CheckBox MonoPrecursorChk 
         Caption         =   "Use Monoisotopic Precursor Ions"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox HisContPepChk 
         Caption         =   "His"
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   3000
         Width           =   2295
      End
      Begin VB.CheckBox MonoProductChk 
         Caption         =   "Use Monoisotopic Product Ions"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.TextBox MinLenTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Text            =   "7"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox MaxLenTxt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Text            =   "25"
         Top             =   740
         Width           =   615
      End
      Begin VB.CheckBox CysContPepChk 
         Caption         =   "Cys"
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   2280
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox MetContPepChk 
         Caption         =   "Met"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   2640
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.Label ExcludeLbl 
         Caption         =   "Exclude Peptides Containing:"
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
         Left            =   240
         TabIndex        =   32
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label MinLenLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Min Peptide Length:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   390
         Width           =   1455
      End
      Begin VB.Label MaxLenLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Max Peptide Length:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   770
         Width           =   1575
      End
   End
   Begin VB.TextBox ProductTxt 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   8400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   5400
      Width           =   3975
   End
   Begin VB.TextBox PrecursorTxt 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   5400
      Width           =   3615
   End
   Begin VB.ListBox PeptideList 
      Appearance      =   0  'Flat
      Height          =   4755
      ItemData        =   "P3 Predictor.frx":0086
      Left            =   480
      List            =   "P3 Predictor.frx":0088
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   5400
      Width           =   3615
   End
   Begin VB.CommandButton ResetCmd 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14880
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton DigestCmd 
      Caption         =   "Digest Protein"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14880
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label ProductIonLbl 
      Caption         =   "Product Ions:"
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
      Left            =   14640
      TabIndex        =   40
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label LibraryLbl 
      Caption         =   "Check Library:"
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
      Left            =   11280
      TabIndex        =   37
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label TransitionsLbl 
      Caption         =   "SRM Transitions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   22
      Top             =   5040
      Width           =   3735
   End
   Begin VB.Label ProductLbl 
      Caption         =   "Product Ion Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   5040
      Width           =   3615
   End
   Begin VB.Label PrecursorLbl 
      Caption         =   "Precursor Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label PeptideLbl 
      Caption         =   "Peptide Sequences"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   5040
      Width           =   3615
   End
   Begin VB.Label ProteinLbl 
      Caption         =   "Paste Protein Sequence"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "P3Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileError As Boolean
Dim AAcount(1 To 20, 1 To 11) As Variant
Dim precursormzarray() As Integer
Dim MassiveArray() As Integer
Dim LibraryStatsFile() As Variant
Const Program = "P3 Predictor"
Const Version = 0.5






Public Sub Form_Load()

'----------------------------------------------------------
' Load data into AAcount array
' 1,1 to 20,1 are the single letter amino acid strings
' 1,2 to 20,2 are the corresponding # of residues in the
'       peptides (not added until program is run)
' 1,3 to 20,3 are the # of carbons in the residue
' 1,4 to 20,4 are the # of hydrogens in the residue
' 1,5 to 20,5 are the # of nitrogens in the residue
' 1,6 to 20,6 are the # of oxygens in the residue
' 1,7 to 20,7 are the # of sulfurs in the residue
' 1,8 to 20,8 are the residue monoisotopic mass
' 1,9 to 20,9 are the residue average mass
' 1,10 to 20,10 are the retention coefficients for amino acids
' 1,11 to 20,11 are the N-term retention coefficients
'----------------------------------------------------------

AAcount(1, 1) = "A"
    AAcount(1, 3) = 3   ' carbons
    AAcount(1, 4) = 5   ' hydrogens
    AAcount(1, 5) = 1   ' nitrogens
    AAcount(1, 6) = 1   ' oxygens
    AAcount(1, 7) = 0   ' sulfurs
    AAcount(1, 8) = 71.03711
    AAcount(1, 9) = 71.0788
    AAcount(1, 10) = 0.8    'AA retention coefficients
    AAcount(1, 11) = -1.5   'N-term retention coefficients

AAcount(2, 1) = "R"
    AAcount(2, 3) = 6
    AAcount(2, 4) = 12
    AAcount(2, 5) = 4
    AAcount(2, 6) = 1
    AAcount(2, 7) = 0
    AAcount(2, 8) = 156.10111
    AAcount(2, 9) = 156.1876
    AAcount(2, 10) = -1.3
    AAcount(2, 11) = 8
    
AAcount(3, 1) = "N"
    AAcount(3, 3) = 4
    AAcount(3, 4) = 6
    AAcount(3, 5) = 2
    AAcount(3, 6) = 2
    AAcount(3, 7) = 0
    AAcount(3, 8) = 114.04293
    AAcount(3, 9) = 114.1039
    AAcount(3, 10) = -1.2
    AAcount(3, 11) = 5
    
AAcount(4, 1) = "D"
    AAcount(4, 3) = 4
    AAcount(4, 4) = 5
    AAcount(4, 5) = 1
    AAcount(4, 6) = 3
    AAcount(4, 7) = 0
    AAcount(4, 8) = 115.02694
    AAcount(4, 9) = 115.0886
    AAcount(4, 10) = -0.5
    AAcount(4, 11) = 9
    
AAcount(5, 1) = "C"     'This elemental composition is for carboxyamidomethylated cysteine
    AAcount(5, 3) = 5
    AAcount(5, 4) = 8
    AAcount(5, 5) = 2
    AAcount(5, 6) = 2
    AAcount(5, 7) = 1
    AAcount(5, 8) = 160.0306
    AAcount(5, 9) = 160.1968
    AAcount(5, 10) = -0.8
    AAcount(5, 11) = 4
    
AAcount(6, 1) = "Q"
    AAcount(6, 3) = 5
    AAcount(6, 4) = 8
    AAcount(6, 5) = 2
    AAcount(6, 6) = 2
    AAcount(6, 7) = 0
    AAcount(6, 8) = 128.05858
    AAcount(6, 9) = 128.1308
    AAcount(6, 10) = -0.9
    AAcount(6, 11) = 1
    
AAcount(7, 1) = "E"
    AAcount(7, 3) = 5
    AAcount(7, 4) = 7
    AAcount(7, 5) = 1
    AAcount(7, 6) = 3
    AAcount(7, 7) = 0
    AAcount(7, 8) = 129.04259
    AAcount(7, 9) = 129.1155
    AAcount(7, 10) = 0
    AAcount(7, 11) = 7
    
AAcount(8, 1) = "G"
    AAcount(8, 3) = 2
    AAcount(8, 4) = 3
    AAcount(8, 5) = 1
    AAcount(8, 6) = 1
    AAcount(8, 7) = 0
    AAcount(8, 8) = 57.02146
    AAcount(8, 9) = 57.052
    AAcount(8, 10) = -0.9
    AAcount(8, 11) = 5

AAcount(9, 1) = "H"
    AAcount(9, 3) = 6
    AAcount(9, 4) = 7
    AAcount(9, 5) = 3
    AAcount(9, 6) = 1
    AAcount(9, 7) = 0
    AAcount(9, 8) = 137.05891
    AAcount(9, 9) = 137.1412
    AAcount(9, 10) = -1.3
    AAcount(9, 11) = 4

AAcount(10, 1) = "I"
    AAcount(10, 3) = 6
    AAcount(10, 4) = 11
    AAcount(10, 5) = 1
    AAcount(10, 6) = 1
    AAcount(10, 7) = 0
    AAcount(10, 8) = 113.08406
    AAcount(10, 9) = 113.1595
    AAcount(10, 10) = 8.4
    AAcount(10, 11) = -8
    
AAcount(11, 1) = "L"
    AAcount(11, 3) = 6
    AAcount(11, 4) = 11
    AAcount(11, 5) = 1
    AAcount(11, 6) = 1
    AAcount(11, 7) = 0
    AAcount(11, 8) = 113.08406
    AAcount(11, 9) = 113.1595
    AAcount(11, 10) = 9.6
    AAcount(11, 11) = -9
    
AAcount(12, 1) = "K"
    AAcount(12, 3) = 6
    AAcount(12, 4) = 12
    AAcount(12, 5) = 2
    AAcount(12, 6) = 1
    AAcount(12, 7) = 0
    AAcount(12, 8) = 128.09496
    AAcount(12, 9) = 128.1742
    AAcount(12, 10) = -1.9
    AAcount(12, 11) = 4.6
    
AAcount(13, 1) = "M"
    AAcount(13, 3) = 5
    AAcount(13, 4) = 9
    AAcount(13, 5) = 1
    AAcount(13, 6) = 1
    AAcount(13, 7) = 1
    AAcount(13, 8) = 131.04049
    AAcount(13, 9) = 131.1986
    AAcount(13, 10) = 5.8
    AAcount(13, 11) = -5.5
    
AAcount(14, 1) = "F"
    AAcount(14, 3) = 9
    AAcount(14, 4) = 9
    AAcount(14, 5) = 1
    AAcount(14, 6) = 1
    AAcount(14, 7) = 0
    AAcount(14, 8) = 147.06841
    AAcount(14, 9) = 147.1766
    AAcount(14, 10) = 10.5
    AAcount(14, 11) = -7

AAcount(15, 1) = "P"
    AAcount(15, 3) = 5
    AAcount(15, 4) = 7
    AAcount(15, 5) = 1
    AAcount(15, 6) = 1
    AAcount(15, 7) = 0
    AAcount(15, 8) = 97.05276
    AAcount(15, 9) = 97.1167
    AAcount(15, 10) = 0.2
    AAcount(15, 11) = 4
    
AAcount(16, 1) = "S"
    AAcount(16, 3) = 3
    AAcount(16, 4) = 5
    AAcount(16, 5) = 1
    AAcount(16, 6) = 2
    AAcount(16, 7) = 0
    AAcount(16, 8) = 87.03204
    AAcount(16, 9) = 87.0782
    AAcount(16, 10) = -0.8
    AAcount(16, 11) = 5

AAcount(17, 1) = "T"
    AAcount(17, 3) = 4
    AAcount(17, 4) = 7
    AAcount(17, 5) = 1
    AAcount(17, 6) = 2
    AAcount(17, 7) = 0
    AAcount(17, 8) = 101.04768
    AAcount(17, 9) = 101.1051
    AAcount(17, 10) = 0.4
    AAcount(17, 11) = 5
    
AAcount(18, 1) = "W"
    AAcount(18, 3) = 11
    AAcount(18, 4) = 10
    AAcount(18, 5) = 2
    AAcount(18, 6) = 1
    AAcount(18, 7) = 0
    AAcount(18, 8) = 186.07931
    AAcount(18, 9) = 186.2133
    AAcount(18, 10) = 11
    AAcount(18, 11) = -4
    
AAcount(19, 1) = "Y"
    AAcount(19, 3) = 9
    AAcount(19, 4) = 9
    AAcount(19, 5) = 1
    AAcount(19, 6) = 2
    AAcount(19, 7) = 0
    AAcount(19, 8) = 163.06333
    AAcount(19, 9) = 163.176
    AAcount(19, 10) = 4
    AAcount(19, 11) = -3
    
AAcount(20, 1) = "V"
    AAcount(20, 3) = 5
    AAcount(20, 4) = 9
    AAcount(20, 5) = 1
    AAcount(20, 6) = 1
    AAcount(20, 7) = 0
    AAcount(20, 8) = 99.06841
    AAcount(20, 9) = 99.1326
    AAcount(20, 10) = 5
    AAcount(20, 11) = -5.5
    
    
'Add Organisms to the Organism Combobox
EnzymeCombo.AddItem ("Trypsin [KR | P]")
EnzymeCombo.AddItem ("Trypsin [KR | -] ")
EnzymeCombo.AddItem ("Chymotrypsin [YWFM | P]")
EnzymeCombo.AddItem ("CNBr [M | -]")
EnzymeCombo.AddItem ("Elastase [ALIV | P]")
EnzymeCombo.AddItem ("Lys-C [K | P]")
EnzymeCombo.AddItem ("No Cleavage")

EnzymeCombo.Text = "Trypsin [KR | P]"
    
    
'Add Organisms to the Organism Combobox
OrganismCombo.AddItem ("None")
OrganismCombo.AddItem ("yeast")
OrganismCombo.AddItem ("worm")
OrganismCombo.AddItem ("human")
OrganismCombo.AddItem ("mouse")
OrganismCombo.AddItem ("rat")
OrganismCombo.AddItem ("fly")
OrganismCombo.AddItem ("E. coli")

OrganismCombo.Text = "None"

'Add Choices to the ProductIon ListBox
SelectProductIonList.AddItem ("Precursor + y1")
SelectProductIonList.AddItem ("Precursor + y2")
SelectProductIonList.AddItem ("Precursor + y3")
SelectProductIonList.AddItem ("Precursor + b1")
SelectProductIonList.AddItem ("Precursor + b2")
SelectProductIonList.AddItem ("Precursor + b3")
SelectProductIonList.AddItem ("y-ions N-term to Pro")
SelectProductIonList.AddItem ("b-ions N-term to Pro")
SelectProductIonList.AddItem ("All y-ions")
SelectProductIonList.AddItem ("All b-ions")

SelectProductIonList.Selected(6) = True
SelectProductIonList.Selected(1) = True
SelectProductIonList.Selected(0) = True



End Sub


Private Sub DigestCmd_Click()
    
Dim i As Long
Dim ii As Long
Dim iii As Integer
Dim AA As String
Dim StartPep As Long
Dim EndPep As Long
Dim PeptideLen As Long
Dim Peptide As String
Dim ArrayElementTemp As Long
Dim Sequence As String
Dim NumTryPep As Long
Dim RaggedEnds As Boolean
Dim PassPeptide As Boolean
Dim StartAA As Long
Dim Library As String
Dim inlib As Boolean
Dim CleavageSpecificity() As String
Dim ExceptProline As Boolean
Dim Cleave As Boolean



Screen.MousePointer = vbHourglass

Sequence = UCase(Strip_String(ProteinRTB.Text))
ProteinRTB.SelStart = 0
ProteinRTB.SelLength = Len(ProteinRTB.Text)
ProteinRTB.SelColor = vbBlack
ProteinRTB.SelFontName = "courier"
ProteinRTB.SelFontSize = 10

'Sequence = UCase(Strip_String(Proteintxt.Text))

   
    If OrganismCombo.List(OrganismCombo.ListIndex) = "worm" Then
        Library = App.Path & "\worm.lib.stats"
    ElseIf OrganismCombo.List(OrganismCombo.ListIndex) = "yeast" Then
        Library = App.Path & "\yeast.lib.stats"
    ElseIf OrganismCombo.List(OrganismCombo.ListIndex) = "human" Then
        Library = App.Path & "\human.lib.stats"
    ElseIf OrganismCombo.List(OrganismCombo.ListIndex) = "fly" Then
        Library = App.Path & "\fly.lib.stats"
    ElseIf OrganismCombo.List(OrganismCombo.ListIndex) = "mouse" Then
        Library = App.Path & "\mouse.lib.stats"
    ElseIf OrganismCombo.List(OrganismCombo.ListIndex) = "rat" Then
        Library = App.Path & "\rat.lib.stats"
    ElseIf OrganismCombo.List(OrganismCombo.ListIndex) = "E. coli" Then
        Library = App.Path & "\ecoli.lib.stats"
    End If
    
    
    If EnzymeCombo.ListIndex = -1 Or EnzymeCombo.ListIndex = 0 Then  'If trypsin [KR | P]
        ReDim CleavageSpecificity(0 To 1)
        ExceptProline = True
        CleavageSpecificity(0) = "K"
        CleavageSpecificity(1) = "R"
    
    ElseIf EnzymeCombo.ListIndex = 1 Then       ' if trypsin [KR | -]
        ReDim CleavageSpecificity(0 To 1)
        ExceptProline = False
        CleavageSpecificity(0) = "K"
        CleavageSpecificity(1) = "R"
    
    ElseIf EnzymeCombo.ListIndex = 2 Then       ' if chymotrypsin [YWFM | P]
        ReDim CleavageSpecificity(0 To 3)
        ExceptProline = True
        CleavageSpecificity(0) = "Y"
        CleavageSpecificity(1) = "W"
        CleavageSpecificity(2) = "F"
        CleavageSpecificity(3) = "M"
        
    ElseIf EnzymeCombo.ListIndex = 3 Then       ' CNBr [M | -]
        ReDim CleavageSpecificity(0 To 0)
        ExceptProline = False
        CleavageSpecificity(0) = "M"
    
    ElseIf EnzymeCombo.ListIndex = 4 Then       ' Elastase [ALIV | P]
        ReDim CleavageSpecificity(0 To 3)
        ExceptProline = True
        CleavageSpecificity(0) = "A"
        CleavageSpecificity(1) = "L"
        CleavageSpecificity(2) = "I"
        CleavageSpecificity(3) = "V"
    
    ElseIf EnzymeCombo.ListIndex = 5 Then       'Lys-C [K | P]
        ReDim CleavageSpecificity(0 To 0)
        ExceptProline = True
        CleavageSpecificity(0) = "K"
        
        
    End If
    
    
    
    

ProteinRTB.Text = Sequence
'Proteintxt.Text = Sequence

StartAA = ExcludeFirstAAtxt.Text + 1

If StartAA < 1 Or StartAA > Len(Sequence) Then StartAA = Len(Sequence)

PeptideList.Clear

    
    StartPep = StartAA
    NumTryPep = 0
    RaggedEnds = False
    PassPeptide = True
    
    If EnzymeCombo.List(EnzymeCombo.ListIndex) = "No Cleavage" Then
        PeptideList.AddItem (Sequence)
    
    Else
    
       If Sequence <> "" Then
            For i = StartAA To Len(Sequence)
        
                AA = Mid(Sequence, i, 1)
                'Identify Tryptic Peptides
                Cleave = False
                                
                For iii = 0 To UBound(CleavageSpecificity)
                    If UCase(AA) = CleavageSpecificity(iii) Then
                        If ExceptProline = True Then
                            If Mid(Sequence, i + 1, 1) <> "P" Then Cleave = True
                        Else
                            Cleave = True
                        End If
                    End If
                    
                Next
                
                If Cleave = True Or i = Len(Sequence) Then
                    
                        PeptideLen = i - StartPep + 1
                        Peptide = Mid(Sequence, StartPep, PeptideLen)
                
                        PassPeptide = CheckPeptideFeatures(Peptide)
                        If RaggedEndsChk.Value = 1 Then
                            EndPep = StartPep + PeptideLen - 1
                            RaggedEnds = CheckRaggedEnds(Sequence, StartPep, EndPep)
                        End If
                
                        If RaggedEnds = False And PassPeptide = True Then
                            PeptideList.AddItem (Peptide)
                                                
                            Call AddColor(StartPep - 1, PeptideLen)
                            NumTryPep = NumTryPep + 1
                        End If
                
                        StartPep = i + 1
                    
            
                End If
     
            Next
        End If
    End If
    
        If OrganismCombo.List(OrganismCombo.ListIndex) <> "None" Then
            If Dir$(Library) <> "" Then
                PeptideList.Visible = False
                For ii = 0 To PeptideList.ListCount - 1
                    inlib = CheckLibrary(PeptideList.List(ii))
                    If inlib = True Then PeptideList.Selected(ii) = True
                Next ii
                PeptideList.Visible = True
            End If
        End If
    
    
    PeptideLbl.Caption = "Peptide Sequences: " & NumTryPep
    Screen.MousePointer = vbArrow
End Sub

Function CheckRaggedEnds(Sequence As String, StartPep As Long, EndPep As Long) As Boolean
Dim RaggedEnds As Boolean

RaggedEnds = False

' Check for possible ragged ends
    If StartPep > 1 Then
        If Mid(Sequence, EndPep + 1, 1) = "R" Then
            RaggedEnds = True
        ElseIf Mid(Sequence, EndPep + 1, 1) = "K" Then
            RaggedEnds = True
        ElseIf Mid(Sequence, StartPep - 2, 1) = "K" Then
            RaggedEnds = True
        ElseIf Mid(Sequence, StartPep - 2, 1) = "R" Then
            RaggedEnds = True
        Else
            RaggedEnds = False
        End If
    End If
                    
                    
CheckRaggedEnds = RaggedEnds

End Function

Function CheckPeptideFeatures(Peptide As String) As Boolean
Dim AA As String
Dim i As Long
Dim PeptidePass As Boolean
Dim RPRet As Double
Dim test As Boolean

PeptidePass = True
test = False
RPRet = 0

        
    'Check for Cys
    If CysContPepChk.Value = 1 Then
        If InStr(Peptide, "C") > 0 Then PeptidePass = False
    End If
    
    'Check for Met
    If MetContPepChk.Value = 1 Then
        If InStr(Peptide, "M") > 0 Then PeptidePass = False
    End If
    
    'Check for His
    If HisContPepChk.Value = 1 Then
        If InStr(Peptide, "H") > 0 Then PeptidePass = False
    End If
    
    'Check for glycosylation motif
    If GlycosContPepChk.Value = 1 Then
        test = Peptide Like "*N?[ST]*"
        If test = True Then PeptidePass = False
    End If
    
    'Check for internal R-P or K-P
    If RPnKPContPepChk.Value = 1 Then
        test = Peptide Like "*[KR]P*"
        If test = True Then PeptidePass = False
    End If
           
    
    ' Check the length
    If Len(Peptide) < Int(MinLenTxt.Text) Or Len(Peptide) > Int(MaxLenTxt.Text) Then PeptidePass = False
                    

        
    CheckPeptideFeatures = PeptidePass
                    
                


End Function


Public Sub AddColor(Start As Long, Length As Long)
    
    ProteinRTB.SelStart = Start
    ProteinRTB.SelLength = Length
    ProteinRTB.SelColor = vbRed
End Sub

Function Format_sequence(Sequence) As String
Dim temp_sequence As String
Dim i As Long
Dim segment As Integer
Dim count As Long

segment = 10
 
    count = 0
    For i = 1 To Len(Sequence)
        If count = segment Then
            temp_sequence = temp_sequence + " " + Mid(Sequence, i, 1)
            count = 1
        Else
            temp_sequence = temp_sequence + Mid(Sequence, i, 1)
            count = count + 1
        End If
        
        
    Next
    

    Format_sequence = temp_sequence
End Function


Public Sub PeptideMass(Peptide As String, MonoMZ() As Double, AvgMZ() As Double, Optional iontype As Integer = 0)
' iontype defines whether the mass should be a precursor ion (0), y-ion (1), or b-ion (2)
Dim A As Long, i As Long, AA As String
Dim MonoTemp As Double, AvgTemp As Double
Dim CtermMono As Double, NtermMono As Double
Dim CtermAvg As Double, NtermAvg As Double

CtermMono = 17.00274
CtermAvg = 17.00734

NtermMono = 1.007825
NtermAvg = 1.00794

ReDim MonoMZ(0 To 4) As Double
ReDim AvgMZ(0 To 4) As Double
    
    'Clear AAcount
    For A = 1 To 20
        AAcount(A, 2) = 0
    Next
    
    'Calculate amino acid composition
    For i = 1 To Len(Peptide)
        AA = Mid(Peptide, i, 1)
        For A = 1 To 20
            If AAcount(A, 1) = UCase(AA) Then
                AAcount(A, 2) = AAcount(A, 2) + 1
                Exit For
            End If
        Next
    Next
    
        MonoTemp = 0
        AvgTemp = 0
        For A = 1 To 20
            MonoTemp = MonoTemp + AAcount(A, 2) * AAcount(A, 8)
            AvgTemp = AvgTemp + AAcount(A, 2) * AAcount(A, 9)
        Next
        
        If iontype = 0 Then
            MonoTemp = MonoTemp + CtermMono + NtermMono
            AvgTemp = AvgTemp + CtermAvg + NtermAvg
        ElseIf iontype = 1 Then
            MonoTemp = MonoTemp + 17.00274 + 1.007825
            AvgTemp = AvgTemp + 17.00734 + 1.00794
        ElseIf iontype = 2 Then
            MonoTemp = MonoTemp
            AvgTemp = AvgTemp
        End If
            
        MonoMZ(0) = MonoTemp
        AvgMZ(0) = AvgTemp
        
        MonoMZ(1) = MonoTemp + 1.007825
        AvgMZ(1) = AvgTemp + 1.00794
        
        MonoMZ(2) = (MonoTemp + 1.007828 * 2) / 2
        AvgMZ(2) = (AvgTemp + 1.00794 * 2) / 2
        
        MonoMZ(3) = (MonoTemp + 1.007828 * 3) / 3
        AvgMZ(3) = (AvgTemp + 1.00794 * 3) / 3
        
        MonoMZ(4) = (MonoTemp + 1.007828 * 4) / 4
        AvgMZ(4) = (AvgTemp + 1.00794 * 4) / 4

End Sub

Function Strip_String(PassData As Variant) As Variant
Dim tempstring As String

'Removes and carrage returns, line feeds, and spaces surrounding a string
tempstring = Replace(PassData, vbCr, "")
tempstring = Replace(tempstring, vbLf, "")
tempstring = Replace(tempstring, " ", "")
Strip_String = Trim(tempstring)



End Function

Private Sub OrganismCombo_Click()
Dim Library As String
   
    Screen.MousePointer = vbHourglass
   
    If OrganismCombo.List(OrganismCombo.ListIndex) = "worm" Then
        Library = App.Path & "\worm.lib.stats"
    ElseIf OrganismCombo.List(OrganismCombo.ListIndex) = "yeast" Then
        Library = App.Path & "\yeast.lib.stats"
    ElseIf OrganismCombo.List(OrganismCombo.ListIndex) = "human" Then
        Library = App.Path & "\human.lib.stats"
    ElseIf OrganismCombo.List(OrganismCombo.ListIndex) = "fly" Then
        Library = App.Path & "\fly.lib.stats"
    ElseIf OrganismCombo.List(OrganismCombo.ListIndex) = "mouse" Then
        Library = App.Path & "\mouse.lib.stats"
    ElseIf OrganismCombo.List(OrganismCombo.ListIndex) = "rat" Then
        Library = App.Path & "\rat.lib.stats"
    ElseIf OrganismCombo.List(OrganismCombo.ListIndex) = "E. coli" Then
        Library = App.Path & "\ecoli.lib.stats"
    End If
    
    If OrganismCombo.List(OrganismCombo.ListIndex) <> "None" Then
        If Dir$(Library) <> "" Then
            LibraryStatsFile() = ImportDelimitedFile(Library, FileError, vbTab)
        Else
            MsgBox Library & " does not exist", 64, "Notice"
        End If
    End If
    
    Screen.MousePointer = vbArrow
End Sub

Private Sub PeptideList_Click()
Dim Peptide As String
Dim MonoMZ() As Double
Dim AvgMZ() As Double
Dim ProductIons() As Variant
Dim PepLen As Long
Dim RPRet As Double
Dim inlib As Boolean
Dim Library As String

    Screen.MousePointer = vbHourglass
    Peptide = PeptideList.List(PeptideList.ListIndex)
    PepLen = Len(Peptide)
    inlib = False
    
    If OrganismCombo.List(OrganismCombo.ListIndex) = "worm" Then
        Library = App.Path & "\worm.lib.stats"
    ElseIf OrganismCombo.List(OrganismCombo.ListIndex) = "yeast" Then
        Library = App.Path & "\yeast.lib.stats"
    ElseIf OrganismCombo.List(OrganismCombo.ListIndex) = "human" Then
        Library = App.Path & "\human.lib.stats"
    ElseIf OrganismCombo.List(OrganismCombo.ListIndex) = "fly" Then
        Library = App.Path & "\fly.lib.stats"
    ElseIf OrganismCombo.List(OrganismCombo.ListIndex) = "mouse" Then
        Library = App.Path & "\mouse.lib.stats"
    ElseIf OrganismCombo.List(OrganismCombo.ListIndex) = "rat" Then
        Library = App.Path & "\rat.lib.stats"
    ElseIf OrganismCombo.List(OrganismCombo.ListIndex) = "E. coli" Then
        Library = App.Path & "\ecoli.lib.stats"
    End If
    
    ReDim ProductIons(1 To PepLen, 0 To 4)
    
    Call PeptideMass(Peptide, MonoMZ(), AvgMZ(), 0)
    RPRet = CalcC18Ret(Peptide)
    
    If OrganismCombo.List(OrganismCombo.ListIndex) <> "None" Then
        If Dir$(Library) <> "" Then
            inlib = CheckLibrary(Peptide)
        End If
    End If
    
    Call PrintPrecursorTxtBox(Peptide, MonoMZ(), AvgMZ(), RPRet, inlib)
    
    Call CalcProductIons(Peptide, ProductIons())
    Call PrintProductTxtBox(Peptide, ProductIons())
    
    
    Screen.MousePointer = vbArrow
End Sub

Private Sub AddAllSelectedCmd_Click()
Dim Peptide As String
Dim MonoMZ() As Double
Dim AvgMZ() As Double
Dim Precursor As Double
Dim Product As Double
Dim ProductIons() As Variant
Dim i As Long, ii As Long, iii As Long
Dim RTStart As Double
Dim RTEnd As Double
Dim RPRet As Double
Dim CE As Double
Dim NumAboveY As Integer
Dim NumAboveB As Integer
Dim ProlineCount As Integer
Dim ProlinePeak(100) As Double
Dim y1 As Double
Dim y2 As Double
Dim y3 As Double
Dim b1 As Double
Dim b2 As Double
Dim b3 As Double


For i = 0 To PeptideList.ListCount - 1
    If PeptideList.Selected(i) = True Then
        Peptide = PeptideList.List(i)
        ReDim ProductIons(1 To Len(Peptide), 0 To 4)
        Call PeptideMass(Peptide, MonoMZ(), AvgMZ(), 0)
        Call CalcProductIons(Peptide, ProductIons())
        
        CE = Round(AvgMZ(2) * Val(CESlpTxt.Text) + Val(CEIntTxt.Text), 1)
        RPRet = CalcC18Ret(Peptide)
        RTStart = (RPRet - Val(RTIntTxt.Text)) / Val(RTSlpTxt.Text) - (Val(RTWidthTxt.Text) / 2)
        RTEnd = (RPRet - Val(RTIntTxt.Text)) / Val(RTSlpTxt.Text) + (Val(RTWidthTxt.Text) / 2)
   
        
        NumAboveY = 0
        NumAboveB = 0
        ProlineCount = 0
        
        If MonoPrecursorChk.Value = 1 Then
            Precursor = Round(MonoMZ(2), 2)
        Else
            Precursor = Round(AvgMZ(2), 2)
        End If
        
        For ii = 1 To UBound(ProductIons)
            If SelectProductIonList.Selected(8) = True Then   ' If all y-ions selected
                If MonoProductChk.Value = 1 Then
                    Product = Round(ProductIons(ii, 1), 2)
                Else
                    Product = Round(ProductIons(ii, 2), 2)
                End If
                
                If ii < UBound(ProductIons) Then
                    SRMOutputList.AddItem (Precursor & ", " & Product & ", " & CE & ", , " & Peptide & ", Protein")
                End If
            Else
                ' Get the individual y-ions
                If MonoProductChk.Value = 1 Then
                    If ProductIons(ii, 1) > Precursor Then
                        NumAboveY = NumAboveY + 1
                        If NumAboveY = 1 Then y1 = Round(ProductIons(ii, 1), 2)
                        If NumAboveY = 2 Then y2 = Round(ProductIons(ii, 1), 2)
                        If NumAboveY = 3 Then y3 = Round(ProductIons(ii, 1), 2)
                    End If
                    
                    ' Get Y-ions N-term to proline
                    If SelectProductIonList.Selected(6) = True And NumAboveY < 1 Then
                        If Mid(Peptide, Len(Peptide) - ii + 1, 1) = "P" Then
                            ProlinePeak(ProlineCount) = Round(ProductIons(ii, 1), 2)
                            ProlineCount = ProlineCount + 1
                        End If
                    ElseIf SelectProductIonList.Selected(6) = True And NumAboveY > 3 Then
                        If Mid(Peptide, Len(Peptide) - ii + 1, 1) = "P" Then
                            ProlinePeak(ProlineCount) = Round(ProductIons(ii, 1), 2)
                            ProlineCount = ProlineCount + 1
                        End If
                    End If
                    
                    
                Else
                    If ProductIons(ii, 2) > Precursor Then
                        NumAboveY = NumAboveY + 1
                        If NumAboveY = 1 Then y1 = Round(ProductIons(ii, 2), 2)
                        If NumAboveY = 2 Then y2 = Round(ProductIons(ii, 2), 2)
                        If NumAboveY = 3 Then y3 = Round(ProductIons(ii, 2), 2)
                    End If
                    
                    ' Get Y-ions N-term to proline
                    If SelectProductIonList.Selected(6) = True And NumAboveY < 1 Then
                        If Mid(Peptide, Len(Peptide) - ii + 1, 1) = "P" Then
                            ProlinePeak(ProlineCount) = Round(ProductIons(ii, 2), 2)
                            ProlineCount = ProlineCount + 1
                        End If
                    ElseIf SelectProductIonList.Selected(6) = True And NumAboveY > 3 Then
                        If Mid(Peptide, Len(Peptide) - ii + 1, 1) = "P" Then
                            ProlinePeak(ProlineCount) = Round(ProductIons(ii, 2), 2)
                            ProlineCount = ProlineCount + 1
                        End If
                    End If
                End If
                
                
            End If
            
            If SelectProductIonList.Selected(9) = True Then   ' If all b-ions selected
                If MonoProductChk.Value = 1 Then
                    Product = Round(ProductIons(ii, 3), 2)
                Else
                    Product = Round(ProductIons(ii, 4), 2)
                End If
                
                'SRMOutputList.AddItem (Precursor & ", " & Product & ", " & CE & ", , " & Peptide & ", Protein")
                SRMOutputList.AddItem (Precursor & ", " & Product & ", " & CE & ", " & RTStart & ", " & RTEnd & ", " & Peptide)
                
            Else      ' Get the individual b-ions
                If MonoProductChk.Value = 1 Then
                    If ProductIons(ii, 3) > Precursor Then
                        NumAboveB = NumAboveB + 1
                        If NumAboveB = 1 Then b1 = Round(ProductIons(ii, 3), 2)
                        If NumAboveB = 2 Then b2 = Round(ProductIons(ii, 3), 2)
                        If NumAboveB = 3 Then b3 = Round(ProductIons(ii, 3), 2)
                    End If
                    
                    ' Get B-ions N-term to proline
                    If SelectProductIonList.Selected(7) = True And NumAboveB < 1 Then
                        If Mid(Peptide, ii, 1) = "P" Then
                            ProlinePeak(ProlineCount) = Round(ProductIons(ii - 1, 3), 2)
                            ProlineCount = ProlineCount + 1
                        End If
                    ElseIf SelectProductIonList.Selected(7) = True And NumAboveB > 3 Then
                        If Mid(Peptide, ii, 1) = "P" Then
                            ProlinePeak(ProlineCount) = Round(ProductIons(ii - 1, 3), 2)
                            ProlineCount = ProlineCount + 1
                        End If
                        
                    End If
                Else
                    If ProductIons(ii, 4) > Precursor Then
                        NumAboveB = NumAboveB + 1
                        If NumAboveB = 1 Then b1 = Round(ProductIons(ii, 4), 2)
                        If NumAboveB = 2 Then b2 = Round(ProductIons(ii, 4), 2)
                        If NumAboveB = 3 Then b3 = Round(ProductIons(ii, 4), 2)
                    End If
                    
                    ' Get B-ions N-term to proline
                    If SelectProductIonList.Selected(7) = True And NumAboveB < 1 Then
                        If Mid(Peptide, ii, 1) = "P" Then
                            ProlinePeak(ProlineCount) = Round(ProductIons(ii - 1, 4), 2)
                            ProlineCount = ProlineCount + 1
                        End If
                    ElseIf SelectProductIonList.Selected(7) = True And NumAboveB > 3 Then
                        If Mid(Peptide, ii, 1) = "P" Then
                            ProlinePeak(ProlineCount) = Round(ProductIons(ii - 1, 4), 2)
                            ProlineCount = ProlineCount + 1
                                                
                        End If
                    End If
                End If
                    
            End If
            
        Next
        
        ' Output the individual y-ions
        If SelectProductIonList.Selected(8) = False Then
            If SelectProductIonList.Selected(0) = True Then
                'SRMOutputList.AddItem (Round(Precursor, 2) & ", " & y1 & ", " & CE & ", , " & Peptide & ", Protein")
                SRMOutputList.AddItem (Round(Precursor, 2) & ", " & y1 & ", " & CE & ", " & RTStart & ", " & RTEnd & ", " & Peptide)
            
            End If
            If SelectProductIonList.Selected(1) = True Then
                'SRMOutputList.AddItem (Round(Precursor, 2) & ", " & y2 & ", " & CE & ", , " & Peptide & ", Protein")
                SRMOutputList.AddItem (Round(Precursor, 2) & ", " & y2 & ", " & CE & ", " & RTStart & ", " & RTEnd & ", " & Peptide)
           
            End If
            If SelectProductIonList.Selected(2) = True Then
                'SRMOutputList.AddItem (Round(Precursor, 2) & ", " & y3 & ", " & CE & ", , " & Peptide & ", Protein")
                SRMOutputList.AddItem (Round(Precursor, 2) & ", " & y3 & ", " & CE & ", " & RTStart & ", " & RTEnd & ", " & Peptide)
           
            End If
            
        End If
        
        ' Output the individual b-ions
        If SelectProductIonList.Selected(9) = False Then
            If SelectProductIonList.Selected(3) = True Then
                'SRMOutputList.AddItem (Round(Precursor, 2) & ", " & b1 & ", " & CE & ", , " & Peptide & ", Protein")
                SRMOutputList.AddItem (Round(Precursor, 2) & ", " & b1 & ", " & CE & ", " & RTStart & ", " & RTEnd & ", " & Peptide)
           
            End If
            If SelectProductIonList.Selected(4) = True Then
                'SRMOutputList.AddItem (Round(Precursor, 2) & ", " & b2 & ", " & CE & ", , " & Peptide & ", Protein")
                SRMOutputList.AddItem (Round(Precursor, 2) & ", " & b2 & ", " & CE & ", " & RTStart & ", " & RTEnd & ", " & Peptide)
           
            End If
            If SelectProductIonList.Selected(5) = True Then
                'SRMOutputList.AddItem (Round(Precursor, 2) & ", " & b3 & ", " & CE & ", , " & Peptide & ", Protein")
                SRMOutputList.AddItem (Round(Precursor, 2) & ", " & b3 & ", " & CE & ", " & RTStart & ", " & RTEnd & ", " & Peptide)
           
            End If
            
        End If
        
        'Output Proline Peaks
        If ProlineCount > 0 Then
            For ii = 0 To ProlineCount - 1
                'SRMOutputList.AddItem (Precursor & ", " & ProlinePeak(ii) & ", " & CE & ", , " & Peptide & ", Protein")
                SRMOutputList.AddItem (Precursor & ", " & ProlinePeak(ii) & ", " & CE & ", " & RTStart & ", " & RTEnd & ", " & Peptide)
           
            Next
        End If
    End If


Next


 

    


    
End Sub



Function CheckLibrary(Peptide As String) As Boolean
Dim FileError As Boolean
Dim i As Long
Dim begin As Long
Dim PeptideFound As Boolean

    'LibraryStatsFile() = ImportDelimitedFile(Library, FileError, vbTab)
    
    PeptideFound = False
    
    begin = 0
    Do Until Strip_String(LibraryStatsFile(begin)(0)) = "id"
        begin = begin + 1
    Loop
    
    For i = begin To UBound(LibraryStatsFile)
        If UCase(Peptide) = UCase(Strip_String(LibraryStatsFile(i)(6))) Then
            PeptideFound = True
            i = UBound(LibraryStatsFile)
        End If
        
    Next

    CheckLibrary = PeptideFound
End Function




Private Sub RemoveAllCmd_Click()
    SRMOutputList.Clear
    
    
End Sub

Private Sub RemoveCmd_Click()
Dim Index As Integer

Index = SRMOutputList.ListIndex

If Index > -1 Then
    SRMOutputList.RemoveItem (Index)
    If SRMOutputList.List(Index) > "" Then
        SRMOutputList.Selected(Index) = True
    End If
End If
    
End Sub

Function CalcC18Ret(Peptide As String) As Double
Dim i As Long
Dim SumRc As Double
Dim R1Nt As Double
Dim R2Nt As Double
Dim R3Nt As Double
Dim Kl As Double
Dim C18Ret As Double
Dim N As Long

    N = Len(Peptide)
    C18Ret = 0
    
    If N < 10 Then
        Kl = 1 - 0.027 * (10 - N)
    ElseIf N > 20 Then
        Kl = 1 - 0.014 * (N - 20)
    Else
        Kl = 1
    End If
    
    
    SumRc = 0
    For i = 1 To 20
        SumRc = SumRc + AAcount(i, 2) * AAcount(i, 10)
        
        If Mid(Peptide, 1, 1) = AAcount(i, 1) Then R1Nt = AAcount(i, 11)
        If Mid(Peptide, 2, 1) = AAcount(i, 1) Then R2Nt = AAcount(i, 11)
        If Mid(Peptide, 3, 1) = AAcount(i, 1) Then R3Nt = AAcount(i, 11)
        
    Next
    
    C18Ret = Kl * (SumRc + (0.42 * R1Nt) + (0.22 * R2Nt) + (0.05 * R3Nt))
    
    If C18Ret >= 38 Then
        C18Ret = C18Ret - 0.3 * (C18Ret - 38)
    End If
    
    CalcC18Ret = C18Ret

End Function



Public Sub PrintProductTxtBox(Peptide As String, ProductIons() As Variant)
Dim i As Long
Dim temptxt As String
    
    temptxt = "Seq #  B-Ion" & vbTab & "  #   Y-ion" & vbCrLf
    
    If MonoProductChk.Value = 1 Then
        For i = 1 To Len(Peptide)
            temptxt = temptxt & Mid(Peptide, i, 1) & "   " & i & ": " & Round(ProductIons(i, 3), 2) & vbTab & "  " & _
                Len(Peptide) - i + 1 & ": " & Round(ProductIons(Len(Peptide) - i + 1, 1), 2) & vbCrLf
        Next
    Else
        For i = 1 To Len(Peptide)
            temptxt = temptxt & Mid(Peptide, i, 1) & "   " & i & ": " & Round(ProductIons(i, 4), 2) & vbTab & "  " & _
                Len(Peptide) - i + 1 & ": " & Round(ProductIons(Len(Peptide) - i + 1, 2), 2) & vbCrLf
        Next
    End If
    
    ProductTxt.Text = temptxt

End Sub


Public Sub CalcProductIons(Peptide As String, ProductIons() As Variant)
Dim MonoMZ() As Double
Dim AvgMZ() As Double
Dim Yion As String
Dim Bion As String
Dim i As Long
    
    For i = 1 To Len(Peptide)
            ProductIons(i, 0) = i
            
        'Calculate Y ion m/z
            Yion = Right(Peptide, i)
            Call PeptideMass(Yion, MonoMZ(), AvgMZ(), 1)
            ProductIons(i, 1) = MonoMZ(1)
            ProductIons(i, 2) = AvgMZ(1)
            
        'Calculate B ion m/z
            Bion = Left(Peptide, i)
            Call PeptideMass(Bion, MonoMZ(), AvgMZ(), 2)
            ProductIons(i, 3) = MonoMZ(1)
            ProductIons(i, 4) = AvgMZ(1)
    Next
            

End Sub


Public Sub PrintPrecursorTxtBox(Peptide As String, MonoMZ() As Double, AvgMZ() As Double, C18Ret As Double, inlib As Boolean)
Dim temptext As String


    temptext = "Peptide: " & Peptide & vbCrLf
    temptext = temptext & vbCrLf
    temptext = temptext & "Monoisotopic Mass" & vbCrLf
    temptext = temptext & vbTab & "M = " & Round(MonoMZ(0), 3) & vbCrLf
    temptext = temptext & vbTab & "M+H = " & Round(MonoMZ(1), 3) & vbCrLf
    temptext = temptext & vbTab & "(M+2H)/2 = " & Round(MonoMZ(2), 3) & vbCrLf
    temptext = temptext & vbTab & "(M+3H)/3 = " & Round(MonoMZ(3), 3) & vbCrLf
    temptext = temptext & vbCrLf
    temptext = temptext & " Average Mass" & vbCrLf
    temptext = temptext & vbTab & "M = " & Round(AvgMZ(0), 3) & vbCrLf
    temptext = temptext & vbTab & "M+H = " & Round(AvgMZ(1), 3) & vbCrLf
    temptext = temptext & vbTab & "(M+2H)/2 = " & Round(AvgMZ(2), 3) & vbCrLf
    temptext = temptext & vbTab & "(M+3H)/3 = " & Round(AvgMZ(3), 3) & vbCrLf & vbCrLf
    temptext = temptext & "Collision Energy (+1) = " & Round(AvgMZ(1) * Val(CESlpTxt.Text) + Val(CEIntTxt.Text), 1) & vbCrLf
    temptext = temptext & "Collision Energy (+2) = " & Round(AvgMZ(2) * Val(CESlpTxt.Text) + Val(CEIntTxt.Text), 1) & vbCrLf
    temptext = temptext & "Collision Energy (+3) = " & Round(AvgMZ(3) * Val(CESlpTxt.Text) + Val(CEIntTxt.Text), 1) & vbCrLf
    temptext = temptext & vbCrLf & vbCrLf
    temptext = temptext & "Hydrophobicity Retention Factor = " & Round(C18Ret, 2)
    
    If inlib = True Then
        temptext = temptext & vbCrLf & vbCrLf & "Peptide found in library"
    End If
    PrecursorTxt.Text = temptext
    
End Sub






Private Sub ResetCmd_Click()
    ProteinRTB.Text = ""
    'Proteintxt.Text
    PeptideList.Clear
    PrecursorTxt.Text = ""
    ProductTxt.Text = ""
    
    
End Sub

Private Sub SaveCSVCmd_Click()
Dim filename As String
Dim fnum As Integer
Dim i As Integer


' This subroutine uses the Windows' common dialog
' box to allow the user to save data to a specific file.

dbsaveresults.CancelError = True


On Error GoTo dbCancel

If SRMOutputList.ListCount < 1 Then GoTo dbCancel

    dbsaveresults.Flags = &H4
    dbsaveresults.Filter = "CSV Files (*.csv) | *.csv | All Files (*.*) | *.*"
    dbsaveresults.DefaultExt = "csv"
    
    dbsaveresults.ShowSave
    
    filename = dbsaveresults.filename
    
    fnum = FreeFile()
    Open filename For Output As #fnum
    
        For i = 0 To SRMOutputList.ListCount
            Print #fnum, SRMOutputList.List(i)
        Next i
        
       
    Close #fnum
  

dbCancel:
    'The user pressed cancel so ignore file selection


End Sub

Function ReadFileContents(filename As String, FileError As Boolean) As String
Dim fnum As Integer, isOpen As Boolean
Dim i As Long

On Error GoTo Error_Handler

fnum = FreeFile()
Open filename For Input As #fnum
isOpen = True


ReadFileContents = Input(LOF(fnum), fnum)

Error_Handler:
If isOpen = True Then Close #fnum
If Err Then MsgBox "File " & filename & " Not Found", 64, "Notice"
If Err Then FileError = True
End Function

Function ImportDelimitedFile(filename As String, FileError As Boolean, Optional delimiter As String = vbTab) As Variant()
Dim lines() As String
Dim i As Long
Dim temp As Integer

FileError = False
lines() = Split(ReadFileContents(filename, FileError), vbLf)

If FileError = True Then
    GoTo done
Else
    If Dir$(filename) <> "" Then
    
    For i = 0 To UBound(lines)
        If Len(lines(i)) = 0 Then lines(i) = vbNullChar
    Next
    
    lines() = Filter(lines(), vbNullChar, False)
    ReDim Values(0 To UBound(lines)) As Variant
    For i = 0 To UBound(lines)
        Values(i) = Split(lines(i), delimiter)
        
    Next
    
    ImportDelimitedFile = Values()
    End If
End If
done:
End Function


