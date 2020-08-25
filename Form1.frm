VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Program Loundry Creative Wangi"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13575
   BeginProperty Font 
      Name            =   "Cagar"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF00FF&
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   9645
   ScaleWidth      =   13575
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":25316
      Height          =   1575
      Left            =   960
      TabIndex        =   35
      Top             =   8040
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   2778
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "NoFaktur"
         Caption         =   "NoFaktur"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "NamaPelanggan"
         Caption         =   "NamaPelanggan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Alamat"
         Caption         =   "Alamat"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "TanggalTransaksi"
         Caption         =   "TanggalTransaksi"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "TanggalPengambilan"
         Caption         =   "TanggalPengambilan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "JenisCucian"
         Caption         =   "JenisCucian"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Harga"
         Caption         =   "Harga"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Berat"
         Caption         =   "Berat"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Total"
         Caption         =   "Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Diskon"
         Caption         =   "Diskon"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "TotalBayar"
         Caption         =   "TotalBayar"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   2085.166
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1200
      Top             =   8160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=uts_loundry"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "uts_loundry"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from loundry_wangi"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cagar"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Lihat Data"
      BeginProperty Font 
         Name            =   "Cagar"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   32
      Top             =   7560
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   4320
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   4320
      TabIndex        =   14
      Text            =   "Text3"
      Top             =   2640
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H80000007&
      Height          =   360
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4440
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   5520
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "Cagar"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9120
      MaskColor       =   &H000000FF&
      Picture         =   "Form1.frx":2532B
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Cagar"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9120
      Picture         =   "Form1.frx":2745D
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Batal"
      BeginProperty Font 
         Name            =   "Cagar"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9120
      Picture         =   "Form1.frx":294DA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF00FF&
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "Cagar"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9120
      Picture         =   "Form1.frx":2EA24
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF0000&
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Cagar"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9120
      Picture         =   "Form1.frx":30A2E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   1695
   End
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   7320
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   -2147483641
      Format          =   "Rp ###,###,##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   6720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   -2147483641
      Format          =   "Rp ###,###,##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   6120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   -2147483641
      Format          =   "Rp ###,###,##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   4920
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cagar"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "Rp ###,###,##"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   3840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   147783683
      CurrentDate     =   43055
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   3240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   147783683
      CurrentDate     =   43055
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   1080
      Picture         =   "Form1.frx":32AA9
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Wangi"
      BeginProperty Font 
         Name            =   "Cagar"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   6480
      TabIndex        =   34
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Creative"
      BeginProperty Font 
         Name            =   "Cagar"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4800
      TabIndex        =   33
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Pelanggan"
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
      Left            =   1200
      TabIndex        =   31
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   30
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Transaksi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   29
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Pengambilan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   28
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Cucian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   27
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Berat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   26
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga"
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
      Left            =   1200
      TabIndex        =   25
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Diskon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   24
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "/Kg/Helai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   23
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Kg"
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
      Left            =   7320
      TabIndex        =   22
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "5%"
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
      Left            =   7320
      TabIndex        =   21
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   20
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Bayar"
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
      Left            =   1200
      TabIndex        =   19
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No Faktur "
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
      Left            =   1200
      TabIndex        =   18
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   5
      FillColor       =   &H000000FF&
      Height          =   6735
      Left            =   1080
      Top             =   1200
      Width           =   7575
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   5
      Height          =   6255
      Left            =   8880
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Loundry"
      BeginProperty Font 
         Name            =   "Cagar"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3240
      TabIndex        =   17
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Program Loundry By YudaBayu Prabowo"
      BeginProperty Font 
         Name            =   "Cagar"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Combo1 = "Kain Baju/Celana" Then
MaskEdBox1 = "5000"
End If
If Combo1 = "Bad Cover" Then
MaskEdBox1 = "20000"
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Command1_Click()
sql = "insert into loundry_wangi (NoFaktur,NamaPelanggan,Alamat,TanggalTransaksi,TanggalPengambilan,JenisCucian,Harga,Berat,Total,Diskon,TotalBayar) " & _
                       "Values('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Format(DTPicker1, "dd-MMM-yyyy") & "','" & Format(DTPicker2, "dd-MMM-yyyy") & "','" & Combo1 & "','" & MaskEdBox1 & "','" & Text4 & "','" & MaskEdBox2 & "','" & MaskEdBox3 & "','" & MaskEdBox4 & "')"
con.Execute (sql)

Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
MaskEdBox1 = ""
MaskEdBox2 = ""
MaskEdBox3 = ""
MaskEdBox4 = ""
DTPicker1 = Date
DTPicker2 = Date
Adodc1.Refresh
Text1.SetFocus
MsgBox "Data telah di Simpan!", vbInformation + vbOKOnly = vbIgnore


End Sub
Private Sub Command2_Click()
sql = "insert into loundry_wangi (NoFaktur,NamaPelanggan,Alamat,TanggalTransaksi,TanggalPengambilan,JenisCucian,Harga,Berat,Total,Diskon,TotalBayar) " & _
                       "Values('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Format(DTPicker1, "dd-MMM-yyyy") & "','" & Format(DTPicker2, "dd-MMM-yyyy") & "','" & Combo1 & "','" & MaskEdBox1 & "','" & Text4 & "','" & MaskEdBox2 & "','" & MaskEdBox3 & "','" & MaskEdBox4 & "')"
con.Execute (sql)

Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
DTPicker1 = Date
DTPicker2 = Date
MaskEdBox1 = ""
MaskEdBox2 = ""
MaskEdBox3 = ""
MaskEdBox4 = ""
Adodc1.Refresh
Text1.SetFocus
Adodc1.Refresh
Text1.SetFocus
End Sub

Private Sub Command3_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
DTPicker1 = Date
DTPicker2 = Date
MaskEdBox1 = ""
MaskEdBox2 = ""
MaskEdBox3 = ""
MaskEdBox4 = ""
Adodc1.Refresh
Text1.SetFocus
End Sub

Private Sub Command4_Click()
Dim X As String
X = MsgBox(("Anda Yakin data ingin di hapus?"), vbYesNo + vbCritical)
If X = vbYes Then
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveFirst
DataGrid1.ReBind
DataGrid1.Refresh
MsgBox "Data  telah di Hapus!", vbInformation + vbOKOnly = vbIgnore
End If
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Form_Activate()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
DTPicker1 = Date
DTPicker2 = Date
MaskEdBox1 = ""
MaskEdBox2 = ""
MaskEdBox3 = ""
MaskEdBox4 = ""
Adodc1.Refresh
Text1.SetFocus
End Sub

Private Sub Form_Load()
If con.State = adStateClosed Then
connect

Combo1.AddItem "Kain Baju/Celana"
Combo1.AddItem "Bad Cover"
End If
End Sub

Private Sub MaskEdBox2_Change()
If Val(Text4) > 5 Then
MaskEdBox3 = (MaskEdBox2) * (5 / 100)
Else
MaskEdBox3 = "Anda Tidak Mendapat Diskon"
End If
End Sub

Private Sub MaskEdBox3_Change()
MaskEdBox4 = Val(MaskEdBox2) - Val(MaskEdBox3)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
sql = "select * from loundry_wangi where  NoFaktur='" & Text1 & "'"
Set rs = con.Execute(sql)
    If Not rs.EOF Then
    MsgBox ("data sudah ada")
    Text2 = rs!NamaPelanggan
    Text3 = rs!Alamat
    DTPicker1 = rs!TanggalTransaksi
    DTPicker2 = rs!TanggalPengambilan
    Combo1 = rs!JenisCucian
    MaskEdBox1 = rs!Harga
    Text4 = rs!Berat
    MaskEdBox2 = rs!Total
    MaskEdBox3 = rs!Diskon
    MaskEdBox4 = rs!TotalBayar
    Else
    MsgBox ("data tidak ada lanjutkan input data baru")
    Text2.SetFocus
    End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
Text3.SetFocus
End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo1.SetFocus
End If
End Sub

Private Sub Text4_Change()
MaskEdBox2 = Val(MaskEdBox1) * Val(Text4)
End Sub

Private Sub Text6_Change()

End Sub

Private Sub Text7_Change()

End Sub

Private Sub Text8_Change()

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub
