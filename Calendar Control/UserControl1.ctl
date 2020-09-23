VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.UserControl RCalendar 
   ClientHeight    =   3345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3465
   LockControls    =   -1  'True
   PropertyPages   =   "UserControl1.ctx":0000
   ScaleHeight     =   3345
   ScaleWidth      =   3465
   ToolboxBitmap   =   "UserControl1.ctx":0032
   Begin VB.Frame Frame1 
      BackColor       =   &H00AA820A&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   0
      TabIndex        =   3
      Top             =   690
      Width           =   3465
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   53
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   495
         TabIndex        =   52
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   990
         TabIndex        =   51
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1485
         TabIndex        =   50
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1980
         TabIndex        =   49
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2475
         TabIndex        =   48
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   2970
         TabIndex        =   47
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   46
         Top             =   510
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   495
         TabIndex        =   45
         Top             =   510
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   990
         TabIndex        =   44
         Top             =   510
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   1485
         TabIndex        =   43
         Top             =   510
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   1980
         TabIndex        =   42
         Top             =   510
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   2475
         TabIndex        =   41
         Top             =   510
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   2970
         TabIndex        =   40
         Top             =   510
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   0
         TabIndex        =   39
         Top             =   765
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   495
         TabIndex        =   38
         Top             =   765
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   990
         TabIndex        =   37
         Top             =   765
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   1485
         TabIndex        =   36
         Top             =   765
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   1980
         TabIndex        =   35
         Top             =   765
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   2475
         TabIndex        =   34
         Top             =   765
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   2970
         TabIndex        =   33
         Top             =   765
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   0
         TabIndex        =   32
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   495
         TabIndex        =   31
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   990
         TabIndex        =   30
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   1485
         TabIndex        =   29
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   1980
         TabIndex        =   28
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   2475
         TabIndex        =   27
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   2970
         TabIndex        =   26
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   0
         TabIndex        =   25
         Top             =   1275
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   495
         TabIndex        =   24
         Top             =   1275
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   990
         TabIndex        =   23
         Top             =   1275
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   1485
         TabIndex        =   22
         Top             =   1275
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   1980
         TabIndex        =   21
         Top             =   1275
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   33
         Left            =   2475
         TabIndex        =   20
         Top             =   1275
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   34
         Left            =   2970
         TabIndex        =   19
         Top             =   1275
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Sun"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Mon"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   495
         TabIndex        =   17
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Tue"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   990
         TabIndex        =   16
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Wed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1485
         TabIndex        =   15
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Thu"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1980
         TabIndex        =   14
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Fri"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2475
         TabIndex        =   13
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "Sat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   2970
         TabIndex        =   12
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   41
         Left            =   2970
         TabIndex        =   11
         Top             =   1530
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   40
         Left            =   2475
         TabIndex        =   10
         Top             =   1530
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   39
         Left            =   1980
         TabIndex        =   9
         Top             =   1530
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   38
         Left            =   1485
         TabIndex        =   8
         Top             =   1530
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   990
         TabIndex        =   7
         Top             =   1530
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   495
         TabIndex        =   6
         Top             =   1530
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   35
         Left            =   0
         TabIndex        =   5
         Top             =   1530
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000040&
         Height          =   495
         Left            =   555
         Shape           =   2  'Oval
         Top             =   1965
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Today :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   2085
         Width           =   720
      End
   End
   Begin VB.ComboBox cboyear 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1965
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   1245
   End
   Begin VB.ComboBox cbomonth 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "UserControl1.ctx":0344
      Left            =   0
      List            =   "UserControl1.ctx":036C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   1680
      Top             =   2040
   End
   Begin ComCtl2.UpDown UpDownMonth 
      Height          =   360
      Left            =   1455
      TabIndex        =   0
      Top             =   0
      Width           =   255
      _ExtentX        =   423
      _ExtentY        =   635
      _Version        =   327681
      Value           =   1
      Max             =   12
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown UpDownYear 
      Height          =   360
      Left            =   3210
      TabIndex        =   54
      Top             =   0
      Width           =   255
      _ExtentX        =   423
      _ExtentY        =   635
      _Version        =   327681
      Value           =   1
      Max             =   9999
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   55
      Top             =   360
      Width           =   3465
   End
   Begin VB.Menu mnuSetDate 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuSetSystemDate 
         Caption         =   "Set System Date"
      End
   End
End
Attribute VB_Name = "RCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ********************************************************
'         Custom Calendar Control by Raman Ver 1.0
'         Please don't forget to vote if u like it
'*********************************************************


'*********************************************************
' The PrintCalendar Algoritm generates Calendar of any month mainly
' by finding extra no. of days counted beyond complete no.
' of weeks contained in the no. of years upto the current
' month of the current year
' 0 extra days means the current month starts with Sunday
' 1 extra day means the current month starts with Monday
' and so on
'*********************************************************

Option Explicit

' This variable stores the index of the label whose
' caption is to be set as current day so that it remains
' with different formatting even if month is changed
Dim lblindex As Integer

' This variable stores the index of the label that has
' the caption as selected by user using mouse so that
' it remains highlighted ever if month is changed
Dim lblhighlightindex As Integer

' This variable stores the date selected by user
Dim selecteddate As Date

' This variable indicates whether any date has been selected
Dim flag As Boolean
' This variable stores the caption of label that has
' been right clicked by user
Dim captionrightclick As Integer

' This variable indicates whether after right clicking
' the date has been set
Dim dateset As Boolean
Dim selcolor As OLE_COLOR ' Stores BackGround Color of selected date
Dim tcolor As OLE_COLOR ' Stores ForeGround Color of today
Dim bcolor As OLE_COLOR ' Stores BackGround Color of Calendar
Dim dcolor As OLE_COLOR ' Stores ForeGround Color of days

' Current height and width of usercontrol
Dim curheight As Single, curwidth As Single

' This variable stores actual width of month combo box
' in order to minimise the round off error caused when
' width of combo box is multiplied by width factor
' when resizing the control
' Similar is the case with other such variables
Dim cbomonthactualwidth As Single
Dim cboyearactualwidth As Single
Dim frame1actualwidth As Single, frame1actualheight As Single
Dim lbl2actualwidth As Single, lbl2actualheight As Single
Dim lbl4actualwidth As Single, lbl4actualheight As Single
Dim lbl3actualfontsize As Single
' This array stores names of days corresponding to day numbers
Dim NameOfDay(6) As String * 3
Public Event Click()
Public Event DblClick()
Public Enum DayOfWeek
    Sun
    Mon
    Tue
    Wed
    Thu
    Fri
    Sat
End Enum
Private Sub cboyear_Click()
    Dim i
    For i = 0 To 41
        Label1(i).Caption = ""
    Next i
    PrintCalendar cboyear.ListIndex + 1, cbomonth.ListIndex + 1
    UpDownYear.Value = cboyear.ListIndex + 1
End Sub

Private Sub cbomonth_Click()
    Dim i
    For i = 0 To 41
        Label1(i).Caption = ""
    Next i
    PrintCalendar cboyear.ListIndex + 1, cbomonth.ListIndex + 1
    UpDownMonth.Value = cbomonth.ListIndex + 1
End Sub

Private Function isLeapYear(Year As Integer) As Integer
    If ((Year Mod 400) = 0) Then
        isLeapYear = 1
    ElseIf ((Year Mod 100) = 0) Then
        isLeapYear = 0
    ElseIf ((Year Mod 4) = 0) Then
        isLeapYear = 1
    Else
        isLeapYear = 0
    End If
End Function
Private Sub PrintCalendar(Yr As Integer, Mnth As Integer)
    Dim months As Variant
    ' This array contains no. of days beyond complete no.
    ' of weeks contained in every month of a year
    ' By default Feb. is taken as 28-day month
    months = Array(3, 0, 3, 2, 3, 2, 3, 3, 2, 3, 2, 3)

    Dim temp As Integer, signi_years As Integer, extra_days As Integer
    Dim i As Integer, no_of_days As Integer
    ' Following three steps calculate no. of years
    ' beyond the greatest mulitple of 400
    ' This is done in order to facilitate the algorithm
    ' as every 400 years end with a complete no. of weeks
    ' leaving no extra days
    temp = Yr \ 400 '(Remember:- '\' means integer division
    temp = temp * 400
    temp = Yr - temp
    ' Following step further helps simplify the algoritm
    ' by calculating no. of extra days in no. of years which is
    ' muliple of 100 among those remaining years as found
    ' by previous three steps
    ' Every 100 years contain 5 days beyond complete no.
    ' of weeks so the algorithm finds total no. of extra
    ' days by multiplying the no. by 5
    extra_days = (5 * (temp \ 100)) Mod 7
    ' Following two steps calculate the remaining no.
    ' of years(excluding current year)
    temp = temp - (temp \ 100) * 100
    signi_years = temp - 1
    ' Following two steps calculate no. of extra days in those
    ' remaining no. of years(excluding current year)
    ' As we know every 4th year is a leap year yielding 2
    ' extra days and every normal year yields 1 extra day
    temp = signi_years \ 4
    extra_days = (extra_days + (temp * 2 + signi_years - temp)) Mod 7
    ' Following steps calculate no. of extra days
    ' yielded upto the start of the current month(or
    ' upto the end of the previous month)
    ' The algorithm also takes care if current year is a
    ' leap year
    For i = 0 To Mnth - 2
        extra_days = (extra_days + months(i)) Mod 7
    Next i
    If (isLeapYear(Yr)) Then
            extra_days = extra_days + 1
    End If
    extra_days = (extra_days + 1) Mod 7
    ' Calculate no. of days in each month in order to
    ' generate the Calendar
    Select Case (Mnth)
        Case 1, 3, 5, 7, 8, 10, 12: no_of_days = 31
        Case 4, 6, 9, 11:  no_of_days = 30
        Case 2:
                If (isLeapYear(Yr)) Then
                    no_of_days = 29
                Else
                    no_of_days = 28
                End If
    End Select
    
    ' Adjust for first day to be displayed on calendar
    extra_days = (extra_days + (7 - FirstDay)) Mod 7
    For i = 1 To no_of_days
         Label1(i + extra_days - 1).Caption = i
    Next i
    Label4.Caption = cbomonth.Text & " " & cboyear.Text
    Label3.Caption = "Today : " & Format(Date, "mmm dd, yyyy")
    ' Remove any formatting which was previously applied
    Label1(lblindex).ForeColor = dcolor
    Label1(lblhighlightindex).BackColor = Me.BackGroundColor
    ' Find the index of the label whose caption would be
    ' chosen for the current day and apply formatting to
    ' it in order to highlight the current day
    lblindex = VBA.Day(Date) + extra_days - 1 ' VBA.Day avoids confusing with control's Day property
    ' Change the forecolor only if the current month and year
    ' are equal to month and year corresponding to system date
    If VBA.Month(Date) = (cbomonth.ListIndex + 1) And VBA.Year(Date) = cboyear.Text Then
        Label1(lblindex).ForeColor = tcolor
    End If
    ' If any date was selected by user keep it highlighted
    ' this can be done by checking if current month and year
    ' are equal to those selected by user previously
    If flag = True And VBA.Month(selecteddate) = (cbomonth.ListIndex + 1) And VBA.Year(selecteddate) = cboyear.Text Then
        Label1_MouseUp lblhighlightindex, vbLeftButton, 0, 0, 0
    End If
End Sub

Private Sub Frame1_Click()
    RaiseEvent Click
End Sub

Private Sub Frame1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Label1_Click(Index As Integer)
    RaiseEvent Click
End Sub

Private Sub Label1_DblClick(Index As Integer)
    RaiseEvent DblClick
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Label1(Index).Caption = "" Then Exit Sub
    Label1(Index).ToolTipText = cbomonth.Text & " " & Label1(Index).Caption & ", " & cboyear.Text
End Sub

Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Do not allow selecting blank labels
    If Label1(Index).Caption = "" Then Exit Sub
    
    If Button = vbRightButton Then
        captionrightclick = Label1(Index).Caption
        PopupMenu mnuSetDate
        If dateset Then
            Label1(lblindex).ForeColor = dcolor
            Label1(Index).ForeColor = tcolor
            Label3.Caption = "Today : " & Format(Date, "mmm dd, yyyy")
            lblindex = Index
            dateset = False
            
            ' Call this sub procedure recursively in order to
            ' highlight the caption after date has been set
            Label1(Index).BackColor = Me.BackGroundColor
            Label1_MouseUp Index, vbLeftButton, 0, 0, 0
        End If
    ElseIf Button = vbLeftButton Then
        If Label1(Index).BackColor = Me.BackGroundColor Then
            ' If selected label is not formatted then format it and
            ' remove formatting from any previously selected
            ' label
            Label1(lblhighlightindex).BackColor = Me.BackGroundColor
            Label1(Index).BackColor = selcolor
            selecteddate = CDate(Label1(Index).Caption & "/" & cbomonth.Text & "/" & cboyear.Text)
            ' Calculate index of newly selected label
            lblhighlightindex = Index
            flag = True
        Else
            Label1(Index).BackColor = Me.BackGroundColor
            flag = False
        End If
    End If
End Sub

Private Sub Label2_Click(Index As Integer)
    RaiseEvent Click
End Sub

Private Sub Label2_DblClick(Index As Integer)
    RaiseEvent DblClick
End Sub

Private Sub Label3_Click()
    ' Reset the Calendar
    cboyear.ListIndex = VBA.Year(Date) - 1
    cbomonth.ListIndex = VBA.Month(Date) - 1
    RaiseEvent Click
 End Sub

Private Sub Label3_DblClick()
    RaiseEvent DblClick
End Sub


Private Sub mnuSetSystemDate_Click()
    Date = DateSerial(cboyear.ListIndex + 1, cbomonth.ListIndex + 1, captionrightclick)
    dateset = True
End Sub

Private Sub Timer1_Timer()
    ' Timer continuously checks if system date has or has been changed
    ' (like when next day has occured or date changed manually)
    
    ' Perform necessary action only when date does not equal
    ' that displayed on label3(this is to avoid unnecessary
    ' printing of calendar)
    If Mid(Label3.Caption, 9) <> Format(Date, "mmm dd, yyyy") Then
        If cbomonth.ListIndex <> VBA.Month(Date) - 1 Or cboyear.ListIndex <> VBA.Year(Date) - 1 Then
            ' Click events for these comboboxes will automatically
            ' be generated when their listindices are changed
            ' so no need to explicitly call PrintCalendar function
            cbomonth.ListIndex = VBA.Month(Date) - 1
            cboyear.ListIndex = VBA.Year(Date) - 1
        Else
            Dim i
            For i = 0 To 41
                Label1(i).Caption = ""
            Next i
            PrintCalendar VBA.Year(Date), VBA.Month(Date)
        End If
    End If
End Sub

Private Sub updownmonth_Change()
    cbomonth.ListIndex = UpDownMonth.Value - 1
End Sub

Private Sub updownyear_Change()
    cboyear.ListIndex = UpDownYear.Value - 1
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    Dim i
    For i = 1 To 9999
        cboyear.AddItem i
    Next i
    tcolor = vbRed
    selcolor = RGB(255, 224, 192)
    bcolor = vbWhite
    dcolor = vbBlack
    cboyear.ListIndex = VBA.Year(Date) - 1
    cbomonth.ListIndex = VBA.Month(Date) - 1
    UpDownMonth.Value = cbomonth.ListIndex + 1
    UpDownYear.Value = cboyear.ListIndex + 1
    curheight = UserControl.ScaleHeight
    curwidth = UserControl.ScaleWidth
    cbomonthactualwidth = cbomonth.Width
    cboyearactualwidth = cboyear.Width
    frame1actualwidth = Frame1.Width
    frame1actualheight = Frame1.Height
    lbl2actualwidth = Label2(0).Width
    lbl2actualheight = Label2(0).Height
    lbl4actualwidth = Label4.Width
    lbl4actualheight = Label4.Height
    NameOfDay(0) = "Sun"
    NameOfDay(1) = "Mon"
    NameOfDay(2) = "Tue"
    NameOfDay(3) = "Wed"
    NameOfDay(4) = "Thu"
    NameOfDay(5) = "Fri"
    NameOfDay(6) = "Sat"
End Sub

Private Sub UserControl_Resize()
    Dim i As Integer
    With UserControl
        If .Height < 50 Then
            .Height = 50
            Exit Sub
        End If
    End With
    ResizeControls
End Sub

Public Property Get BackGroundColor() As OLE_COLOR
Attribute BackGroundColor.VB_Description = "Sets/Retrieves the Background Color"
Attribute BackGroundColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackGroundColor = bcolor
End Property

Public Property Let BackGroundColor(ByVal NewColor As OLE_COLOR)
    Dim i As Integer
    bcolor = NewColor
    For i = 0 To 41
        Label1(i).BackColor = bcolor
    Next i
    PropertyChanged "BackGroundColor"
End Property


Public Property Get MonthBackColor() As OLE_COLOR
Attribute MonthBackColor.VB_Description = "Sets/Retrieves the BackGround Color of Month Combo Box"
Attribute MonthBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    MonthBackColor = cbomonth.BackColor
End Property

Public Property Let MonthBackColor(ByVal NewColor As OLE_COLOR)
    cbomonth.BackColor = NewColor
    PropertyChanged "MonthBackColor"
End Property

Public Property Get YearBackColor() As OLE_COLOR
Attribute YearBackColor.VB_Description = "Sets/Retrieves the BackGround Color of Year Combo Box"
Attribute YearBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    YearBackColor = cboyear.BackColor
End Property

Public Property Let YearBackColor(ByVal NewColor As OLE_COLOR)
    cboyear.BackColor = NewColor
    PropertyChanged "YearBackColor"
End Property

Public Property Get SelectionColor() As OLE_COLOR
Attribute SelectionColor.VB_Description = "Sets/Retrieves the BackGround Color of selected date"
Attribute SelectionColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    SelectionColor = selcolor
End Property

Public Property Let SelectionColor(ByVal NewColor As OLE_COLOR)
    selcolor = NewColor
    
    ' If any date is selected then highlight it
    If flag = True And VBA.Month(selecteddate) = (cbomonth.ListIndex + 1) And VBA.Year(selecteddate) = cboyear.Text Then
        Label1(lblhighlightindex).BackColor = selcolor
    End If
    PropertyChanged "SelectionColor"
End Property

Public Property Get HeadingBackColor() As OLE_COLOR
Attribute HeadingBackColor.VB_Description = "Sets/Retrieves the BackGround Color of day headings"
Attribute HeadingBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HeadingBackColor = Label2(0).BackColor
End Property

Public Property Let HeadingBackColor(ByVal NewColor As OLE_COLOR)
    Dim i As Integer
    For i = 0 To 6
        Label2(i).BackColor = NewColor
    Next i
    PropertyChanged "HeadingBackColor"
End Property
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackGroundColor", BackGroundColor, vbWhite
    PropBag.WriteProperty "MonthBackColor", MonthBackColor, &HC0C0FF
    PropBag.WriteProperty "YearBackColor", YearBackColor, &HC0C0FF
    PropBag.WriteProperty "SelectionColor", SelectionColor, RGB(255, 224, 192)
    PropBag.WriteProperty "HeadingBackColor", HeadingBackColor, &HC0C000
    PropBag.WriteProperty "HeadingColor", HeadingColor, vbBlack
    PropBag.WriteProperty "TodayColor", TodayColor, vbRed
    PropBag.WriteProperty "DayColor", DayColor, vbBlack
    PropBag.WriteProperty "TitleColor", TitleColor, vbBlack
    PropBag.WriteProperty "DayFont", DayFont
    PropBag.WriteProperty "HeadingFont", HeadingFont
    PropBag.WriteProperty "TitleFont", TitleFont
    PropBag.WriteProperty "MonthYearFont", MonthYearFont
    PropBag.WriteProperty "GridLines", GridLines, True
    PropBag.WriteProperty "FirstDay", FirstDay, 0
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BackGroundColor = PropBag.ReadProperty("BackGroundColor", vbWhite)
    MonthBackColor = PropBag.ReadProperty("MonthBackColor", &HC0C0FF)
    YearBackColor = PropBag.ReadProperty("YearBackColor", &HC0C0FF)
    SelectionColor = PropBag.ReadProperty("SelectionColor", RGB(255, 224, 192))
    HeadingBackColor = PropBag.ReadProperty("HeadingBackColor", &HC0C000)
    HeadingColor = PropBag.ReadProperty("HeadingColor", vbBlack)
    TodayColor = PropBag.ReadProperty("TodayColor", vbRed)
    DayColor = PropBag.ReadProperty("DayColor", vbBlack)
    TitleColor = PropBag.ReadProperty("TitleColor", vbBlack)
    Set DayFont = PropBag.ReadProperty("DayFont")
    Set HeadingFont = PropBag.ReadProperty("HeadingFont")
    Set TitleFont = PropBag.ReadProperty("TitleFont")
    Set MonthYearFont = PropBag.ReadProperty("MonthYearFont")
    GridLines = PropBag.ReadProperty("GridLines", True)
    FirstDay = PropBag.ReadProperty("FirstDay", 0)
End Sub

Private Sub ResizeControls()
    Dim heightfactor As Single, widthfactor As Single
    
    ' Calculate change in height of usercontrol
    heightfactor = UserControl.ScaleHeight / curheight
    widthfactor = UserControl.ScaleWidth / curwidth
    
    ' Store the changed height and width
    curheight = UserControl.ScaleHeight
    curwidth = UserControl.ScaleWidth
    
    ' Calculate left, top, width and height of required
    ' controls according to width and height change factors
    cbomonthactualwidth = cbomonthactualwidth * widthfactor
    cbomonth.Width = cbomonthactualwidth
    UpDownMonth.Height = cbomonth.Height
    UpDownMonth.Move cbomonth.Left + cbomonth.Width
    cboyearactualwidth = cboyearactualwidth * widthfactor
    cboyear.Width = cboyearactualwidth
    cboyear.Move UserControl.ScaleWidth - UpDownYear.Width - cboyear.Width
    UpDownYear.Height = cboyear.Height
    UpDownYear.Move cboyear.Left + cboyear.Width
    Label4.Move Label4.Left, cbomonth.Top + cbomonth.Height, Label4.Width * widthfactor, Label4.Height * heightfactor
    frame1actualwidth = frame1actualwidth * widthfactor
    frame1actualheight = frame1actualheight * heightfactor
    Frame1.Move Frame1.Left, Label4.Top + Label4.Height, frame1actualwidth, frame1actualheight
    Dim i As Integer
    lbl2actualwidth = lbl2actualwidth * widthfactor
    lbl2actualheight = lbl2actualheight * heightfactor
    Label2(0).Move Label2(0).Left, Label2(0).Top, lbl2actualwidth, lbl2actualheight
    For i = 1 To 6
        Label2(i).Move Label2(0).Left + i * Label2(0).Width, Label2(0).Top, Label2(0).Width, Label2(0).Height
    Next i
    For i = 0 To 41
        Label1(i).Move Label2(0).Left + (i Mod 7) * Label2(0).Width, Label2(0).Top + ((i \ 7) + 1) * Label2(0).Height, Label2(0).Width, Label2(0).Height
    Next i
    
    ' Center Shape1 horizontally and vertically
    Shape1.Move (UserControl.ScaleWidth - Shape1.Width) / 2, (Frame1.Height - (Label1(35).Top + Label1(35).Height) - Shape1.Height) / 2 + Label1(35).Top + Label1(35).Height
       
    ' Center Label3 horizontally and vertically
    Label3.Move (Shape1.Width - Label3.Width) / 2 + Shape1.Left, (Shape1.Height - Label3.Height) / 2 + Shape1.Top
End Sub

Public Property Get TodayColor() As OLE_COLOR
Attribute TodayColor.VB_Description = "Sets/Retrieves the ForeGround Color of today"
Attribute TodayColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TodayColor = tcolor
End Property

Public Property Let TodayColor(ByVal NewColor As OLE_COLOR)
    tcolor = NewColor
    If VBA.Month(Date) = (cbomonth.ListIndex + 1) And VBA.Year(Date) = cboyear.Text Then
        Label1(lblindex).ForeColor = tcolor
    End If
    PropertyChanged "TodayColor"
End Property

Public Property Get Day() As Byte
Attribute Day.VB_Description = "Returns selected day of currently displayed month"
Attribute Day.VB_ProcData.VB_Invoke_Property = ";Text"
    ' If the currently displayed month and year is that
    ' which was selected by user then return the day
    ' else return 0
    ' since user can change month or year after
    ' selecting the date
    If flag = True And VBA.Month(selecteddate) = (cbomonth.ListIndex + 1) And VBA.Year(selecteddate) = cboyear.Text Then
        Day = Label1(lblhighlightindex).Caption
    Else
        Day = 0
    End If
End Property
Public Property Let Day(ByVal NewDay As Byte)
    MsgBox "Day Property is Read Only", vbInformation, "RCalendar"
End Property

Public Property Get Month() As Byte
Attribute Month.VB_Description = "Returns number corresponding to currently displayed month"
Attribute Month.VB_ProcData.VB_Invoke_Property = ";Text"
    Month = cbomonth.ListIndex + 1
End Property

Public Property Let Month(ByVal NewMonth As Byte)
    MsgBox "Month Property is Read Only", vbInformation, "RCalendar"
End Property
Public Property Get Year() As Integer
Attribute Year.VB_Description = "Returns currently displayed year"
Attribute Year.VB_ProcData.VB_Invoke_Property = ";Text"
    Year = Val(cboyear.Text)
End Property
Public Property Let Year(ByVal NewYear As Integer)
    MsgBox "Year Property is Read Only", vbInformation, "RCalendar"
End Property

Public Property Get HeadingColor() As OLE_COLOR
Attribute HeadingColor.VB_Description = "Sets/Retrieves the ForeGround Color of day headings"
Attribute HeadingColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HeadingColor = Label2(0).ForeColor
End Property

Public Property Let HeadingColor(ByVal NewColor As OLE_COLOR)
    Dim i As Integer
    For i = 0 To 6
        Label2(i).ForeColor = NewColor
    Next i
    PropertyChanged "HeadingColor"
End Property

Public Property Get DayColor() As OLE_COLOR
Attribute DayColor.VB_Description = "Sets/Retrieves the ForeGround Color of days"
Attribute DayColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DayColor = dcolor
End Property

Public Property Let DayColor(ByVal NewColor As OLE_COLOR)
    Dim i As Integer
    dcolor = NewColor
    For i = 0 To 41
        Label1(i).ForeColor = dcolor
    Next i
    If VBA.Month(Date) = Me.Month And VBA.Year(Date) = Me.Year Then
        Label1(lblindex).ForeColor = tcolor
    End If
    PropertyChanged "DayColor"
End Property

Public Property Get Value() As Date
Attribute Value.VB_Description = "Retrurns currently selected date or else returns current system date"
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Misc"
    If flag = False Then
        Value = Date
    Else
        Value = selecteddate
    End If
End Property

Public Property Let Value(ByVal NewDate As Date)
    MsgBox "Value Property is Read Only", vbInformation, "RCalendar"
End Property
Public Property Get DayFont() As StdFont
Attribute DayFont.VB_Description = "Sets/Retrieves Font of Days"
Attribute DayFont.VB_ProcData.VB_Invoke_Property = ";Font"
    Set DayFont = Label1(0).Font
End Property

Public Property Set DayFont(ByVal NewFont As StdFont)
    Dim i As Integer
    For i = 0 To 41
        Set Label1(i).Font = NewFont
    Next i
    PropertyChanged "DayFont"
End Property

Public Property Get HeadingFont() As StdFont
Attribute HeadingFont.VB_Description = "Sets/Retrieves Font of day headings"
Attribute HeadingFont.VB_ProcData.VB_Invoke_Property = ";Font"
    Set HeadingFont = Label2(0).Font
End Property

Public Property Set HeadingFont(ByVal NewFont As StdFont)
    Dim i As Integer
    For i = 0 To 6
        Set Label2(i).Font = NewFont
    Next i
    PropertyChanged "HeadingFont"
End Property

Public Property Get TitleFont() As StdFont
Attribute TitleFont.VB_Description = "Sets/Retrieves Font of Calendar Title"
Attribute TitleFont.VB_ProcData.VB_Invoke_Property = ";Font"
    Set TitleFont = Label4.Font
End Property

Public Property Set TitleFont(ByVal NewFont As StdFont)
    Set Label4.Font = NewFont
    PropertyChanged "TitleFont"
End Property

Public Property Get TitleColor() As OLE_COLOR
Attribute TitleColor.VB_Description = "Sets/Retrieves the ForeGround Color of Calendar Title"
Attribute TitleColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TitleColor = Label4.ForeColor
End Property

Public Property Let TitleColor(ByVal NewColor As OLE_COLOR)
    Label4.ForeColor = NewColor
    PropertyChanged "TitleColor"
End Property


Public Property Get MonthYearFont() As StdFont
Attribute MonthYearFont.VB_Description = "Sets/Retrieves the font of Month and Year Combo Boxes"
Attribute MonthYearFont.VB_ProcData.VB_Invoke_Property = ";Font"
    Set MonthYearFont = cbomonth.Font
End Property

Public Property Set MonthYearFont(ByVal NewFont As StdFont)
    Set cbomonth.Font = NewFont
    Set cboyear.Font = NewFont
    ResizeControls
    PropertyChanged "MonthYearFont"
End Property


Public Property Get GridLines() As Boolean
Attribute GridLines.VB_Description = "Determines whether Calendar should display Gridlines or not"
Attribute GridLines.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute GridLines.VB_MemberFlags = "200"
    GridLines = CBool(Label1(0).BorderStyle)
End Property

Public Property Let GridLines(ByVal NewValue As Boolean)
    Dim i As Integer
    For i = 0 To 41
        Label1(i).BorderStyle = Abs(NewValue)
    Next i
    PropertyChanged "GridLines"
End Property

Public Property Get FirstDay() As DayOfWeek
Attribute FirstDay.VB_ProcData.VB_Invoke_Property = ";Text"
    Select Case Label2(0).Caption
        Case "Sun": FirstDay = Sun
        Case "Mon": FirstDay = Mon
        Case "Tue": FirstDay = Tue
        Case "Wed": FirstDay = Wed
        Case "Thu": FirstDay = Thu
        Case "Fri": FirstDay = Fri
        Case "Sat": FirstDay = Sat
    End Select
End Property
Public Property Let FirstDay(ByVal NewDay As DayOfWeek)
    Dim i As Integer
    For i = 0 To 6
        Label2(i).Caption = NameOfDay((NewDay + i) Mod 7)
    Next i
    For i = 0 To 41
        Label1(i).Caption = ""
    Next i
    PrintCalendar Me.Year, Me.Month
    PropertyChanged "FirstDay"
End Property
