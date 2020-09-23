VERSION 5.00
Begin VB.Form fFont 
   Caption         =   "Form1"
   ClientHeight    =   2460
   ClientLeft      =   1395
   ClientTop       =   2565
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   4395
   Begin Project1.UserControl1 UserControl11 
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3625
      BeginProperty Font {7DDE0CCE-EB33-46FA-8F84-179117453CDD} 
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "To test the font object, return to design time and goto the custom properties of the usercontrol that's on this form."
      Height          =   855
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "fFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

