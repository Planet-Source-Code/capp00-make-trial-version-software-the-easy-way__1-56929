VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000FF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trial-Version"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "Label2"
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
      Left            =   480
      TabIndex        =   1
      Top             =   2520
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Label1.Caption = "This is Trial-Version software. To get full functionality, please register by contacting the software provider for a non-expiring version." & vbCrLf & "Thank You"
'=====================
'Call sub
'=====================
    CheckExpiration
End Sub


Public Sub CheckExpiration()
    Dim eXpdate As Date
    Dim preDate As Date
    Dim x As Integer
    eXpdate = "12/1/2004"    ' date you want to end the trial version
    preDate = "10/25/2004"   ' date you want to detect if system has rolled back time
    x = eXpdate - Date
    
'======================
'Date checking
'======================
    
    If Date >= eXpdate Then
        Label2.Caption = "This trial-version has expired" & vbCrLf & "please register to continue"
        Exit Sub

    ElseIf Date < preDate Then
        Label2.Caption = "The system clock has been rolled back" & vbCrLf & "The program will no longer function"
        Exit Sub

    ElseIf Date < eXpdate Then
        Label2.Caption = "Program expires on: " & eXpdate & vbCrLf & x & " day(s) remaining"
        Exit Sub
    End If

'======================
'enter code to go to next form or other events
'i.e.   frmMain.show
'i.e.   unload me
'======================

End Sub
