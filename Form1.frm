VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel Background Accept"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   2040
   End
   Begin VB.TextBox Text1 
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   2640
      Width           =   10455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Begin Accepting"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Begin Accepting, then use a browser to browse to http://localhost:5555/"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label LblStatus 
      Caption         =   "Idle, Not Listening"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   1200
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Status:"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim acceptSocket As ChilkatSocket
Attribute acceptSocket.VB_VarHelpID = -1
Dim connectedSocket As ChilkatSocket
Attribute connectedSocket.VB_VarHelpID = -1
Dim acceptConnTask As ChilkatTask

Private acceptInProgress As Boolean
Private connectionAvailable As Boolean

Private numStatusDots As Integer


' Begin accepting connections on port 5555
Private Sub Command1_Click()

    If (acceptInProgress = True) Then
        Exit Sub
    End If
    connectionIsAvailable = False
    acceptErrorText = ""
    
    ' If acceptSocket was previously existing, make sure it is closed so that we can re-bind and listen.
    If (Not (acceptSocket Is Nothing)) Then
        acceptSocket.Close 10
        Set acceptSocket = Nothing
    End If
    
    LblStatus.Caption = "Idle, Not Listening"
    
    Set acceptSocket = New ChilkatSocket
    
    maxBacklog = 25
    success = acceptSocket.BindAndListen(5555, maxBacklog)
    If (success <> 1) Then
        Text1.Text = acceptSocket.LastErrorText
        Exit Sub
    End If
    
    Set acceptConnTask = acceptSocket.AcceptNextConnectionAsync(0)
    acceptConnTask.Run
    
    acceptInProgress = True
    LblStatus.Caption = "Listening.."
    numStatusDots = 2
    
    ' Start a timer that will fire every 1/10th of a second to check
    ' for a received connection.
    Timer1.Interval = 100
    Timer1.Enabled = True
End Sub


' Cancel the task.
Private Sub Command2_Click()
    If (acceptInProgress = True) Then
        If (Not acceptConnTask Is Nothing) Then
            acceptConnTask.Cancel
            
            acceptInProgress = False
            Set acceptConnTask = Nothing
            
            Timer1.Enabled = False
            LblStatus.Caption = "Listening canceled."
        End If
    End If
End Sub


Private Sub readConnectedSocket()

    ' Create a ChilkatSocket object for the accepted connection,
    ' and load it with the socket connection..
    Set connectedSocket = New ChilkatSocket
    connectedSocket.LoadTaskResult acceptConnTask
    
    Dim startLine As String
    Dim header As String
    startLine = connectedSocket.ReceiveUntilMatch(vbCrLf)
    header = connectedSocket.ReceiveUntilMatch(vbCrLf & vbCrLf)
    Text1.Text = startLine & header
    
    ' Send our response..
    connectedSocket.SendString "HTTP/1.1 200 OK" & vbCrLf
    connectedSocket.SendString "Content-Type: text/html" & vbCrLf
    body = "<html><body>Hello from VB6!</body></html>"
    connectedSocket.SendString "Content-Length: " & Len(body) & vbCrLf & vbCrLf & body
    
    connectedSocket.Close (10)
    Set connectedSocket = Nothing
    
End Sub


' Called every 1/10th of a second when waiting for a connection.
Private Sub Timer1_Timer()
    If (acceptConnTask Is Nothing) Then
        ' Nothing to do...
        Timer1.Enabled = False
        acceptInProgress = False
    Else
        If (acceptConnTask.StatusInt > 4) Then
            ' The task has successfully finished, was canceled, or failed.
            If (acceptConnTask.StatusInt < 7) Then
                Text1.Text = "The AcceptNextConnection background task was canceled."
                
                ' Make sure we can re-bind and listen if needed...
                If (Not (acceptSocket Is Nothing)) Then
                    acceptSocket.Close 10
                    Set acceptSocket = Nothing
                End If
            Else
                ' The task completed, with success or failure..
                If (acceptConnTask.TaskSuccess = 1) Then
                    ' A connection was received.
                    LblStatus.Caption = "Connection Accepted!"
                    Text1.Text = acceptConnTask.ResultErrorText
                    
                    connectionAvailable = True
                    readConnectedSocket
                Else
                    ' There was some kind of failure..
                    Text1.Text = acceptConnTask.ResultErrorText
                End If
            End If
            Timer1.Enabled = False
            acceptInProgress = False
        Else
            ' Still waiting for the incoming connection..
            If (numStatusDots = 2) Then
                LblStatus.Caption = "Listening..."
                numStatusDots = 3
            Else
                LblStatus.Caption = "Listening.."
                numStatusDots = 2
            End If
               
        End If
    End If
    
End Sub

Private Sub Form_Load()
    ' Unlock Chilkat once for all Chilkat objects.
    Dim glob As New ChilkatGlobal
    success = glob.UnlockBundle("30-day trial")
    If (success <> 1) Then
        MsgBox "Failed to unlock"
        Text1.Text = glob.LastErrorText
    End If
    
    'glob.ThreadPoolLogPath = "C:/VB6_projects/Accept-Connection-Async/threadPoolLog.txt"
    connectionAvailable = False
    acceptInProgress = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If (Not (acceptSocket Is Nothing)) Then
        acceptSocket.Close 10
        Set acceptSocket = Nothing
    End If
        
    Dim glob As New ChilkatGlobal
    ' Make sure no background tasks remain.
    glob.FinalizeThreadPool
End Sub
