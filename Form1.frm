VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Audio Waveform Generator"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4590
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHundreds 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   600
      TabIndex        =   18
      Top             =   1320
      Width           =   250
   End
   Begin VB.VScrollBar scrHundreds 
      Height          =   520
      Left            =   600
      Max             =   0
      Min             =   9
      TabIndex        =   20
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox txtThousands 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   360
      TabIndex        =   19
      Top             =   1320
      Width           =   250
   End
   Begin VB.TextBox txtTens 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      TabIndex        =   17
      Top             =   1320
      Width           =   250
   End
   Begin VB.TextBox txtUnits 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   16
      Top             =   1320
      Width           =   250
   End
   Begin VB.TextBox txtDecimal 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   1320
      Width           =   265
   End
   Begin VB.VScrollBar scrThousands 
      Height          =   520
      Left            =   360
      Max             =   0
      Min             =   9
      TabIndex        =   14
      Top             =   1200
      Width           =   255
   End
   Begin VB.VScrollBar VScroll4 
      Height          =   30
      Left            =   600
      TabIndex        =   13
      Top             =   1200
      Width           =   255
   End
   Begin VB.VScrollBar scrTens 
      Height          =   520
      Left            =   840
      Max             =   0
      Min             =   9
      TabIndex        =   12
      Top             =   1200
      Width           =   255
   End
   Begin VB.VScrollBar scrUnits 
      Height          =   520
      Left            =   1080
      Max             =   0
      Min             =   9
      TabIndex        =   11
      Top             =   1200
      Width           =   255
   End
   Begin VB.VScrollBar scrDecimal 
      Height          =   520
      Left            =   1320
      Max             =   0
      Min             =   9
      TabIndex        =   10
      Top             =   1200
      Width           =   265
   End
   Begin VB.Frame Frame3 
      Caption         =   "Function"
      Height          =   1095
      Left            =   2040
      TabIndex        =   7
      Top             =   840
      Width           =   2415
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   495
         Left            =   1200
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Waveform"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.OptionButton optNoise 
         Caption         =   "Noise"
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optSaw 
         Caption         =   "Saw"
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optTriangle 
         Caption         =   "Triangle"
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optSquare 
         Caption         =   "Square"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optSine 
         Caption         =   "Sine"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frequency"
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A simple audio waveform generator using DirectX8
'Text boxes and vertical scrollers are used to
'create an array of up/down buttons for frequency
'selection.

'Lower frequency limit is set to 20 Hz.

'Upper frequency limit is 9999.9 Hz which is the
'practical limit when using a 44,100 Hz samplerate.

Option Explicit

'Create the DirectSound8 Object

Dim dx As New DirectX8
Dim ds As DirectSound8
Dim dsBuffer As DirectSoundSecondaryBuffer8

'Declare the variables that will be used

Dim frequency As Double
Dim increment As Double
Dim fileName As String
Dim fileSize As Double
Dim sample As Long
Dim period As Double
Dim state As Integer
Dim bufferptr As Long
Dim inputValue As Double

Const Pi = 3.141592654
Const sampleRate = 44100
Const amplitude = 127

'Initialize Waveform and Frequency Selectors - Note this
'can also be done in the Properties Editor. I did it this
'way to allow seeing all the initial settings at a glance.

Private Sub Form_Load()
    
    'Initialize the Waveform Selection option
    
    optSine.Value = True
    
    'Initialize the Freqency Selector buttons
    
    scrDecimal.Max = 0
    scrDecimal.Min = 9
    scrDecimal.Value = 0
    txtDecimal.Enabled = False
    txtDecimal.Text = "." & Str(scrDecimal.Value / 10)
    scrUnits.Max = 0
    scrUnits.Min = 9
    scrUnits.Value = 0
    txtUnits.Enabled = False
    txtUnits.Text = Str(scrUnits.Value)
    scrTens.Max = 0
    scrTens.Min = 9
    scrTens.Value = 4
    txtTens.Enabled = False
    txtTens.Text = Str(scrTens.Value)
    scrHundreds.Max = 0
    scrHundreds.Min = 9
    scrHundreds.Value = 4
    txtHundreds.Enabled = False
    txtHundreds.Text = Str(scrHundreds.Value)
    scrThousands.Max = 0
    scrThousands.Min = 9
    scrThousands.Value = 0
    txtThousands.Enabled = False
    txtThousands.Text = ""
    
    Me.Show
    On Local Error Resume Next
    Set ds = dx.DirectSoundCreate("")
    If Err.Number <> 0 Then
        MsgBox "Unable to start DirectSound"
        End
    End If
    ds.SetCooperativeLevel Me.hWnd, DSSCL_PRIORITY
    
    'Set the default startup frequency
    
    conFreq
    
End Sub

'Cleanup on program exit

Private Sub cmdExit_Click()

    Cleanup
    Unload Me
    
End Sub

'Dispose of the DirectSound Object and its buffer

Private Sub Cleanup()

    If Not (dsBuffer Is Nothing) Then dsBuffer.Stop
    Set dsBuffer = Nothing
    Set ds = Nothing
    Set dx = Nothing
    
End Sub

'Select Decimal and insert decimal point

Private Sub scrDecimal_Change()

    txtDecimal.Text = "." & Str(scrDecimal.Value)
    
    conFreq
    
End Sub

'Select Units Value

Private Sub scrUnits_Change()

    txtUnits.Text = Str(scrUnits.Value)
    
    conFreq
    
End Sub

'Select Tens Value

Private Sub scrTens_Change()

    txtTens.Text = Str(scrTens.Value)
    
    conFreq
    
End Sub

'Select Hundreds and mask leading zeroes

Private Sub scrHundreds_Change()

    txtHundreds.Text = Str(scrHundreds.Value)
    If scrHundreds.Value = 0 And scrThousands.Value = 0 Then
        txtHundreds.Text = ""
    End If

    conFreq
    
End Sub

'Select Thousands and mask leading zero

Private Sub scrThousands_Change()

    txtThousands.Text = Str(scrThousands.Value)
    If scrThousands.Value = 0 Then
        txtThousands.Text = ""
    End If
    
    conFreq
    
End Sub

'Concatenate the frequency selector settings

Private Sub conFreq()


    frequency = (scrThousands.Value * 1000) + (scrHundreds.Value * 100) + (scrTens.Value * 10) + scrUnits.Value + (scrDecimal.Value / 10)
        If frequency < 20 Then
            MsgBox "Frequency cannot be lower than 20 Hz."
            frequency = 20
            txtTens.Text = "2"
            txtUnits.Text = "0"
            txtDecimal.Text = ".0"
        End If
        
    
End Sub

Private Sub cmdGenerate_Click()

    If optSine.Value = True Then
        sineWave
    End If
    
    If optSquare.Value = True Then
        squareWave
    End If
    
    If optTriangle.Value = True Then
        triangleWave
    End If
    
    If optSaw.Value = True Then
        sawWave
    End If
    
    If optNoise = True Then
        noise
    End If
    
End Sub

Private Sub sineWave()

    makeFile                                            'Create the file and write header
    
    bufferptr = 45                                      'Offset to beginning of waveform
        increment = Pi / (sampleRate / frequency)
        For inputValue = 0 To (2 * Pi) Step increment   'Step around the circle
            sample = Int(amplitude * Sin(inputValue))   'Calculate the sample value
            Put #1, bufferptr, sample                   'Write sample value to file
            bufferptr = bufferptr + 1                   'Increment buffer pointer
        Next inputValue                                 'Loop to the next sample
        
    closeFile                                           'Fill in the rest of the file data
                                                        'and close the file.
    
End Sub

Private Sub squareWave()

    makeFile
    
    bufferptr = 45
    period = (sampleRate / frequency)
    state = 1
        If state = 1 Then                   'Positive half cycle
            For inputValue = 0 To period
                sample = amplitude * state
                Put #1, bufferptr, sample
                bufferptr = bufferptr + 1
            Next inputValue
        End If
        
     state = -1
        If state = -1 Then                  'Negative half cycle
            For inputValue = 0 To period
                sample = amplitude * state
                Put #1, bufferptr, sample
                bufferptr = bufferptr + 1
            Next inputValue
        End If
        
    closeFile
    
End Sub

Private Sub sawWave()

    makeFile
    
    bufferptr = 45
        period = sampleRate / (frequency / 2)
        For inputValue = 0 To period
            sample = Int(2 * amplitude * (inputValue / period))
            Put #1, bufferptr, sample
            bufferptr = bufferptr + 1
        Next inputValue
        
    closeFile
    
End Sub

Private Sub triangleWave()

    makeFile
    
    state = 0
    bufferptr = 45
        period = sampleRate / frequency
        If state = 0 Then
        For inputValue = 0 To period / 2    'Generate Positive Slope
            sample = Int(2 * amplitude * (inputValue / period))
            Put #1, bufferptr, sample
            bufferptr = bufferptr + 1
        Next inputValue
        
        state = 1
        End If
        If state = 1 Then
        For inputValue = 0 To period        'Generate Negative Slope
            sample = Int((amplitude - 2 * amplitude) - 2 * amplitude * (inputValue - period) / period)
            Put #1, bufferptr, sample
            bufferptr = bufferptr + 1
        Next inputValue
        
        state = 2
        End If
        If state = 2 Then
        For inputValue = 0 To period / 2    'Positive Slope to finish cycle
            sample = Int(amplitude + (2 * amplitude * (inputValue / period)))
            Put #1, bufferptr, sample
            bufferptr = bufferptr + 1
        Next inputValue
        End If
        
    closeFile
    
End Sub

Private Sub noise()

    Randomize                               'Seed random # generator
    
    makeFile
    
    bufferptr = 45
    period = sampleRate
        For inputValue = 0 To period        'Create 44,100 random samples
            sample = Rnd(amplitude) * 254
            Put #1, bufferptr, sample
            bufferptr = bufferptr + 1
        Next inputValue
        
    closeFile
    
End Sub

'Create the .wav file and write header data

Private Sub makeFile()

    fileName = App.Path & "\temp.wav"
    
    Kill fileName                   'REM this line if file does not exist
    
    Open fileName For Binary Access Write As #1
        Put #1, 1, "RIFF"           '"RIFF" header
        Put #1, 5, CInt(0)          'Filesize - 8, will write later
        Put #1, 9, "WAVEfmt "       '"WAVEfmt " header - not space after fmt
        Put #1, 17, CLng(16)        'Lenth of format data
        Put #1, 21, CInt(1)         'Wave type PCM
        Put #1, 23, CInt(1)         '1 channel
        Put #1, 25, CLng(44100)     '44.1 kHz SampleRate
        Put #1, 29, CLng(88200)     '(SampleRate * BitsPerSample * Channels) / 8
        Put #1, 33, CInt(2)         '(BitsPerSample * Channels) / 8
        Put #1, 35, CInt(16)        'BitsPerSample
        Put #1, 37, "data"          '"data" Chunkheader
        Put #1, 41, CInt(0)         'Filesize - 44, will write later

End Sub

'Get the file length, write it into the header and close the file.

Private Sub closeFile()

    fileSize = LOF(1)
    Put #1, 5, CLng(fileSize - 8)
    Put #1, 41, CLng(fileSize - 44)
    Close #1
    
    Play
    
End Sub

'Define the DirectSound8 buffer, create it and set the play mode

Private Sub Play()

    Dim bufferDesc As DSBUFFERDESC
    bufferDesc.lFlags = DSBCAPS_STATIC Or DSBCAPS_STICKYFOCUS
    fileName = App.Path & "\temp.wav"
    Set dsBuffer = ds.CreateSoundBufferFromFile(fileName, bufferDesc)
    dsBuffer.Play DSBPLAY_LOOPING
    
End Sub

'Stop playing and clear the DirectSound8 buffer

Private Sub cmdStop_Click()

    dsBuffer.Stop
    Set dsBuffer = Nothing
    
End Sub
