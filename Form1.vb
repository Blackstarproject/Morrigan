Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Input
Imports C1.Win.C1Gauge
Imports Microsoft.VisualBasic.Devices
Imports NAudio.CoreAudioApi
Imports NAudio.Dsp
Imports NAudio.Wave

Public Class Morrigan

    'Volume Control
    Private Const WM_APPCOMMAND As Integer = &H319
    Private Const APPCOMMAND_VOLUME_MUTE As Integer = &H80000
    Private Const APPCOMMAND_VOLUME_UP As Integer = &HA0000
    Private Const APPCOMMAND_VOLUME_DOWN As Integer = &H90000
    'End Control:>>>
    Private TargetDT As DateTime
    Private CountDownFrom As TimeSpan = TimeSpan.FromMinutes(60)


    Private pixelsPerBuffer As Integer
    Private Shared spec_height As Integer
    Private Shared spec_width As Integer = 600
    Private Shared unanalyzed_max_sec As Double 'maximum amount of unanalyzed audio to maintain in memory
    Private Shared buffers_captured As Integer = 0 'total number of audio buffers filled
    Private Shared buffers_remaining As Integer = 0 'number of buffers which have yet to be analyzed
    Private Shared unanalyzed_values As New List(Of Short)() ' audio data lives here waiting to be analyzed
    Private Shared spec_data As List(Of List(Of Double)) ' columns are time points, rows are frequency points

    Private Shared rand As New Random()

    Private Const WM_NCHITTEST As Integer = &H84
    Private Const HTCLIENT As Integer = &H1
    Private Const HTCAPTION As Integer = &H2

    ' sound card settings
    Private rate As Integer
    Private buffer_update_hz As Integer

    ' spectrogram and FFT settings
    Private fft_size As Integer

    Private borderForm As New Form

    'NAudio for recording
    Private waveSource As WaveIn = Nothing

    Private waveFile As WaveFileWriter = Nothing


    Private Const WM_SYSCOMMAND As Integer = 274
    Private Const SC_MAXIMIZE As Integer = 61488

    'Movable form without border
    Dim drag As Boolean

    Dim mousex As Integer

    Dim mousey As Integer




    Private _data As New List(Of String)()



    Private Sub Form1_Load(sender As Object,
                           e As EventArgs) Handles MyBase.Load


        Dim worker As New BackgroundWorker()
        AddHandler worker.DoWork, AddressOf WorkerOnDoWork

        AddHandler worker.RunWorkerCompleted, AddressOf WorkerOnRunWorkerCompleted
        worker.RunWorkerAsync()


        'Application docking using a panel. Parent to child process
        Dim proc As Process = Process.Start("C:\Users\rytho\source\repos\Morrigan\Morrigan\Mara.exe")
        proc.WaitForInputIdle()
        Thread.Sleep(1000)
        SetParent(proc.MainWindowHandle,
                  Panel1.Handle)
        Panel1.Dock = DockStyle.None
        Thread.Sleep(1000)
        SendMessage(proc.MainWindowHandle,
                    WM_SYSCOMMAND,
                    SC_MAXIMIZE,
                    0)
        'Rounded Form
        With Me
            .FormBorderStyle = FormBorderStyle.None
            .Region = New Region(RoundedRectangle(.ClientRectangle, 50))
        End With
        With borderForm
            .ShowInTaskbar = False
            .FormBorderStyle = FormBorderStyle.None
            .StartPosition = FormStartPosition.Manual
            .BackColor = Color.Black
            .Opacity = 0.25
            Dim r As Rectangle = Bounds
            r.Inflate(2, 2)
            .Bounds = r
            .Region = New Region(RoundedRectangle(.ClientRectangle, 50))
            r = New Rectangle(3, 3, Width - 4, Height - 4)
            .Region.Exclude(RoundedRectangle(r, 48))
            '.Show(Me)
        End With
        PictureBox2.Visible = False
        PictureBox11.Visible = False

    End Sub

    Private Function IsPrime(n As Long) As Boolean
        For i As Integer = 2 To Math.Sqrt(n)
            If n Mod i = 0 Then

                Return False
            End If
        Next

        For i As Integer = 0 To 100
            Dim weaver As New MemoryWeaver()
            Thread.Sleep(10)
            _data.Add(i)
            _data.Add(n)
            Call Task.Run(Sub()

                              Invoke(New Action(Sub()
                                                    Dim t As New Thread(Sub()

                                                                            Dim data = New Short(fft_size - 1) {}
                                                                            data = unanalyzed_values.GetRange(0, fft_size).ToArray()
                                                                            spec_data.RemoveAt(0)
                                                                            Dim fft_buffer = New Complex(fft_size - 1) {}
                                                                            For o = 0 To fft_size - 1
                                                                                fft_buffer(o).X = CSng(unanalyzed_values(i) _
                                                    * FastFourierTransform.HammingWindow(o, fft_size))
                                                                                fft_buffer(o).Y = 0
                                                                            Next
                                                                            FastFourierTransform.FFT(True, Math.Log(fft_size, 2.0), fft_buffer)
                                                                            For o = 0 To spec_data(spec_data.Count - 1).Count - 1
                                                                                ' should this be sqrt(X^2+Y^2)?
                                                                                Dim val As Double
                                                                                val = CDbl(fft_buffer(o).X) + CDbl(fft_buffer(i).Y)
                                                                                val = Math.Abs(val)

                                                                            Next
                                                                            Dim new_data As New List(Of Double)()
                                                                            new_data.Reverse()
                                                                            spec_data.Insert(spec_data.Count, new_data)
                                                                            unanalyzed_values.RemoveRange(
                                                    0,
                                                    fft_size / pixelsPerBuffer)
                                                                        End Sub)
                                                    t.Start()
                                                    t.Join()
                                                End Sub))
                          End Sub)


        Next

        Return True
    End Function
    Private Sub WorkerOnDoWork(sender As Object, e As DoWorkEventArgs)
        Timer1.Start()
        Timer2.Start()
        Timer3.Start()
    End Sub
    Private Sub WorkerOnRunWorkerCompleted(sender As Object, runWorkerCompletedEventArgs As RunWorkerCompletedEventArgs)
        'Work has finished. Launch new UI from here.
        Timer1.Dispose()
        Timer2.Dispose()
        Timer3.Dispose()
    End Sub
    'Borderless form control
    Protected Overrides Sub WndProc(ByRef m As Message)
        MyBase.WndProc(m)

        If m.Msg = WM_NCHITTEST AndAlso m.Result = HTCLIENT Then
            m.Result = HTCAPTION
        End If
    End Sub
    'Rounded corners on Form>>>>
    Private Function RoundedRectangle(rect As RectangleF,
                                      diam As Single) As Drawing2D.GraphicsPath
        Dim path As New Drawing2D.GraphicsPath
        path.AddArc(rect.Left, rect.Top, diam, diam, 180, 90)
        path.AddArc(rect.Right - diam, rect.Top, diam, diam, 270, 90)
        path.AddArc(rect.Right - diam, rect.Bottom - diam, diam, diam, 0, 90)
        path.AddArc(rect.Left, rect.Bottom - diam, diam, diam, 90, 90)
        path.CloseFigure()
        Return path
    End Function
    Private Sub Form1_Paint(sender As Object,
                            e As PaintEventArgs) Handles Me.Paint
        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        Dim r As New Rectangle(1, 1, Width - 2, Height - 2)

        Using pn As New Pen(Color.Black, 2)
            Dim path As Drawing2D.GraphicsPath = RoundedRectangle(r, 48)
            e.Graphics.DrawPath(pn, path)
        End Using
    End Sub
    'End Rounded Corners>>>

    'Multi-Threading is vital when programming and knowing how to threading.
    'This allows your program to operate smoothly. Background Workers can be used for lighter task.
    'Be sure your theads aren't running into the same methods together, it can slow down or freeze your system.
    Private Sub Thread1()
        Dim t As New Thread(Sub()
                                If Not Directory.Exists(path:=$"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\WM") Then
                                    Directory.CreateDirectory(path:=$"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\WM")
                                End If
                            End Sub)
        t.Start()
        t.Join()
    End Sub

    Private Sub Thread2()
        Dim t As New Thread(Sub()
                                rate = 44100 'sound card configuration
                                buffer_update_hz = 20
                                pixelsPerBuffer = 10
                                unanalyzed_max_sec = 2.5
                                fft_size = 4096 ' must be a multiple of 2
                                spec_height = fft_size / 2
                                spec_data = New List(Of List(Of Double))()
                                Dim data_empty As New List(Of Double)()
                                For i = 0 To spec_height - 1
                                    data_empty.Add(0)
                                Next
                                For i = 0 To spec_width - 1
                                    spec_data.Add(data_empty)
                                Next
                            End Sub)
        t.Start()
        t.Join()
    End Sub

    Private Sub Thread4(args)
        Call Task.Run(Sub()

                          Invoke(New Action(Sub()
                                                Dim t As New Thread(Sub()
                                                                        buffers_captured += 1
                                                                        buffers_remaining += 1
                                                                    End Sub)
                                                t.Start()
                                                t.Join()
                                            End Sub))
                      End Sub)

    End Sub
    Private Sub Thread5()

    End Sub
    'Record
    Private Sub Button1_Click(sender As Object,
                              e As EventArgs) Handles Button1.Click
        Call Task.Run(Sub()
                          Thread.Sleep(10)
                          Invoke(New Action(Sub()
                                                Thread2() 'Sound Card Configuration
                                                Dim waveIn = New WaveIn With {
                                                                               .DeviceNumber = 0
                                                                           }
                                                AddHandler waveIn.DataAvailable, AddressOf Audio_buffer_captured
                                                waveIn.WaveFormat = New WaveFormat(rate, 1)
                                                waveIn.BufferMilliseconds = 1000 / buffer_update_hz
                                                waveIn.StartRecording()

                                                waveSource = New WaveIn() With {
                                                                                                           .WaveFormat = New WaveFormat(44100, 1)
                                                                                                       }
                                                AddHandler waveSource.DataAvailable, New EventHandler(Of WaveInEventArgs)(AddressOf WaveSource_DataAvailable)
                                                AddHandler waveSource.RecordingStopped, New EventHandler(Of StoppedEventArgs)(AddressOf WaveSource_RecordingStopped)

                                                waveFile = New WaveFileWriter($"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\WM\ScarlettJournal.wav", waveSource.WaveFormat)
                                                waveSource.StartRecording()
                                                Timer1.Enabled = True
                                                PictureBox11.Visible = True
                                                PictureBox2.Visible = True
                                            End Sub))
                      End Sub)

    End Sub

    Public Sub BoostRAM()
        Call Task.Run(Sub()
                          Thread.Sleep(50)
                          Invoke(New Action(Sub()
                                                Dim FUEL, used As String
                                                Dim BOOST As New ComputerInfo()
                                                Dim mem As ULong = ULong.Parse(BOOST.AvailablePhysicalMemory.ToString())
                                                FUEL = (mem / (1024 * 1024) & " GB").ToString() 'changed + to &
                                                Dim mem1 As ULong = ULong.Parse(BOOST.TotalPhysicalMemory.ToString()) _
                                                    - ULong.Parse(BOOST.AvailablePhysicalMemory.ToString())
                                                used = (mem1 / (1024 * 1024) & " GB").ToString() 'changed + to &
                                            End Sub))
                      End Sub)
    End Sub


    'Stop Recording
    Private Sub Button2_Click(sender As Object,
                              e As EventArgs) Handles Button2.Click

        Call Task.Run(Sub()
                          Thread.Sleep(1000)
                          Invoke(New Action(Sub()
                                                Application.ExitThread()
                                                PictureBox2.Visible = False
                                                Label1.Hide()
                                                Label2.Hide()
                                                Label3.Hide()
                                                PictureBox11.Visible = False
                                                Shell("wscript.exe C:\Users\rytho\source\repos\Morrigan\Morrigan\caller.vbs", vbNormalFocus)
                                            End Sub))
                      End Sub)
    End Sub

    Private Sub WaveSource_DataAvailable(sender As Object,
                                         e As WaveInEventArgs)
        Call Task.Run(Sub()
                          Thread.Sleep(50)
                          Invoke(New Action(Sub()
                                                If waveFile IsNot Nothing Then
                                                    waveFile.Write(e.Buffer, 0, e.BytesRecorded)
                                                    waveFile.Flush()
                                                End If
                                            End Sub))
                      End Sub)
    End Sub

    'Disposal when recording ends.
    Private Sub WaveSource_RecordingStopped(sender As Object,
                                          e As StoppedEventArgs)
        Call Task.Run(Sub()
                          Thread.Sleep(50)
                          Invoke(New Action(Sub()
                                                If waveSource IsNot Nothing Then
                                                    waveSource.Dispose()
                                                    waveSource = Nothing
                                                End If
                                                If waveFile IsNot Nothing Then
                                                    waveFile.Dispose()
                                                    waveFile = Nothing
                                                End If
                                            End Sub))
                      End Sub)
    End Sub

    Private Sub Analyze_values()
        Call Task.Run(Sub()

                          Invoke(New Action(Sub()
                                                If fft_size = 0 Then Return
                                                If unanalyzed_values.Count < fft_size Then Return
                                                Label1.Text = String.Format("Analysis: {0}", unanalyzed_values.Count)
                                                While unanalyzed_values.Count >= fft_size
                                                    Analyze_chunk()
                                                End While
                                                Label1.Text = String.Format("Analysis: up to date")
                                            End Sub))
                      End Sub)

    End Sub
    Private Sub Analyze_chunk()
        Call Task.Run(Sub()
                          Thread.Sleep(50)
                          Invoke(New Action(Sub()
                                                Thread5()
                                            End Sub))
                      End Sub)

    End Sub
    Private Sub Audio_buffer_captured(sender As Object,
                                      args As WaveInEventArgs)
        Call Task.Run(Sub()
                          Thread.Sleep(50)
                          Invoke(New Action(Sub()
                                                Thread4(args)
                                                Dim values = New Short((args.Buffer.Length / 2) - 1) {}
                                                For i As Integer = 0 To args.BytesRecorded - 1 Step 2
                                                    values(i / 2) = CShort(args.Buffer(i + 1) << 8 Or args.Buffer(i + 0))
                                                Next

                                                ' add these values to the growing list, but ensure it doesn't get too big
                                                unanalyzed_values.AddRange(values)

                                                Dim unanalyzed_max_count = CInt(unanalyzed_max_sec) * rate

                                                If unanalyzed_values.Count > unanalyzed_max_count Then
                                                    unanalyzed_values.RemoveRange(0, unanalyzed_values.Count - unanalyzed_max_count)
                                                End If
                                                Label1.Text = String.Format("Buffers captured: {0}", buffers_captured)
                                                Label2.Text = String.Format("Buffer size: {0}", values.Length)
                                                Label3.Text = String.Format("Unanalyzed values: {0}", unanalyzed_values.Count)
                                            End Sub))
                      End Sub)

    End Sub

    Private Sub Timer1_Tick(sender As Object,
                            e As EventArgs)
        Call Task.Run(Sub()
                          Thread.Sleep(30)
                          Invoke(New Action(Sub()
                                                Analyze_values()
                                            End Sub))
                      End Sub)

    End Sub

    Private Sub Morrigan_FormClosing(sender As Object,
                                     e As FormClosingEventArgs) Handles MyBase.FormClosing
        Call Task.Run(Sub()
                          Thread.Sleep(1000)
                          Invoke(New Action(Sub()
                                                Timer1.Dispose()
                                                Timer2.Dispose()
                                                Timer3.Dispose()
                                                Shell("wscript.exe C:\Users\rytho\source\repos\Morrigan\Morrigan\caller.vbs", vbNormalFocus)
                                            End Sub))
                      End Sub)

    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object,
                                         e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        BoostRAM()
    End Sub



    'Flash PictureBox
    Private Async Sub Flash()
        While True
            Await Task.Delay(1200)
            PictureBox7.Visible = Not PictureBox7.Visible
        End While
    End Sub
    Private Async Sub Flash1()
        While True
            Await Task.Delay(1200)
            PictureBox8.Visible = Not PictureBox8.Visible
        End While
    End Sub
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        If PictureBox7.Visible Then
            PictureBox7.BackColor = Color.DarkGray
            Flash()
        End If
    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        If PictureBox8.Visible Then
            PictureBox8.BackColor = Color.DarkGray
            Flash1()
        End If
    End Sub
    'End Flash:>>>


    <DllImport("user32.dll")>
    Public Shared Function SendMessageW(hWnd As IntPtr,
Msg As Integer,
wParam As IntPtr,
lParam As IntPtr) As IntPtr

    End Function
    Declare Auto Function SetParent Lib "user32.dll" (hWndChild As IntPtr,
                                                      hWndNewParent As IntPtr) As Integer
    Declare Auto Function SendMessage Lib "user32.dll" (hWnd As IntPtr,
                                                        Msg As Integer,
                                                        wParam As Integer,
                                                        lParam As Integer) As Integer

    Private Sub PictureBox10_Click(sender As Object, e As EventArgs) Handles PictureBox10.Click
        Dim userName As String = Environment.UserName
        Dim savePath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim dateString As String = Date.Now.ToString("yyyyMMddHHmmss")
        Dim captureSavePath As String = String.Format($"{{0}}\WM\{{1}}\capture_{{2}}.png", savePath, userName, dateString)
        ' This line is modified for multiple screens, also takes into account different screen size (if any)
        Dim bmp As New Bitmap(
                Screen.AllScreens.Sum(Function(s As Screen) s.Bounds.Width),
                Screen.AllScreens.Max(Function(s As Screen) s.Bounds.Height))
        Dim gfx As Graphics = Graphics.FromImage(bmp)
        ' This line is modified to take everything based on the size of the bitmap
        gfx.CopyFromScreen(SystemInformation.VirtualScreen.X,
               SystemInformation.VirtualScreen.Y,
               0, 0, SystemInformation.VirtualScreen.Size)
        ' Oh, create the directory if it doesn't exist
        Directory.CreateDirectory(Path.GetDirectoryName(captureSavePath))
        bmp.Save(captureSavePath)

    End Sub
    Private Sub PictureBox1_MouseDown(sender As Object, e As Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseDown
        drag = True 'Sets the variable drag to true.
        mousex = Cursor.Position.X - Left() 'Sets variable mousex
        mousey = Cursor.Position.Y - Top 'Sets variable mousey
    End Sub

    Private Sub PictureBox1_MouseMove(sender As Object, e As Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove
        If drag Then
            Top = Cursor.Position.Y - mousey
            Left = Cursor.Position.X - mousex
        End If
    End Sub

    Private Sub PictureBox1_MouseUp(sender As Object, e As Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseUp
        drag = False 'Sets drag to false, so the form does not move according to the code in MouseMove
    End Sub

    'Volume up
    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click
        Dim repeat = 10
        For i = 0 To repeat - 1
            SendMessageW(Me.Handle, WM_APPCOMMAND, Me.Handle, CType(APPCOMMAND_VOLUME_UP, IntPtr))
        Next
    End Sub

    'Volume down
    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click
        Dim repeat = 10
        For i = 0 To repeat - 1
            SendMessageW(Me.Handle, WM_APPCOMMAND, Me.Handle, CType(APPCOMMAND_VOLUME_DOWN, IntPtr))
        Next
    End Sub


    Private Sub BackgroundWorker2_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Call Task.Run(Sub()
                          Thread.Sleep(3000)
                          Invoke(New Action(Sub()
                                                Thread1()
                                            End Sub))
                      End Sub)
    End Sub

End Class
