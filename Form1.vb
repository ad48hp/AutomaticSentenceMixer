Imports System.IO
Imports System.Xml
'Imports Fiddler

Public Class Form1
    Public krok As Integer
    Public myurl As String
    Public done As Boolean
    Public mnam2 As String
    Public nodezz As New List(Of Nodez)
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'FiddlerApplication.Shutdown()
        Dim mywriter As New StreamWriter("options")
        mywriter.WriteLine(TextBox8.Text)
        mywriter.WriteLine(TextBox2.Text)
        mywriter.WriteLine(TextBox3.Text)
        mywriter.WriteLine(TextBox5.Text)
        mywriter.WriteLine(TextBox6.Text)
        mywriter.WriteLine(TextBox7.Text)
        mywriter.WriteLine(TextBox1.Text)
        mywriter.Close()
    End Sub

    Private Sub SortOptions(v1 As Boolean, v2 As Boolean)
        If v1 Then
            TextBox1.Text = TextBox1.Text.Replace(",", " ").Replace(".", "")
            While TextBox1.Text.Contains("  ")
                TextBox1.Text = TextBox1.Text.Replace("  ", " ")
            End While
        End If
        If v2 Then
            TextBox2.Text = TextBox2.Text.Replace("http://www.youtube.com", "").Replace("https://www.youtube.com", "")
            While TextBox2.Text.StartsWith("//")
                TextBox2.Text = TextBox2.Text.Substring(1, TextBox2.Text.Length - 1)
            End While
            If Not TextBox2.Text.StartsWith("/") Then
                TextBox2.Text = "/" + TextBox2.Text
            End If
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Shown
        If Not Directory.Exists(Application.StartupPath + "/output") Then
            Directory.CreateDirectory(Application.StartupPath + "/output")
        End If
        If File.Exists(Application.StartupPath + "/options") Then
            Dim myreader As New StreamReader("options")
            TextBox8.Text = myreader.ReadLine()
            TextBox2.Text = myreader.ReadLine()
            TextBox3.Text = myreader.ReadLine()
            TextBox5.Text = myreader.ReadLine()
            TextBox6.Text = myreader.ReadLine()
            TextBox7.Text = myreader.ReadLine()
            While Not myreader.EndOfStream
                Dim myline As String
                myline = myreader.ReadLine()
                If Not myline = "" Then
                    TextBox1.Text += myline
                    TextBox1.Text += Environment.NewLine
                End If
            End While
            myreader.Close()
        End If
        For Each item In Me.Controls
            If TypeOf item Is TextBox AndAlso DirectCast(item, TextBox).Multiline Then
                AddHandler DirectCast(item, TextBox).KeyDown, AddressOf slttext
            End If
        Next
        TextBox1.SelectionStart = TextBox1.Text.Length
        TextBox1.SelectionLength = 0
        'AddHandler FiddlerApplication.AfterSessionComplete, AddressOf sesscomp
        'FiddlerApplication.Startup(8888, False, True, False)
    End Sub

    Private Sub slttext(sender As Object, e As KeyEventArgs)
        If e.Control AndAlso e.KeyCode = Keys.A Then
            Dim txt As TextBox
            txt = DirectCast(sender, TextBox)
            txt.SelectionStart = 0
            txt.SelectionLength = txt.Text.Length
        End If
    End Sub

    'Private Sub sesscomp(suss As Session)
    'If krok = 0 AndAlso suss.fullUrl.Contains("youtube.com/api/timedtext") Then
    'Dim mywriter69 As New StreamWriter(TextBox3.Text + "/" + mnam2 + ".xml")
    'suss.GetRequestBodyAsString()
    'mywriter69.Close()
    'End If
    'End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim mnpath As String
        mnpath = Application.StartupPath + "/output"
        If Not Directory.Exists(TextBox3.Text) Then
            If MessageBox.Show("Path doesn't exist. Do you want to create it ?", "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) Then
                Directory.CreateDirectory(TextBox3.Text)
            End If
        End If
        If TextBox8.Text = "" OrElse Not File.Exists(TextBox8.Text) Then
            MessageBox.Show("FFMPEG path is invalid.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        If TextBox5.Text = "" Then
            TextBox5.Text = "0"
        End If
        If TextBox6.Text = "" Then
            TextBox6.Text = "0"
        End If
        TextBox5.Text = TextBox5.Text.Replace(".", ",")
        TextBox6.Text = TextBox6.Text.Replace(".", ",")
        If Not Double.TryParse(TextBox5.Text, 0) Then
            MessageBox.Show("Start time isn't a number.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        If Not Double.TryParse(TextBox6.Text, 0) Then
            MessageBox.Show("End time isn't a number.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        If Not Integer.TryParse(TextBox7.Text, 0) OrElse Not Integer.Parse(TextBox7.Text, 0) > 0 Then
            MessageBox.Show("Max. number isn't an integer bigger than 0.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        If Directory.Exists(TextBox3.Text) AndAlso Not TextBox8.Text = "" AndAlso File.Exists(TextBox8.Text) AndAlso Double.TryParse(TextBox5.Text, 1) AndAlso Double.TryParse(TextBox6.Text, 1) AndAlso Double.TryParse(TextBox7.Text, 1) Then
            krok = 0
            If Not TextBox2.Text = "" Then
                'WebBrowser1.Navigate("https://youtube.com" + TextBox2.Text)
            End If
            Dim mdr As FileInfo() = New DirectoryInfo(TextBox3.Text).GetFiles()
            For Each item In mdr
                If item.Extension.ToLower().Replace(".", "") = "mp4" Then
                    Dim mnam As String
                    mnam = Path.GetFileNameWithoutExtension(item.FullName)
                    While Not File.Exists(TextBox3.Text + "/" + mnam + ".xml")
                        mnam2 = mnam
                        myurl = "https://youtube.com/watch?v=" + mnam
                        krok = 0
                        WebBrowser1.Navigate(myurl)
                        While Not krok = 1
                            Application.DoEvents()
                        End While
                    End While
                    Dim myxml As XDocument = XDocument.Load(TextBox3.Text + "/" + mnam + ".xml")
                    Dim mynodes As List(Of XElement)
                    mynodes = myxml.Descendants().Where(Function(z) z.Name.LocalName.ToLower() = "s").ToList()
                    For Each ndz In mynodes
                        Dim mt As Long
                        mt = 0
                        Dim mt2 As Long
                        mt2 = 0
                        If Not ndz.Attribute("t") Is Nothing AndAlso Not ndz.Attribute("t").Value = "" Then
                            mt += Long.Parse(ndz.Attribute("t").Value)
                        End If
                        Dim ndz2 As XElement
                        ndz2 = ndz.NodesAfterSelf().OfType(Of XElement).FirstOrDefault(Function(z) z.Name.LocalName = "s")
                        If Not ndz2 Is Nothing AndAlso Not ndz2.Attribute("t") Is Nothing AndAlso Not ndz2.Attribute("t").Value = "" Then
                            mt2 += Long.Parse(ndz2.Attribute("t").Value)
                        End If
                        mt += Long.Parse(ndz.Parent.Attribute("t"))
                        mt2 += Long.Parse(ndz.Parent.Attribute("t"))
                        Dim myndz As New Nodez(mt, mt2, ndz.Value.ToLower().Replace(" ", ""))
                        nodezz.Add(myndz)
                    Next
                    Dim ttxt As String()
                    ttxt = TextBox1.Text.ToLower().Split(New String() {" "}, StringSplitOptions.RemoveEmptyEntries)
                    Dim i31 As Integer
                    i31 = 1
                    For Each mndz In ttxt
                        Dim nodezzi As List(Of Nodez)
                        nodezzi = New List(Of Nodez)
                        Dim curi As Integer
                        curi = 1
                        For Each nodezi2 In nodezz
                            If nodezi2.text = mndz.ToLower().Replace(" ", "") Then
                                nodezzi.Add(nodezi2)
                                curi += 1
                            End If
                            If curi > Integer.Parse(TextBox7.Text) Then
                                Exit For
                            End If
                        Next
                        For Each nodezi In nodezzi
                            While File.Exists(mnpath + "/" + i31.ToString() + ".mp4")
                                i31 += 1
                            End While
                            Dim mstartinfo As New ProcessStartInfo
                            mstartinfo.FileName = TextBox8.Text
                            Dim tm1 As TimeSpan
                            Dim tm2 As TimeSpan
                            tm1 = TimeSpan.FromMilliseconds(nodezi.time)
                            tm1.Add(TimeSpan.FromSeconds(Double.Parse(TextBox5.Text)))
                            tm2.Add(TimeSpan.FromSeconds(Double.Parse(TextBox5.Text) + Double.Parse(TextBox6.Text)))
                            'tm2 = TimeSpan.FromMilliseconds(nodezi.endtime - nodezi.time)
                            'If tm2.TotalMilliseconds <= 0 Then
                            tm2 = TimeSpan.FromMilliseconds(1024)
                            'End If
                            Dim myname As String
                            If Integer.Parse(TextBox7.Text) = 1 Then
                                myname = i31.ToString()
                            Else
                                Dim i36 As Integer
                                i36 = 1
                                While File.Exists(mnpath + "/" + nodezi.text + i36.ToString() + ".mp4")
                                    i36 += 1
                                End While
                                myname = nodezi.text + i36.ToString()
                            End If
                            mstartinfo.Arguments = "-ss " + Math.Floor(tm1.TotalHours).ToString().PadLeft(2, "0") + ":" + tm1.Minutes.ToString().PadLeft(2, "0") + ":" + tm1.Seconds.ToString().PadLeft(2, "0") + "." + tm1.Milliseconds.ToString().PadLeft(2, "0") + " -i " + ChrW(34) + item.FullName + ChrW(34) + " -c copy -t " + Math.Floor(tm2.TotalHours).ToString().PadLeft(2, "0") + ":" + tm2.Minutes.ToString().PadLeft(2, "0") + ":" + tm2.Seconds.ToString().PadLeft(2, "0") + "." + tm2.Milliseconds.ToString().PadLeft(2, "0") + " " + ChrW(34) + mnpath + "/" + myname + ".mp4" + ChrW(34)
                            Process.Start(mstartinfo).WaitForExit()
                        Next
                    Next
                    If False Then
                        Dim mstartinfo2 As New ProcessStartInfo
                        mstartinfo2.FileName = Application.StartupPath + "/ffmpeg.exe"
                        mstartinfo2.WorkingDirectory = Application.StartupPath
                        Dim i87 As Integer
                        i87 = 1
                        While i87 < i31
                            mstartinfo2.Arguments = ""
                            For i42 = i87 To i87 + 4
                                If i87 <= i31 Then
                                    mstartinfo2.Arguments += "-i " + i42.ToString() + ".mp4 "
                                End If

                            Next
                            mstartinfo2.Arguments += "-filter_complex " + ChrW(34)
                            For i42 = i87 To i87 + 4
                                If i87 <= i31 Then
                                    mstartinfo2.Arguments += "[" + (i42 - 1).ToString() + ":v]"
                                End If
                            Next
                            mstartinfo2.Arguments += "hstack,format=yuv420p[v];"
                            For i42 = i87 To i87 + 4
                                If i87 <= i31 Then
                                    mstartinfo2.Arguments += "[" + (i42 - 1).ToString() + ":a]"
                                End If
                            Next
                            mstartinfo2.Arguments += "amerge[a]" + ChrW(34) + " -map " + ChrW(34) + "[v]" + ChrW(34) + " -map " + ChrW(34) + "[a]" + ChrW(34) + " -c:v libx264 -crf 18 -ac 2 editz.mp4"
                            Process.Start(mstartinfo2).WaitForExit()
                            i87 += 4 + 1
                        End While
                    End If
                End If
            Next
            krok = 1
        End If
    End Sub

    Private Sub Form1_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load
        WebBrowser1.ScriptErrorsSuppressed = True
        AddHandler WebBrowser1.DocumentCompleted, AddressOf doccomp
    End Sub

    Private Sub doccomp(sender As Object, e As WebBrowserDocumentCompletedEventArgs)
        Select Case krok
            Case 0
                If e.Url.AbsoluteUri.Contains("youtube.com/watch?v=" + mnam2) Then
                    Dim mstop As New Stopwatch
                    Dim bt1 As HtmlElement
                    Dim bt2 As HtmlElement
                    bt1 = Nothing
                    bt2 = Nothing
                    mstop.Reset()
                    mstop.Start()
                    While mstop.ElapsedMilliseconds < 2000
                        Application.DoEvents()
                    End While
                    For Each item31 As HtmlElement In WebBrowser1.Document.All
                        If item31.GetAttribute("classname").Contains("ytp-play-button") Then
                            bt1 = item31
                        End If
                        If item31.GetAttribute("classname").Contains("ytp-subtitles-button") Then
                            bt2 = item31
                        End If
                    Next
                    done = False
                    Dim i As Integer
                    i = 1
                    While Not done
                        bt1.InvokeMember("click")
                        mstop.Reset()
                        mstop.Start()
                        While mstop.ElapsedMilliseconds < 2000
                            Application.DoEvents()
                        End While
                        bt2.InvokeMember("click")
                        mstop.Reset()
                        mstop.Start()
                        While mstop.ElapsedMilliseconds < 4600
                            Application.DoEvents()
                        End While
                        i += 1
                        If i > 5 Then
                            done = True
                        End If
                    End While
                    krok = 1
                End If
        End Select
    End Sub
End Class
