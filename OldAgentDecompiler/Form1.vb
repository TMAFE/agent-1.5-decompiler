Imports System.IO
Imports OpenMcdf
Imports System.Drawing
Imports System.Drawing.Imaging

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ofd As New OpenFileDialog() With {
            .Filter = "Agent 1.5 Character (*.acs;*.acf)|*.acs;*.acf|All files (*.*)|*.*",
            .Title = "Open Agent 1.5 Character"
        }

        If ofd.ShowDialog() <> DialogResult.OK Then
            Return
        End If

        Dim charPath = ofd.FileName
        Dim outDir = Path.GetDirectoryName(charPath)

        Try
            Dim palette As Color() = Agent15Extract.TryLoadCustomPalette(outDir)
            Dim ext = Path.GetExtension(charPath).ToLowerInvariant()
            If ext = ".acs" Then
                ProcessAcsContainer(charPath, outDir, palette)
            ElseIf ext = ".acf" Then
                ProcessLooseAcfAndAafs(charPath, outDir, palette)
            Else
                MessageBox.Show("Please select an Agent 1.5 .ACS or .ACF file.")
            End If

            MessageBox.Show("Extraction finished.")
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub ' this code ATTEMPTS to find the embedded ACF and AAF inside a 1.5 ACS. This will NOT work with compressed 1.5 chars.

    Private Sub ProcessAcsContainer(acsPath As String, outputDir As String, palette As Color())
        Using cf As New CompoundFile(acsPath)
            Dim root = cf.RootStorage


            Dim acfStream As CFStream = Nothing
            Try
                acfStream = root.GetStream("char.acf")
            Catch

            End Try

            If acfStream IsNot Nothing Then
                Dim acfBytes = acfStream.GetData()
                File.WriteAllBytes(Path.Combine(outputDir, "char_extracted.acf"), acfBytes)
            End If
            root.VisitEntries(
                Sub(item As CFItem)
                    Dim entry = TryCast(item, CFStream)
                    If entry Is Nothing Then Return

                    If entry.Name.EndsWith(".aaf", StringComparison.OrdinalIgnoreCase) Then
                        Dim aafBytes = entry.GetData()
                        Dim animName = Path.GetFileNameWithoutExtension(entry.Name)

                        ' this portion attempts WAV extraction
                        Dim wavOut = Path.Combine(outputDir, animName & ".wav")
                        Agent15Extract.ExtractWavFromAaf(aafBytes, wavOut)

                        ' NOTE: THIS EXTRACTS IT AS A GRAYSCALE BMP since I have not found a way to reverse the file for the original pallette HOWEVER I did make a workaround
                        ' where you can place a file called custompallette.bmp into the same directory as the character.
                        ' For this to work at its best, take a screenshot of the restpose frame of the 1.5 agent and save it as a 256 BMP 
                        ' and place it in the same directory as the file

                        Dim bmpOut = Path.Combine(outputDir, animName & ".bmp")
                        Try
                            Agent15Extract.ExtractBitmapFromAaf(aafBytes, bmpOut, palette)
                        Catch ex As Exception
                            Debug.WriteLine("BMP extraction failed for " & animName & ": " & ex.Message)
                        End Try
                    End If
                End Sub,
                True
            )
        End Using
    End Sub

    Private Sub ProcessLooseAcfAndAafs(acfPath As String, outputDir As String, palette As Color())
        Dim acfBytes = File.ReadAllBytes(acfPath)
        File.WriteAllBytes(Path.Combine(outputDir, "char_extracted.acf"), acfBytes)

        ' Looks for  any .aaf in the same directory
        For Each aafPath In Directory.GetFiles(outputDir, "*.aaf")
            Dim aafBytes = File.ReadAllBytes(aafPath)
            Dim animName = Path.GetFileNameWithoutExtension(aafPath)

            Dim wavOut = Path.Combine(outputDir, animName & ".wav")
            Agent15Extract.ExtractWavFromAaf(aafBytes, wavOut)

            Dim bmpOut = Path.Combine(outputDir, animName & ".bmp")
            Agent15Extract.ExtractBitmapFromAaf(aafBytes, bmpOut, palette)
        Next
    End Sub

End Class

'  Agent 1.5 AAF extractors - WAV and 8bpp indexed BMP
Public Module Agent15Extract

    '  WAV extractor

    Public Sub ExtractWavFromAaf(aafBytes As Byte(), outputPath As String)
        Dim riffSig As Byte() = System.Text.Encoding.ASCII.GetBytes("RIFF")
        Dim waveSig As Byte() = System.Text.Encoding.ASCII.GetBytes("WAVE")

        Dim idx As Integer = IndexOfBytes(aafBytes, riffSig, 0)
        If idx < 0 Then
            Throw New InvalidDataException("RIFF header not found in AAF.")
        End If

        ' Check for "WAVE" at idx+8
        If Not BytesEqual(aafBytes, idx + 8, waveSig) Then
            Throw New InvalidDataException("WAVE chunk not found after RIFF header.")
        End If

        Dim size As Integer = BitConverter.ToInt32(aafBytes, idx + 4)
        Dim totalLen As Integer = size + 8

        If idx + totalLen > aafBytes.Length Then
            Throw New InvalidDataException("RIFF chunk size goes past end of file.")
        End If

        Dim wav As Byte() = New Byte(totalLen - 1) {}
        Buffer.BlockCopy(aafBytes, idx, wav, 0, totalLen)

        File.WriteAllBytes(outputPath, wav)
    End Sub

    '  Extract a single 8-bit frame as a BMP
    '  If palette Is Nothing -> 256-level grayscale (look at my previous comment!)
    '  otherwise use provided 256-entry Color() palette
    Public Sub ExtractBitmapFromAaf(aafBytes As Byte(),
                                    outputPath As String,
                                    Optional palette As Color() = Nothing)

        Dim riffSig As Byte() = System.Text.Encoding.ASCII.GetBytes("RIFF")
        Dim riffIdx As Integer = IndexOfBytes(aafBytes, riffSig, 0)
        If riffIdx < 0 Then
            Throw New InvalidDataException("RIFF not found in AAF.")
        End If

        Dim riffSize As Integer = BitConverter.ToInt32(aafBytes, riffIdx + 4)
        Dim audioEnd As Integer = riffIdx + 8 + riffSize

        If audioEnd + 6 > aafBytes.Length Then
            Throw New InvalidDataException("AAF too short after RIFF block.")
        End If

        Dim frameCount As UShort = BitConverter.ToUInt16(aafBytes, audioEnd)
        Dim frameSize As Integer = BitConverter.ToInt32(aafBytes, audioEnd + 2)

        If frameCount < 1 Then
            Throw New InvalidDataException("No frames declared in AAF.")
        End If
        If frameSize <= 0 Then
            Throw New InvalidDataException("Invalid frame size in AAF.")
        End If

        Dim pixelOffset As Integer = audioEnd + 2 + 4
        If pixelOffset + frameSize > aafBytes.Length Then
            Throw New InvalidDataException("Frame size goes past end of file.")
        End If

        Dim pixels(frameSize - 1) As Byte
        Buffer.BlockCopy(aafBytes, pixelOffset, pixels, 0, frameSize)

        Dim side As Integer = CInt(Math.Truncate(Math.Sqrt(frameSize)))
        If side * side <> frameSize Then
            Throw New InvalidDataException($"Frame size {frameSize} is not a perfect square; cannot infer dimensions.")
        End If

        Using bmp As New Bitmap(side, side, PixelFormat.Format8bppIndexed)
            Dim pal As ColorPalette = bmp.Palette

            If palette IsNot Nothing AndAlso palette.Length >= 256 Then
                For i = 0 To 255
                    pal.Entries(i) = palette(i)
                Next
            Else

                For i = 0 To 255
                    pal.Entries(i) = Color.FromArgb(i, i, i)
                Next
            End If

            bmp.Palette = pal

            Dim rect As New Rectangle(0, 0, side, side)
            Dim bmpData = bmp.LockBits(rect, ImageLockMode.WriteOnly, PixelFormat.Format8bppIndexed)

            Dim stride As Integer = bmpData.Stride
            Dim scanline As Integer = side

            For y As Integer = 0 To side - 1
                Dim destOffset As Integer = bmpData.Scan0.ToInt64() + (side - 1 - y) * stride
                Dim srcOffset As Integer = y * scanline
                System.Runtime.InteropServices.Marshal.Copy(pixels, srcOffset, CType(destOffset, IntPtr), scanline)
            Next

            bmp.UnlockBits(bmpData)
            bmp.Save(outputPath, ImageFormat.Bmp)
        End Using
    End Sub

    '  Try to load palette from custompallette.bmp in the same folder
    '  If it exists and is indexed (<=256 colors), use its palette
    ' NOTE: It will use the custompallette I mentioned twice previously BUT it will fallback to grayscale if not present
    Public Function TryLoadCustomPalette(baseDir As String) As Color()
        If String.IsNullOrEmpty(baseDir) OrElse Not Directory.Exists(baseDir) Then
            Return Nothing
        End If

        Dim customPath As String = Path.Combine(baseDir, "custompallette.bmp")
        If Not File.Exists(customPath) Then
            Return Nothing
        End If

        Return LoadPaletteFromBmp(customPath)
    End Function

    Private Function LoadPaletteFromBmp(bmpPath As String) As Color()
        Try
            Using bmp As New Bitmap(bmpPath)
                Dim pal As ColorPalette = bmp.Palette
                If pal Is Nothing OrElse pal.Entries Is Nothing OrElse pal.Entries.Length = 0 Then
                    ' If the user accidentally saved as 24-bit truecolor instead of indexed,this breaks stuff so return Nothing.
                    Return Nothing
                End If

                Dim count As Integer = Math.Min(256, pal.Entries.Length)
                Dim colors(255) As Color
                For i = 0 To count - 1
                    colors(i) = pal.Entries(i)
                Next
                For i = count To 255
                    colors(i) = Color.Black
                Next

                Return colors
            End Using
        Catch
            Return Nothing
        End Try
    End Function

    Public Function IndexOfBytes(haystack As Byte(), needle As Byte(), start As Integer) As Integer
        For i = start To haystack.Length - needle.Length
            Dim ok As Boolean = True
            For j = 0 To needle.Length - 1
                If haystack(i + j) <> needle(j) Then
                    ok = False
                    Exit For
                End If
            Next
            If ok Then Return i
        Next
        Return -1
    End Function

    Private Function BytesEqual(arr As Byte(), offset As Integer, needle As Byte()) As Boolean
        If offset + needle.Length > arr.Length Then Return False
        For i = 0 To needle.Length - 1
            If arr(offset + i) <> needle(i) Then Return False
        Next
        Return True
    End Function

End Module
