Imports System.IO

Module X360Zip
    Public Structure zipHeaders
        Public Structure localFileEntry
            Implements IComparable(Of localFileEntry)
            Shared signature As Byte() = {&H50, &H4B, &H3, &H4}
            Public minVersion As UShort
            Public genBitFlag As UShort
            Public compressionMethod As UShort
            Public lastModTime As UShort
            Public lastModDate As UShort
            Public CRC32 As UInt32
            Public compressedSize As UInt32
            Public uncompressedSize As UInt32
            Public fileNameLength As UShort
            Public exFieldLength As UShort
            Public FileName As String
            Public exField As String
            Public Data As Byte()
            Public positionInFile As UInt32
            Public Function CompareTo(ByVal other As localFileEntry) As Integer Implements System.IComparable(Of localFileEntry).CompareTo 'Lets us sort it
                Dim myBytes As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(FileName)
                Dim otherBytes As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(other.FileName)
                Dim length As Integer = myBytes.Length
                If otherBytes.Length > myBytes.Length Then length = otherBytes.Length
                ReDim Preserve myBytes(length)
                ReDim Preserve otherBytes(length)
                For i As Integer = 0 To length
                    If myBytes(i) > otherBytes(i) Then Return 1
                    If myBytes(i) < otherBytes(i) Then Return -1
                Next
                Return 0
            End Function
        End Structure
        Public Structure centralFileHeader
            Shared signature As Byte() = {&H50, &H4B, &H1, &H2}
            Public versionMadeBy As UShort
            Public minVersion As UShort
            Public genBitFlag As UShort
            Public compressionMethod As UShort
            Public lastModTime As UShort
            Public lastModDate As UShort
            Public CRC32 As UInt32
            Public compressedSize As UInt32
            Public uncompressedSize As UInt32
            Public fileNameLength As UShort
            Public exFieldLength As UShort
            Public fileCommentLength As UShort
            Public diskStart As UShort
            Public intFileAttrib As UShort
            Public extFileAttrib As UInt32
            Public localHeaderLoc As UInt32
            Public FileName As String
            Public exField As String
            Public comment As String
        End Structure
        Public Structure centralMainHeader
            Shared signature As Byte() = {&H50, &H4B, &H5, &H6}
            Public curDiskNum As UShort
            Public startDiskNum As UShort
            Public numCentralRecordsOnDisk As UShort
            Public numCentralRecords As UShort
            Public centralDirSize As UInt32
            Public centralDirOffset As UInt32
            Public zipCommentLen As UShort
            Public zipComment As String
            Public offset As UInt32
        End Structure
    End Structure

    Public Class zipFile
        Dim file As Byte()
        Dim mainHeader As zipHeaders.centralMainHeader
        Public centralFileHeaders As zipHeaders.centralFileHeader()
        Public localFileEntries As zipHeaders.localFileEntry()
        Public Function readUShort(ByRef pos As UInt32) 'get a ushort from the zip headers
            Dim result As UShort
            result = file(pos + 1)
            result = (result << 8) + file(pos)
            pos += 2
            Return result
        End Function
        Public Function readUInt32(ByRef pos As UInt32) 'Get a uint32 from the zip headers
            Dim result As UInt32
            result = file(pos + 3) 'Mmm, endianness
            result = (result << 8) + file(pos + 2)
            result = (result << 8) + file(pos + 1)
            result = (result << 8) + file(pos)
            pos += 4
            Return result
        End Function
        Public Function readBytes(ByRef pos As UInt32, ByVal byteCount As Integer) 'Get a block of bytes from the zip headers
            Dim Bytes As Byte()
            ReDim Bytes(byteCount - 1)
            For i As Integer = 0 To byteCount - 1
                Bytes(i) = file(pos + i)
            Next
            pos += byteCount
            Return Bytes
        End Function
        Public Function readString(ByRef pos As UInt32, ByVal strSize As Integer) 'Get a string from the zip headers
            Return System.Text.ASCIIEncoding.ASCII.GetString(readBytes(pos, strSize))
        End Function

        Public Sub New(ByVal FileName As String)
            Trace.WriteLine("Start load!")
            loadFile(FileName)
        End Sub
        Private Function loadFile(ByVal FileName As String) As Boolean
            file = IO.File.ReadAllBytes(FileName) 'Load filename. Should add an if exists... check
            mainHeader.offset = getMainHeaderStart()
            If (parseMainHeader()) Then
                parseCentralDirectory()
                parseLocalFileEntries()
            End If
            Return True
        End Function

        Private Function getMainHeaderStart() As UInt32
            Dim offset As UInt32 = 0
            Dim match As Boolean = False 'Flag for signature search
            For i As Integer = file.Length - 1 To 0 Step -1
                If i + zipHeaders.centralMainHeader.signature.Length < file.Length Then 'Don't go out of bounds
                    match = True
                    For j As Integer = 0 To zipHeaders.centralMainHeader.signature.Length - 1
                        If Not file(i + j) = zipHeaders.centralMainHeader.signature(j) Then
                            match = False
                            Exit For
                        End If
                    Next
                End If
                If match Then 'If we found the signature
                    Trace.WriteLine("Central Directory Header found at 0x" & Hex(i))
                    offset = i
                    Exit For
                End If
            Next

            If match = False Then 'No match for central directory
                'TODO: How to handle?
            End If
            Return offset
        End Function
        Private Function parseMainHeader() As Boolean
            If mainHeader.offset = 0 Then Return False 'No main header found(call getMainHeaderStart first? Is this a valid zip?)
            Dim curOffset As UInt32 = mainHeader.offset
            curOffset += zipHeaders.centralMainHeader.signature.Length 'Offset by signature length. Don't care 'bout it.
            mainHeader.curDiskNum = readUShort(curOffset)
            mainHeader.startDiskNum = readUShort(curOffset)
            mainHeader.numCentralRecordsOnDisk = readUShort(curOffset)
            mainHeader.numCentralRecords = readUShort(curOffset)
            mainHeader.centralDirSize = readUInt32(curOffset)
            mainHeader.centralDirOffset = readUInt32(curOffset)
            mainHeader.zipCommentLen = readUShort(curOffset)
            mainHeader.zipComment = readString(curOffset, mainHeader.zipCommentLen)
            Trace.WriteLine("There are " & mainHeader.centralDirSize & " files in this archive.")
            Return True
        End Function

        Private Function parseCentralDirectory() As Boolean
            ReDim centralFileHeaders(mainHeader.numCentralRecords - 1) 'This should be the same as numCentralRecordsOnDisk for x360 files.
            Dim curPos As UInt32 = mainHeader.centralDirOffset
            For i As Integer = 0 To mainHeader.numCentralRecords - 1
                getNextDirectoryFile(curPos)
                centralFileHeaders(i) = parseDirectoryFileHeader(curPos)
            Next
            Return True
        End Function
        Sub getNextDirectoryFile(ByRef curPos) 'Moves curPos to the next directory file
            Dim match As Boolean = False 'Flag for signature search
            For i As Integer = curPos - 1 To file.Length
                If i + zipHeaders.centralFileHeader.signature.Length < file.Length Then 'Don't go out of bounds
                    match = True
                    For j As Integer = 0 To zipHeaders.centralFileHeader.signature.Length - 1
                        If Not file(i + j) = zipHeaders.centralFileHeader.signature(j) Then
                            match = False
                            Exit For
                        End If
                    Next
                End If
                If match Then 'we found the signature
                    curPos = i + zipHeaders.centralFileHeader.signature.Length
                    Exit For
                End If
            Next
        End Sub
        Function parseDirectoryFileHeader(ByRef curPos As UInt32) As zipHeaders.centralFileHeader
            parseDirectoryFileHeader.versionMadeBy = readUShort(curPos)
            parseDirectoryFileHeader.minVersion = readUShort(curPos)
            parseDirectoryFileHeader.genBitFlag = readUShort(curPos)
            parseDirectoryFileHeader.compressionMethod = readUShort(curPos)
            parseDirectoryFileHeader.lastModTime = readUShort(curPos)
            parseDirectoryFileHeader.lastModDate = readUShort(curPos)
            parseDirectoryFileHeader.CRC32 = readUInt32(curPos)
            parseDirectoryFileHeader.compressedSize = readUInt32(curPos)
            parseDirectoryFileHeader.uncompressedSize = readUInt32(curPos)
            parseDirectoryFileHeader.fileNameLength = readUShort(curPos)
            parseDirectoryFileHeader.exFieldLength = readUShort(curPos)
            parseDirectoryFileHeader.fileCommentLength = readUShort(curPos)
            parseDirectoryFileHeader.diskStart = readUShort(curPos)
            parseDirectoryFileHeader.intFileAttrib = readUShort(curPos)
            parseDirectoryFileHeader.extFileAttrib = readUInt32(curPos)
            parseDirectoryFileHeader.localHeaderLoc = readUInt32(curPos)
            parseDirectoryFileHeader.FileName = readString(curPos, parseDirectoryFileHeader.fileNameLength)
            parseDirectoryFileHeader.exField = "" 'Lets not do this. The x360 zips lack this field.
            parseDirectoryFileHeader.comment = "" 'This appears to be missing too.
        End Function

        Private Function parseLocalFileEntries() As Boolean 'Bet this one takes a long time (and some RAM). :D
            ReDim localFileEntries(mainHeader.numCentralRecords - 1)
            For i As Integer = 0 To mainHeader.numCentralRecords - 1
                localFileEntries(i) = createLocalFileEntryFromCentralFileHeader(centralFileHeaders(i))
            Next
            Return True
        End Function
        Function createLocalFileEntry(ByRef curPos As UInt32) As zipHeaders.localFileEntry
            curPos += zipHeaders.localFileEntry.signature.Length
            createLocalFileEntry.minVersion = readUShort(curPos)
            createLocalFileEntry.genBitFlag = readUShort(curPos)
            createLocalFileEntry.compressionMethod = readUShort(curPos)
            createLocalFileEntry.lastModTime = readUShort(curPos)
            createLocalFileEntry.lastModDate = readUShort(curPos)
            createLocalFileEntry.CRC32 = readUInt32(curPos)
            createLocalFileEntry.compressedSize = readUInt32(curPos)
            createLocalFileEntry.uncompressedSize = readUInt32(curPos)
            createLocalFileEntry.fileNameLength = readUShort(curPos)
            createLocalFileEntry.exFieldLength = readUShort(curPos)
            createLocalFileEntry.FileName = readString(curPos, createLocalFileEntry.fileNameLength)
            createLocalFileEntry.exField = readString(curPos, createLocalFileEntry.exFieldLength)
            createLocalFileEntry.Data = readBytes(curPos, createLocalFileEntry.uncompressedSize)
        End Function
        Function createLocalFileEntryFromCentralFileHeader(ByVal header As zipHeaders.centralFileHeader) As zipHeaders.localFileEntry
            Return createLocalFileEntry(header.localHeaderLoc)
        End Function

        Function getCentralFileHeaderIndex(ByVal header As zipHeaders.centralFileHeader)
            For i As Integer = 0 To centralFileHeaders.Length
                If header.FileName = centralFileHeaders(i).FileName Then Return i
            Next
            Return -1 'Not found
        End Function
        Function getLocalFileEntryIndex(ByVal header As zipHeaders.localFileEntry)
            For i As Integer = 0 To localFileEntries.Length
                If header.FileName = localFileEntries(i).FileName Then Return i
            Next
            Return -1 'Not found
        End Function
        Function getLocalFileEntryFromCentralFileHeader(ByVal header As zipHeaders.centralFileHeader)
            Return localFileEntries(getCentralFileHeaderIndex(header))
        End Function
        Function getCentralFileHeaderFromLocalFileEntry(ByVal header As zipHeaders.localFileEntry)
            Return centralFileHeaders(getLocalFileEntryIndex(header))
        End Function

        Public Function getFileFromCentralFileHeader(ByVal header As zipHeaders.centralFileHeader) As Byte()
            Return createLocalFileEntryFromCentralFileHeader(header).Data
        End Function
        Public Function getFileFromLocalFileEntry(ByVal header As zipHeaders.localFileEntry) As Byte()
            Return header.Data
        End Function

        Public Sub extractTo(ByVal folderName As String)
            Dim text As String = ""
            For i As Integer = 0 To centralFileHeaders.Length - 1
                Dim fileName As String = folderName & centralFileHeaders(i).FileName.Replace("/", "\")
                Dim filePath As String = Mid(fileName, 1, fileName.LastIndexOf("\"))
                IO.Directory.CreateDirectory(filePath)
                IO.File.WriteAllBytes(fileName, getFileFromCentralFileHeader(centralFileHeaders(i)))
            Next
        End Sub
    End Class

    Public Class zipFolder
        Dim _folderName As String
        Dim Directories As String()
        Dim Files As String()
        Dim mainHeader As zipHeaders.centralMainHeader
        Dim localFileData As Byte()
        Dim centralFileData As Byte()
        Dim mainHeaderData As Byte()
        Dim fileData As Byte()
        Public centralFileHeaders As zipHeaders.centralFileHeader()
        Public localFileEntries As zipHeaders.localFileEntry()

        Private Function writeUShort(ByRef file As Byte(), ByRef pos As UInt32, ByVal Number As UShort) 'Write a ushort to file
            Dim result As UShort
            file(pos) = Number And &HFF
            file(pos + 1) = (Number >> 8) And &HFF
            pos += 2
            Return result
        End Function
        Private Function writeUInt32(ByRef file As Byte(), ByRef pos As UInt32, ByVal Number As UInt32) 'Write a uint32 to file
            Dim result As UInt32
            file(pos) = Number And &HFF
            file(pos + 1) = (Number >> 8) And &HFF
            file(pos + 2) = (Number >> 16) And &HFF
            file(pos + 3) = (Number >> 24) And &HFF
            pos += 4
            Return result
        End Function
        Private Sub writeBytes(ByRef file As Byte(), ByRef pos As UInt32, ByVal bytes As Byte()) 'Write a block of bytes to file
            If IsNothing(bytes) Then Exit Sub
            For i As Integer = 0 To bytes.Length - 1
                file(pos + i) = bytes(i)
            Next
            pos += bytes.Length
        End Sub
        Private Sub writeString(ByRef file As Byte(), ByRef pos As UInt32, ByVal Str As String) 'Write a string to file
            If Not IsNothing(Str) Then writeBytes(file, pos, System.Text.ASCIIEncoding.ASCII.GetBytes(Str))
        End Sub
        Private Function getCRC32(ByRef Buffer As Byte()) As UInt32 'http://dotnet-snippets.com/dns/calculate-crc32-hash-from-file-SID587.aspx
            Try
                Dim CRC32Result As Integer = &HFFFFFFFF
                Dim CRC32Table(256) As Integer
                Dim DWPolynomial As Integer = &HEDB88320
                Dim DWCRC As Integer
                Dim i As Integer, j As Integer, n As Integer

                'Create CRC32 Table
                For i = 0 To 255
                    DWCRC = i
                    For j = 8 To 1 Step -1
                        If (DWCRC And 1) Then
                            DWCRC = ((DWCRC And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                            DWCRC = DWCRC Xor DWPolynomial
                        Else
                            DWCRC = ((DWCRC And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                        End If
                    Next j
                    CRC32Table(i) = DWCRC
                Next i

                'Calcualting CRC32 Hash
                For i = 0 To Buffer.Length - 1
                    n = (CRC32Result And &HFF) Xor Buffer(i)
                    CRC32Result = ((CRC32Result And &HFFFFFF00) \ &H100) And &HFFFFFF
                    CRC32Result = CRC32Result Xor CRC32Table(n)
                Next i
                Return Not (CRC32Result)
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Public Sub New(ByVal folderName As String)
            loadFolder(folderName)
        End Sub

        Private Sub loadFolder(ByVal folderName As String)
            _folderName = folderName
            listAllFiles()
            createLocalFileEntries()
            createLocalFileData()
            createCentralFileHeaders()
            createCentralFileData()
            createMainCentralHeader()
            createMainHeaderData() 'Only called here to get how big it's going to be.
            createFileData()
            Trace.WriteLine("Loaded!")
        End Sub
        Private Sub listAllFiles()
            ReDim Files(-1)
            Directories = Directory.GetDirectories(_folderName, "*", SearchOption.AllDirectories)
            ReDim Preserve Directories(Directories.Length)
            Directories(Directories.Length - 1) = _folderName
            For Each folder As String In Directories
                Dim curPos As Integer = Files.Length
                Dim _files As String() = Directory.GetFiles(folder)
                ReDim Preserve Files(Files.Length + _files.Length - 1)
                For i As Integer = 0 To _files.Length - 1
                    Files(curPos + i) = Mid(_files(i), _folderName.Length + 1)
                Next
            Next
        End Sub

        Private Sub createLocalFileEntries()
            ReDim localFileEntries(Files.Length - 1)
            For i As Integer = 0 To Files.Length - 1
                If Not IsNothing(Files(i)) Then
                    localFileEntries(i).FileName = Files(i)
                    localFileEntries(i).fileNameLength = Files(i).Length
                    localFileEntries(i).Data = IO.File.ReadAllBytes(_folderName & Files(i))
                    localFileEntries(i).CRC32 = getCRC32(localFileEntries(i).Data)
                    localFileEntries(i).compressedSize = localFileEntries(i).Data.Length
                    localFileEntries(i).uncompressedSize = localFileEntries(i).Data.Length
                    localFileEntries(i).exFieldLength = 0
                    localFileEntries(i).compressionMethod = 0
                    localFileEntries(i).minVersion = &HA
                    localFileEntries(i).lastModDate = 0
                    localFileEntries(i).lastModTime = 0
                    localFileEntries(i).genBitFlag = 0
                End If
            Next
            Array.Sort(localFileEntries)
        End Sub
        Private Sub createLocalFileData()
            Dim totalLength As UInt32 = 0
            Dim curPos As UInt32 = 0
            For i As Integer = 0 To localFileEntries.Length - 1
                totalLength = totalLength + 30
                totalLength = totalLength + localFileEntries(i).fileNameLength
                totalLength = totalLength + localFileEntries(i).exFieldLength
                totalLength = totalLength + localFileEntries(i).uncompressedSize
            Next
            ReDim localFileData(totalLength + 1 + (localFileEntries.Length * &H800))
            For i As Integer = 0 To localFileEntries.Length - 1
                localFileEntries(i).positionInFile = curPos
                writeBytes(localFileData, curPos, zipHeaders.localFileEntry.signature)
                writeUShort(localFileData, curPos, localFileEntries(i).minVersion)
                writeUShort(localFileData, curPos, localFileEntries(i).genBitFlag)
                writeUShort(localFileData, curPos, localFileEntries(i).compressionMethod)
                writeUShort(localFileData, curPos, localFileEntries(i).lastModTime)
                writeUShort(localFileData, curPos, localFileEntries(i).lastModDate)
                writeUInt32(localFileData, curPos, localFileEntries(i).CRC32)
                writeUInt32(localFileData, curPos, localFileEntries(i).compressedSize)
                writeUInt32(localFileData, curPos, localFileEntries(i).uncompressedSize)
                writeUShort(localFileData, curPos, localFileEntries(i).fileNameLength)
                Dim exFieldLength As Integer = curPos + 2 + localFileEntries(i).fileNameLength
                'exFieldLength = &H800 - (exFieldLength Mod &H800)
                'totalLength = totalLength + exFieldLength
                'Dim exFieldBytes As Byte()
                'ReDim exFieldBytes(exFieldLength - 1)
                localFileEntries(i).exFieldLength = 0
                writeUShort(localFileData, curPos, 0)
                writeString(localFileData, curPos, localFileEntries(i).FileName)
                'writeBytes(localFileData, curPos, exFieldBytes)
                writeBytes(localFileData, curPos, localFileEntries(i).Data)
            Next
            ReDim Preserve localFileData(totalLength + 1)
        End Sub

        Private Sub createCentralFileHeaders()
            ReDim centralFileHeaders(localFileEntries.Length - 1)
            For i As Integer = 0 To Files.Length - 1
                If Not IsNothing(Files(i)) Then
                    centralFileHeaders(i).FileName = localFileEntries(i).FileName
                    centralFileHeaders(i).fileNameLength = localFileEntries(i).fileNameLength
                    centralFileHeaders(i).CRC32 = localFileEntries(i).CRC32
                    centralFileHeaders(i).compressedSize = localFileEntries(i).compressedSize
                    centralFileHeaders(i).uncompressedSize = localFileEntries(i).uncompressedSize
                    centralFileHeaders(i).minVersion = localFileEntries(i).minVersion
                    centralFileHeaders(i).versionMadeBy = &H14
                    centralFileHeaders(i).localHeaderLoc = localFileEntries(i).positionInFile
                    centralFileHeaders(i).exFieldLength = localFileEntries(i).exFieldLength

                End If
            Next
        End Sub
        Private Sub createCentralFileData()
            Dim totalLength As UInt32 = 0
            Dim curPos As UInt32 = 0
            For i As Integer = 0 To centralFileHeaders.Length - 1
                totalLength = totalLength + 46
                totalLength = totalLength + centralFileHeaders(i).fileNameLength
                totalLength = totalLength + centralFileHeaders(i).fileCommentLength
            Next
            ReDim centralFileData(totalLength)
            For i As Integer = 0 To centralFileHeaders.Length - 1
                writeBytes(centralFileData, curPos, zipHeaders.centralFileHeader.signature)
                writeUShort(centralFileData, curPos, centralFileHeaders(i).versionMadeBy)
                writeUShort(centralFileData, curPos, centralFileHeaders(i).minVersion)
                writeUShort(centralFileData, curPos, centralFileHeaders(i).genBitFlag)
                writeUShort(centralFileData, curPos, centralFileHeaders(i).compressionMethod)
                writeUShort(centralFileData, curPos, centralFileHeaders(i).lastModTime)
                writeUShort(centralFileData, curPos, centralFileHeaders(i).lastModDate)
                writeUInt32(centralFileData, curPos, centralFileHeaders(i).CRC32)
                writeUInt32(centralFileData, curPos, centralFileHeaders(i).compressedSize)
                writeUInt32(centralFileData, curPos, centralFileHeaders(i).uncompressedSize)
                writeUShort(centralFileData, curPos, centralFileHeaders(i).fileNameLength)
                writeUShort(centralFileData, curPos, centralFileHeaders(i).exFieldLength)
                writeUShort(centralFileData, curPos, centralFileHeaders(i).fileCommentLength)
                writeUShort(centralFileData, curPos, centralFileHeaders(i).diskStart)
                writeUShort(centralFileData, curPos, centralFileHeaders(i).intFileAttrib)
                writeUInt32(centralFileData, curPos, centralFileHeaders(i).extFileAttrib)
                writeUInt32(centralFileData, curPos, centralFileHeaders(i).localHeaderLoc)
                writeString(centralFileData, curPos, centralFileHeaders(i).FileName)
                writeString(centralFileData, curPos, centralFileHeaders(i).exField)
                writeString(centralFileData, curPos, centralFileHeaders(i).comment)
            Next
        End Sub

        Private Sub createMainCentralHeader()
            mainHeader.numCentralRecords = localFileEntries.Length
            mainHeader.numCentralRecordsOnDisk = mainHeader.numCentralRecords
            mainHeader.centralDirOffset = localFileData.Length - 1
            mainHeader.centralDirSize = centralFileData.Length
            mainHeader.zipComment = "XZP2 2048"
            mainHeader.zipCommentLen = &H20
        End Sub
        Private Sub createMainHeaderData()
            ReDim mainHeaderData(23 + mainHeader.zipCommentLen)
            Dim curPos As UInt32 = 0
            writeBytes(mainHeaderData, curPos, zipHeaders.centralMainHeader.signature)
            writeUShort(mainHeaderData, curPos, mainHeader.curDiskNum)
            writeUShort(mainHeaderData, curPos, mainHeader.startDiskNum)
            writeUShort(mainHeaderData, curPos, mainHeader.numCentralRecordsOnDisk)
            writeUShort(mainHeaderData, curPos, mainHeader.numCentralRecords)
            writeUInt32(mainHeaderData, curPos, mainHeader.centralDirSize)
            writeUInt32(mainHeaderData, curPos, mainHeader.centralDirOffset)
            Dim comment As Byte()
            comment = System.Text.ASCIIEncoding.ASCII.GetBytes(mainHeader.zipComment)
            ReDim Preserve comment(mainHeader.zipCommentLen + 1)
            writeUShort(mainHeaderData, curPos, mainHeader.zipCommentLen)
            writeBytes(mainHeaderData, curPos, comment)
        End Sub

        Private Sub createFileData()
            Dim curPos = 0
            Dim totalLength As UInt32 = localFileData.Length + centralFileData.Length + mainHeaderData.Length - 3
            ReDim fileData(localFileData.Length + centralFileData.Length + mainHeaderData.Length - 3 + 5000)
            For i As Integer = 0 To localFileData.Length - 1
                fileData(curPos) = localFileData(i)
                curPos += 1
            Next
            Dim padding As Integer = curPos
            padding = &H800 - (padding Mod &H800)
            totalLength += padding
            Dim exFieldBytes As Byte()
            ReDim exFieldBytes(padding - 1)
            writeBytes(fileData, curPos, exFieldBytes)
            mainHeader.centralDirOffset = curPos
            For i As Integer = 0 To centralFileData.Length - 1
                fileData(curPos) = centralFileData(i)
                curPos += 1
            Next
            padding = curPos
            padding = &H800 - (padding Mod &H800)
            totalLength += padding
            ReDim exFieldBytes(padding - 1)
            writeBytes(fileData, curPos, exFieldBytes)
            mainHeader.centralDirSize += padding
            createMainHeaderData() 'Called here to update the offset after padding to the central directory.
            For i As Integer = 0 To mainHeaderData.Length - 1
                fileData(curPos) = mainHeaderData(i)
                curPos += 1
            Next
            ReDim Preserve fileData(totalLength)
        End Sub
        Public Sub saveTo(ByVal filename As String)
            IO.File.WriteAllBytes(filename, fileData)
        End Sub
    End Class
End Module
