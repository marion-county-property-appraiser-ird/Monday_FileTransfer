Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.Collections.ObjectModel
Imports System.IO.Compression
Imports Microsoft.VisualBasic.FileIO

Module Module1

    ' Private fromFile As String = String.Empty
    ' Private fromFolder As String = String.Empty '"C:\Users\ssherouse\Documents\MerlinDataExtracts\"
    Private folderDate As String = DateTime.Now.Year & DateTime.Now.Month.ToString.PadLeft(2, "0") & DateTime.Now.Day.ToString.PadLeft(2, "0")
    Private fridayDate As String

    Sub Main()
        Console.Title = "MONDAY ROUTINE"
        Console.WriteLine("")
        Console.WriteLine("Please enter last Friday's Date (yyyymmdd)")
        fridayDate = Console.ReadLine().Trim

        Console.WriteLine()
        Console.WriteLine("What are you doing? (Pick a Number)")
        Console.WriteLine("  1. Beginning Routine  OR  2. Ending Routine")
        Dim choice As String = Console.ReadLine()
        If choice.ToUpper = "1" Then
            'DO THIS AT THE START OF MONDAY ROUTINE

            'this copies files from rawdata\utility to datarecs 
            Console.WriteLine("Copying files from RawData to Datarecs" & fridayDate & " and Renaming the Extension to .txt")
            CopyTemplates()
            Console.WriteLine("Datarecs copy complete")

            'this zips mcpaextracts folder in datarecs and copies it to FTP for county
            Console.WriteLine("Now Zipping MCPAEXTRACTS....please wait")
            ZipMCPAEXTRACTS()
            Console.WriteLine("MCPAEXTRACTS copy complete")

            'this moves last weeks files to the back up forlder for the cd
            Console.WriteLine("Now moving County/CD files to backup folder....please wait")
            MoveCDFiles()
        Else
            'DO AT END OF ROUTINE

            'this zips oas file and copies it to FTP for county and mcpa cd folder
            Console.WriteLine("Now Zipping PARCELDATA for County/CD....please wait")
            ZipForFTP_CD()

            'this copies all papoly files and copies them to mcpa cd folder
            Console.WriteLine("Now Copying papoly files from County folder to CD folder....please wait")
            CopyCDFiles()

            'this zips mcpadata folder in datarecs and copies it to FTP for county
            Console.WriteLine("Now Zipping MCPADATA for County FTP....please wait")
            ZipMCPADATA()

            'this zips mcpapolygons folder in datarecs and copies it to FTP for county
            Console.WriteLine("Now Zipping MCPAPOLYGONS for County FTP....please wait")
            ZipMCPAPOLYGONS()

            'this copies all specific files from dated folder to Merlin server for map update
            Console.WriteLine("Now Copying ALL files from Dated folder to W:\MiscLayerData\Parcel for Map Update....please wait")
            CopyForMapUpdate()

            'this copies all specific files from dated folder to Joes folder on the k drive
            Console.WriteLine("Now Copying ALL files from Dated folder to K:\Polygons\MCPA\2003_MassImpVacAprsl\SSS_JOE\FreeancePAPOLY....please wait")
            CopyToJoe()
        End If

        Console.WriteLine("Process Finished")
    End Sub

    Private Sub CopyTemplates()
        Dim i As Integer
        Dim sFile As String = Nothing
        Dim newFile As String = Nothing
        Dim workingfiles As New ArrayList

        Dim files As ReadOnlyCollection(Of String)
        files = My.Computer.FileSystem.GetFiles("\\MCPAFILESERVER\K Drive\Polygons\RawData\Utility Files\" & fridayDate, FileIO.SearchOption.SearchTopLevelOnly, "*.*")
        For i = 0 To files.Count - 1
            sFile = files.Item(i).ToString
            newFile = Path.GetFileNameWithoutExtension(sFile)

            If Path.GetExtension(sFile).ToUpper = ".CSV" Then
                If newFile.Contains("AddressExtractView") Or newFile.Contains("DescriptionExtractView") Or newFile.Contains("MiscImprovementExtractView") Or newFile.Contains("NameExtractView") Or newFile.Contains("ParentParcelExtractView") Or newFile.Contains("SitusAddressExtractView") Then
                    If Not MoveFiles(sFile, "\\MCPAFILESERVER\K Drive\Polygons\MCPA\" & fridayDate & "\Datarecs" & fridayDate & "\MCPAEXTRACT\" & newFile & ".txt") Then Return
                ElseIf newFile.Contains("LandExtractView") Or newFile.Contains("MasterParcelExtractView") Or newFile.Contains("MobileHomeExtractView") Then
                    If Not CopyFiles(sFile, "\\MCPAFILESERVER\K Drive\Polygons\MCPA\" & fridayDate & "\Datarecs" & fridayDate & "\MCPAEXTRACT\" & newFile & ".txt", False) Then Return
                    If Not CopyFiles(sFile, "\\MCPAFILESERVER\K Drive\Polygons\MCPA\" & fridayDate & "\Datarecs" & fridayDate & "\" & newFile & ".txt", False) Then Return
                Else
                    If Not CopyFiles(sFile, "\\MCPAFILESERVER\K Drive\Polygons\MCPA\" & fridayDate & "\Datarecs" & fridayDate & "\" & newFile & ".txt", False) Then Return
                End If
            ElseIf Path.GetExtension(sFile).ToUpper = ".TXT" Then
                Path.GetFileName(sFile)
                CopyFiles(sFile, "\\MCPAFILESERVER\K Drive\Polygons\MCPA\" & fridayDate & "\Datarecs" & fridayDate & "\" & Path.GetFileName(sFile), False)
            End If
        Next

        If files.Count = 0 Then
            Console.WriteLine("No files available to copy")
        End If
    End Sub

    Private Sub MoveCDFiles()
        Dim startPath As String = "\\MCPAFILESERVER\K Drive\Polygons\MCPA\Bill_BCC_MCPA-CD\MCPA-CD\"
        Dim buPath As String = "\\MCPAFILESERVER\K Drive\Polygons\MCPA\Bill_BCC_MCPA-CD\MCPA-CD\backup\"
        Dim i As Integer
        Dim sFile As String = Nothing
        Dim newFile As String = Nothing
        Dim files As ReadOnlyCollection(Of String)

        'THIS MOVES THE FILES TO THE BACKUP FOLDER
        files = My.Computer.FileSystem.GetFiles(startPath, FileIO.SearchOption.SearchTopLevelOnly, "*.*")
        For i = 0 To files.Count - 1
            sFile = files.Item(i).ToString
            newFile = Path.GetFileNameWithoutExtension(sFile)

            If newFile.Contains("papoly") Then
                If Not MoveFiles(sFile, buPath & newFile) Then Return
            End If

            If newFile.Contains("OAS") And newFile IsNot "OAS" & fridayDate & ".dbf" Then
                If Not MoveFiles(sFile, buPath & newFile) Then Return
            End If
        Next
    End Sub

    Private Sub CopyCDFiles()
        Dim startPath As String = "\\MCPAFILESERVER\K Drive\Polygons\MCPA\Bill_BCC_MCPA-CD\MCPA-CD\"
        Dim endPath As String = "\\MCPAFILESERVER\K Drive\Mcpa\"
        Dim i As Integer
        Dim sFile As String = Nothing
        Dim newFile As String = Nothing
        Dim files As ReadOnlyCollection(Of String)

        'COPIES ALL PAPOLY FILES FROM BILL FOLDER TO CD FOLDER
        files = My.Computer.FileSystem.GetFiles(startPath, FileIO.SearchOption.SearchTopLevelOnly, "*.*")
        For i = 0 To files.Count - 1
            sFile = files.Item(i).ToString
            newFile = Path.GetFileNameWithoutExtension(sFile)

            If newFile.Contains("papoly") Then
                If Not CopyFiles(sFile, endPath & newFile, True) Then Return
            End If
        Next
    End Sub

    Private Sub ZipMCPAEXTRACTS()
        Dim startPath As String = "\\MCPAFILESERVER\K Drive\Polygons\MCPA\" & fridayDate & "\Datarecs" & fridayDate & "\MCPAEXTRACT\"
        Dim zipPathDest As String = "\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\MCPAEXTRACT.zip"

        If Not DeleteFiles(zipPathDest) Then Return

        ZipFile.CreateFromDirectory(startPath, zipPathDest, CompressionLevel.Optimal, False)
    End Sub

    Private Sub ZipForFTP_CD()
        Dim folderName As String = "\\MCPAFILESERVER\K Drive\Polygons\MCPA\Bill_BCC_MCPA-CD\MCPA-CD\"
        Dim zipFileOAS As String = "OAS" & fridayDate.Substring(2, fridayDate.Length - 2) & ".dbf"
        Dim zipPath1 As String = folderName & "parceldata.zip"
        Dim zipPath2 As String = "\\MCPAFILESERVER\K Drive\Mcpa\parceldata.zip"

        'create a new entry in a zip archive from an existing file and extract the archive contents
        'TO FTP FOR COUNTY
        Using archive As ZipArchive = ZipFile.Open(zipPath1, ZipArchiveMode.Update)
            'need to remove the old one
            For x As Integer = 0 To archive.Entries.Count - 1
                archive.Entries.Item(x).Delete()
            Next

            'this adds the new file to zip file
            archive.CreateEntryFromFile(folderName & zipFileOAS, zipFileOAS, CompressionLevel.Fastest)
        End Using

        'to MCPA CD
        Using archive As ZipArchive = ZipFile.Open(zipPath2, ZipArchiveMode.Update)
            'need to remove the old one
            For x As Integer = 0 To archive.Entries.Count - 1
                archive.Entries.Item(x).Delete()
            Next

            'this adds the new file to zip file
            archive.CreateEntryFromFile(folderName & zipFileOAS, zipFileOAS, CompressionLevel.Fastest)
        End Using
    End Sub

    Private Sub ZipMCPADATA()
        Dim folderName As String = "\\MCPAFILESERVER\K Drive\Polygons\MCPA\Bill_BCC_MCPA-CD\"
        Dim zipFile1 As String = "MCPADATA.dbf"
        Dim zipFile2 As String = "MCPADATA.cpg"
        Dim zipPath1 As String = folderName & "MCPADATA.zip"

        'create a new entry in a zip archive from an existing file and extract the archive contents
        Using archive As ZipArchive = ZipFile.Open(zipPath1, ZipArchiveMode.Update)
            'need to remove the old one
            For x As Integer = 0 To archive.Entries.Count - 1
                archive.Entries.Item(x).Delete()
            Next

            'this adds the new file to zip file
            archive.CreateEntryFromFile(folderName & zipFile1, zipFile1, CompressionLevel.Fastest)
            archive.CreateEntryFromFile(folderName & zipFile2, zipFile2, CompressionLevel.Fastest)
        End Using

        'COPIES THE ZIP FILE TO FTP FOR COUNTY
        Dim zipPathDest As String = "\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\MCPADATA.zip"

        If Not DeleteFiles(zipPathDest) Then Return
        If Not CopyFiles(zipPath1, zipPathDest, True) Then Return
    End Sub

    Private Sub ZipMCPAPOLYGONS()
        Dim folderName As String = "\\MCPAFILESERVER\K Drive\Polygons\MCPA\Bill_BCC_MCPA-CD\"
        Dim zipFile1 As String = "MCPAPOLYGONS.ADX"
        Dim zipFile2 As String = "MCPAPOLYGONS.cpg"
        Dim zipFile3 As String = "MCPAPOLYGONS.dbf"
        Dim zipFile4 As String = "MCPAPOLYGONS.prj"
        Dim zipFile5 As String = "MCPAPOLYGONS.sbn"
        Dim zipFile6 As String = "MCPAPOLYGONS.sbx"
        Dim zipFile7 As String = "MCPAPOLYGONS.shp"
        Dim zipFile8 As String = "MCPAPOLYGONS.shx"

        Dim zipPath1 As String = folderName & "MCPAPOLYGONS.zip"

        'create a new entry in a zip archive from an existing file and extract the archive contents
        Using archive As ZipArchive = ZipFile.Open(zipPath1, ZipArchiveMode.Update)
            'need to remove the old one

            For x As Integer = 0 To archive.Entries.Count - 1
                archive.Entries.Item(x).Delete()
            Next

            'this adds the new file to zip file
            archive.CreateEntryFromFile(folderName & zipFile1, zipFile1, CompressionLevel.Fastest)
            archive.CreateEntryFromFile(folderName & zipFile2, zipFile2, CompressionLevel.Fastest)
            archive.CreateEntryFromFile(folderName & zipFile3, zipFile3, CompressionLevel.Fastest)
            archive.CreateEntryFromFile(folderName & zipFile4, zipFile4, CompressionLevel.Fastest)
            archive.CreateEntryFromFile(folderName & zipFile5, zipFile5, CompressionLevel.Fastest)
            archive.CreateEntryFromFile(folderName & zipFile6, zipFile6, CompressionLevel.Fastest)
            archive.CreateEntryFromFile(folderName & zipFile7, zipFile7, CompressionLevel.Fastest)
            archive.CreateEntryFromFile(folderName & zipFile8, zipFile8, CompressionLevel.Fastest)
        End Using

        'COPIES THE ZIP FILE TO FTP FOR COUNTY
        Dim zipPathDest As String = "\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\MCPAPOLYGONS.zip"

        If Not DeleteFiles(zipPathDest) Then Return
        If Not CopyFiles(zipPath1, zipPathDest, True) Then Return
    End Sub

    Private Sub CopyForMapUpdate()
        Dim startDEST As String = "\\MCPAFILESERVER\K Drive\Polygons\MCPA\"
        Dim endDEST As String = "\\MERLIN\Merlin\MiscLayerData\Parcel\"
        Dim i, k As Integer
        Dim sFile As String = Nothing
        Dim newFile As String = Nothing
        Dim files As ReadOnlyCollection(Of String)

        'DELETE FILES FROM W:\MiscLayerData\Parcel 
        files = My.Computer.FileSystem.GetFiles(endDEST, FileIO.SearchOption.SearchTopLevelOnly, "*.*")
        For k = 0 To files.Count - 1
            sFile = files.Item(k).ToString
            newFile = Path.GetFileNameWithoutExtension(sFile)

            If newFile.Contains("2019qi") Or newFile.Contains("2019qv") Or newFile.Contains("2020qi") Or newFile.Contains("2020qv") _
            Or newFile.Contains("COMM") Or newFile.Contains("Exempt21") Or newFile.Contains("Freepapoly") Or newFile.Contains("Futurepapoly") _
            Or newFile.Contains("FutureFreepapoly") Or newFile.Contains("HXUndeliverable") Or newFile.Contains("LandModel") _
            Or newFile.Contains("lndchg20") Or newFile.Contains("lrateac") Or newFile.Contains("lratenoac") Or newFile.Contains("N Parcel") _
            Or newFile.Contains("NewAg") Or newFile.Contains("No_lndRTEchg_18_to_20_All") Or newFile.Contains("nooasis") _
            Or newFile.Contains("papoly") Or newFile.Contains("parcel") Or newFile.Contains("Permits") Or newFile.Contains("RoadType") _
            Or newFile.Contains("sinkhole") Or newFile.Contains("SubType") Or newFile.Contains("tang") Or newFile.Contains("Valchg20") Then
                If Not DeleteFiles(sFile) Then Return
            End If
        Next

        'Copying ALL files from Dated folder
        files = My.Computer.FileSystem.GetFiles(startDEST & fridayDate, FileIO.SearchOption.SearchTopLevelOnly, "*.*")
        For i = 0 To files.Count - 1
            sFile = files.Item(i).ToString
            newFile = Path.GetFileNameWithoutExtension(sFile)

            If newFile.Contains("2019qi") Or newFile.Contains("2019qv") Or newFile.Contains("2020qi") Or newFile.Contains("2020qv") _
            Or newFile.Contains("COMM") Or newFile.Contains("Exempt21") Or newFile.Contains("Freepapoly") Or newFile.Contains("Futurepapoly") _
            Or newFile.Contains("FutureFreepapoly") Or newFile.Contains("HXUndeliverable") Or newFile.Contains("LandModel") _
            Or newFile.Contains("lndchg20") Or newFile.Contains("lrateac") Or newFile.Contains("lratenoac") Or newFile.Contains("N Parcel") _
            Or newFile.Contains("NewAg") Or newFile.Contains("No_lndRTEchg_18_to_20_All") Or newFile.Contains("nooasis") _
            Or newFile.Contains("papoly") Or newFile.Contains("parcel") Or newFile.Contains("Permits") Or newFile.Contains("RoadType") _
            Or newFile.Contains("sinkhole") Or newFile.Contains("SubType") Or newFile.Contains("tang") Or newFile.Contains("Valchg20") Then
                If Not CopyFiles(sFile, endDEST & Path.GetFileName(sFile), False) Then Return
            End If
        Next

        If files.Count = 0 Then
            Console.WriteLine("No files available to copy")
        End If
    End Sub

    Private Sub CopyToJoe()
        Dim startDEST As String = "\\MCPAFILESERVER\K Drive\Polygons\MCPA\"
        Dim endDEST As String = "\\\MCPAFILESERVER\K Drive\Polygons\MCPA\2003_MassImpVacAprsl\SSS_JOE\FreeancePAPOLY\"
        Dim i As Integer
        Dim sFile As String = Nothing
        Dim newFile As String = Nothing
        Dim files As ReadOnlyCollection(Of String)

        'Copying ALL files from Dated folder
        files = My.Computer.FileSystem.GetFiles(startDEST & fridayDate, FileIO.SearchOption.SearchTopLevelOnly, "*.*")
        For i = 0 To files.Count - 1
            sFile = files.Item(i).ToString
            newFile = Path.GetFileNameWithoutExtension(sFile)

            If newFile.Contains("2019qi") Or newFile.Contains("2019qv") Or newFile.Contains("2020qi") Or newFile.Contains("2020qv") _
            Or newFile.Contains("COMM") Or newFile.Contains("Exempt21") Or newFile.Contains("Freepapoly") Or newFile.Contains("Futurepapoly") _
            Or newFile.Contains("FutureFreepapoly") Or newFile.Contains("LandModel") Or newFile.Contains("lndchg20") Or newFile.Contains("lrateac") _
            Or newFile.Contains("lratenoac") Or newFile.Contains("N Parcel") Or newFile.Contains("NewAg") Or newFile.Contains("No_lndRTEchg_18_to_20_All") _
            Or newFile.Contains("nooasis") Or newFile.Contains("papoly") Or newFile.Contains("parcel") Or newFile.Contains("Parcel_LNDUSE_G") _
            Or newFile.Contains("Permits") Or newFile.Contains("tang") Or newFile.Contains("Valchg20") Then
                If Not CopyFiles(sFile, endDEST & Path.GetFileName(sFile), True) Then Return
            End If
        Next

        If files.Count = 0 Then
            Console.WriteLine("No files available to copy")
        End If
    End Sub

#Region "MoveCopyDeleteVerify"
    Private Function CopyFiles(ByVal fromFolder As String, ByVal toFolder As String, overWriteFiles As Boolean) As Boolean
        If VerifyFileExist(fromFolder) = True Then
            If VerifyFileExist(toFolder) = False Then
                My.Computer.FileSystem.CopyFile(fromFolder, toFolder, overWriteFiles)
            Else
                Console.WriteLine("A file with the name " & fromFolder & " already exists.")
            End If
        Else
            MsgBox("The file " & fromFolder & " you are attempting to copy does not exist.")
            Return False
        End If

        Return True
    End Function

    Private Function MoveFiles(ByVal fromFolder As String, ByVal toFolder As String) As Boolean
        If VerifyFileExist(fromFolder) = True Then
            If VerifyFileExist(toFolder) = False Then
                My.Computer.FileSystem.MoveFile(fromFolder, toFolder)
            Else
                Console.WriteLine("A file with the name " & fromFolder & " already exists.")
            End If
        Else
            MsgBox("The file " & fromFolder & " you are attempting to move does not exist.")
            Return False
        End If

        Return True
    End Function

    Private Function DeleteFiles(ByVal fileName As String) As Boolean
        If VerifyFileExist(fileName) Then
            My.Computer.FileSystem.DeleteFile(fileName)
        Else
            MsgBox("The file " & fileName & " you are attempting to delete does not exist.")
            Return False
        End If

        Return True
    End Function

    Private Function VerifyFileExist(ByVal strFilePath As String) As Boolean
        Dim fileExists As Boolean
        fileExists = My.Computer.FileSystem.FileExists(strFilePath)

        Return fileExists
    End Function

    Private Function VerifyPathExist(ByVal strPath As String) As Boolean
        Dim folderExists As Boolean
        folderExists = My.Computer.FileSystem.DirectoryExists(strPath)

        Return folderExists
    End Function
#End Region

End Module
