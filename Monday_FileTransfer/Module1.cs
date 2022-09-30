using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Compression;
using Microsoft.VisualBasic;

namespace Monday_FileTransfer
{

    static class Module1
    {

        // Private fromFile As String = String.Empty
        // Private fromFolder As String = String.Empty '"C:\Users\ssherouse\Documents\MerlinDataExtracts\"
        private static string folderDate = DateTime.Now.Year + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0');
        private static string fridayDate;
        private static string datedfolder = @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\";
        private static string datedfolderWDate;

        public static void Main()
        {
            Console.Title = "MONDAY ROUTINE";
            Console.WriteLine("");
            Console.WriteLine("Please enter last Friday's Date (yyyymmdd)");
            fridayDate = Console.ReadLine().Trim();

            datedfolderWDate = datedfolder + fridayDate + @"\";

            Console.WriteLine();
            Console.WriteLine("What are you doing? (Pick a Number)");
            Console.WriteLine("  1. Beginning Routine  OR  2. Ending Routine");
            string choice = Console.ReadLine();
            if (choice.ToUpper() == "1")
            {
                // DO THIS AT THE START OF MONDAY ROUTINE

                // this copies files from rawdata\utility to datarecs 
                Console.WriteLine("Copying files from RawData to Datarecs" + fridayDate + " and Renaming the Extension to .txt");
                CopyExtracts();
                Console.WriteLine("Datarecs copy complete");

                // this zips mcpaextracts folder in datarecs and copies it to FTP for county
                Console.WriteLine("Now Zipping MCPAEXTRACTS....please wait");
                ZipMCPAEXTRACTS();
                Console.WriteLine("MCPAEXTRACTS copy complete");

                // this moves last weeks files to the back up forlder for the cd
                Console.WriteLine("Now moving County/CD files to backup folder....please wait");
                MoveCDFiles();
            }
            else
            {
                // DO AT END OF ROUTINE

                // this zips oas file and copies it to FTP for county and mcpa cd folder
                Console.WriteLine("Now Zipping PARCELDATA for County/CD....please wait");
                ZipForFTP_CD();

                // this copies all papoly files and copies them to mcpa cd folder
                Console.WriteLine("Now Copying papoly files from County folder to CD folder....please wait");
                CopyCDFiles();

                // this zips mcpadata folder in datarecs and copies it to FTP for county
                Console.WriteLine("Now Zipping MCPADATA for County FTP....please wait");
                ZipMCPADATA();

                // this zips mcpapolygons folder in datarecs and copies it to FTP for county
                Console.WriteLine("Now Zipping MCPAPOLYGONS for County FTP....please wait");
                ZipMCPAPOLYGONS();

                // this zips all futurepapoly files and copies it to FTP for county
                Console.WriteLine("Now Zipping FUTUREPAPOLY for County FTP....please wait");
                ZipToCountyFTP("Futurepapoly", datedfolderWDate, @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\");

                // this copies all specific files from dated folder to Merlin server for map update
                Console.WriteLine(@"Now Copying ALL files from Dated folder to W:\MiscLayerData\Parcel for Map Update....please wait");
                CopyForMapUpdate();

                // this copies all specific files from dated folder to Joes folder on the k drive
                Console.WriteLine(@"Now Copying ALL files from Dated folder to K:\Polygons\MCPA\2003_MassImpVacAprsl\SSS_JOE\FreeancePAPOLY....please wait");
                CopyToJoe();

                Console.WriteLine("Do you need to copy MONTHLY Files?  Y or N");
                string choice2 = Console.ReadLine();
                if (choice2.ToUpper() == "Y")
                {
                    // this copies all specific files mcpapolygons folder in datarecs and copies it to FTP for county
                    Console.WriteLine("Now Copying MONTHLY files....please wait");
                    CopyForMonthlyUpdate();

                    // this zips monthly files in unchanged_Polys folder and copies it to FTP for county
                    Console.WriteLine("Now Zipping Monthly files to County FTP....please wait");

                    ZipToCountyFTP("Blocks", @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\unchanged_Polys\", @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\");
                    ZipToCountyFTP("Condo", @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\unchanged_Polys\", @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\");
                    ZipToCountyFTP("County_Boundary", @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\unchanged_Polys\", @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\");
                    ZipToCountyFTP("Govtlot", @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\unchanged_Polys\", @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\");
                    ZipToCountyFTP("grants", @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\unchanged_Polys\", @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\");
                    ZipToCountyFTP("Lot_Replat", @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\unchanged_Polys\", @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\");
                    ZipToCountyFTP("lots", @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\unchanged_Polys\", @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\");
                    ZipToCountyFTP("MAP_INDEX", @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\unchanged_Polys\", @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\");
                    ZipToCountyFTP("millage_groups", @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\unchanged_Polys\", @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\");
                    ZipToCountyFTP("municipalities", @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\unchanged_Polys\", @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\");
                    ZipToCountyFTP("section", @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\unchanged_Polys\", @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\");
                    ZipToCountyFTP("subhist", @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\unchanged_Polys\", @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\");
                    ZipToCountyFTP("subbdy", @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\unchanged_Polys\", @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\");
                    ZipToCountyFTP("Township", @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\unchanged_Polys\", @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\");
                    ZipToCountyFTP("water", @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\unchanged_Polys\", @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\");

                    // ***ZiP these & copy only the zipped file to V:\MCPA
                    // Blocks
                    // Condo
                    // County_Boundary
                    // Govtlot
                    // grants
                    // Lot_Replat
                    // lots
                    // MAP_INDEX
                    // millage_groups
                    // municipalities
                    // section
                    // subhist
                    // subbdy
                    // Township
                    // water

                }
            }

            Console.WriteLine("Process Finished");
        }
        #region Zip Files
        private static void ZipMCPAEXTRACTS()
        {
            string startPath = datedfolder + fridayDate + @"\Datarecs" + fridayDate + @"\MCPAEXTRACT\";
            string zipPathDest = @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\MCPAEXTRACT.zip";

            if (!DeleteFiles(zipPathDest))
                return;

            ZipFile.CreateFromDirectory(startPath, zipPathDest, CompressionLevel.Optimal, false);
        }
        private static void ZipForFTP_CD()
        {
            string folderName = @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\Bill_BCC_MCPA-CD\MCPA-CD\";
            string zipFileOAS = "OAS" + fridayDate.Substring(2, fridayDate.Length - 2) + ".dbf";
            string zipPath1 = folderName + "parceldata.zip";
            string zipPath2 = @"\\MCPAFILESERVER\K Drive\Mcpa\parceldata.zip";

            // create a new entry in a zip archive from an existing file and extract the archive contents
            // TO FTP FOR COUNTY
            using (var archive = ZipFile.Open(zipPath1, ZipArchiveMode.Update))
            {
                // need to remove the old one
                for (int x = 0, loopTo = archive.Entries.Count - 1; x <= loopTo; x++)
                    archive.Entries[0].Delete();

                // this adds the new file to zip file
                archive.CreateEntryFromFile(folderName + zipFileOAS, zipFileOAS, CompressionLevel.Fastest);
            }

            // to MCPA CD
            using (var archive = ZipFile.Open(zipPath2, ZipArchiveMode.Update))
            {
                // need to remove the old one
                for (int x = 0, loopTo1 = archive.Entries.Count - 1; x <= loopTo1; x++)
                    archive.Entries[0].Delete();

                // this adds the new file to zip file
                archive.CreateEntryFromFile(folderName + zipFileOAS, zipFileOAS, CompressionLevel.Fastest);
            }
        }
        private static void ZipMCPADATA()
        {
            string folderName = @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\Bill_BCC_MCPA-CD\";
            string zipFile1 = "MCPADATA.dbf";
            string zipFile2 = "MCPADATA.cpg";
            string zipPath1 = folderName + "MCPADATA.zip";

            // create a new entry in a zip archive from an existing file and extract the archive contents
            using (var archive = ZipFile.Open(zipPath1, ZipArchiveMode.Update))
            {
                // need to remove the old one
                for (int x = 0, loopTo = archive.Entries.Count - 1; x <= loopTo; x++)
                    archive.Entries[0].Delete();

                // this adds the new file to zip file
                archive.CreateEntryFromFile(folderName + zipFile1, zipFile1, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(folderName + zipFile2, zipFile2, CompressionLevel.Fastest);
            }

            // COPIES THE ZIP FILE TO FTP FOR COUNTY
            string zipPathDest = @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\MCPADATA.zip";

            if (!DeleteFiles(zipPathDest))
                return;
            if (!CopyFiles(zipPath1, zipPathDest, true))
                return;
        }
        private static void ZipMCPAPOLYGONS()
        {
            string folderName = @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\Bill_BCC_MCPA-CD\";
            string zipFile1 = "MCPAPOLYGONS.ADX";
            string zipFile2 = "MCPAPOLYGONS.cpg";
            string zipFile3 = "MCPAPOLYGONS.dbf";
            string zipFile4 = "MCPAPOLYGONS.prj";
            string zipFile5 = "MCPAPOLYGONS.sbn";
            string zipFile6 = "MCPAPOLYGONS.sbx";
            string zipFile7 = "MCPAPOLYGONS.shp";
            string zipFile8 = "MCPAPOLYGONS.shx";

            string zipPath1 = folderName + "MCPAPOLYGONS.zip";

            // create a new entry in a zip archive from an existing file and extract the archive contents
            using (var archive = ZipFile.Open(zipPath1, ZipArchiveMode.Update))
            {
                // need to remove the old one

                for (int x = 0, loopTo = archive.Entries.Count - 1; x <= loopTo; x++)
                    archive.Entries[0].Delete();

                // this adds the new file to zip file
                archive.CreateEntryFromFile(folderName + zipFile1, zipFile1, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(folderName + zipFile2, zipFile2, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(folderName + zipFile3, zipFile3, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(folderName + zipFile4, zipFile4, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(folderName + zipFile5, zipFile5, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(folderName + zipFile6, zipFile6, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(folderName + zipFile7, zipFile7, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(folderName + zipFile8, zipFile8, CompressionLevel.Fastest);
            }

            // COPIES THE ZIP FILE TO FTP FOR COUNTY
            string zipPathDest = @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\MCPAPOLYGONS.zip";

            if (!DeleteFiles(zipPathDest))
                return;
            if (!CopyFiles(zipPath1, zipPathDest, true))
                return;
        }
        // Private Sub ZipFUTUREPAPOLY()
        // ' Dim folderName As String = datedfolder & fridayDate & "\"
        // Dim zipFile2 As String = "Futurepapoly.cpg"
        // Dim zipFile3 As String = "Futurepapoly.dbf"
        // Dim zipFile4 As String = "Futurepapoly.prj"
        // Dim zipFile5 As String = "Futurepapoly.sbn"
        // Dim zipFile6 As String = "Futurepapoly.sbx"
        // Dim zipFile7 As String = "Futurepapoly.shp"
        // Dim zipFile8 As String = "Futurepapoly.shx"
        // Dim zipPath1 As String = "\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\Futurepapoly.zip"

        // 'create a new entry in a zip archive from an existing file and extract the archive contents
        // Using archive As ZipArchive = ZipFile.Open(zipPath1, ZipArchiveMode.Update)
        // 'need to remove the old one

        // For x As Integer = 0 To archive.Entries.Count - 1
        // archive.Entries.Item(0).Delete()
        // Next

        // 'this adds the new file to zip file
        // archive.CreateEntryFromFile(datedfolderWDate & zipFile2, zipFile2, CompressionLevel.Fastest)
        // archive.CreateEntryFromFile(datedfolderWDate & zipFile3, zipFile3, CompressionLevel.Fastest)
        // archive.CreateEntryFromFile(datedfolderWDate & zipFile4, zipFile4, CompressionLevel.Fastest)
        // archive.CreateEntryFromFile(datedfolderWDate & zipFile5, zipFile5, CompressionLevel.Fastest)
        // archive.CreateEntryFromFile(datedfolderWDate & zipFile6, zipFile6, CompressionLevel.Fastest)
        // archive.CreateEntryFromFile(datedfolderWDate & zipFile7, zipFile7, CompressionLevel.Fastest)
        // archive.CreateEntryFromFile(datedfolderWDate & zipFile8, zipFile8, CompressionLevel.Fastest)
        // End Using
        // End Sub
        private static void ZipMONTHLY(string filename)
        {
            string zipFile2 = filename + ".cpg";
            string zipFile3 = filename + ".dbf";
            string zipFile4 = filename + ".prj";
            string zipFile5 = filename + ".sbn";
            string zipFile6 = filename + ".sbx";
            string zipFile7 = filename + ".shp";
            string zipFile8 = filename + ".shx";
            string zipPath1 = @"\\Mcpaserver6\ims\ApacheFTP\res\home\mcbcc\MCPA\" + filename + ".zip";

            // create a new entry in a zip archive from an existing file and extract the archive contents
            using (var archive = ZipFile.Open(zipPath1, ZipArchiveMode.Update))
            {
                // need to remove the old one

                for (int x = 0, loopTo = archive.Entries.Count - 1; x <= loopTo; x++)
                    archive.Entries[0].Delete();

                // this adds the new file to zip file
                archive.CreateEntryFromFile(datedfolderWDate + zipFile2, zipFile2, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(datedfolderWDate + zipFile3, zipFile3, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(datedfolderWDate + zipFile4, zipFile4, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(datedfolderWDate + zipFile5, zipFile5, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(datedfolderWDate + zipFile6, zipFile6, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(datedfolderWDate + zipFile7, zipFile7, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(datedfolderWDate + zipFile8, zipFile8, CompressionLevel.Fastest);
            }
        }

        private static void ZipToCountyFTP(string filename, string sourceFolder, string zipLoc)
        {
            // Dim wDrive As String = "\\MERLIN\Merlin\MiscLayerData\Parcel\Monthly updates"

            string zipFile1 = filename + ".cpg";
            string zipFile2 = filename + ".dbf";
            string zipFile3 = filename + ".prj";
            string zipFile4 = filename + ".sbn";
            string zipFile5 = filename + ".sbx";
            string zipFile6 = filename + ".shp";
            string zipFile7 = filename + ".shx";

            // create a new entry in a zip archive from an existing file and extract the archive contents
            using (var archive = ZipFile.Open(zipLoc + filename + ".zip", ZipArchiveMode.Update))
            {
                // need to remove the old one

                for (int x = 0, loopTo = archive.Entries.Count - 1; x <= loopTo; x++)
                    archive.Entries[0].Delete();

                // this adds the new file to zip file
                archive.CreateEntryFromFile(sourceFolder + zipFile1, zipFile1, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(sourceFolder + zipFile2, zipFile2, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(sourceFolder + zipFile3, zipFile3, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(sourceFolder + zipFile4, zipFile4, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(sourceFolder + zipFile5, zipFile5, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(sourceFolder + zipFile6, zipFile6, CompressionLevel.Fastest);
                archive.CreateEntryFromFile(sourceFolder + zipFile7, zipFile7, CompressionLevel.Fastest);
            }
        }

        #endregion

        #region CopyFiles

        private static void CopyExtracts()
        {
            int i;
            string sFile = null;
            string newFile = null;
            var workingfiles = new ArrayList();

            ReadOnlyCollection<string> files;
            files = My.MyProject.Computer.FileSystem.GetFiles(@"\\MCPAFILESERVER\K Drive\Polygons\RawData\Utility Files\" + fridayDate, Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, "*.*");
            var loopTo = files.Count - 1;
            for (i = 0; i <= loopTo; i++)
            {
                sFile = files[i].ToString();
                newFile = Path.GetFileNameWithoutExtension(sFile);

                // TO MCPAEXTRACTS    |  TO DATARECS FOR HEATHER
                // ADDRESS          |    COMM_SALES
                // DESC             |    EXEMPTION
                // LAND             |    LAND
                // MASTERPARCEL     |    LANDMODEL
                // MISCIMPR         |    MAPPING226
                // MOBILEHOME       |    MERLIN226TNG
                // NAME             |    PERMIT
                // PARENTPARCEL     |    SALESMAP
                // SITUS            |  


                if (Path.GetExtension(sFile).ToUpper() == ".CSV")
                {
                    if (newFile.Contains("AddressExtractView") | newFile.Contains("DescriptionExtractView") | newFile.Contains("MasterParcelExtractView") | newFile.Contains("MiscImprovementExtractView") | newFile.Contains("MobileHomeExtractView") | newFile.Contains("NameExtractView") | newFile.Contains("ParentParcelExtractView") | newFile.Contains("SitusAddressExtractView"))
                    {
                        if (!CopyFiles(sFile, datedfolder + fridayDate + @"\Datarecs" + fridayDate + @"\MCPAEXTRACT\" + newFile + ".txt", false))
                            return;
                    }
                    // this sends only to mcpaextract folder
                    else if (newFile.Contains("LandExtractView"))
                    {
                        if (!CopyFiles(sFile, datedfolder + fridayDate + @"\Datarecs" + fridayDate + @"\MCPAEXTRACT\" + newFile + ".txt", false))
                            return;
                        if (!CopyFiles(sFile, datedfolder + fridayDate + @"\Datarecs" + fridayDate + @"\" + newFile + ".txt", false))
                            return;
                    }
                    // this copies to datarecs and mcpaextract folder
                    else if (!CopyFiles(sFile, datedfolder + fridayDate + @"\Datarecs" + fridayDate + @"\" + newFile + ".txt", false))
                        // this copies to datarecs only
                        return;
                }
                else if (Path.GetExtension(sFile).ToUpper() == ".TXT")
                {
                    Path.GetFileName(sFile);
                    CopyFiles(sFile, datedfolder + fridayDate + @"\Datarecs" + fridayDate + @"\" + Path.GetFileName(sFile), false);
                }
            }

            if (files.Count == 0)
            {
                Console.WriteLine("No files available to copy");
            }
        }

        private static void MoveCDFiles()
        {
            string startPath = @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\Bill_BCC_MCPA-CD\MCPA-CD\";
            string buPath = @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\Bill_BCC_MCPA-CD\MCPA-CD\backup\";
            int i;
            string sFile = null;
            string newFile = null;
            ReadOnlyCollection<string> files;

            // 'THIS DELETES THE PAPOLY FILES FROM THE BACKUP FOLDER
            files = My.MyProject.Computer.FileSystem.GetFiles(buPath, Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, "*.*");
            var loopTo = files.Count - 1;
            for (i = 0; i <= loopTo; i++)
            {
                sFile = files[i].ToString();
                newFile = Path.GetFileNameWithoutExtension(sFile);

                if (newFile.Contains("papoly"))
                {
                    if (!DeleteFiles(buPath + Path.GetFileName(sFile)))
                        return;
                }
            }

            // THIS MOVES THE FILES TO THE BACKUP FOLDER
            files = My.MyProject.Computer.FileSystem.GetFiles(startPath, Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, "*.*");
            var loopTo1 = files.Count - 1;
            for (i = 0; i <= loopTo1; i++)
            {
                sFile = files[i].ToString();
                newFile = Path.GetFileNameWithoutExtension(sFile);

                if (newFile.Contains("papoly"))
                {
                    if (!MoveFiles(sFile, buPath + Path.GetFileName(sFile)))
                        return;
                }

                if (newFile.Contains("OAS") & !ReferenceEquals(newFile, "OAS" + fridayDate + ".dbf"))
                {
                    if (!MoveFiles(sFile, buPath + Path.GetFileName(sFile)))
                        return;
                }
            }
        }

        private static void CopyCDFiles()
        {
            string startPath = @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\Bill_BCC_MCPA-CD\MCPA-CD\";
            string endPath = @"\\MCPAFILESERVER\K Drive\Mcpa\";
            int i;
            string sFile = null;
            string newFile = null;
            ReadOnlyCollection<string> files;

            // COPIES ALL PAPOLY FILES FROM BILL FOLDER TO CD FOLDER
            files = My.MyProject.Computer.FileSystem.GetFiles(startPath, Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, "*.*");
            var loopTo = files.Count - 1;
            for (i = 0; i <= loopTo; i++)
            {
                sFile = files[i].ToString();
                newFile = Path.GetFileNameWithoutExtension(sFile);

                if (newFile.Contains("papoly"))
                {
                    if (!CopyFiles(sFile, endPath + Path.GetFileName(sFile), true))
                        return;
                }
            }
        }

        private static void CopyForMapUpdate()
        {
            // Dim startDEST As String = datedfolder
            string endDEST = @"\\MERLIN\Merlin\MiscLayerData\Parcel\";
            int i, k;
            string sFile = null;
            string newFile = null;
            ReadOnlyCollection<string> files;

            // DELETE FILES FROM W:\MiscLayerData\Parcel 
            files = My.MyProject.Computer.FileSystem.GetFiles(endDEST, Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, "*.*");
            var loopTo = files.Count - 1;
            for (k = 0; k <= loopTo; k++)
            {
                sFile = files[k].ToString();
                newFile = Path.GetFileNameWithoutExtension(sFile);

                if (newFile.Contains("2020qi") | newFile.Contains("2020qv") | newFile.Contains("2021qi") | newFile.Contains("2021qv") | newFile.Contains("COMM") | newFile.Contains("Exempt21") | newFile.Contains("Freepapoly") | newFile.Contains("Futurepapoly") | newFile.Contains("FutureFreepapoly") | newFile.Contains("LandModel") | newFile.Contains("lndchg21") | newFile.Contains("lrateac") | newFile.Contains("lratenoac") | newFile.Contains("N Parcel") | newFile.Contains("NewAg") | newFile.Contains("No_lndRTEchg_19_to_21_All") | newFile.Contains("nooasis") | newFile.Contains("papoly") | newFile.Contains("parcel") | newFile.Contains("tang") | newFile.Contains("Valchg21"))



                {
                    if (!DeleteFiles(sFile))
                        return;
                }
            }

            // Copying ALL files from Dated folder
            files = My.MyProject.Computer.FileSystem.GetFiles(datedfolder + fridayDate, Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, "*.*");
            var loopTo1 = files.Count - 1;
            for (i = 0; i <= loopTo1; i++)
            {
                sFile = files[i].ToString();
                newFile = Path.GetFileNameWithoutExtension(sFile);

                if (newFile.Contains("2020qi") | newFile.Contains("2020qv") | newFile.Contains("2021qi") | newFile.Contains("2021qv") | newFile.Contains("COMM") | newFile.Contains("Exempt21") | newFile.Contains("Freepapoly") | newFile.Contains("Futurepapoly") | newFile.Contains("FutureFreepapoly") | newFile.Contains("LandModel") | newFile.Contains("lndchg21") | newFile.Contains("lrateac") | newFile.Contains("lratenoac") | newFile.Contains("N Parcel") | newFile.Contains("NewAg") | newFile.Contains("No_lndRTEchg_19_to_21_All") | newFile.Contains("nooasis") | newFile.Contains("papoly") | newFile.Contains("parcel") | newFile.Contains("tang") | newFile.Contains("Valchg21"))



                {
                    if (!CopyFiles(sFile, endDEST + Path.GetFileName(sFile), false))
                        return;
                }
            }

            if (files.Count == 0)
            {
                Console.WriteLine("No files available to copy");
            }
        }
        private static void CopyToJoe()
        {
            string endDEST = @"\\\MCPAFILESERVER\K Drive\Polygons\MCPA\2003_MassImpVacAprsl\SSS_JOE\FreeancePAPOLY\";
            int i;
            string sFile = null;
            string newFile = null;
            ReadOnlyCollection<string> files;

            // Copying ALL files from Dated folder
            files = My.MyProject.Computer.FileSystem.GetFiles(datedfolder + fridayDate, Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, "*.*");
            var loopTo = files.Count - 1;
            for (i = 0; i <= loopTo; i++)
            {
                sFile = files[i].ToString();
                newFile = Path.GetFileNameWithoutExtension(sFile);

                if (newFile.Contains("2020qi") | newFile.Contains("2020qv") | newFile.Contains("2021qi") | newFile.Contains("2021qv") | newFile.Contains("COMM") | newFile.Contains("Exempt21") | newFile.Contains("Freepapoly") | newFile.Contains("Futurepapoly") | newFile.Contains("FutureFreepapoly") | newFile.Contains("LandModel") | newFile.Contains("lndchg21") | newFile.Contains("lrateac") | newFile.Contains("lratenoac") | newFile.Contains("N Parcel") | newFile.Contains("NewAg") | newFile.Contains("No_lndRTEchg_19_to_21_All") | newFile.Contains("nooasis") | newFile.Contains("papoly") | newFile.Contains("parcel") | newFile.Contains("Parcel_LNDUSE_G") | newFile.Contains("tang") | newFile.Contains("Valchg21"))




                {
                    if (!CopyFiles(sFile, endDEST + Path.GetFileName(sFile), true))
                        return;
                }
            }

            if (files.Count == 0)
            {
                Console.WriteLine("No files available to copy");
            }
        }
        private static void CopyForMonthlyUpdate()
        {
            string origLoc = @"\\MCPAFILESERVER\K Drive\Polygons\MCPA\unchanged_Polys\";
            string wDrive = @"\\MERLIN\Merlin\MiscLayerData\Parcel\Monthly updates\";
            string cdDrive = @"\\MCPAFILESERVER\K Drive\MCPA\";
            int i, k;
            string sFile = null;
            string newFile = null;
            ReadOnlyCollection<string> files;

            // DELETE FILES FROM W:\MiscLayerData\Parcel\Monthly updates
            files = My.MyProject.Computer.FileSystem.GetFiles(wDrive, Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, "*.*");
            var loopTo = files.Count - 1;
            for (k = 0; k <= loopTo; k++)
            {
                sFile = files[k].ToString();
                newFile = Path.GetFileNameWithoutExtension(sFile);

                if (newFile.Contains("BILLBOARDS") | newFile.Contains("Blocks") | newFile.Contains("Condo") | newFile.Contains("County_Boundary") | newFile.Contains("Future_Land_Use") | newFile.Contains("Govtlot") | newFile.Contains("grants") | newFile.Contains("Lot_Replat") | newFile.Contains("lots") | newFile.Contains("MAP_INDEX") | newFile.Contains("municipalities") | newFile.Contains("section") | newFile.Contains("subhist") | newFile.Contains("subbdy") | newFile.Contains("Township") | newFile.Contains("water"))

                {
                    if (!DeleteFiles(sFile))
                        return;
                }
            }

            // Copying ALL files from K:\Polygons\MCPA\unchanged_Polys\ to W:\MiscLayerData\Parcel\Monthly updates
            files = My.MyProject.Computer.FileSystem.GetFiles(origLoc, Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, "*.*");
            var loopTo1 = files.Count - 1;
            for (i = 0; i <= loopTo1; i++)
            {
                sFile = files[i].ToString();
                newFile = Path.GetFileNameWithoutExtension(sFile);

                if (newFile.Contains("BILLBOARDS") | newFile.Contains("Blocks") | newFile.Contains("Condo") | newFile.Contains("County_Boundary") | newFile.Contains("Future_Land_Use") | newFile.Contains("Govtlot") | newFile.Contains("grants") | newFile.Contains("Lot_Replat") | newFile.Contains("lots") | newFile.Contains("MAP_INDEX") | newFile.Contains("municipalities") | newFile.Contains("section") | newFile.Contains("subhist") | newFile.Contains("subbdy") | newFile.Contains("Township") | newFile.Contains("water"))

                {
                    if (!CopyFiles(sFile, wDrive + Path.GetFileName(sFile), false))
                        return;
                }
            }

            if (files.Count == 0)
            {
                Console.WriteLine("No files available to copy");
            }

            // DELETES FILES FROM K:\MCPA
            files = My.MyProject.Computer.FileSystem.GetFiles(cdDrive, Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, "*.*");
            var loopTo2 = files.Count - 1;
            for (k = 0; k <= loopTo2; k++)
            {
                sFile = files[k].ToString();
                newFile = Path.GetFileNameWithoutExtension(sFile);

                if (newFile.Contains("Blocks") | newFile.Contains("Condo") | newFile.Contains("County_Boundary") | newFile.Contains("Govtlot") | newFile.Contains("grants") | newFile.Contains("Lot_Replat") | newFile.Contains("lots") | newFile.Contains("MAP_INDEX") | newFile.Contains("millage_groups") | newFile.Contains("municipalities") | newFile.Contains("section") | newFile.Contains("subhist") | newFile.Contains("subbdy") | newFile.Contains("Township") | newFile.Contains("water"))

                {
                    if (!DeleteFiles(sFile))
                        return;
                }
            }

            // Copying ALL files to K:\MCPA
            files = My.MyProject.Computer.FileSystem.GetFiles(origLoc, Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, "*.*");
            var loopTo3 = files.Count - 1;
            for (i = 0; i <= loopTo3; i++)
            {
                sFile = files[i].ToString();
                newFile = Path.GetFileNameWithoutExtension(sFile);

                if (newFile.Contains("Blocks") | newFile.Contains("Condo") | newFile.Contains("County_Boundary") | newFile.Contains("Govtlot") | newFile.Contains("grants") | newFile.Contains("Lot_Replat") | newFile.Contains("lots") | newFile.Contains("MAP_INDEX") | newFile.Contains("millage_groups") | newFile.Contains("municipalities") | newFile.Contains("section") | newFile.Contains("subhist") | newFile.Contains("subbdy") | newFile.Contains("Township") | newFile.Contains("water"))

                {
                    if (!CopyFiles(sFile, cdDrive + Path.GetFileName(sFile), false))
                        return;
                }
            }


            // Copy all files from K:\Polygons\MCPA\unchanged_Polys\ to W:\MiscLayerData\Parcel\Monthly updates

            // ***ZiP these & copy only the zipped file to V:\MCPA
            // Blocks
            // Condo
            // County_Boundary
            // Future_Land_Use
            // Govtlot
            // grants
            // Lot_Replat
            // lots
            // MAP_INDEX
            // millage_groups
            // municipalities
            // section
            // subhist
            // subbdy
            // Township
            // water

            // Copy all files from K:\Polygons\MCPA\unchanged_Polys\ to K:\Mcpa
            // Blocks
            // Condo
            // County_Boundary
            // lots
            // section
            // subhist
            // subbdy
            // Township
            // water
        }

        #endregion

        #region MoveCopyDeleteVerify
        private static bool CopyFiles(string fromFolder, string toFolder, bool overWriteFiles)
        {
            if (VerifyFileExist(fromFolder) == true)
            {
                // If VerifyFileExist(toFolder) = False Then
                // If overWriteFiles = True Then
                My.MyProject.Computer.FileSystem.CopyFile(fromFolder, toFolder, overWriteFiles);
            }
            // Else
            // Console.WriteLine("A file with the name " & fromFolder & " already exists.")
            // End If
            else
            {
                Interaction.MsgBox("The file " + fromFolder + " you are attempting to copy does not exist.");
                return false;
            }

            return true;
        }

        private static bool MoveFiles(string fromFolder, string toFolder)
        {
            if (VerifyFileExist(fromFolder) == true)
            {
                if (VerifyFileExist(toFolder) == false)
                {
                    My.MyProject.Computer.FileSystem.MoveFile(fromFolder, toFolder);
                }
                else
                {
                    Console.WriteLine("A file with the name " + fromFolder + " already exists.");
                }
            }
            else
            {
                Interaction.MsgBox("The file " + fromFolder + " you are attempting to move does not exist.");
                return false;
            }

            return true;
        }

        private static bool DeleteFiles(string fileName)
        {
            if (VerifyFileExist(fileName))
            {
                File.Delete(fileName);
            }
            else
            {
                Interaction.MsgBox("The file " + fileName + " you are attempting to delete does not exist.");
                return false;
            }

            return true;
        }

        private static bool VerifyFileExist(string strFilePath)
        {
            bool fileExists;
            fileExists = File.Exists(strFilePath);

            return fileExists;
        }

        private static bool VerifyPathExist(string strPath)
        {
            bool folderExists;
            folderExists = Directory.Exists(strPath);

            return folderExists;
        }
        #endregion

    }
}