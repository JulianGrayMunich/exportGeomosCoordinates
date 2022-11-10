using databaseAPI;
using GNAgeneraltools;
using GNAspreadsheettools;


using OfficeOpenXml;
using System.Data.SqlClient;
using System.Configuration;

using Microsoft.Win32;
using static OfficeOpenXml.ExcelErrorValue;
using System.Globalization;



namespace CoordinateExporter
{
    class Program
    {
        public static void Main(string[] args)
        {

#pragma warning disable CS0162
#pragma warning disable CS0164
#pragma warning disable CS0168
#pragma warning disable CS8600
#pragma warning disable CS8601
#pragma warning disable CS8602
#pragma warning disable CS8604
#pragma warning disable CA1416

            gnaTools gnaT = new gnaTools();
            dbAPI gnaDBAPI = new dbAPI();
            spreadsheetAPI gnaSpreadsheetAPI = new spreadsheetAPI();

            //==== Console settings
            Console.OutputEncoding = System.Text.Encoding.Unicode;
            CultureInfo culture;

            //==== System config variables
            string strSoftwareLicenseTag = ConfigurationManager.AppSettings["SoftwareLicenseTag"];

            string strMonitoringSystemsName = ConfigurationManager.AppSettings["MonitoringSystemsName"];


            string strDBconnection = ConfigurationManager.ConnectionStrings["DBconnectionString"].ConnectionString;
            string strProjectTitle = ConfigurationManager.AppSettings["ProjectTitle"];
            string strExcelPath = ConfigurationManager.AppSettings["ExcelPath"];
            string strExcelFile = ConfigurationManager.AppSettings["ExcelFile"];
            string strFTPSubdirectory = ConfigurationManager.AppSettings["FTPSubdirectory"];
            string strReferenceWorksheet = ConfigurationManager.AppSettings["ReferenceWorksheet"];
            string strSurveyWorksheet = ConfigurationManager.AppSettings["SurveyWorksheet"];
            string strFirstDataRow = ConfigurationManager.AppSettings["FirstDataRow"];

            string strCheckWorksheetsExist = ConfigurationManager.AppSettings["checkWorksheetsExist"];


            string strPrepareCoordinateExportWorkbook = ConfigurationManager.AppSettings["PrepareCoordinateExportWorkbook"];
            string strCoordinateOrder = ConfigurationManager.AppSettings["CoordinateOrder"];
            string strOutputFileExtension = ConfigurationManager.AppSettings["OutputFileExtension"];
            string strCSVseparator = ConfigurationManager.AppSettings["CSVseparator"];
            string strCSVformat = ConfigurationManager.AppSettings["CSVformat"];
            string strIncludeHeader = ConfigurationManager.AppSettings["includeHeader"];
            string strReplacementNames = ConfigurationManager.AppSettings["ReplacementNames"];
            string strIncludeToRdata = ConfigurationManager.AppSettings["includeToRdata"];

            string strManualBlockStart = ConfigurationManager.AppSettings["manualBlockStart"];
            string strManualBlockEnd = ConfigurationManager.AppSettings["manualBlockEnd"];

            string strHistoricBlockStart = ConfigurationManager.AppSettings["historicBlockStart"];
            string strHistoricBlockEnd = ConfigurationManager.AppSettings["historicBlockEnd"];

            string strTimeBlockType = ConfigurationManager.AppSettings["TimeBlockType"];
            string strTimeOffsetHrs = ConfigurationManager.AppSettings["TimeOffsetHrs"];
            string strBlockSizeHrs = ConfigurationManager.AppSettings["BlockSizeHrs"];

            string strSendEmails = ConfigurationManager.AppSettings["SendEmails"];
            string strEmailLogin = ConfigurationManager.AppSettings["EmailLogin"];
            string strEmailPassword = ConfigurationManager.AppSettings["EmailPassword"];

            string strExitFlag = "";

            Coordinates[] coordinate = new Coordinates[5000];       // this class is defined in spreadsheetAPS.cs
            string[,] strSensorID = new string[5000, 2];
            string[,] strPointDeltas = new string[5000, 2];
            string[] strCSVdata = new string[1000000];

            string strTimeStart = "";
            string strTimeBlockStartLocal = "";
            string strTimeBlockEndLocal = "";
            string strTimeBlockStartUTC = "";
            string strTimeBlockEndUTC = "";

            string strMasterWorkbookFullPath = strExcelPath + strExcelFile;
            string strOutputFile = strFTPSubdirectory + strProjectTitle.Replace(" ", "_") + "_" + DateTime.UtcNow.ToString("yyyyMMdd_HHmm") + "." + strOutputFileExtension;

            int idErow = 2;   // the row for dE in the reference worksheet
            int idEcol = 6;   // the col for dE in the reference worksheet
            int iRowCount = 0;
            int iHistoricBlockCounter = 0;
            int iBlockSizeHrs = Convert.ToInt16(strBlockSizeHrs);

            double Eref = 0.0;
            double Nref = 0.0;
            double Href = 0.0;

            int iStart = Convert.ToInt32(strFirstDataRow);
            int iObservationCounter = 0;
            int iCSVcounter = 0;
            string strObservationFlag = "No";       // When a new coordinate is found then this changes to Yes and the CSV file is created

            string strPointName = " ";
            string strHeaderLine = " ";
            string strMaxwell1 = " ";
            string strMaxwell2 = " ";
            double dN, dE, dH;
            double dblN = 0.0;
            double dblE = 0.0;
            double dblH = 0.0;
            double dblToRoffset, dblToR;








            //==== Set the EPPlus license
            ExcelPackage.LicenseContext = LicenseContext.Commercial;


            //==== [Main program]===========================================================================================

            gnaT.WelcomeMessage("exportGeomosCoordinates 20221110");
            gnaT.checkLicenseValidity(strSoftwareLicenseTag, strProjectTitle, strEmailLogin, strEmailPassword, strSendEmails);


            //==== Environment check

            Console.WriteLine("");
            Console.WriteLine("1. Check system environment");
            Console.WriteLine("     Project: " + strProjectTitle);
            Console.WriteLine("     Master workbook: " + strMasterWorkbookFullPath);

            gnaDBAPI.testDBconnection(strDBconnection);

            if (strCheckWorksheetsExist == "Yes")
            {
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strReferenceWorksheet);
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strSurveyWorksheet);
            }
            else
            {
                Console.WriteLine("     Existance of workbook & worksheets is not checked");
            }


            // Generate the current time stamp 


            switch (strTimeBlockType)
            {
                case ("Manual"):
                    strTimeBlockStartUTC = gnaT.convertLocalToUTC(strManualBlockStart);
                    strTimeBlockEndUTC = gnaT.convertLocalToUTC(strManualBlockEnd);
                    strTimeBlockStartLocal = strManualBlockStart.Replace("'", "").Trim();
                    strTimeBlockStartLocal = " '" + strTimeBlockStartLocal + ":00' ";
                    strTimeBlockEndLocal = strManualBlockEnd.Replace("'", "").Trim();
                    strTimeBlockEndLocal = " '" + strTimeBlockEndLocal + ":00' ";
                    break;
                case ("Schedule"):
                    double dblStartTimeOffset = -1.0 * (Convert.ToDouble(strTimeOffsetHrs));
                    double dblEndTimeOffset = dblStartTimeOffset - (Convert.ToDouble(strBlockSizeHrs));
                    strTimeBlockStartLocal = " '" + DateTime.Now.AddHours(dblEndTimeOffset).ToString("yyyy-MM-dd HH:mm:ss") + "' ";
                    strTimeBlockEndLocal = " '" + DateTime.Now.AddHours(dblStartTimeOffset).ToString("yyyy-MM-dd HH:mm:ss") + "' ";
                    strTimeBlockStartUTC = gnaT.convertLocalToUTC(strTimeBlockStartLocal);
                    strTimeBlockEndUTC = gnaT.convertLocalToUTC(strTimeBlockEndLocal);
                    break;
                default:
                    dblStartTimeOffset = -1.0 * (Convert.ToDouble(strTimeOffsetHrs));
                    dblEndTimeOffset = dblStartTimeOffset - (Convert.ToDouble(strBlockSizeHrs));
                    strTimeBlockStartLocal = " '" + DateTime.Now.AddHours(dblEndTimeOffset).ToString("yyyy-MM-dd HH:mm:ss") + "' ";
                    strTimeBlockEndLocal = " '" + DateTime.Now.AddHours(dblStartTimeOffset).ToString("yyyy-MM-dd HH:mm:ss") + "' ";
                    strTimeBlockStartUTC = gnaT.convertLocalToUTC(strTimeBlockStartLocal);
                    strTimeBlockEndUTC = gnaT.convertLocalToUTC(strTimeBlockEndLocal);
                    break;
            }


            DateTime now = DateTime.UtcNow;
            string strDateTime = DateTime.Now.ToString("yyyyMMdd_HHmm");
            string strDateTimeUTC = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm");   //2022-07-26 13:45:15
            string strTimeNowUTC = "'" + DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm") + "'";
            string strDataTimeStamp = strDateTimeUTC;

            Console.WriteLine("\n   Time block: " + strTimeBlockType);
            Console.WriteLine("     " + strTimeBlockStartLocal.Replace("'", "").Trim() + " (local)");
            Console.WriteLine("     " + strTimeBlockEndLocal.Replace("'", "").Trim() + " (local)");
            Console.WriteLine("");

            Console.WriteLine("2. Extract point names");
            string[] strPointNames = gnaSpreadsheetAPI.readPointNames(strMasterWorkbookFullPath, strSurveyWorksheet, strFirstDataRow);

            Console.WriteLine("3. Extract SensorID");
            strSensorID = gnaDBAPI.getGeomosSensorID(strDBconnection, strPointNames, strMonitoringSystemsName);

            Console.WriteLine("4. Write SensorID to workbook");
            gnaSpreadsheetAPI.writeSensorID(strMasterWorkbookFullPath, strSurveyWorksheet, strSensorID, strFirstDataRow);

            if (strPrepareCoordinateExportWorkbook == "Yes")
            {
                // coordinates are read for a timeblock and written to the reference worksheet to enable the corrections to be computed

                Console.WriteLine("5. Prepare reference deltas");
                Console.WriteLine("       Extract mean deltas from DB");
                strTimeBlockStartUTC = gnaT.convertLocalToUTC(strManualBlockStart);
                strTimeBlockEndUTC = gnaT.convertLocalToUTC(strManualBlockEnd);
                strPointDeltas = gnaDBAPI.getGeomosMeanDeltas(strDBconnection, strMonitoringSystemsName, strTimeBlockStartUTC, strTimeBlockEndUTC, strSensorID);
                Console.WriteLine("       Write mean deltas to worksheet");
                gnaSpreadsheetAPI.writeDeltas(strMasterWorkbookFullPath, strReferenceWorksheet, strPointDeltas, idErow, idEcol, strTimeBlockStartUTC, strTimeBlockEndUTC, strCoordinateOrder);
            }
            else
            {
                do
                {

                    if (strTimeBlockType == "Historic")
                    {
                        iHistoricBlockCounter++;
                        var answer = gnaT.generateHistoricTimeBlockStartEnd(strHistoricBlockStart, strHistoricBlockEnd, iBlockSizeHrs, iHistoricBlockCounter);
                        strTimeBlockStartUTC = answer.Item1;
                        strTimeBlockEndUTC = answer.Item2;

                        string strTemp = strTimeBlockStartUTC.Replace("'", "");
                        strTemp = strTemp.Trim();
                        strTemp = strTemp.Replace(":", "");
                        strTemp = strTemp.Replace("/", "");
                        strTemp = strTemp.Replace("-", "");
                        strTemp = strTemp.Replace(" ", "_");
                        strOutputFile = strFTPSubdirectory + strProjectTitle.Replace(" ", "_") + "_" + strTemp + "." + strOutputFileExtension;

                        DateTime HistoricTimeblockEnd = DateTime.Parse(strHistoricBlockEnd.Replace("'", ""), System.Globalization.CultureInfo.InvariantCulture);
                        DateTime currentTimeblockStart = DateTime.Parse(strTimeBlockStartUTC.Replace("'", ""), System.Globalization.CultureInfo.InvariantCulture);
                        DateTime currentTimeblockEnd = DateTime.Parse(strTimeBlockEndUTC.Replace("'", ""), System.Globalization.CultureInfo.InvariantCulture);

                        if (currentTimeblockStart > HistoricTimeblockEnd)
                        { goto Finish; }
                        else if (currentTimeblockEnd > HistoricTimeblockEnd)
                        {
                            strTimeBlockEndUTC = strHistoricBlockEnd;
                        }

                        strTemp = ("  " + strTimeBlockStartUTC + " to " + strTimeBlockEndUTC).Replace("'", "");
                        Console.WriteLine(strTemp);

                    }

                    FileInfo newFile = new FileInfo(strMasterWorkbookFullPath);

                    int iRow = 0;
                    strPointName = "blank";

                    using (ExcelPackage package = new ExcelPackage(newFile))
                    {

                        ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strReferenceWorksheet];

                        Console.WriteLine("6. Extract reference values from Reference worksheet");
                        iRow = Convert.ToInt16(strFirstDataRow);
                        int iLastRow = -1;
                        strPointName = "blank";
                        // Extract data from Reference worksheet
                        do
                        {
                            // Read in strPointData: 1:SensorID, 2:Name, 3:Eref,4:Nref,5:Href,
                            // 9:dEcorr, 10:dNcorr, 11:dHcorr, 15:Replacement Name, 16:TimeStamp,Type
                            iLastRow++;
                            coordinate[iLastRow] = new Coordinates();
                            coordinate[iLastRow].SensorID = Convert.ToString(namedWorksheet.Cells[iRow, 1].Value);
                            coordinate[iLastRow].Name = Convert.ToString(namedWorksheet.Cells[iRow, 2].Value);
                            coordinate[iLastRow].Eref = Convert.ToDouble(namedWorksheet.Cells[iRow, 3].Value);
                            coordinate[iLastRow].Nref = Convert.ToDouble(namedWorksheet.Cells[iRow, 4].Value);
                            coordinate[iLastRow].Href = Convert.ToDouble(namedWorksheet.Cells[iRow, 5].Value);
                            coordinate[iLastRow].dEcorr = Convert.ToDouble(namedWorksheet.Cells[iRow, 9].Value);
                            coordinate[iLastRow].dNcorr = Convert.ToDouble(namedWorksheet.Cells[iRow, 10].Value);
                            coordinate[iLastRow].dHcorr = Convert.ToDouble(namedWorksheet.Cells[iRow, 11].Value);
                            coordinate[iLastRow].ToRoffset = Convert.ToDouble(namedWorksheet.Cells[iRow, 15].Value);
                            coordinate[iLastRow].ReplacementName = Convert.ToString(namedWorksheet.Cells[iRow, 17].Value);
                            coordinate[iLastRow].Type = Convert.ToString(namedWorksheet.Cells[iRow, 19].Value);
                            coordinate[iLastRow].Timestamp = Convert.ToString(namedWorksheet.Cells[iRow, 41].Value);

                            iRow++;
                        } while (coordinate[iLastRow].Name != "");

                        //iLastRow--;
                        iRowCount = iLastRow;
                        coordinate[iLastRow].Name = "TheEnd";   // Point Name

                        string strStart = coordinate[0].Timestamp;
                        strPointName = "blank";
                        strHeaderLine = "";
                        int iIndex = 0;

                        // Step through each strPointName and extract the deltas.
                        // assign the results to strCSVdata[iObservationCounter]

                        do
                        {
                            if (strReplacementNames == "Yes")
                            {
                                strPointName = coordinate[iIndex].ReplacementName;
                            }
                            else
                            {
                                strPointName = coordinate[iIndex].Name;
                            }

                            if ((coordinate[iIndex].Name == "TheEnd") || (coordinate[iIndex].SensorID == "Missing")) goto Jump;

                            if (strTimeBlockType == "Manual")
                            {
                                strTimeBlockStartUTC = gnaT.convertLocalToUTC(strManualBlockStart);
                                strTimeBlockEndUTC = gnaT.convertLocalToUTC(strManualBlockEnd);
                            }
                            else if (strTimeBlockType == "Schedule")
                            {
                                strTimeBlockStartUTC = "'" + coordinate[iIndex].Timestamp + "'";
                                strTimeBlockEndUTC = strTimeNowUTC;
                            }

                            Eref = coordinate[iIndex].Eref;
                            Nref = coordinate[iIndex].Nref;
                            Href = coordinate[iIndex].Href;

                            strTimeStart = coordinate[iIndex].Timestamp;
                            string strType = coordinate[iIndex].Type;
                            string strCSVstring = "";

                            string[,] strDeltas = new string[5000, 4]; //PointName, dN, dE, dH, dateTime

                            SqlConnection conn = null;
                            conn = new SqlConnection(strDBconnection);
                            conn.Open();

                            //string SensorID = coordinate[iIndex].SensorID;

                            if (strTimeBlockStartUTC == "") strTimeBlockStartUTC = "'2022-01-01 00:00'";

                            //extract the deltas from the DB between strTimeStart and strTimeNow

                            try
                            {
                                string SQLaction = @"
                                SELECT * FROM dbo.Results
                                WHERE Results.PointID = @PointID
                                AND Results.Type = 0 
                                AND Results.Epoch BETWEEN " 
                                + strTimeBlockStartUTC + " AND " + strTimeBlockEndUTC +
                                " ORDER BY Results.Epoch";

                                string strTemp = SQLaction;
                                SqlCommand cmd = new SqlCommand(SQLaction, conn);

                                // define the parameter used in the command object and add to the command
                                cmd.Parameters.Add(new SqlParameter("@PointID", coordinate[iIndex].SensorID));

                                // Define the data reader
                                SqlDataReader dataReader = cmd.ExecuteReader();

                                // Create the header line in case it is needed
                                switch (strCSVformat)
                                {
                                    case "Datum":
                                        //,LP14,,,,06:54.5,,5000.054069,1044.349157,93.40703182,,,,,,,,,,,,,,,,,,,,,,,,,
                                        if (strCoordinateOrder == "NEH")
                                        {
                                            strHeaderLine = "blank,Name,blank,blank,blank,UTCtime,blank,N,E,H";
                                        }
                                        else
                                        {
                                            strHeaderLine = "blank,Name,blank,blank,blank,UTCtime,blank,E,N,H";
                                        }
                                        break;
                                    case "Dywidag":
                                        //,LP14,,,,06:54.5,,5000.054069,1044.349157,93.40703182,,,,,,,,,,,,,,,,,,,,,,,,,
                                        if (strCoordinateOrder == "NEH")
                                        {
                                            strHeaderLine = "blank,Name,blank,blank,blank,UTCtime,blank,N,E,H";
                                        }
                                        else
                                        {
                                            strHeaderLine = "blank,Name,blank,blank,blank,UTCtime,blank,E,N,H";
                                        }
                                        break;
                                    case "Standard":
                                        // Write to CSV file: Standard format, no spaces
                                        if (strCoordinateOrder == "NEH")
                                        {
                                            strHeaderLine = "Name,UTCtime,N,E,H,Type";
                                        }
                                        else
                                        {
                                            strHeaderLine = "Name,UTCtime,E,N,H,Type";
                                        }
                                        break;
                                    case "MissionOS":
                                        // Write to CSV file: Standard format, no spaces
                                        if (strCoordinateOrder == "NEH")
                                        {
                                            strHeaderLine = "Name,UTCtime,Nref,Eref,Href,dN,dE,dH,N,E,H,Type";
                                        }
                                        else
                                        {
                                            strHeaderLine = "Name,UTCtime,Eref,Nref,Href,dE,dN,dH,E,N,H,Type";
                                        }
                                        break;
                                    default:
                                        // Write to CSV file: Standard format, no spaces
                                        if (strCoordinateOrder == "NEH")
                                        {
                                            strHeaderLine = "Name,UTCtime,N,E,H";
                                        }
                                        else
                                        {
                                            strHeaderLine = "Name,UTCtime,E,N,H";
                                        }
                                        break;
                                }

                                if (strIncludeToRdata == "Yes")
                                {
                                    strHeaderLine = strHeaderLine + ",Prism Offset,Top of Rail";
                                }

                                // Now read through the results and assign them to the strCSVdata[iObservationCounter] array

                                iCSVcounter = 0;
                                while (dataReader.Read())
                                {
                                    coordinate[iIndex].Timestamp = strDataTimeStamp;
                                    dN = Math.Round(Convert.ToDouble(dataReader["dN"]) + coordinate[iIndex].dNcorr, 4);
                                    dE = Math.Round(Convert.ToDouble(dataReader["dE"]) + coordinate[iIndex].dEcorr, 4);
                                    dH = Math.Round(Convert.ToDouble(dataReader["dH"]) + coordinate[iIndex].dHcorr, 4);
                                    dblN = Nref + dN;
                                    dblE = Eref + dE;
                                    dblH = Href + dH;
                                    dblToRoffset = coordinate[iIndex].ToRoffset;
                                    dblToR = Math.Round(dblH + dblToRoffset, 4);

                                    if (strReplacementNames == "Yes")
                                    {
                                        strPointName = coordinate[iIndex].ReplacementName;
                                    }
                                    else
                                    {
                                        strPointName = coordinate[iIndex].Name;
                                    }

                                    // Maxwell components
                                    if (strCoordinateOrder == "NEH")
                                    {
                                        strMaxwell1 = Nref.ToString("0.0000") + strCSVseparator + Eref.ToString("0.0000") + strCSVseparator + Href.ToString("0.0000");
                                        strMaxwell2 = dN.ToString("0.0000") + strCSVseparator + dE.ToString("0.0000") + strCSVseparator + dH.ToString("0.0000");
                                        strMaxwell1 = strMaxwell1 + "," + strMaxwell2;
                                        strMaxwell2 = dblN.ToString("0.0000") + strCSVseparator + dblE.ToString("0.0000") + strCSVseparator + dblH.ToString("0.0000") + strCSVseparator + strType;
                                    }
                                    else
                                    {
                                        strMaxwell1 = Eref.ToString("0.0000") + strCSVseparator + Nref.ToString("0.0000") + strCSVseparator + Href.ToString("0.0000");
                                        strMaxwell2 = dE.ToString("0.0000") + strCSVseparator + dN.ToString("0.0000") + strCSVseparator + dH.ToString("0.0000");
                                        strMaxwell1 = strMaxwell1 + "," + strMaxwell2;
                                        strMaxwell2 = dblE.ToString("0.0000") + strCSVseparator + dblN.ToString("0.0000") + strCSVseparator + dblH.ToString("0.0000") + strCSVseparator + strType;
                                    }

                                    // name, timestamp enh
                                    culture = CultureInfo.CreateSpecificCulture("en-GB");

                                    string strTimeStamp = "";

                                    if (strTimeBlockType == "Manual")
                                    {
                                        strTimeStamp = strManualBlockEnd;
                                    }
                                    else if (strTimeBlockType == "Schedule")
                                    {
                                        strTimeStamp = strDataTimeStamp;
                                    }
                                    else if (strTimeBlockType == "Historic")
                                    {
                                        strTimeStamp = Convert.ToString(dataReader["EndTimeUTC"]);
                                        strTimeStamp.Trim();
                                        strObservationFlag = "Yes";
                                    }

                                    strTimeStamp.Trim();
                                    if ((strTimeStamp.Substring(2, 1) == "/") || (strTimeStamp.Substring(3, 1) == "/"))
                                    {
                                        strTimeStamp = gnaT.formatTimestampMissionOS(strTimeStamp);
                                    };


                                    string strObservationTime = Convert.ToString(dataReader["EndTimeUTC"]).Trim();

                                    if ((strObservationTime.Substring(2, 1) == "/") || (strObservationTime.Substring(3, 1) == "/"))
                                    {
                                        strObservationTime = gnaT.formatTimestampMissionOS(strObservationTime);
                                    };

                                    if (strCoordinateOrder == "NEH")
                                    {
                                        strCSVstring = dblN.ToString("0.0000") + strCSVseparator + dblE.ToString("0.0000") + strCSVseparator + dblH.ToString("0.0000");
                                    }
                                    else
                                    {
                                        strCSVstring = dblE.ToString("0.0000") + strCSVseparator + dblN.ToString("0.0000") + strCSVseparator + dblH.ToString("0.0000");
                                    }

                                    switch (strCSVformat)
                                    {
                                        case "Datum":
                                            //,LP14,,,,06:54.5,,5000.054069,1044.349157,93.40703182,,,,,,,,,,,,,,,,,,,,,,,,,
                                            strCSVstring = "," + strPointName + ",,,," + strObservationTime + ",," + strCSVstring;
                                            break;
                                        case "Dywidag":
                                            //,LP14,,,,06:54.5,,5000.054069,1044.349157,93.40703182,,,,,,,,,,,,,,,,,,,,,,,,,
                                            strCSVstring = "," + strPointName + ",,,," + strObservationTime + ",," + strCSVstring;
                                            break;
                                        case "Standard":
                                            // Write to CSV file: Standard format, no spaces
                                            strCSVstring = strPointName + strCSVseparator + strObservationTime + strCSVseparator + strCSVstring + strCSVseparator + strType;
                                            break;
                                        case "MissionOS":
                                            // Write to CSV file: Standard format, no spaces
                                            strCSVstring = strPointName + strCSVseparator + strObservationTime + strCSVseparator + strMaxwell1 + strCSVseparator + strMaxwell2;
                                            break;
                                        default:
                                            // Write to CSV file: Standard format, no spaces
                                            strCSVstring = strPointName + strCSVseparator + strObservationTime + strCSVseparator + strCSVstring;
                                            break;
                                    }

                                    if (strIncludeToRdata == "Yes")
                                    {
                                        strCSVstring = strCSVstring + strCSVseparator + dblToRoffset.ToString("0.0000") + strCSVseparator + dblToR.ToString("0.0000");
                                    }

                                    strCSVdata[iObservationCounter] = strCSVstring;
                                    iObservationCounter++;
                                    iCSVcounter++;
                                }

                                // Close the dataReader
                                if (dataReader != null)
                                {
                                    dataReader.Close();
                                }
                            }
                            catch (System.Data.SqlClient.SqlException ex)
                            {
                                Console.WriteLine("getPointDeltas: DB Connection Failed: ");
                                Console.WriteLine(ex);
                                Console.ReadKey();
                            }

                            finally
                            {
                                conn.Dispose();
                                conn.Close();
                            }

                            if (iCSVcounter > 0)
                            {
                                strObservationFlag = "Yes";                 // a new coordinate has been found so a CSV file must be created
                            }

Jump:
                            iIndex++;

                            strExitFlag = coordinate[iIndex].Name;

                        } while (strExitFlag != "TheEnd");

                    }

                    // Write the results for this point to the CSV file

                    if (strObservationFlag == "Yes")
                    {
                        Console.WriteLine("8. Create coordinate CSV file " + strOutputFile);

                        if (!File.Exists(strOutputFile))
                        {
                            string strFileName = strOutputFile;

                            using (StreamWriter writetext = new StreamWriter(strFileName, false))
                            {
                                if (strIncludeHeader == "Yes") writetext.WriteLine(strHeaderLine);

                                for (int i = 0; i < iObservationCounter; i++)
                                {
                                    writetext.WriteLine(strCSVdata[i]);
                                }
                                writetext.Close();
                            }
                        }


                        if (strTimeBlockType == "Schedule")
                        {
                            Console.WriteLine("9. Update the time stamp in the master spreadsheet..");
                            // Update the timestamp in the spreadsheet

                            FileInfo newFile2 = new FileInfo(strMasterWorkbookFullPath);

                            using (ExcelPackage package = new ExcelPackage(newFile2))
                            {

                                ExcelWorksheet namedWorksheet = package.Workbook.Worksheets[strReferenceWorksheet];

                                iRow = Convert.ToInt16(strFirstDataRow);
                                for (int i = 0; i < iRowCount; i++)
                                {
                                    if (coordinate[i].Timestamp != "")
                                    {
                                        namedWorksheet.Cells[iRow, 41].Value = coordinate[i].Timestamp;
                                    }  // Time of latest observation
                                    iRow++;
                                }
                                package.Save();

                            }
                            goto TheEnd;
                        }
                    }
                    else if (strTimeBlockType != "Historic")
                    {
                        string strTemp = strTimeBlockStartUTC;
                        strTemp = strTemp.Replace("'", "");
                        Console.WriteLine("9. No new observations since (UTC) " + strTemp);
                    }

TheEnd:

// Update the activity file

                    string strProgramEnd = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string strRecordFile = strExcelPath + "CoordinateExportActivityRecord.txt";

                    if (strTimeBlockType == "Manual")
                    {
                        strProgramEnd = strProgramEnd + " manual extraction for time " + strManualBlockStart;
                    }

                    if (!File.Exists(strRecordFile))
                    {
                        string strFileName = strRecordFile;
                        using (StreamWriter writetext = new StreamWriter(strFileName, false))
                        {
                            writetext.WriteLine(strProgramEnd);
                            writetext.Close();
                        }
                    }
                    else
                    {
                        using (StreamWriter writetext = new StreamWriter(strRecordFile, append: true))
                        {
                            writetext.WriteLine(strProgramEnd);
                            writetext.Close();
                        }
                    }

                    if (strTimeBlockType == "Schedule") goto Finish;
                    Console.WriteLine(" ");

                } while (strExitFlag != "TheEnd");



                //strTimeBlockStartUTC != "The End"


            }


ThatsAllFolks:

Finish:
            string strMessage = "Export Coordinates: " + strProjectTitle + " (" + strDateTime + ")" + " (FTP folder)";
            string strRootFolder = @"C:\";
            gnaT.updateSystemLogFile(strRootFolder, strMessage);

            Console.WriteLine("");
            Console.WriteLine("Task complete");
            Console.WriteLine("");


            string strFreezeScreen = ConfigurationManager.AppSettings["freezeScreen"];

            if (strFreezeScreen == "Yes")
            {
                Console.WriteLine("freezeScreen set to Yes");
                Console.WriteLine("press key to exit..");
                Console.ReadKey();
            }

            Console.WriteLine("Geomos coordinate export completed...");
            Environment.Exit(0);




        }
    }
}