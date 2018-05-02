using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

using System.Text.RegularExpressions;
using Manatee.Trello;
using Manatee.Trello.ManateeJson;
using Manatee.Trello.WebApi;
using Excel = Microsoft.Office.Interop.Excel;


namespace MspUpdate
{
    public class TrelloConnection
    {
        Dictionary<string, List<string>> cardSorted = new Dictionary<string, List<string>>();
        Boolean Dbg;
        ManateeFactory JsonFactory;
        string MspExe;
        List<string> noLabelTickets = new List<string>();
        Microsoft.Office.Interop.Excel._Worksheet oShtExec;
        Microsoft.Office.Interop.Excel._Worksheet oShtCrdsFrmTrllo;
        Microsoft.Office.Interop.Excel._Worksheet oShtKdsFrmTrllo;
        Microsoft.Office.Interop.Excel._Worksheet oShtLgTrlloScn;
        Microsoft.Office.Interop.Excel._Worksheet oShtExprtMspWrkitms;
        Microsoft.Office.Interop.Excel._Worksheet oShtLgMspWrkitms;
        Microsoft.Office.Interop.Excel._Worksheet oShtLgTrlloCrds;
        Microsoft.Office.Interop.Excel._Worksheet oShtLgMspTsks;
        Microsoft.Office.Interop.Excel._Worksheet oShtExprtMspTsks;
        Microsoft.Office.Interop.Excel._Worksheet oShtBgCrOpnCntDta;
        

        Microsoft.Office.Interop.Excel._Worksheet oShtExprtMspTsksWOkToIncld;
        Microsoft.Office.Interop.Excel._Worksheet oShtExprtMsrs;
        Microsoft.Office.Interop.Excel._Worksheet oSht;
        Microsoft.Office.Interop.Excel._Worksheet oShtAllCrds;
        Microsoft.Office.Interop.Excel._Worksheet oShtTrlloScnErrr;
        Microsoft.Office.Interop.Excel._Worksheet oShtExprt10K;
        Microsoft.Office.Interop.Excel._Worksheet oShtTsks;

        Microsoft.Office.Interop.Excel.Application oXL;
        Microsoft.Office.Interop.Excel._Workbook oWB;
        WebApiClientProvider RestClientProvider;
        ManateeSerializer serializer;
        string Tm;
        bool UpdtMspActls;
        bool UpdtMspMsrs;
        bool UpdtMspPrjctd;



        public void CteReadBoard(
            string Prjct,
            string XlsTmpltPth,
            string XlsFlPth,
            string CnfgFlPth,
            Configuration Cnfg,
            String TmStrt
        )
        {
            if (Cnfg.DbgUsr)
            {
                Console.Write("\r\nEntering CTEReadBoard");
            }

            // Manatee configuration
            var serializer = new ManateeSerializer();
            var JsonFactory = new ManateeFactory();
            var RestClientProvider = new WebApiClientProvider();
            TrelloConfiguration.JsonFactory = JsonFactory;
            TrelloConfiguration.RestClientProvider = RestClientProvider;
            TrelloConfiguration.Serializer = serializer;
            TrelloConfiguration.Deserializer = serializer;
            //TrelloConfiguration.JsonFactory = new ManateeFactory();
            //TrelloConfiguration.RestClientProvider = new WebApiClientProvider();
            TrelloAuthorization.Default.AppKey = Cnfg.TrlloAppKy;
            TrelloAuthorization.Default.UserToken = Cnfg.TrlloUsrTkn;

            //Start Excel and get Application object.
            oXL = new Microsoft.Office.Interop.Excel.Application();
            //oXL.Visible = true;

            try
            {
                oXL.Visible = true;
            }
            catch (Exception Excptn)
            {
                Console.Write("\r\nError making xls visible: " + Excptn);
            }

            // Open the template xls and save under new name.
            oWB = (Microsoft.Office.Interop.Excel._Workbook)oXL.Workbooks.Open(XlsTmpltPth);
            oXL.UserControl = false;
            oXL.DisplayAlerts = false;
            oWB.SaveAs(XlsFlPth, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            oXL.DisplayAlerts = true;

            if (Cnfg.DbgUsr)
            {
                Console.Write("\r\nAfter xls saved");
            }

            if (Cnfg.DbgUsr)
            {
                Console.Write("\r\nAfter Excel started");
            }

            switch (Cnfg.PtsHrs.ToUpper())
            {
                case "POINTS":
                    // Create xls sheets
                    oShtExec = oWB.Worksheets["Exec"];
                    oShtExec.Cells[2, 2] = Prjct;
                    oShtExec.Cells[3, 2] = CnfgFlPth;
                    oShtExec.Cells[4, 2] = TmStrt;

                    oShtExprtMsrs = oWB.Worksheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                    oShtExprtMsrs.Name = "Export Measures";
                    oShtExprtMsrs.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                    oShtExprtMsrs.Columns[1].ColumnWidth = 30;
                    oShtExprtMsrs.Columns[2].ColumnWidth = 15;
                    oShtExprtMsrs.Columns[3].ColumnWidth = 15;
                    oShtExprtMsrs.Cells[1, 1] = "Measure Name";
                    oShtExprtMsrs.Cells[1, 2] = "Date";
                    oShtExprtMsrs.Cells[1, 3] = "Measure Value";


                    oShtExprtMspTsksWOkToIncld = oWB.Worksheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                    oShtExprtMspTsksWOkToIncld.Name = "Export MSP Tasks";
                    oShtExprtMspTsksWOkToIncld.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                    oShtExprtMspTsksWOkToIncld.Columns[1].ColumnWidth = 30;
                    oShtExprtMspTsksWOkToIncld.Columns[2].ColumnWidth = 15;
                    oShtExprtMspTsksWOkToIncld.Columns[3].ColumnWidth = 15;
                    oShtExprtMspTsksWOkToIncld.Cells[1, 1] = "Measure Name";
                    oShtExprtMspTsksWOkToIncld.Cells[1, 2] = "Date";
                    oShtExprtMspTsksWOkToIncld.Cells[1, 3] = "Measure Value";

                    oShtExprtMspWrkitms = oWB.Worksheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                    oShtExprtMspWrkitms.Name = "Export MSP Workitems";
                    oShtExprtMspWrkitms.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                    oShtExprtMspWrkitms.Columns[1].ColumnWidth = 15;
                    oShtExprtMspWrkitms.Columns[2].ColumnWidth = 15;
                    oShtExprtMspWrkitms.Columns[3].ColumnWidth = 20;
                    oShtExprtMspWrkitms.Columns[4].ColumnWidth = 8;
                    oShtExprtMspWrkitms.Columns[5].ColumnWidth = 10;
                    oShtExprtMspWrkitms.Columns[6].ColumnWidth = 25;
                    oShtExprtMspWrkitms.Columns[7].ColumnWidth = 8;
                    oShtExprtMspWrkitms.Columns[8].ColumnWidth = 12;
                    oShtExprtMspWrkitms.Columns[9].ColumnWidth = 8;
                    oShtExprtMspWrkitms.Columns[10].ColumnWidth = 5;
                    oShtExprtMspWrkitms.Cells[1, 1] = "Board";
                    oShtExprtMspWrkitms.Cells[1, 2] = "List";
                    oShtExprtMspWrkitms.Cells[1, 3] = "Card Name";
                    oShtExprtMspWrkitms.Cells[1, 4] = "Priority";
                    oShtExprtMspWrkitms.Cells[1, 5] = "Task Type";
                    oShtExprtMspWrkitms.Cells[1, 6] = "Labels";
                    oShtExprtMspWrkitms.Cells[1, 7] = "Points";
                    oShtExprtMspWrkitms.Cells[1, 8] = "Is Open";
                    oShtExprtMspWrkitms.Cells[1, 9] = "Card ID";
                    oShtExprtMspWrkitms.Cells[1, 10] = "Card URL";

                    //oShtLgMspWrkitms = oWB.Worksheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                    //oShtLgMspWrkitms.Name = "Log MSP Workitems";
                    //oShtLgMspWrkitms.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                    //oShtLgMspWrkitms.Columns[1].ColumnWidth = 5;
                    //oShtLgMspWrkitms.Columns[2].ColumnWidth = 30;
                    //oShtLgMspWrkitms.Columns[3].ColumnWidth = 40;
                    //oShtLgMspWrkitms.Cells[1, 1] = "ID";
                    //oShtLgMspWrkitms.Cells[1, 2] = "Name";
                    //oShtLgMspWrkitms.Cells[1, 3] = "Description";

                    oShtLgTrlloCrds = oWB.Worksheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                    oShtLgTrlloCrds.Name = "Log Trello Cards";
                    oShtLgTrlloCrds.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                    oShtLgTrlloCrds.Columns[1].ColumnWidth = 30;
                    oShtLgTrlloCrds.Columns[2].ColumnWidth = 20;
                    oShtLgTrlloCrds.Columns[3].ColumnWidth = 40;
                    oShtLgTrlloCrds.Columns[4].ColumnWidth = 20;
                    oShtLgTrlloCrds.Cells[1, 1] = "Row on Cards From Trello";
                    oShtLgTrlloCrds.Cells[1, 2] = "Code Section";
                    oShtLgTrlloCrds.Cells[1, 3] = "Description";
                    oShtLgTrlloCrds.Cells[1, 4] = "Trello Card ID";

                    oShtLgTrlloScn = oWB.Worksheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                    oShtLgTrlloScn.Name = "Log Trello Scan";
                    oShtLgTrlloScn.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                    oShtLgTrlloScn.Columns[1].ColumnWidth = 20;
                    oShtLgTrlloScn.Columns[2].ColumnWidth = 20;
                    oShtLgTrlloScn.Columns[3].ColumnWidth = 40;
                    oShtLgTrlloScn.Columns[4].ColumnWidth = 20;
                    oShtLgTrlloScn.Cells[1, 1] = "List";
                    oShtLgTrlloScn.Cells[1, 2] = "Card";
                    oShtLgTrlloScn.Cells[1, 3] = "Task";
                    oShtLgTrlloScn.Cells[1, 4] = "Exception";

                    oShtCrdsFrmTrllo = oWB.Worksheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                    oShtCrdsFrmTrllo.Name = "Cards From Trello";
                    oShtCrdsFrmTrllo.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                    oShtCrdsFrmTrllo.Columns[1].ColumnWidth = 20;
                    oShtCrdsFrmTrllo.Columns[2].ColumnWidth = 20;
                    oShtCrdsFrmTrllo.Columns[3].ColumnWidth = 20;
                    oShtCrdsFrmTrllo.Columns[4].ColumnWidth = 10;
                    oShtCrdsFrmTrllo.Columns[5].ColumnWidth = 10;
                    oShtCrdsFrmTrllo.Columns[6].ColumnWidth = 15;
                    oShtCrdsFrmTrllo.Columns[7].ColumnWidth = 8;
                    oShtCrdsFrmTrllo.Columns[8].ColumnWidth = 5;
                    oShtCrdsFrmTrllo.Columns[9].ColumnWidth = 5;
                    oShtCrdsFrmTrllo.Columns[10].ColumnWidth = 15;
                    oShtCrdsFrmTrllo.Cells[1, 1] = "Board";
                    oShtCrdsFrmTrllo.Cells[1, 2] = "List";
                    oShtCrdsFrmTrllo.Cells[1, 3] = "Card Name";
                    oShtCrdsFrmTrllo.Cells[1, 4] = "Card Type";
                    oShtCrdsFrmTrllo.Cells[1, 5] = "Priority";
                    oShtCrdsFrmTrllo.Cells[1, 6] = "Labels";
                    oShtCrdsFrmTrllo.Cells[1, 7] = "Points";
                    oShtCrdsFrmTrllo.Cells[1, 8] = "Card ID";
                    oShtCrdsFrmTrllo.Cells[1, 9] = "Card URL";
                    oShtCrdsFrmTrllo.Cells[1, 10] = "KD Label";

                    oShtKdsFrmTrllo = oWB.Worksheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                    oShtKdsFrmTrllo.Name = "KDs From Trello";
                    oShtKdsFrmTrllo.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                    oShtKdsFrmTrllo.Columns[1].ColumnWidth = 30;
                    oShtKdsFrmTrllo.Columns[2].ColumnWidth = 20;
                    oShtKdsFrmTrllo.Columns[3].ColumnWidth = 10;
                    oShtKdsFrmTrllo.Columns[4].ColumnWidth = 15;
                    oShtKdsFrmTrllo.Columns[5].ColumnWidth = 15;
                    oShtKdsFrmTrllo.Columns[6].ColumnWidth = 30;
                    oShtKdsFrmTrllo.Columns[7].ColumnWidth = 10;
                    oShtKdsFrmTrllo.Columns[8].ColumnWidth = 20;
                    oShtKdsFrmTrllo.Columns[3].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    oShtKdsFrmTrllo.Columns[4].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    oShtKdsFrmTrllo.Cells[1, 1] = "KD Name";
                    oShtKdsFrmTrllo.Cells[1, 2] = "KD Label";
                    oShtKdsFrmTrllo.Cells[1, 3] = "KD Value";
                    oShtKdsFrmTrllo.Cells[1, 4] = "KD TimeCritical";
                    oShtKdsFrmTrllo.Cells[1, 5] = "KD Status";
                    oShtKdsFrmTrllo.Cells[1, 6] = "KD Card ID";
                    oShtKdsFrmTrllo.Cells[1, 7] = "KD Card URL";
                    oShtKdsFrmTrllo.Cells[1, 8] = "Exceptions";

                    // Parse workitem cards for points
                    TrelloParsePoints(Prjct, XlsTmpltPth, XlsFlPth, CnfgFlPth, Cnfg, TmStrt);

                    // Parse KD cards for KD info
                    TrelloParseKds(Prjct, XlsTmpltPth, XlsFlPth, CnfgFlPth, Cnfg);

                    // Save workbook
                    oWB.Save();

                    // Update Project Online
                    // DateTime DtUpdt = new DateTime(2017, 1, 3);
                    // string PrjctMsp = "BruceP Test 2";
                    if (UpdtMspActls || UpdtMspPrjctd || UpdtMspMsrs)
                    {
                        oXL.Run("Update_Project_Online_Points", MspExe);
                    }

                    break;

                //case "HOURS":
                //    // Create xls sheets
                //    oShtExec = oWB.Worksheets["Exec"];
                //    oShtExec.Cells[2, 2] = Prjct;
                //    oShtExec.Cells[3, 2] = CnfgFlPth;
                //    oShtExec.Cells[4, 2] = TmStrt;

                //    oShtExprtMsrs = oWB.Worksheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                //    oShtExprtMsrs.Name = "Export Measures";
                //    oShtExprtMsrs.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                //    oShtExprtMsrs.Columns[1].ColumnWidth = 40;
                //    oShtExprtMsrs.Columns[2].ColumnWidth = 10;
                //    oShtExprtMsrs.Columns[3].ColumnWidth = 15;
                //    oShtExprtMsrs.Cells[1, 1] = "Measure Name";
                //    oShtExprtMsrs.Cells[1, 2] = "Date";
                //    oShtExprtMsrs.Cells[1, 3] = "Measure Value";

                //    oShtExprtMspTsksWOkToIncld = oWB.Worksheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                //    oShtExprtMspTsksWOkToIncld.Name = "Export MSP Tasks w OkToIncld";
                //    oShtExprtMspTsksWOkToIncld.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                //    oShtExprtMspTsksWOkToIncld.Columns[1].ColumnWidth = 8;
                //    oShtExprtMspTsksWOkToIncld.Columns[2].ColumnWidth = 8;
                //    oShtExprtMspTsksWOkToIncld.Columns[3].ColumnWidth = 20;
                //    oShtExprtMspTsksWOkToIncld.Columns[4].ColumnWidth = 15;
                //    oShtExprtMspTsksWOkToIncld.Columns[5].ColumnWidth = 20;
                //    oShtExprtMspTsksWOkToIncld.Columns[6].ColumnWidth = 8;
                //    oShtExprtMspTsksWOkToIncld.Columns[7].ColumnWidth = 6;
                //    oShtExprtMspTsksWOkToIncld.Columns[8].ColumnWidth = 6;
                //    oShtExprtMspTsksWOkToIncld.Columns[9].ColumnWidth = 6;
                //    oShtExprtMspTsksWOkToIncld.Columns[10].ColumnWidth = 20;
                //    oShtExprtMspTsksWOkToIncld.Columns[11].ColumnWidth = 10;
                //    oShtExprtMspTsksWOkToIncld.Cells[1, 1] = "Condition";
                //    oShtExprtMspTsksWOkToIncld.Cells[1, 2] = "OkToIncld";
                //    oShtExprtMspTsksWOkToIncld.Cells[1, 3] = "Card";
                //    oShtExprtMspTsksWOkToIncld.Cells[1, 4] = "Phase";
                //    oShtExprtMspTsksWOkToIncld.Cells[1, 5] = "Task";
                //    oShtExprtMspTsksWOkToIncld.Cells[1, 6] = "Assigned";
                //    oShtExprtMspTsksWOkToIncld.Cells[1, 7] = "Hrs Ttl";
                //    oShtExprtMspTsksWOkToIncld.Cells[1, 8] = "Hrs Actl";
                //    oShtExprtMspTsksWOkToIncld.Cells[1, 9] = "Hrs Rmng";
                //    oShtExprtMspTsksWOkToIncld.Cells[1, 10] = "Labels";
                //    oShtExprtMspTsksWOkToIncld.Cells[1, 11] = "Card URL";

                //    oShtExprtMspTsks = oWB.Worksheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                //    oShtExprtMspTsks.Name = "Export MSP Tasks";
                //    oShtExprtMspTsks.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                //    oShtExprtMspTsks.Columns[1].ColumnWidth = 20;
                //    oShtExprtMspTsks.Columns[2].ColumnWidth = 15;
                //    oShtExprtMspTsks.Columns[3].ColumnWidth = 20;
                //    oShtExprtMspTsks.Columns[4].ColumnWidth = 8;
                //    oShtExprtMspTsks.Columns[5].ColumnWidth = 6;
                //    oShtExprtMspTsks.Columns[6].ColumnWidth = 6;
                //    oShtExprtMspTsks.Columns[7].ColumnWidth = 6;
                //    oShtExprtMspTsks.Columns[8].ColumnWidth = 20;
                //    oShtExprtMspTsks.Columns[9].ColumnWidth = 10;
                //    oShtExprtMspTsks.Cells[1, 1] = "Work Item for Task";
                //    oShtExprtMspTsks.Cells[1, 2] = "Phase";
                //    oShtExprtMspTsks.Cells[1, 3] = "Task";
                //    oShtExprtMspTsks.Cells[1, 4] = "Assigned";
                //    oShtExprtMspTsks.Cells[1, 5] = "Hrs Ttl";
                //    oShtExprtMspTsks.Cells[1, 6] = "Hrs Actl";
                //    oShtExprtMspTsks.Cells[1, 7] = "Hrs Rmng";
                //    oShtExprtMspTsks.Cells[1, 8] = "Labels";
                //    oShtExprtMspTsks.Cells[1, 9] = "Card URL";

                //    oShtExprt10K = oWB.Worksheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                //    oShtExprt10K.Name = "Export 10K";
                //    oShtExprt10K.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                //    oShtExprt10K.Columns[1].ColumnWidth = 20;
                //    oShtExprt10K.Columns[2].ColumnWidth = 10;
                //    oShtExprt10K.Columns[3].ColumnWidth = 20;
                //    oShtExprt10K.Columns[4].ColumnWidth = 15;
                //    oShtExprt10K.Columns[5].ColumnWidth = 5;
                //    oShtExprt10K.Columns[6].ColumnWidth = 4;
                //    oShtExprt10K.Columns[7].ColumnWidth = 5;
                //    oShtExprt10K.Columns[8].ColumnWidth = 6;
                //    oShtExprt10K.Columns[9].ColumnWidth = 16;
                //    oShtExprt10K.Columns[10].ColumnWidth = 10;
                //    oShtExprt10K.Cells[1, 1] = "Card Name";
                //    oShtExprt10K.Cells[1, 2] = "Workphase Name";
                //    oShtExprt10K.Cells[1, 3] = "Task Name";
                //    oShtExprt10K.Cells[1, 4] = "Assigned";
                //    oShtExprt10K.Cells[1, 5] = "Hrs Actl";
                //    oShtExprt10K.Cells[1, 6] = "Hrs Rmnng";
                //    oShtExprt10K.Cells[1, 7] = "Labels";
                //    oShtExprt10K.Cells[1, 8] = "Error Flag";
                //    oShtExprt10K.Cells[1, 9] = "Error Desc";
                //    oShtExprt10K.Cells[1, 10] = "Card URL";

                //    oShtBgCrOpnCntDta = oWB.Worksheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                //    oShtBgCrOpnCntDta.Name = "BG CR Open Count Data";
                //    oShtBgCrOpnCntDta.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                //    oShtBgCrOpnCntDta.Columns[1].ColumnWidth = 20;
                //    oShtBgCrOpnCntDta.Columns[2].ColumnWidth = 30;
                //    oShtBgCrOpnCntDta.Columns[3].ColumnWidth = 40;
                //    oShtBgCrOpnCntDta.Columns[4].ColumnWidth = 20;
                //    oShtBgCrOpnCntDta.Columns[5].ColumnWidth = 30;
                //    oShtBgCrOpnCntDta.Columns[6].ColumnWidth = 40;
                //    oShtBgCrOpnCntDta.Cells[1, 1] = "Measure";
                //    oShtBgCrOpnCntDta.Cells[1, 2] = "Card Name";
                //    oShtBgCrOpnCntDta.Cells[1, 3] = "Board Name";
                //    oShtBgCrOpnCntDta.Cells[1, 4] = "List Name";
                //    oShtBgCrOpnCntDta.Cells[1, 5] = "Labels";
                //    oShtBgCrOpnCntDta.Cells[1, 6] = "Card URL";

                //    oShtLgMspTsks = oWB.Worksheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                //    oShtLgMspTsks.Name = "Log MSP Tasks";
                //    oShtLgMspTsks.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                //    oShtLgMspTsks.Columns[1].ColumnWidth = 8;
                //    oShtLgMspTsks.Columns[2].ColumnWidth = 40;
                //    oShtLgMspTsks.Cells[1, 1] = "Task ID";
                //    oShtLgMspTsks.Cells[1, 2] = "Task Name";
                //    oShtLgMspTsks.Cells[1, 3] = "Description";

                //    oShtTrlloScnErrr = oWB.Worksheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                //    oShtTrlloScnErrr.Name = "Log Trello Tasks";
                //    oShtTrlloScnErrr.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                //    oShtTrlloScnErrr.Columns[1].ColumnWidth = 10;
                //    oShtTrlloScnErrr.Columns[2].ColumnWidth = 30;
                //    oShtTrlloScnErrr.Columns[3].ColumnWidth = 40;
                //    oShtTrlloScnErrr.Cells[1, 1] = "Trello Task #";
                //    oShtTrlloScnErrr.Cells[1, 2] = "Code Section";
                //    oShtTrlloScnErrr.Cells[1, 3] = "Description";

                //    oShtAllCrds = oWB.Worksheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                //    oShtAllCrds.Name = "All Cards";
                //    oShtAllCrds.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                //    oShtAllCrds.Columns[1].ColumnWidth = 20;
                //    oShtAllCrds.Columns[2].ColumnWidth = 30;
                //    oShtAllCrds.Columns[3].ColumnWidth = 40;
                //    oShtAllCrds.Columns[4].ColumnWidth = 25;
                //    oShtAllCrds.Cells[1, 1] = "Board";
                //    oShtAllCrds.Cells[1, 2] = "List";
                //    oShtAllCrds.Cells[1, 3] = "Card Name";
                //    oShtAllCrds.Cells[1, 4] = "Card ID";
                //    oShtAllCrds.Cells[1, 5] = "Labels";

                //    oShtTsks = oWB.Worksheets.Add(oWB.Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                //    oShtTsks.Name = "Tasks";
                //    oShtTsks.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                //    oShtTsks.Columns[1].ColumnWidth = 10;
                //    oShtTsks.Columns[2].ColumnWidth = 10;
                //    oShtTsks.Columns[3].ColumnWidth = 20;
                //    oShtTsks.Columns[4].ColumnWidth = 10;
                //    oShtTsks.Columns[5].ColumnWidth = 10;
                //    oShtTsks.Columns[6].ColumnWidth = 10;
                //    oShtTsks.Columns[7].ColumnWidth = 5;
                //    oShtTsks.Columns[8].ColumnWidth = 5;
                //    oShtTsks.Columns[9].ColumnWidth = 4;
                //    oShtTsks.Columns[10].ColumnWidth = 10;
                //    oShtTsks.Columns[11].ColumnWidth = 5;
                //    oShtTsks.Columns[12].ColumnWidth = 5;
                //    oShtTsks.Columns[13].ColumnWidth = 5;
                //    oShtTsks.Columns[14].ColumnWidth = 6;
                //    oShtTsks.Columns[15].ColumnWidth = 6;
                //    oShtTsks.Columns[16].ColumnWidth = 6;
                //    oShtTsks.Columns[17].ColumnWidth = 6;
                //    oShtTsks.Columns[18].ColumnWidth = 6;
                //    oShtTsks.Columns[19].ColumnWidth = 6;
                //    oShtTsks.Columns[20].ColumnWidth = 16;
                //    oShtTsks.Columns[21].ColumnWidth = 6;
                //    oShtTsks.Columns[22].ColumnWidth = 10;
                //    oShtTsks.Columns[23].ColumnWidth = 10;
                //    oShtTsks.Columns[24].ColumnWidth = 10;
                //    oShtTsks.Cells[1, 1] = "Board Name";
                //    oShtTsks.Cells[1, 2] = "List Name";
                //    oShtTsks.Cells[1, 3] = "Card Name";
                //    oShtTsks.Cells[1, 4] = "Priority";
                //    oShtTsks.Cells[1, 5] = "Workphase Name";
                //    oShtTsks.Cells[1, 6] = "Task Name";
                //    oShtTsks.Cells[1, 7] = "Assigned";
                //    oShtTsks.Cells[1, 8] = "Hrs Actl";
                //    oShtTsks.Cells[1, 9] = "Hrs Rmnng";
                //    oShtTsks.Cells[1, 10] = "Role";
                //    oShtTsks.Cells[1, 11] = "Labels";
                //    oShtTsks.Cells[1, 12] = "EstmtO";
                //    oShtTsks.Cells[1, 13] = "EstmtL";
                //    oShtTsks.Cells[1, 14] = "EstmtP";
                //    oShtTsks.Cells[1, 15] = "EstmtE";
                //    oShtTsks.Cells[1, 16] = "Card ID";
                //    oShtTsks.Cells[1, 17] = "Checklist ID";
                //    oShtTsks.Cells[1, 18] = "Task ID";
                //    oShtTsks.Cells[1, 19] = "Error Flag";
                //    oShtTsks.Cells[1, 20] = "Error Desc";
                //    oShtTsks.Cells[1, 21] = "CheckItem Name";
                //    oShtTsks.Cells[1, 22] = "Card URL";
                //    oShtTsks.Cells[1, 23] = "Card Last Activity";
                //    oShtTsks.Cells[1, 24] = "Sort";

                //    TrelloParseHours(Prjct, XlsTmpltPth, XlsFlPth, CnfgFlPth, Cnfg, TmStrt);

                //    // Save workbook
                //    oWB.Save();

                //    // Update Project Online
                //    // DateTime DtUpdt = new DateTime(2017, 1, 3);
                //    // string PrjctMsp = "BruceP Test 2";
                //    if (UpdtMspActls || UpdtMspPrjctd || UpdtMspMsrs)
                //    {
                //        oXL.Run("Update_Project_Online_Hours", MspExe);
                //    }

                    //break;
            }

            // Write time completed
            Tm = DateTime.Now.ToString("hh:mm:ss");
            Console.Write("\r\nMSP update completed at " + Tm);
            oShtExec.Cells[11, 2] = Tm;

            // Close xls
            oWB.Save();
            oWB.Close();
            oXL.Quit();


        }

        public void TrelloParsePoints (
            string Prjct,
            string XlsTmpltPth,
            string XlsFlPth,
            string CnfgFlPth,
            Configuration Cnfg,
            String TmStrt
        )
        {
            if (Cnfg.DbgUsr)
            {
                Console.Write("\r\nEntering TrelloParsePoints");


            }

            if (Cnfg.DbgUsr)
            {
                Console.Write("\r\nAfter worksheets defined");
            }

            // Variables
            string[] BffrTkns = new string[] { "" };
            String BrdNm = "";
            List<string> Brds = new List<string>();
            string CrdId = "";
            DateTime CrdLstActvty;
            string CrdNm = "";
            string CrdPrty = "";
            List<string> Crds = new List<string>();
            string CrdTyp = "";
            string CrdUrl = "";
            bool DbgIncldThsCrd;
            DataTable FldsKd = new DataTable();
            DataTable FldsWrkItm = new DataTable();
            DateTime IncldCrdsChngdAftr;
            int iRwAllCrds = 1;
            int iRw1;
            string KdLbl = "";
            string Lbls;
            string Lst;
            List<string> LstsExcldd = new List<string>();
            List<string> LstsIncldd = new List<string>();
            MatchCollection Mtch;
            bool NtMrkrCrd;
            bool NtTmpltCrd;
            List<string> PmLns = new List<string>();
            int Pts = -99;
            string PtsHrs = "";
            string Rl = "";
            string[] RlLst = new string[] { "ROLE", "BE", "CM", "CSS", "DO", "FE", "MC", "QA", "UX", "PM", "PO" };
            var StryTsk = new Dictionary<string, string>();
            List<string> Tkns = new List<string>();
            string[] Tkns1 = new string[] { "" };
            string[] Tkns2 = new string[] { "" };
            string[] Tkns3 = new string[] { "" };
            List<string> TrlloLstsInclddInpt = new List<string>();
            List<string> TrlloLstsExclddInpt = new List<string>();

            FldsKd.Columns.Add("KdNm", typeof(string));
            FldsKd.Columns.Add("KdLbl", typeof(string));
            FldsKd.Columns.Add("KdVlu", typeof(int));
            FldsKd.Columns.Add("KdTm", typeof(int));
            FldsKd.Columns.Add("KdId", typeof(string));

            FldsWrkItm.Columns.Add("BrdNm", typeof(string));
            FldsWrkItm.Columns.Add("Lst", typeof(string));
            FldsWrkItm.Columns.Add("CrdNm", typeof(string));
            FldsWrkItm.Columns.Add("CrdPrty", typeof(string));
            FldsWrkItm.Columns.Add("Lbls", typeof(string));
            FldsWrkItm.Columns.Add("Pts", typeof(int));
            FldsWrkItm.Columns.Add("CrdId", typeof(string));
            FldsWrkItm.Columns.Add("ErrrFnd", typeof(bool));
            FldsWrkItm.Columns.Add("ErrrTxt", typeof(string));
            FldsWrkItm.Columns.Add("CrdUrl", typeof(string));
            FldsWrkItm.Columns.Add("CrdLstActvty", typeof(DateTime));

            //Parms
            Brds = Cnfg.Brds;
            MspExe = Cnfg.MspExe;
            UpdtMspActls = Cnfg.UpdtMspActls;
            UpdtMspMsrs = Cnfg.UpdtMspMsrs;
            UpdtMspPrjctd = Cnfg.UpdtMspPrjctd;
            PtsHrs = Cnfg.PtsHrs;
            //PstAllChckLstItms = Cnfg.PstAllChckLstItms;
            IncldCrdsChngdAftr = Cnfg.IncldCrdsChngdAftr;
            TrlloLstsInclddInpt = Cnfg.TrlloLstsInclddInpt;
            TrlloLstsExclddInpt = Cnfg.TrlloLstsExclddInpt;

            // Debug: Cards to be read for debug
            Dbg = false;
            //Crds.Add("57e046ca970fb81e9e789ea6");

            // Loop for each board
            foreach (string BrdId in Brds)
            {
                // Board data elements
                var Brd = new Board(BrdId);
                BrdNm = Brd.Name;

                // Find Trello lists to be included and excluded.  Write them to xls exec tab.
                bool IncldLst;
                if (TrlloLstsInclddInpt.Count == 0 && TrlloLstsExclddInpt.Count == 0)
                {
                    foreach (List BrdLst in Brd.Lists)
                    {
                        Cnfg.TrlloLstsIncldd.Add(BrdLst.Name);
                    }
                }

                if (TrlloLstsInclddInpt.Count != 0 && TrlloLstsExclddInpt.Count == 0)
                {
                    foreach (List BrdLst in Brd.Lists)
                    {
                        IncldLst = false;
                        foreach (string Tlst in TrlloLstsInclddInpt)
                        {
                            if (BrdLst.Name == Tlst)
                            {
                                IncldLst = true;
                            }
                        }

                        if (IncldLst)
                        {
                            Cnfg.TrlloLstsIncldd.Add(BrdLst.Name);
                        }
                        else
                        {
                            Cnfg.TrlloLstsExcldd.Add(BrdLst.Name);
                        }
                    }
                }

                if (TrlloLstsInclddInpt.Count == 0 && TrlloLstsExclddInpt.Count != 0)
                {
                    foreach (List BrdLst in Brd.Lists)
                    {
                        IncldLst = true;
                        foreach (string Tlst in TrlloLstsExclddInpt)
                        {
                            if (BrdLst.Name == Tlst)
                            {
                                IncldLst = false;
                            }
                        }

                        if (IncldLst)
                        {
                            Cnfg.TrlloLstsIncldd.Add(BrdLst.Name);
                        }
                        else
                        {
                            Cnfg.TrlloLstsExcldd.Add(BrdLst.Name);
                        }
                    }
                }

                if (TrlloLstsInclddInpt.Count != 0 && TrlloLstsExclddInpt.Count != 0)
                {
                    foreach (List BrdLst in Brd.Lists)
                    {
                        IncldLst = false;
                        foreach (string Tlst in TrlloLstsInclddInpt)
                        {
                            if (BrdLst.Name == Tlst)
                            {
                                IncldLst = true;
                            }
                        }

                        foreach (string Tlst in TrlloLstsExclddInpt)
                        {
                            if (BrdLst.Name == Tlst)
                            {
                                IncldLst = false;
                            }
                        }

                        if (IncldLst)
                        {
                            Cnfg.TrlloLstsIncldd.Add(BrdLst.Name);
                        }
                        else
                        {
                            Cnfg.TrlloLstsExcldd.Add(BrdLst.Name);
                        }
                    }
                }

                // Find Trello lists that are open and closed
                foreach (List BrdLst in Brd.Lists)
                {
                    // Add Release and Hotfix lists to ListsNotOpen
                    if (BrdLst.Name.Length >= 6)
                    {
                        if (BrdLst.Name.Substring(0, 6).ToUpper() == "HOTFIX")
                        {
                            Cnfg.TrlloLstsNtOpn.Add(BrdLst.Name);
                        }
                    }
                    if (BrdLst.Name.Length >= 7)
                    {
                        if (BrdLst.Name.Substring(0, 7).ToUpper() == "RELEASE")
                        {
                            Cnfg.TrlloLstsNtOpn.Add(BrdLst.Name);
                        }
                    }

                    if (!Cnfg.TrlloLstsNtOpn.Contains(BrdLst.Name))
                    {
                        Cnfg.TrlloLstsOpn.Add(BrdLst.Name);
                    }
                }

                // Write lists of lists to xls tab Exec.
                int RwLbls = 14;
                oShtExec.Cells[RwLbls, 1] = "Lists included in the Trello scan";
                iRw1 = RwLbls;
                foreach (string LstNm in Cnfg.TrlloLstsIncldd)
                {
                    iRw1++;
                    oShtExec.Cells[iRw1, 1] = LstNm;
                }

                oShtExec.Cells[RwLbls, 2] = "Lists excluded from the Trello scan";
                iRw1 = RwLbls;
                foreach (string LstNm in Cnfg.TrlloLstsExcldd)
                {
                    iRw1++;
                    oShtExec.Cells[iRw1, 2] = LstNm;
                }

                //oShtExec.Cells[RwLbls, 3] = "Lists contining rejected cards.  Tasks for these cards will be deleted in Project Online.";
                //iRw1 = RwLbls;
                //foreach (string LstNm in Cnfg.TrlloLstsRjctd)
                //{
                //    iRw1++;
                //    oShtExec.Cells[iRw1, 3] = LstNm;
                //}

                oShtExec.Cells[RwLbls, 3] = "Lists containing cards that are open";
                iRw1 = RwLbls;
                foreach (string LstNm in Cnfg.TrlloLstsOpn)
                {
                    iRw1++;
                    oShtExec.Cells[iRw1, 3] = LstNm;
                }

                oShtExec.Cells[RwLbls, 4] = "Lists containing cards that are not open";
                iRw1 = RwLbls;
                foreach (string LstNm in Cnfg.TrlloLstsNtOpn)
                {
                    iRw1++;
                    oShtExec.Cells[iRw1, 4] = LstNm;
                }


                // Loop for each card on board
                int cntCrds = Brd.Cards.Count();
                foreach (var card in Brd.Cards)
                {

                    // Get card info
                    CrdId = card.Id;
                    CrdNm = card.Name;
                    CrdUrl = card.ShortUrl;
                    CrdLstActvty = card.LastActivity.Value.Date;
                    Lst = card.List.Name;

                    // Get labels
                    Lbls = "";
                    foreach (var label in card.Labels)
                    {
                        if ((label.Name != null) && (label.Name != ""))
                        {
                            if (Lbls == "")
                            {
                                Lbls += label.Name;
                            }
                            else
                            {
                                Lbls += ";" + label.Name;
                            }
                        }
                    }

                    // Get card priority
                    CrdPrty = "unknown";
                    if (Lbls.Contains("priority3-low") || Lbls.Contains("priority-low"))
                    {
                        CrdPrty = "low";
                    }
                    if (Lbls.Contains("priority2-med") || Lbls.Contains("priority-medium"))
                    {
                        CrdPrty = "medium";
                    }
                    if (Lbls.Contains("priority1-high") || Lbls.Contains("priority-high"))
                    {
                        CrdPrty = "high";
                    }
                    if (Lbls.Contains("priority0-critical") || Lbls.Contains("priority-critical"))
                    {
                        CrdPrty = "critical";
                    }

                    // Check that card is on an included Trello List.
                    bool TrlloLstFnd = true;
                    if (Cnfg.TrlloLstsIncldd.Count != 0)
                    {
                        TrlloLstFnd = false;
                        foreach (string TrlloLst in Cnfg.TrlloLstsIncldd)
                        {
                            if (card.List.Name.Equals(TrlloLst))
                            {
                                TrlloLstFnd = true;
                            }
                        }
                    }

                    // check that card in included dates.
                    bool InInclddDts = true;

                    if (card.LastActivity <= IncldCrdsChngdAftr)
                    {
                        InInclddDts = false;
                    }

                    // Marker card
                    NtMrkrCrd = true;
                    if (Lbls.Contains("marker"))
                    {
                        NtMrkrCrd = false;
                    }

                    // Template card
                    NtTmpltCrd = true;
                    if (CrdNm.Contains("<TEMPLATE>"))
                    {
                        NtTmpltCrd = false;
                    }

                    // If debug then check if this card is on debug cards list
                    if (Dbg)
                    {
                        DbgIncldThsCrd = Crds.Contains(card.Id);
                    }
                    else
                    {
                        DbgIncldThsCrd = true;
                    }

                    if (DbgIncldThsCrd && TrlloLstFnd && InInclddDts && NtTmpltCrd && NtMrkrCrd)
                    {
                        // Get card type 
                        CrdTyp = "unknown";
                        foreach (var Lbl in card.Labels)
                        {
                            switch (Lbl.Name)
                            {
                                case "story":
                                    CrdTyp = "story";
                                    break;

                                case "change":
                                    CrdTyp = "change";
                                    break;

                                case "bug":
                                    CrdTyp = "bug";
                                    break;
                            }
                        }

                        // Get points 
                        Pts = -99;
                        foreach (var Lbl in card.Labels)
                        {
                            switch (Lbl.Name)
                            {
                                case "pts:01":
                                    Pts = 1;
                                    break;

                                case "pts:02":
                                    Pts = 2;
                                    break;

                                case "pts:03":
                                    Pts = 3;
                                    break;

                                case "pts:05":
                                    Pts = 5;
                                    break;

                                case "pts:08":
                                    Pts = 8;
                                    break;

                                case "pts:13":
                                    Pts = 13;
                                    break;

                                case "pts:20":
                                    Pts = 20;
                                    break;

                                case "pts:40":
                                    Pts = 40;
                                    break;

                                case "pts:80":
                                    Pts = 80;
                                    break;
                            }
                        }

                        // if Pts not found then default
                        if (Pts == -99)
                        {
                            if (CrdTyp == "story")
                            {
                                Pts = 6;
                            }
                            else
                            {
                                Pts = 4;
                            }
                        }

                        // Get KD label
                        KdLbl = "";
                        foreach (var label in card.Labels)
                        {
                            if ((label.Name != null) && (label.Name != ""))
                            {
                                if (label.Name.Contains("kd label"))
                                {
                                    BffrTkns = Regex.Split(label.Name, ":");
                                    KdLbl = BffrTkns[1].Trim();
                                }
                            }
                        }

                        // Write card to sheet All Cards
                        oShtCrdsFrmTrllo.Activate();
                        iRwAllCrds++;
                        oShtCrdsFrmTrllo.Cells[iRwAllCrds, 1] = Brd.Name;
                        oShtCrdsFrmTrllo.Cells[iRwAllCrds, 2] = card.List.Name;
                        oShtCrdsFrmTrllo.Cells[iRwAllCrds, 3] = card.Name;
                        oShtCrdsFrmTrllo.Cells[iRwAllCrds, 4] = CrdTyp;
                        oShtCrdsFrmTrllo.Cells[iRwAllCrds, 5] = CrdPrty;
                        oShtCrdsFrmTrllo.Cells[iRwAllCrds, 6] = Lbls;
                        if (Pts != -99)
                        {
                            oShtCrdsFrmTrllo.Cells[iRwAllCrds, 7] = Pts;

                        }
                        oShtCrdsFrmTrllo.Cells[iRwAllCrds, 8] = card.Id;
                        oShtCrdsFrmTrllo.Cells[iRwAllCrds, 9] = CrdUrl;
                        oShtCrdsFrmTrllo.Cells[iRwAllCrds, 10] = KdLbl;
                    }

                } // End foreach card

            }


            // All boards completed.  

            // Write time Trello scan done
            Tm = DateTime.Now.ToString("hh:mm:ss");
            Console.Write("\r\nTrello workitem scan done at " + Tm);
            oShtExec.Cells[5, 2] = Tm;

        }

        public void TrelloParseKds(
            string Prjct,
            string XlsTmpltPth,
            string XlsFlPth,
            string CnfgFlPth,
            Configuration Cnfg
        )
        {
            if (Cnfg.DbgUsr)
            {
                Console.Write("\r\nEntering TrelloParseKds");
            }

            Microsoft.Office.Interop.Excel.Range oRng;
            //object misvalue = System.Reflection.Missing.Value;

            if (Cnfg.DbgUsr)
            {
                Console.Write("\r\nAfter worksheets defined");
            }

            // Variables
            string[] BffrTkns = new string[] { "" };
            List<string> KdBrds = new List<string>();
            string CrdId = "";
            DateTime CrdLstActvty;
            string ErrMsg = "";
            string KdBrdNm = "";
            string KdLbl = "";
            string KdNm = "";
            string KdStts = "";
            int KdTm = 1;
            int KdVlu = 1;
            string CrdPrty = "";
            List<string> Crds = new List<string>();
            string CrdUrl = "";
            bool DbgIncldThsCrd;
            DataTable FldsKd = new DataTable();
            DataTable FldsWrkItm = new DataTable();
            DateTime IncldCrdsChngdAftr;
            int iRwAllCrds = 1;
            int iRw1;
            bool KdLblFnd;
            string Lbls;
            string Lst;
            List<string> LstsExcldd = new List<string>();
            List<string> LstsIncldd = new List<string>();
            MatchCollection Mtch;
            bool NtMrkrCrd;
            bool NtTmpltCrd;
            List<string> PmLns = new List<string>();
            string Rl = "";
            string[] RlLst = new string[] { "ROLE", "BE", "CM", "CSS", "DO", "FE", "MC", "QA", "UX", "PM", "PO" };
            var StryTsk = new Dictionary<string, string>();
            List<string> Tkns = new List<string>();
            string[] Tkns1 = new string[] { "" };
            string[] Tkns2 = new string[] { "" };
            string[] Tkns3 = new string[] { "" };
            List<string> KdLstsInclddInpt = new List<string>();
            List<string> KdLstsExclddInpt = new List<string>();

            FldsKd.Columns.Add("KdNm", typeof(string));
            FldsKd.Columns.Add("KdLbl", typeof(string));
            FldsKd.Columns.Add("KdVlu", typeof(int));
            FldsKd.Columns.Add("KdTm", typeof(int));
            FldsKd.Columns.Add("KdId", typeof(string));

            //Parms
            KdBrds = Cnfg.KdBrds;
            MspExe = Cnfg.MspExe;
            IncldCrdsChngdAftr = Cnfg.IncldCrdsChngdAftr;
            KdLstsInclddInpt = Cnfg.KdLstsInclddInpt;
            KdLstsExclddInpt = Cnfg.KdLstsExclddInpt;

            // Loop for each board
            foreach (string KdBrdId in KdBrds)
            {
                // Board data elements
                var KdBrd = new Board(KdBrdId);
                KdBrdNm = KdBrd.Name;

                // Find Trello lists to be included and excluded.  Write them to xls exec tab.
                bool IncldLst;
                if (KdLstsInclddInpt.Count == 0 && KdLstsExclddInpt.Count == 0)
                {
                    foreach (List KdBrdLst in KdBrd.Lists)
                    {
                        Cnfg.KdLstsIncldd.Add(KdBrdLst.Name);
                    }
                }

                if (KdLstsInclddInpt.Count != 0 && KdLstsExclddInpt.Count == 0)
                {
                    foreach (List BrdLst in KdBrd.Lists)
                    {
                        IncldLst = false;
                        foreach (string Tlst in KdLstsInclddInpt)
                        {
                            if (BrdLst.Name == Tlst)
                            {
                                IncldLst = true;
                            }
                        }

                        if (IncldLst)
                        {
                            Cnfg.KdLstsIncldd.Add(BrdLst.Name);
                        }
                        else
                        {
                            Cnfg.KdLstsExcldd.Add(BrdLst.Name);
                        }
                    }
                }

                if (KdLstsInclddInpt.Count == 0 && KdLstsExclddInpt.Count != 0)
                {
                    foreach (List BrdLst in KdBrd.Lists)
                    {
                        IncldLst = true;
                        foreach (string Tlst in KdLstsExclddInpt)
                        {
                            if (BrdLst.Name == Tlst)
                            {
                                IncldLst = false;
                            }
                        }

                        if (IncldLst)
                        {
                            Cnfg.KdLstsIncldd.Add(BrdLst.Name);
                        }
                        else
                        {
                            Cnfg.KdLstsExcldd.Add(BrdLst.Name);
                        }
                    }
                }

                if (KdLstsInclddInpt.Count != 0 && KdLstsExclddInpt.Count != 0)
                {
                    foreach (List BrdLst in KdBrd.Lists)
                    {
                        IncldLst = false;
                        foreach (string Tlst in KdLstsInclddInpt)
                        {
                            if (BrdLst.Name == Tlst)
                            {
                                IncldLst = true;
                            }
                        }

                        foreach (string Tlst in KdLstsExclddInpt)
                        {
                            if (BrdLst.Name == Tlst)
                            {
                                IncldLst = false;
                            }
                        }

                        if (IncldLst)
                        {
                            Cnfg.KdLstsIncldd.Add(BrdLst.Name);
                        }
                        else
                        {
                            Cnfg.KdLstsExcldd.Add(BrdLst.Name);
                        }
                    }
                }

                // Write lists of lists to xls tab Exec.
                int RwTtls = 14;
                oShtExec.Cells[RwTtls, 6] = "KD lists included in the Trello scan";
                iRw1 = RwTtls;
                foreach (string LstNm in Cnfg.KdLstsIncldd)
                {
                    iRw1++;
                    oShtExec.Cells[iRw1, 6] = LstNm;
                }

                oShtExec.Cells[RwTtls, 7] = "KD lists excluded from the Trello scan";
                iRw1 = RwTtls;
                foreach (string LstNm in Cnfg.KdLstsExcldd)
                {
                    iRw1++;
                    oShtExec.Cells[iRw1, 7] = LstNm;
                }


                // Loop for each card on board
                int cntCrds = KdBrd.Cards.Count();
                foreach (var card in KdBrd.Cards)
                {

                    // Get card info
                    CrdId = card.Id;
                    KdNm = card.Name;
                    CrdUrl = card.ShortUrl;
                    //CrdLstActvty = card.LastActivity.Value.Date;
                    Lst = card.List.Name;

                    // Get labels
                    Lbls = "";
                    foreach (var label in card.Labels)
                    {
                        if ((label.Name != null) && (label.Name != ""))
                        {
                            if (Lbls == "")
                            {
                                Lbls += label.Name;
                            }
                            else
                            {
                                Lbls += ";" + label.Name;
                            }
                        }
                    }

                    // Check that card is on an included Trello List.
                    bool KdLstFnd = true;
                    if (Cnfg.KdLstsIncldd.Count != 0)
                    {
                        KdLstFnd = false;
                        foreach (string KdLst in Cnfg.KdLstsIncldd)
                        {
                            if (card.List.Name.Equals(KdLst))
                            {
                                KdLstFnd = true;
                            }
                        }
                    }

                    // If debug then check if this card is on debug cards list
                    if (Dbg)
                    {
                        DbgIncldThsCrd = Crds.Contains(card.Id);
                    }
                    else
                    {
                        DbgIncldThsCrd = true;
                    }

                    if (DbgIncldThsCrd && KdLstFnd)
                    {
                        // Get parms from labels: KD TimeCritical, KD Value, KD status 
                        KdTm = 1;
                        KdVlu = 3;
                        foreach (var Lbl in card.Labels)
                        {
                            switch (Lbl.Name)
                            {
                                case "kd time:01":
                                    KdTm = 1;
                                    break;

                                case "kd time:02":
                                    KdTm = 2;
                                    break;

                                case "kd time:03":
                                    KdTm = 3;
                                    break;

                                case "kd time:05":
                                    KdTm = 5;
                                    break;

                                case "kd time:08":
                                    KdTm = 8;
                                    break;

                                case "kd time:13":
                                    KdTm = 13;
                                    break;

                                case "kd time:20":
                                    KdTm = 20;
                                    break;

                                case "kd time:40":
                                    KdTm = 40;
                                    break;

                                case "kd time:80":
                                    KdTm = 80;
                                    break;

                                case "kd value:01":
                                    KdVlu = 1;
                                    break;

                                case "kd value:02":
                                    KdVlu = 2;
                                    break;

                                case "kd value:03":
                                    KdVlu = 3;
                                    break;

                                case "kd value:05":
                                    KdVlu = 5;
                                    break;

                                case "kd value:08":
                                    KdVlu = 8;
                                    break;

                                case "kd value:13":
                                    KdVlu = 13;
                                    break;

                                case "kd value:20":
                                    KdVlu = 20;
                                    break;

                                case "kd value:40":
                                    KdVlu = 40;
                                    break;

                                case "kd value:80":
                                    KdVlu = 80;
                                    break;

                                case "kd status:cancelled":
                                    KdStts = "cancelled";
                                    break;

                                case "kd status:critical":
                                    KdStts = "critical";
                                    break;

                                case "kd status:done":
                                    KdStts = "done";
                                    break;

                                case "kd status:hold":
                                    KdStts = "hold";
                                    break;

                                case "kd status:on target":
                                    KdStts = "on target";
                                    break;

                                case "kd status:warning":
                                    KdStts = "warning";
                                    break;
                            }

                        }

                        // Get KD Label
                        KdLbl = "";
                        KdLblFnd = false;
                        foreach (var chkList in card.CheckLists)
                        {
                            // Read checklist items
                            foreach (var ChckItm in chkList.CheckItems)
                            {
                                string ChckItmNm = ChckItm.ToString();

                                // Parse tokens
                                BffrTkns = Regex.Split(ChckItmNm.ToString(), ":");
                                if (BffrTkns[0] == "KD Label")
                                {
                                    KdLbl = BffrTkns[1].Trim();
                                    KdLblFnd = true;
                                }
                            }
                        }

                        if (!KdLblFnd)
                        {
                            ErrMsg = "No KD Label defined.";
                        }

                        // Write card to sheet KDs From Trello
                        if (DbgIncldThsCrd && KdLstFnd)
                        {
                            oShtKdsFrmTrllo.Activate();
                            iRwAllCrds++;
                            oShtKdsFrmTrllo.Cells[iRwAllCrds, 1] = KdNm;
                            oShtKdsFrmTrllo.Cells[iRwAllCrds, 2] = KdLbl;
                            oShtKdsFrmTrllo.Cells[iRwAllCrds, 3] = KdVlu;
                            oShtKdsFrmTrllo.Cells[iRwAllCrds, 4] = KdTm;
                            oShtKdsFrmTrllo.Cells[iRwAllCrds, 5] = KdStts;
                            oShtKdsFrmTrllo.Cells[iRwAllCrds, 6] = CrdId;
                            oShtKdsFrmTrllo.Cells[iRwAllCrds, 7] = CrdUrl;
                            if(ErrMsg != "")
                            {
                                //oShtKdsFrmTrllo.Cells[iRwAllCrds, 7] = ErrMsg;
                                ((Excel.Range)oShtKdsFrmTrllo.Cells[iRwAllCrds, 8]).Value2 = ErrMsg;
                                ((Excel.Range)oShtKdsFrmTrllo.Cells[iRwAllCrds, 8
                                    ]).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                            }
                        }
                    }

                } 

            }

            // Write time KD scan done
            Tm = DateTime.Now.ToString("hh:mm:ss");
            Console.Write("\r\nTrello KD scan done at " + Tm);
            oShtExec.Cells[6, 2] = Tm;

        }

        //public void TrelloParseHours(
        //    string Prjct,
        //    string XlsTmpltPth,
        //    string XlsFlPth,
        //    string CnfgFlPth,
        //    Configuration Cnfg,
        //    String TmStrt
        //)
        //{
        //    if (Cnfg.DbgUsr)
        //    {
        //        Console.Write("\r\nEntering TrelloParseHours");
        //    }

        //    if (Cnfg.DbgUsr)
        //    {
        //        Console.Write("\r\nAfter worksheets defined");
        //    }

        //    // Variables
        //    string Assgnd = "";
        //    string[] BffrTkns = new string[] { "" };
        //    DataTable FldsWrkItm = new DataTable();
        //    //Board Brd;
        //    String BrdNm = "";
        //    List<string> Brds = new List<string>();
        //    string CrdId = "";
        //    DateTime CrdLstActvty;
        //    string CrdNm = "";
        //    string CrdPrty = "";
        //    List<string> Crds = new List<string>();
        //    //int CntSctnLn = 0;
        //    string CrdUrl = "";
        //    //string[] DscrptnLns;
        //    bool DbgIncldThsCrd;
        //    bool ErrrFnd = false;
        //    string ErrrTxt = "";
        //    float EstmtE = -99;
        //    float EstmtO = -99;
        //    float EstmtL = -99;
        //    float EstmtP = -99;
        //    float HrsActl = -99;
        //    Boolean HrsPrfxFnd;
        //    float HrsRmng = -99;
        //    DateTime IncldCrdsChngdAftr;
        //    int iRwAllCrds = 1;
        //    //int IndxOfAr;
        //    int iRw1;
        //    string Lbls;
        //    //string LblsCrd;
        //    string Lst;
        //    List<string> LstsExcldd = new List<string>();
        //    List<string> LstsIncldd = new List<string>();
        //    string MspExe;
        //    MatchCollection Mtch;
        //    float Nmbr1;
        //    bool NtMrkrCrd;
        //    bool NtTmpltCrd;
        //    Excel.Range oRng;
        //    //Excel.Range oRngStrt;
        //    //Excel.Range oRngEnd;
        //    List<string> PmLns = new List<string>();
        //    //bool Prsd;
        //    bool PrsnFnd = false;
        //    bool PstAllChckLstItms;
        //    string Rl = "";
        //    //bool RlFnd;
        //    string[] RlLst = new string[] { "ROLE", "BE", "CM", "CSS", "DO", "FE", "MC", "QA", "UX", "PM", "PO" };
        //    //string SctnTyp = "";
        //    string Str1;
        //    //String StryNm;
        //    var StryTsk = new Dictionary<string, string>();
        //    //string Tkn = "";
        //    List<string> Tkns = new List<string>();
        //    string[] Tkns1 = new string[] { "" };
        //    string[] Tkns2 = new string[] { "" };
        //    string[] Tkns3 = new string[] { "" };
        //    string Tm;
        //    //string TrllNm = ""; // Trello username
        //    List<string> TrlloLstsInclddInpt = new List<string>();
        //    List<string> TrlloLstsExclddInpt = new List<string>();
        //    bool TskFnd;
        //    string TskId = "";
        //    string TskNm = "";
        //    string Txt1;
        //    bool UpdtMspActls;
        //    bool UpdtMspMsrs;
        //    bool UpdtMspPrjctd;

        //    FldsWrkItm.Columns.Add("BrdNm", typeof(string));
        //    FldsWrkItm.Columns.Add("Lst", typeof(string));
        //    FldsWrkItm.Columns.Add("CrdNm", typeof(string));
        //    FldsWrkItm.Columns.Add("CrdPrty", typeof(string));
        //    FldsWrkItm.Columns.Add("Lbls", typeof(string));
        //    FldsWrkItm.Columns.Add("WrkPhsNm", typeof(string));
        //    FldsWrkItm.Columns.Add("TskNm", typeof(string));
        //    FldsWrkItm.Columns.Add("Rl", typeof(string));
        //    FldsWrkItm.Columns.Add("Assgnd", typeof(string));
        //    FldsWrkItm.Columns.Add("HrsActl", typeof(float));
        //    FldsWrkItm.Columns.Add("HrsRmnng", typeof(float));
        //    FldsWrkItm.Columns.Add("EstmtO", typeof(float));
        //    FldsWrkItm.Columns.Add("EstmtL", typeof(float));
        //    FldsWrkItm.Columns.Add("EstmtP", typeof(float));
        //    FldsWrkItm.Columns.Add("EstmtE", typeof(float));
        //    FldsWrkItm.Columns.Add("CrdId", typeof(string));
        //    FldsWrkItm.Columns.Add("ChckLstId", typeof(string));
        //    FldsWrkItm.Columns.Add("TskId", typeof(string));
        //    FldsWrkItm.Columns.Add("ErrrFnd", typeof(bool));
        //    FldsWrkItm.Columns.Add("ErrrTxt", typeof(string));
        //    FldsWrkItm.Columns.Add("ChckItmNm", typeof(string));
        //    FldsWrkItm.Columns.Add("CrdUrl", typeof(string));
        //    FldsWrkItm.Columns.Add("CrdLstActvty", typeof(DateTime));

        //    //Parms
        //    Brds = Cnfg.Brds;
        //    MspExe = Cnfg.MspExe;
        //    UpdtMspActls = Cnfg.UpdtMspActls;
        //    UpdtMspMsrs = Cnfg.UpdtMspMsrs;
        //    UpdtMspPrjctd = Cnfg.UpdtMspPrjctd;
        //    PstAllChckLstItms = Cnfg.PstAllChckLstItms;
        //    IncldCrdsChngdAftr = Cnfg.IncldCrdsChngdAftr;
        //    TrlloLstsInclddInpt = Cnfg.TrlloLstsInclddInpt;
        //    TrlloLstsExclddInpt = Cnfg.TrlloLstsExclddInpt;


        //    // Debug: Cards to be read for debug
        //    Boolean Dbg = false;
        //    Crds.Add("57e046ca970fb81e9e789ea6");

        //    // Loop for each board
        //    foreach (string BrdId in Brds)
        //    {
        //        // Board data elements
        //        var Brd = new Board(BrdId);
        //        BrdNm = Brd.Name;

        //        // Find Trello lists to be included and excluded.  Write them to xls exec tab.
        //        bool IncldLst;
        //        if (TrlloLstsInclddInpt.Count == 0 && TrlloLstsExclddInpt.Count == 0)
        //        {
        //            foreach (List BrdLst in Brd.Lists)
        //            {
        //                Cnfg.TrlloLstsIncldd.Add(BrdLst.Name);
        //            }
        //        }

        //        if (TrlloLstsInclddInpt.Count != 0 && TrlloLstsExclddInpt.Count == 0)
        //        {
        //            foreach (List BrdLst in Brd.Lists)
        //            {
        //                IncldLst = false;
        //                foreach (string Tlst in TrlloLstsInclddInpt)
        //                {
        //                    if (BrdLst.Name == Tlst)
        //                    {
        //                        IncldLst = true;
        //                    }
        //                }

        //                if (IncldLst)
        //                {
        //                    Cnfg.TrlloLstsIncldd.Add(BrdLst.Name);
        //                }
        //                else
        //                {
        //                    Cnfg.TrlloLstsExcldd.Add(BrdLst.Name);
        //                }
        //            }
        //        }

        //        if (TrlloLstsInclddInpt.Count == 0 && TrlloLstsExclddInpt.Count != 0)
        //        {
        //            foreach (List BrdLst in Brd.Lists)
        //            {
        //                IncldLst = true;
        //                foreach (string Tlst in TrlloLstsExclddInpt)
        //                {
        //                    if (BrdLst.Name == Tlst)
        //                    {
        //                        IncldLst = false;
        //                    }
        //                }

        //                if (IncldLst)
        //                {
        //                    Cnfg.TrlloLstsIncldd.Add(BrdLst.Name);
        //                }
        //                else
        //                {
        //                    Cnfg.TrlloLstsExcldd.Add(BrdLst.Name);
        //                }
        //            }
        //        }

        //        if (TrlloLstsInclddInpt.Count != 0 && TrlloLstsExclddInpt.Count != 0)
        //        {
        //            foreach (List BrdLst in Brd.Lists)
        //            {
        //                IncldLst = false;
        //                foreach (string Tlst in TrlloLstsInclddInpt)
        //                {
        //                    if (BrdLst.Name == Tlst)
        //                    {
        //                        IncldLst = true;
        //                    }
        //                }

        //                foreach (string Tlst in TrlloLstsExclddInpt)
        //                {
        //                    if (BrdLst.Name == Tlst)
        //                    {
        //                        IncldLst = false;
        //                    }
        //                }

        //                if (IncldLst)
        //                {
        //                    Cnfg.TrlloLstsIncldd.Add(BrdLst.Name);
        //                }
        //                else
        //                {
        //                    Cnfg.TrlloLstsExcldd.Add(BrdLst.Name);
        //                }
        //            }
        //        }

        //        // Find Trello lists that are open and closed.
        //        foreach (List BrdLst in Brd.Lists)
        //        {
        //            if (!Cnfg.TrlloLstsNtOpn.Contains(BrdLst.Name))
        //            {
        //                Cnfg.TrlloLstsOpn.Add(BrdLst.Name);
        //            }
        //        }

        //        // Write lists of lists to xls tab Exec.
        //        oShtExec.Cells[13, 1] = "Lists included in the Trello scan";
        //        iRw1 = 13;
        //        foreach (string LstNm in Cnfg.TrlloLstsIncldd)
        //        {
        //            iRw1++;
        //            oShtExec.Cells[iRw1, 1] = LstNm;
        //        }

        //        oShtExec.Cells[13, 2] = "Lists excluded from the Trello scan";
        //        iRw1 = 13;
        //        foreach (string LstNm in Cnfg.TrlloLstsExcldd)
        //        {
        //            iRw1++;
        //            oShtExec.Cells[iRw1, 2] = LstNm;
        //        }

        //        oShtExec.Cells[13, 3] = "Lists contining rejected cards.  Tasks for these cards will be deleted in Project Online.";
        //        iRw1 = 13;
        //        foreach (string LstNm in Cnfg.TrlloLstsRjctd)
        //        {
        //            iRw1++;
        //            oShtExec.Cells[iRw1, 3] = LstNm;
        //        }

        //        oShtExec.Cells[13, 4] = "Lists containing cards that are open";
        //        iRw1 = 13;
        //        foreach (string LstNm in Cnfg.TrlloLstsOpn)
        //        {
        //            iRw1++;
        //            oShtExec.Cells[iRw1, 4] = LstNm;
        //        }

        //        oShtExec.Cells[13, 5] = "Lists containing cards that are not open";
        //        iRw1 = 13;
        //        foreach (string LstNm in Cnfg.TrlloLstsNtOpn)
        //        {
        //            iRw1++;
        //            oShtExec.Cells[iRw1, 5] = LstNm;
        //        }


        //        // Loop for each card on board
        //        int cntCrds = Brd.Cards.Count();
        //        foreach (var card in Brd.Cards)
        //        {

        //            // Get card info
        //            CrdId = card.Id;
        //            CrdNm = card.Name;
        //            CrdUrl = card.ShortUrl;
        //            CrdLstActvty = card.LastActivity.Value.Date;
        //            Lst = card.List.Name;

        //            // Get labels
        //            Lbls = "";
        //            foreach (var label in card.Labels)
        //            {
        //                if ((label.Name != null) && (label.Name != ""))
        //                {
        //                    if (Lbls == "")
        //                    {
        //                        Lbls += label.Name;
        //                    }
        //                    else
        //                    {
        //                        Lbls += ";" + label.Name;
        //                    }
        //                }
        //            }

        //            // Get card priority
        //            CrdPrty = "unknown";
        //            if (Lbls.Contains("priority3-low") || Lbls.Contains("priority-low"))
        //            {
        //                CrdPrty = "low";
        //            }
        //            if (Lbls.Contains("priority2-med") || Lbls.Contains("priority-medium"))
        //            {
        //                CrdPrty = "medium";
        //            }
        //            if (Lbls.Contains("priority1-high") || Lbls.Contains("priority-high"))
        //            {
        //                CrdPrty = "high";
        //            }
        //            if (Lbls.Contains("priority0-critical") || Lbls.Contains("priority-critical"))
        //            {
        //                CrdPrty = "critical";
        //            }


        //            // Check that card is on an included Trello List.
        //            bool TrlloLstFnd = true;
        //            if (Cnfg.TrlloLstsIncldd.Count != 0)
        //            {
        //                TrlloLstFnd = false;
        //                foreach (string TrlloLst in Cnfg.TrlloLstsIncldd)
        //                {
        //                    if (card.List.Name.Equals(TrlloLst))
        //                    {
        //                        TrlloLstFnd = true;
        //                    }
        //                }
        //            }

        //            // check that card in included dates.
        //            bool InInclddDts = true;

        //            if (card.LastActivity <= IncldCrdsChngdAftr)
        //            {
        //                InInclddDts = false;
        //            }

        //            // Marker card
        //            NtMrkrCrd = true;
        //            if (Lbls.Contains("marker"))
        //            {
        //                NtMrkrCrd = false;
        //            }

        //            // Template3 card
        //            NtTmpltCrd = true;
        //            if (CrdNm.Contains("<TEMPLATE>"))
        //            {
        //                NtTmpltCrd = false;
        //            }

        //            // Write card to sheet All Cards
        //            if (TrlloLstFnd && InInclddDts && NtTmpltCrd && NtMrkrCrd)
        //            {
        //                oShtAllCrds.Activate();
        //                iRwAllCrds++;
        //                oShtAllCrds.Cells[iRwAllCrds, 1] = Brd.Name;
        //                oShtAllCrds.Cells[iRwAllCrds, 2] = card.List.Name;
        //                oShtAllCrds.Cells[iRwAllCrds, 3] = card.Name;
        //                oShtAllCrds.Cells[iRwAllCrds, 4] = card.Id;
        //                oShtAllCrds.Cells[iRwAllCrds, 5] = Lbls;
        //                //Console.WriteLine(card.Name);
        //                //Console.WriteLine(card.Id);
        //            }

        //            // If debug then check if this card is on debug cards list
        //            if (Dbg)
        //            {
        //                DbgIncldThsCrd = Crds.Contains(card.Id);
        //            }
        //            else
        //            {
        //                DbgIncldThsCrd = true;
        //            }

        //            if (DbgIncldThsCrd && TrlloLstFnd && InInclddDts && NtTmpltCrd && NtMrkrCrd)
        //            {
        //                // Read checklists
        //                foreach (var chkList in card.CheckLists)
        //                {
        //                    string WrkPhsNm = chkList.Name;
        //                    string ChckLstId = chkList.Id;

        //                    // Read checklist items
        //                    foreach (var ChckItm in chkList.CheckItems)
        //                    {
        //                        TskId = ChckItm.Id;
        //                        TskFnd = false;
        //                        // string ChckItmNm = ChckItm.Name;

        //                        // Replace "..." with "---" to accomodate cards on Tutor board
        //                        string ChckItmNm = ChckItm.ToString();
        //                        ChckItmNm = ChckItmNm.Replace("...", " --- ");

        //                        // Remove extra spaces from checklist item
        //                        while (ChckItmNm.LastIndexOf("  ") != -1)
        //                        {
        //                            ChckItmNm = ChckItmNm.Replace("  ", " ");
        //                        }

        //                        // Parse tokens
        //                        // Remove spaces in entered hrs
        //                        BffrTkns = Regex.Split(ChckItmNm.ToString(), " ");

        //                        Tkns.Clear();
        //                        HrsPrfxFnd = false;
        //                        foreach (var BffrTkn in BffrTkns)
        //                        {
        //                            if (HrsPrfxFnd)
        //                            {
        //                                Mtch = Regex.Matches(BffrTkn, "[0-9.,]");
        //                                if (BffrTkn.Length == Mtch.Count)
        //                                {
        //                                    Tkns[Tkns.Count - 1] = Tkns[Tkns.Count - 1] + BffrTkn;
        //                                }
        //                                else
        //                                {
        //                                    Tkns.Add(BffrTkn);
        //                                    if (BffrTkn.Contains("ar:") || BffrTkn.Contains("olp:"))
        //                                    {
        //                                        HrsPrfxFnd = true;
        //                                    }
        //                                    else
        //                                    {
        //                                        HrsPrfxFnd = false;
        //                                    }
        //                                }
        //                            }
        //                            else
        //                            {
        //                                Tkns.Add(BffrTkn);
        //                                if (BffrTkn.Contains("ar:") || BffrTkn.Contains("olp:"))
        //                                {
        //                                    HrsPrfxFnd = true;
        //                                }
        //                            }

        //                        }

        //                        // Select checklist items to process
        //                        if (PstAllChckLstItms)
        //                        {
        //                            // All checklist items are tasks
        //                            TskFnd = true;
        //                        }
        //                        else
        //                        {
        //                            // Checklist items containing "ar:" are tasks
        //                            foreach (string Tkn in Tkns)
        //                            {
        //                                if (Tkn.Contains("ar:"))
        //                                {
        //                                    TskFnd = true;
        //                                }
        //                            }
        //                        }

        //                        // If task, get data items
        //                        if (TskFnd)
        //                        {
        //                            //bool ArFnd;
        //                            //bool ArHrsFnd;
        //                            bool ErrrAr;
        //                            bool ErrrOlp;
        //                            //bool OlpFnd;
        //                            //bool OlpHrsFnd;
        //                            //ArFnd = false;
        //                            //ArHrsFnd = false;
        //                            Assgnd = "";
        //                            ErrrAr = false;
        //                            ErrrFnd = false;
        //                            ErrrOlp = false;
        //                            ErrrTxt = "";
        //                            EstmtO = -99;
        //                            EstmtL = -99;
        //                            EstmtP = -99;
        //                            EstmtE = -99;
        //                            HrsActl = -99;
        //                            HrsRmng = -99;
        //                            //OlpFnd = false;
        //                            //OlpHrsFnd = false;
        //                            PrsnFnd = false;
        //                            //RlFnd = false;
        //                            Rl = "";
        //                            string StrHrs = "";
        //                            TskNm = "";
        //                            string TskNmFnd = "not started";

        //                            // Token processing loop
        //                            int iTkn = 0;
        //                            do
        //                            {
        //                                // Initialize token used
        //                                bool TknUsd = false;

        //                                // Look for role
        //                                if (Array.IndexOf(RlLst, Tkns[iTkn].ToUpper()) > -1)
        //                                {
        //                                    //RlFnd = true;
        //                                    Rl = Tkns[iTkn];
        //                                    TknUsd = true;

        //                                    // If accumulating task name, end it
        //                                    if (TskNmFnd == "started")
        //                                    {
        //                                        TskNmFnd = "ended";
        //                                    }
        //                                }

        //                                // Look for assigned
        //                                if (Tkns[iTkn].Contains("@"))
        //                                {
        //                                    PrsnFnd = true;
        //                                    TknUsd = true;
        //                                    Assgnd = Tkns[iTkn];

        //                                    // Remove dashes from assigned (a common mistake)
        //                                    while (Assgnd.LastIndexOf("-") != -1)
        //                                    {
        //                                        Assgnd = Assgnd.Replace("-", "");
        //                                    }

        //                                    // If accumulating task name, end it
        //                                    if (TskNmFnd == "started")
        //                                    {
        //                                        TskNmFnd = "ended";
        //                                    }
        //                                }

        //                                // Actual/Remaining hrs
        //                                if (Tkns[iTkn].Contains("ar:"))
        //                                {
        //                                    //ArFnd = true;
        //                                    TknUsd = true;
        //                                    StrHrs = "";

        //                                    // If accumulating task name, end it
        //                                    if (TskNmFnd == "started")
        //                                    {
        //                                        TskNmFnd = "ended";
        //                                    }

        //                                    // Get string containing numbers
        //                                    if (Tkns[iTkn].Length != 3)
        //                                    {
        //                                        // hrs are in this token
        //                                        //ArHrsFnd = true;

        //                                        // Split to get the part with hrs
        //                                        Tkns1 = Regex.Split(Tkns[iTkn], ":");
        //                                        StrHrs = Tkns1[1];
        //                                    }
        //                                    else
        //                                    {
        //                                        // hrs are in following token.
        //                                        if (iTkn == Tkns.Count - 1)
        //                                        {
        //                                            // At last token so error
        //                                            ErrrAr = true;
        //                                        }
        //                                        else
        //                                        {
        //                                            // Look for next non-blank token
        //                                            {
        //                                                iTkn++;
        //                                                if (iTkn < Tkns.Count)
        //                                                {
        //                                                    if (Tkns[iTkn] != "")
        //                                                    {
        //                                                        StrHrs = Tkns[iTkn];
        //                                                    }
        //                                                }
        //                                            } while (StrHrs == "" && iTkn < Tkns.Count) ;
        //                                        }
        //                                    }

        //                                    //  If string found then parse hrs after removing all spaces
        //                                    if (StrHrs != "")
        //                                    {
        //                                        // Clean up hrs string.  Remove spaces and "...." found in some Tutor tasks
        //                                        Str1 = StrHrs.Replace(" ", "").Replace("....", "");

        //                                        // Split to get hrs numbers
        //                                        // For HrsActl ? = 0
        //                                        // For HrsRmng ? = 5.5h
        //                                        Tkns2 = Regex.Split(Str1, ",");
        //                                        if (Tkns2.Length == 2)
        //                                        {

        //                                            if (float.TryParse(Tkns2[0], out Nmbr1))
        //                                            {
        //                                                HrsActl = float.Parse(Tkns2[0]);
        //                                            }
        //                                            else
        //                                            {
        //                                                if (Tkns2[0] == "?" || Tkns2[0] == "??")
        //                                                {
        //                                                    HrsActl = 0f;
        //                                                }
        //                                                else
        //                                                {
        //                                                    ErrrAr = true;
        //                                                }
        //                                            }

        //                                            if (float.TryParse(Tkns2[1], out Nmbr1))
        //                                            {
        //                                                HrsRmng = float.Parse(Tkns2[1]);
        //                                            }
        //                                            else
        //                                            {
        //                                                if (Tkns2[1] == "?" || Tkns2[1] == "??")
        //                                                {
        //                                                    HrsRmng = 5.5f;
        //                                                }
        //                                                else
        //                                                {
        //                                                    ErrrAr = true;
        //                                                }
        //                                            }
        //                                        }
        //                                        else
        //                                        {
        //                                            ErrrAr = true;
        //                                        }
        //                                    }
        //                                    else
        //                                    {
        //                                        ErrrAr = true;
        //                                    }
        //                                }

        //                                // Estimates
        //                                if (Tkns[iTkn].Contains("olp:"))
        //                                {
        //                                    //OlpFnd = true;
        //                                    TknUsd = true;
        //                                    StrHrs = "";

        //                                    // If accumulating task name, end it
        //                                    if (TskNmFnd == "started")
        //                                    {
        //                                        TskNmFnd = "ended";
        //                                    }

        //                                    // Get string containing numbers
        //                                    if (Tkns[iTkn].Length != 4)
        //                                    {
        //                                        // hrs entries are in this token
        //                                        //OlpHrsFnd = true;
        //                                        Tkns1 = Regex.Split(Tkns[iTkn], ":");
        //                                        StrHrs = Tkns1[1];
        //                                    }
        //                                    else
        //                                    {
        //                                        // hrs are in following token.
        //                                        if (iTkn == Tkns.Count - 1)
        //                                        {
        //                                            // At last token so error
        //                                            ErrrOlp = true;
        //                                        }
        //                                        else
        //                                        {
        //                                            // Look for next non-blank token
        //                                            {
        //                                                iTkn++;
        //                                                if (iTkn < Tkns.Count)
        //                                                {
        //                                                    if (Tkns[iTkn] != "")
        //                                                    {
        //                                                        StrHrs = Tkns[iTkn];
        //                                                    }
        //                                                }
        //                                            } while (StrHrs == "" && iTkn < Tkns.Count) ;
        //                                        }
        //                                    }

        //                                    // If no error then parse hrs after removing all spaces
        //                                    if (!ErrrOlp)
        //                                    {

        //                                        Tkns2 = Regex.Split(StrHrs.Replace(" ", ""), ",");
        //                                        if (Tkns2.Length == 3)
        //                                        {

        //                                            if (float.TryParse(Tkns2[0], out Nmbr1))
        //                                            {
        //                                                EstmtO = float.Parse(Tkns2[0]);
        //                                            }
        //                                            else
        //                                            {
        //                                                if (Tkns2[0] == "?" || Tkns2[0] == "??")
        //                                                {
        //                                                    EstmtO = -99f;
        //                                                }
        //                                                else
        //                                                {
        //                                                    ErrrAr = true;
        //                                                }
        //                                            }

        //                                            if (float.TryParse(Tkns2[1], out Nmbr1))
        //                                            {
        //                                                EstmtL = float.Parse(Tkns2[1]);
        //                                            }
        //                                            else
        //                                            {
        //                                                if (Tkns2[1] == "?" || Tkns2[1] == "??")
        //                                                {
        //                                                    EstmtL = -99f;
        //                                                }
        //                                                else
        //                                                {
        //                                                    ErrrAr = true;
        //                                                }
        //                                            }

        //                                            if (float.TryParse(Tkns2[2], out Nmbr1))
        //                                            {
        //                                                EstmtP = float.Parse(Tkns2[2]);
        //                                            }
        //                                            else
        //                                            {
        //                                                if (Tkns2[2] == "?" || Tkns2[2] == "??")
        //                                                {
        //                                                    EstmtP = -99f;
        //                                                }
        //                                                else
        //                                                {
        //                                                    ErrrAr = true;
        //                                                }
        //                                            }


        //                                            //If we have good olp calculate EstmtE
        //                                            if (!ErrrOlp && EstmtO != -99f && EstmtL != -99f && EstmtP != -99f)
        //                                            {
        //                                                EstmtE = (EstmtL + EstmtO + 4 * EstmtP) / 6;
        //                                            }
        //                                            else
        //                                            {
        //                                                EstmtE = -99f;
        //                                            }

        //                                        }
        //                                        else
        //                                        {
        //                                            ErrrOlp = true;
        //                                        }
        //                                    }
        //                                }

        //                                // Look for dashes.  
        //                                if (Tkns[iTkn].Contains("---"))
        //                                {
        //                                    TknUsd = true;

        //                                    // If accumulating task name, end it
        //                                    if (TskNmFnd == "started")
        //                                    {
        //                                        TskNmFnd = "ended";
        //                                    }
        //                                }

        //                                if (Tkns[iTkn].Contains("--"))
        //                                {
        //                                    TknUsd = true;

        //                                    // If accumulating task name, end it
        //                                    if (TskNmFnd == "started")
        //                                    {
        //                                        TskNmFnd = "ended";
        //                                    }
        //                                }

        //                                // Look for task name = first token not used to last token not used within line.
        //                                if (!TknUsd && !(TskNmFnd == "ended"))
        //                                {

        //                                    if (TskNm == "")
        //                                    {
        //                                        TskNmFnd = "started";
        //                                        TskNm = Tkns[iTkn];
        //                                    }
        //                                    else
        //                                    {
        //                                        TskNm += " " + Tkns[iTkn];
        //                                    }
        //                                }

        //                                iTkn++;

        //                            } while (iTkn < Tkns.Count); // End token processing loop

        //                            // If errors found then fill error text
        //                            // Task name not found
        //                            if (TskNm == "")
        //                            {
        //                                ErrrFnd = true;
        //                                TskNm = "Task name not found";
        //                                Txt1 = "Task name not found";
        //                                if (ErrrTxt == "")
        //                                {
        //                                    ErrrTxt = Txt1;
        //                                }
        //                                else
        //                                {
        //                                    ErrrTxt += ", " + Txt1;
        //                                }
        //                            }


        //                            // Actual hrs with no assigned
        //                            if (!PrsnFnd && HrsActl != -99 && HrsActl != 0)
        //                            {
        //                                ErrrFnd = true;
        //                                Txt1 = "Person not found for task with actual hrs";
        //                                if (ErrrTxt == "")
        //                                {
        //                                    ErrrTxt = Txt1;
        //                                }
        //                                else
        //                                {
        //                                    ErrrTxt += ", " + Txt1;
        //                                }
        //                            }

        //                            // Actual or remaining hrs error
        //                            if (ErrrAr)
        //                            {
        //                                ErrrFnd = true;
        //                                Txt1 = "Actual or Remaining Hrs wrong";
        //                                if (ErrrTxt == "")
        //                                {
        //                                    ErrrTxt = Txt1;
        //                                }
        //                                else
        //                                {
        //                                    ErrrTxt += ", " + Txt1;
        //                                }
        //                            }

        //                            // Estimate error
        //                            if (ErrrOlp)
        //                            {
        //                                ErrrFnd = true;
        //                                Txt1 = "Estimate error";
        //                                if (ErrrTxt == "")
        //                                {
        //                                    ErrrTxt = Txt1;
        //                                }
        //                                else
        //                                {
        //                                    ErrrTxt += ", " + Txt1;
        //                                }
        //                            }

        //                            // Add assignment row to data table
        //                            FldsWrkItm.Rows.Add(BrdNm, Lst, CrdNm, CrdPrty, Lbls, WrkPhsNm, TskNm, Rl, Assgnd, HrsActl, HrsRmng, EstmtO, EstmtL, EstmtP, EstmtE, CrdId, ChckLstId, TskId, ErrrFnd, ErrrTxt, ChckItmNm, CrdUrl, CrdLstActvty);
        //                        } // end if assignment found
        //                    }  // End foreach checklist item
        //                } // End foreach checklist
        //            } // End If DbgIncldThsCrd
        //        } // End foreach card

        //    } // End foreach board

        //    // All boards completed.  
        //    // Write board assignments data table to sheet Tasks
        //    oShtTsks.Activate();
        //    iRw1 = 1;
        //    foreach (DataRow TblRw in FldsWrkItm.Rows)
        //    {
        //        iRw1 += 1;
        //        oShtTsks.Cells[iRw1, 1] = TblRw.Field<string>("BrdNm");
        //        oShtTsks.Cells[iRw1, 2] = TblRw.Field<string>("Lst");
        //        oShtTsks.Cells[iRw1, 3] = TblRw.Field<string>("CrdNm");
        //        oShtTsks.Cells[iRw1, 4] = TblRw.Field<string>("CrdPrty");
        //        oShtTsks.Cells[iRw1, 5] = TblRw.Field<string>("WrkPhsNm");

        //        // Task name: remove double quotes, leading - and +, leading blanks
        //        if (TblRw.Field<string>("TskNm").IndexOf("-") == 0 || TblRw.Field<string>("TskNm").IndexOf("+") == 0)
        //        {
        //            Str1 = TblRw.Field<string>("TskNm").Substring(1, TblRw.Field<string>("TskNm").Length - 1);
        //        }
        //        else
        //        {
        //            Str1 = TblRw.Field<string>("TskNm");
        //        }
        //        while (Str1.IndexOf(" ") == 0)
        //        {
        //            Str1 = Str1.Substring(1, Str1.Length - 1);
        //        }
        //        Str1 = Str1.Replace("\"", string.Empty);
        //        oShtTsks.Cells[iRw1, 6] = Str1;

        //        oShtTsks.Cells[iRw1, 7] = TblRw.Field<string>("Assgnd");
        //        oShtTsks.Cells[iRw1, 8] = TblRw.Field<float>("HrsActl");
        //        oShtTsks.Cells[iRw1, 9] = TblRw.Field<float>("HrsRmnng");
        //        oShtTsks.Cells[iRw1, 10] = TblRw.Field<string>("Rl");
        //        oShtTsks.Cells[iRw1, 11] = TblRw.Field<string>("Lbls");

        //        if (TblRw.Field<float>("EstmtO") != -99f)
        //        {
        //            oShtTsks.Cells[iRw1, 12] = TblRw.Field<float>("EstmtO");
        //        }
        //        else
        //        {
        //            oShtTsks.Cells[iRw1, 12] = null;
        //        }

        //        if (TblRw.Field<float>("EstmtL") != -99f)
        //        {
        //            oShtTsks.Cells[iRw1, 13] = TblRw.Field<float>("EstmtL");
        //        }
        //        else
        //        {
        //            oShtTsks.Cells[iRw1, 13] = null;
        //        }

        //        if (TblRw.Field<float>("EstmtP") != -99f)
        //        {
        //            oShtTsks.Cells[iRw1, 14] = TblRw.Field<float>("EstmtP");
        //        }
        //        else
        //        {
        //            oShtTsks.Cells[iRw1, 14] = null;
        //        }

        //        if (TblRw.Field<float>("EstmtE") != -99f)
        //        {
        //            oShtTsks.Cells[iRw1, 15] = TblRw.Field<float>("EstmtE");
        //        }
        //        else
        //        {
        //            oShtTsks.Cells[iRw1, 15] = null;
        //        }

        //        oShtTsks.Cells[iRw1, 16] = TblRw.Field<string>("CrdId");
        //        oShtTsks.Cells[iRw1, 17] = TblRw.Field<string>("ChckLstId");
        //        oShtTsks.Cells[iRw1, 18] = TblRw.Field<string>("TskId");
        //        oShtTsks.Cells[iRw1, 19] = TblRw.Field<bool>("ErrrFnd");

        //        if (TblRw.Field<string>("ErrrTxt") != "")
        //        {
        //            oShtTsks.Cells[iRw1, 20] = TblRw.Field<string>("ErrrTxt");
        //            // Color row if error
        //            oRng = (Excel.Range)oShtTsks.Range[oShtTsks.Cells[iRw1, 1], oShtTsks.Cells[iRw1, 15]];
        //            oRng.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
        //        }
        //        else
        //        {
        //            oShtTsks.Cells[iRw1, 20] = null;
        //        }

        //        // Checkitem name truncated to 100 chars due to MSP task name length limit
        //        if (TblRw.Field<string>("ChckItmNm").Length <= 250) // was 100 3-17
        //        {
        //            oShtTsks.Cells[iRw1, 21] = TblRw.Field<string>("ChckItmNm");

        //        }
        //        else
        //        {
        //            oShtTsks.Cells[iRw1, 21] = TblRw.Field<string>("ChckItmNm").Substring(0, 250);  // was 100 3-17-2017
        //        }

        //        // Card URL
        //        oShtTsks.Cells[iRw1, 22] = TblRw.Field<string>("CrdUrl");

        //        // Card last activity
        //        oShtTsks.Cells[iRw1, 23] = TblRw.Field<DateTime>("CrdLstActvty");

        //        // Sort field.  Sort checklists in order
        //        //Str1 = TblRw.Field<string>("WrkPhsNm")
        //        oShtTsks.Cells[iRw1, 24] = TblRw.Field<string>("CrdId") + "|"
        //            + TblRw.Field<string>("WrkPhsNm")
        //                .Replace("DEFINE", "1DEFINE")
        //                .Replace("DESIGN", "2DESIGN")
        //                .Replace("DECOMPOSE", "3DECOMPOSE")
        //                .Replace("DEVELOP", "4DEVELOP")
        //                .Replace("CODE", "4CODE")
        //                .Replace("TEST", "5TEST")
        //                .Replace("DOCUMENT", "6DOCUMENT")
        //            + "|" + TblRw.Field<string>("TskNm") + "|" + TblRw.Field<string>("Assgnd");
        //    }

        //    // Sort sheet [Time Records] by story
        //    Excel.Range oLastAACell;
        //    Excel.Range oFirstACell;

        //    oShtTsks.Activate();
        //    oSht = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

        //    //Get complete last Row in Sheet (Not last used just last)     
        //    int intRows = oSht.Rows.Count;

        //    //Get the last cell in col 1
        //    oLastAACell = (Excel.Range)oSht.Cells[intRows, 1];

        //    //Move curser up to the last cell in col 1 that is not blank.  This is the last data row
        //    oLastAACell = oLastAACell.End[Excel.XlDirection.xlUp];

        //    // Move cursor to last col in last data row.  This is one corner of the range.
        //    oLastAACell = (Excel.Range)oSht.Cells[oLastAACell.Row, 24];

        //    //Get First Cell of Data (A2)
        //    oFirstACell = (Excel.Range)oSht.Cells[2, 1];

        //    //Get Entire Range of Data
        //    oRng = (Excel.Range)oSht.Range[oFirstACell, oLastAACell];
        //    //oRng.Select();

        //    //Sort the range by the sort column
        //    oRng.Sort(oRng.Columns[24, Type.Missing], Excel.XlSortOrder.xlAscending);

        //    // Write board assignments data table to sheet Export 10K
        //    oShtExprt10K.Activate();
        //    iRw1 = 1;
        //    foreach (DataRow TblRw in FldsWrkItm.Rows)
        //    {
        //        iRw1 += 1;
        //        oShtExprt10K.Cells[iRw1, 1] = TblRw.Field<string>("CrdNm");
        //        oShtExprt10K.Cells[iRw1, 2] = TblRw.Field<string>("WrkPhsNm");

        //        // Task name: remove double quotes, leading - and +, leading blanks
        //        if (TblRw.Field<string>("TskNm").IndexOf("-") == 0 || TblRw.Field<string>("TskNm").IndexOf("+") == 0)
        //        {
        //            Str1 = TblRw.Field<string>("TskNm").Substring(1, TblRw.Field<string>("TskNm").Length - 1);
        //        }
        //        else
        //        {
        //            Str1 = TblRw.Field<string>("TskNm");
        //        }
        //        while (Str1.IndexOf(" ") == 0)
        //        {
        //            Str1 = Str1.Substring(1, Str1.Length - 1);
        //        }
        //        Str1 = Str1.Replace("\"", string.Empty);
        //        oShtExprt10K.Cells[iRw1, 3] = Str1;

        //        oShtExprt10K.Cells[iRw1, 4] = TblRw.Field<string>("Assgnd");
        //        oShtExprt10K.Cells[iRw1, 5] = TblRw.Field<float>("HrsActl");
        //        oShtExprt10K.Cells[iRw1, 6] = TblRw.Field<float>("HrsRmnng");
        //        oShtExprt10K.Cells[iRw1, 7] = TblRw.Field<string>("Lbls");

        //        oShtExprt10K.Cells[iRw1, 8] = TblRw.Field<bool>("ErrrFnd");

        //        if (TblRw.Field<string>("ErrrTxt") != "")
        //        {
        //            oShtExprt10K.Cells[iRw1, 9] = TblRw.Field<string>("ErrrTxt");
        //            // Color row if error
        //            oRng = (Excel.Range)oShtExprt10K.Range[oShtExprt10K.Cells[iRw1, 1], oShtExprt10K.Cells[iRw1, 15]];
        //            oRng.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
        //        }
        //        else
        //        {
        //            oShtExprt10K.Cells[iRw1, 9] = null;
        //        }

        //        // Card URL
        //        oShtExprt10K.Cells[iRw1, 10] = TblRw.Field<string>("CrdUrl");

        //    }

        //    // Save workbook
        //    oWB.Save();

        //    // Write time workitem scan done
        //    Tm = DateTime.Now.ToString("hh:mm:ss");
        //    Console.Write("\r\nTrello workitem scan done; starting MSP update at " + Tm);
        //    oShtExec.Cells[5, 2] = Tm;

        //}
    }
}

/*
https://developers.trello.com/get-started/start-building -- to get the application key

https://bitbucket.org/gregsdennis/manatee.trello/wiki/Usage

Key:
703231aefbeb477f24ca6871addf1699
Token:
908b5b92fc503cacde7d840d7ad424b2b57175ad2e27dfb254e681f32a73063e
Secret:
246253945bc949c4bb154ae4d1de61c50ab42dd98d56fe1999e1f00b0f6cfd9d
*/
