using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using Manatee.Trello;
using Manatee.Trello.ManateeJson;
using Manatee.Trello.WebApi;
using Excel = Microsoft.Office.Interop.Excel;


namespace MspUpdate
{
    public class TrelloConnection
    {
        Board board = null;
        Dictionary<string, List<string>> cardSorted = new Dictionary<string, List<string>>();
        List<string> noLabelTickets = new List<string>();

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

            var serializer = new ManateeSerializer();
            TrelloConfiguration.Serializer = serializer;
            TrelloConfiguration.Deserializer = serializer;
            TrelloConfiguration.JsonFactory = new ManateeFactory();
            TrelloConfiguration.RestClientProvider = new WebApiClientProvider();
            TrelloAuthorization.Default.AppKey = Cnfg.TrlloAppKy;
            TrelloAuthorization.Default.UserToken = Cnfg.TrlloUsrTkn;

            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel._Worksheet oShtAllCrds;
            Microsoft.Office.Interop.Excel._Worksheet oShtTrlloScnErrr;
            Microsoft.Office.Interop.Excel._Worksheet oShtExec;
            Microsoft.Office.Interop.Excel._Worksheet oShtExprt10K;
            Microsoft.Office.Interop.Excel._Worksheet oShtTsks;

            // Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;

            if (Cnfg.DbgUsr)
            {
                Console.Write("\r\nAfter worksheets defined");
            }

            // Variables
            string Assgnd = "";
            string[] BffrTkns = new string[] { "" };
            DataTable BrdAssgnmnts = new DataTable();
            String BrdNm = "";
            List<string> Brds = new List<string>();
            string CrdId = "";
            DateTime CrdLstActvty;
            string CrdNm = "";
            string CrdPrty = "";
            List<string> Crds = new List<string>();
            //int CntSctnLn = 0;
            string CrdUrl = "";
            //string[] DscrptnLns;
            bool DbgIncldThsCrd;
            bool ErrrFnd = false;
            string ErrrTxt = "";
            float EstmtE = -99;
            float EstmtO = -99;
            float EstmtL = -99;
            float EstmtP = -99;
            float HrsActl = -99;
            Boolean HrsPrfxFnd;
            float HrsRmng = -99;
            DateTime IncldCrdsChngdAftr;
            int iRwAllCrds = 1;
            //int IndxOfAr;
            int iRw1;
            string Lbls;
            //string LblsCrd;
            string Lst;
            string MspExe;
            MatchCollection Mtch;
            float Nmbr1;
            Excel.Range oRng;
            //Excel.Range oRngStrt;
            //Excel.Range oRngEnd;
            List<string> PmLns = new List<string>();
            //bool Prsd;
            bool PrsnFnd = false;
            bool PstAllChckLstItms;
            string Rl = "";
            //bool RlFnd;
            string[] RlLst = new string[] { "ROLE", "BE", "CM", "CSS", "DO", "FE", "MC", "QA", "UX", "PM", "PO" };
            //string SctnTyp = "";
            string Str1;
            //String StryNm;
            var StryTsk = new Dictionary<string, string>();
            //string Tkn = "";
            List<string> Tkns = new List<string>();
            string[] Tkns1 = new string[] { "" };
            string[] Tkns2 = new string[] { "" };
            string[] Tkns3 = new string[] { "" };
            string Tm;
            //string TrllNm = ""; // Trello username
            List<string> TrlloLstsIncldd = new List<string>();
            List<string> TrlloLstsExcldd = new List<string>();
            bool TskFnd;
            string TskId = "";
            string TskNm = "";
            string Txt1;
            bool UpdtMspActls;
            bool UpdtMspMsrs;
            bool UpdtMspPrjctd;

            BrdAssgnmnts.Columns.Add("BrdNm", typeof(string));
            BrdAssgnmnts.Columns.Add("Lst", typeof(string));
            BrdAssgnmnts.Columns.Add("CrdNm", typeof(string));
            BrdAssgnmnts.Columns.Add("CrdPrty", typeof(string));
            BrdAssgnmnts.Columns.Add("Lbls", typeof(string));
            BrdAssgnmnts.Columns.Add("WrkPhsNm", typeof(string));
            BrdAssgnmnts.Columns.Add("TskNm", typeof(string));
            BrdAssgnmnts.Columns.Add("Rl", typeof(string));
            BrdAssgnmnts.Columns.Add("Assgnd", typeof(string));
            BrdAssgnmnts.Columns.Add("HrsActl", typeof(float));
            BrdAssgnmnts.Columns.Add("HrsRmnng", typeof(float));
            BrdAssgnmnts.Columns.Add("EstmtO", typeof(float));
            BrdAssgnmnts.Columns.Add("EstmtL", typeof(float));
            BrdAssgnmnts.Columns.Add("EstmtP", typeof(float));
            BrdAssgnmnts.Columns.Add("EstmtE", typeof(float));
            BrdAssgnmnts.Columns.Add("CrdId", typeof(string));
            BrdAssgnmnts.Columns.Add("ChckLstId", typeof(string));
            BrdAssgnmnts.Columns.Add("TskId", typeof(string));
            BrdAssgnmnts.Columns.Add("ErrrFnd", typeof(bool));
            BrdAssgnmnts.Columns.Add("ErrrTxt", typeof(string));
            BrdAssgnmnts.Columns.Add("ChckItmNm", typeof(string));
            BrdAssgnmnts.Columns.Add("CrdUrl", typeof(string));
            BrdAssgnmnts.Columns.Add("CrdLstActvty", typeof(DateTime));

            //Parms
            Brds = Cnfg.Brds;
            MspExe = Cnfg.MspExe;
            UpdtMspActls = Cnfg.UpdtMspActls;
            UpdtMspMsrs = Cnfg.UpdtMspMsrs;
            UpdtMspPrjctd = Cnfg.UpdtMspPrjctd;
            PstAllChckLstItms = Cnfg.PstAllChckLstItms;
            IncldCrdsChngdAftr = Cnfg.IncldCrdsChngdAftr;
            TrlloLstsIncldd = Cnfg.TrlloLstsIncldd;
            TrlloLstsExcldd = Cnfg.TrlloLstsExcldd;



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

            if (Cnfg.DbgUsr)
            {
                Console.Write("\r\nAfter Excel started");
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

            // Worksheets
            oShtExec = oWB.Worksheets["Exec"];
            oShtExec.Cells[2, 2] = Prjct;
            oShtExec.Cells[3, 2] = CnfgFlPth;
            oShtExec.Cells[4, 2] = TmStrt;

            oShtExprt10K = oWB.Worksheets["Export 10K"];
            oShtExprt10K.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
            oShtExprt10K.Columns[1].ColumnWidth = 20;
            oShtExprt10K.Columns[2].ColumnWidth = 10;
            oShtExprt10K.Columns[3].ColumnWidth = 20;
            oShtExprt10K.Columns[4].ColumnWidth = 15;
            oShtExprt10K.Columns[5].ColumnWidth = 5;
            oShtExprt10K.Columns[6].ColumnWidth = 4;
            oShtExprt10K.Columns[7].ColumnWidth = 5;
            oShtExprt10K.Columns[8].ColumnWidth = 6;
            oShtExprt10K.Columns[9].ColumnWidth = 16;
            oShtExprt10K.Columns[10].ColumnWidth = 10;
            oShtExprt10K.Cells[1, 1] = "Card Name";
            oShtExprt10K.Cells[1, 2] = "Workphase Name";
            oShtExprt10K.Cells[1, 3] = "Task Name";
            oShtExprt10K.Cells[1, 4] = "Assigned";
            oShtExprt10K.Cells[1, 5] = "Hrs Actl";
            oShtExprt10K.Cells[1, 6] = "Hrs Rmnng";
            oShtExprt10K.Cells[1, 7] = "Labels";
            oShtExprt10K.Cells[1, 8] = "Error Flag";
            oShtExprt10K.Cells[1, 9] = "Error Desc";
            oShtExprt10K.Cells[1, 10] = "Card URL";

            oShtTrlloScnErrr = oWB.Worksheets["Trello Scan Errors"];
            oShtTrlloScnErrr.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
            oShtTrlloScnErrr.Columns[1].ColumnWidth = 20;
            oShtTrlloScnErrr.Columns[2].ColumnWidth = 30;
            oShtTrlloScnErrr.Columns[3].ColumnWidth = 40;
            oShtTrlloScnErrr.Cells[1, 1] = "Trello Task #";
            oShtTrlloScnErrr.Cells[1, 2] = "Code Section";
            oShtTrlloScnErrr.Cells[1, 3] = "Error Description";

            oShtAllCrds = oWB.Worksheets["All Cards"];
            oShtAllCrds.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
            oShtAllCrds.Columns[1].ColumnWidth = 20;
            oShtAllCrds.Columns[2].ColumnWidth = 30;
            oShtAllCrds.Columns[3].ColumnWidth = 40;
            oShtAllCrds.Columns[4].ColumnWidth = 25;
            oShtAllCrds.Cells[1, 1] = "Board";
            oShtAllCrds.Cells[1, 2] = "List";
            oShtAllCrds.Cells[1, 3] = "Card Name";
            oShtAllCrds.Cells[1, 4] = "Card ID";
            oShtAllCrds.Cells[1, 5] = "Labels";

            oShtTsks = oWB.Worksheets["Tasks"];
            oShtTsks.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
            oShtTsks.Columns[1].ColumnWidth = 10;
            oShtTsks.Columns[2].ColumnWidth = 10;
            oShtTsks.Columns[3].ColumnWidth = 20;
            oShtTsks.Columns[4].ColumnWidth = 10;
            oShtTsks.Columns[5].ColumnWidth = 10;
            oShtTsks.Columns[6].ColumnWidth = 10;
            oShtTsks.Columns[7].ColumnWidth = 5;
            oShtTsks.Columns[8].ColumnWidth = 5;
            oShtTsks.Columns[9].ColumnWidth = 4;
            oShtTsks.Columns[10].ColumnWidth = 10;
            oShtTsks.Columns[11].ColumnWidth = 5;
            oShtTsks.Columns[12].ColumnWidth = 5;
            oShtTsks.Columns[13].ColumnWidth = 5;
            oShtTsks.Columns[14].ColumnWidth = 6;
            oShtTsks.Columns[15].ColumnWidth = 6;
            oShtTsks.Columns[16].ColumnWidth = 6;
            oShtTsks.Columns[17].ColumnWidth = 6;
            oShtTsks.Columns[18].ColumnWidth = 6;
            oShtTsks.Columns[19].ColumnWidth = 6;
            oShtTsks.Columns[20].ColumnWidth = 16;
            oShtTsks.Columns[21].ColumnWidth = 6;
            oShtTsks.Columns[22].ColumnWidth = 10;
            oShtTsks.Columns[23].ColumnWidth = 10;
            oShtTsks.Columns[24].ColumnWidth = 10;
            oShtTsks.Cells[1, 1] = "Board Name";
            oShtTsks.Cells[1, 2] = "List Name";
            oShtTsks.Cells[1, 3] = "Card Name";
            oShtTsks.Cells[1, 4] = "Priority";
            oShtTsks.Cells[1, 5] = "Workphase Name";
            oShtTsks.Cells[1, 6] = "Task Name";
            oShtTsks.Cells[1, 7] = "Assigned";
            oShtTsks.Cells[1, 8] = "Hrs Actl";
            oShtTsks.Cells[1, 9] = "Hrs Rmnng";
            oShtTsks.Cells[1, 10] = "Role";
            oShtTsks.Cells[1, 11] = "Labels";
            oShtTsks.Cells[1, 12] = "EstmtO";
            oShtTsks.Cells[1, 13] = "EstmtL";
            oShtTsks.Cells[1, 14] = "EstmtP";
            oShtTsks.Cells[1, 15] = "EstmtE";
            oShtTsks.Cells[1, 16] = "Card ID";
            oShtTsks.Cells[1, 17] = "Checklist ID";
            oShtTsks.Cells[1, 18] = "Task ID";
            oShtTsks.Cells[1, 19] = "Error Flag";
            oShtTsks.Cells[1, 20] = "Error Desc";
            oShtTsks.Cells[1, 21] = "CheckItem Name";
            oShtTsks.Cells[1, 22] = "Card URL";
            oShtTsks.Cells[1, 23] = "Card Last Activity";
            oShtTsks.Cells[1, 24] = "Sort";

            foreach (string BrdId in Brds)
            {
                //Read Board and cards in a board
                var board = new Board(BrdId);

                // Board data elements
                try
                {
                    BrdNm = board.Name;
                }
                catch (Manatee.Trello.Exceptions.TrelloInteractionException xcptn)
                {
                    Console.WriteLine("Error: '{0}'", xcptn.Message);
                }


                // Debug: Cards to be read for debug
                Boolean Dbg = false;
                Crds.Add("57e046ca970fb81e9e789ea6");

                int cntCrds = board.Cards.Count();
                foreach (var card in board.Cards)
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
                                Lbls += "," + label.Name;
                            }
                        }
                    }

                    // Get card priority
                    CrdPrty = "unknown";
                    if (Lbls.Contains("priority-high"))
                    {
                        CrdPrty = "high";
                    }
                    if (Lbls.Contains("priority-medium"))
                    {
                        CrdPrty = "medium";
                    }
                    if (Lbls.Contains("priority-low"))
                    {
                        CrdPrty = "low";
                    }


                    // Trello Lists included and excluded
                    bool TrlloLstFnd = true;
                    if (TrlloLstsIncldd.Count != 0)
                    {
                        TrlloLstFnd = false;
                        foreach (string TrlloLst in TrlloLstsIncldd)
                        {
                            if (card.List.Name.Equals(TrlloLst))
                            {
                                TrlloLstFnd = true;
                            }
                        }
                    }

                    if (TrlloLstsExcldd.Count != 0)
                    {
                        if (TrlloLstFnd == true)
                        {
                            foreach (string TrlloLst in TrlloLstsExcldd)
                            {
                                if (card.List.Name.Equals(TrlloLst))
                                {
                                    TrlloLstFnd = false;
                                }
                            }
                        }
                    }

                    // check that card in included dates.
                    bool InInclddDts = true;

                    if (card.LastActivity <= IncldCrdsChngdAftr)
                    {
                        InInclddDts = false;
                    }

                    // Write card to sheet All Cards
                    if (TrlloLstFnd && InInclddDts)
                    {
                        oShtAllCrds.Activate();
                        iRwAllCrds++;
                        oShtAllCrds.Cells[iRwAllCrds, 1] = board.Name;
                        oShtAllCrds.Cells[iRwAllCrds, 2] = card.List.Name;
                        oShtAllCrds.Cells[iRwAllCrds, 3] = card.Name;
                        oShtAllCrds.Cells[iRwAllCrds, 4] = card.Id;
                        oShtAllCrds.Cells[iRwAllCrds, 5] = Lbls;
                        //Console.WriteLine(card.Name);
                        //Console.WriteLine(card.Id);
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

                    if (DbgIncldThsCrd && TrlloLstFnd && InInclddDts)
                    {
                        // Read checklists
                        foreach (var chkList in card.CheckLists)
                        {
                            string WrkPhsNm = chkList.Name;
                            string ChckLstId = chkList.Id;

                            // Read checklist items
                            foreach (var ChckItm in chkList.CheckItems)
                            {
                                TskId = ChckItm.Id;
                                TskFnd = false;
                                // string ChckItmNm = ChckItm.Name;

                                // Replace "..." with "---" to accomodate cards on Tutor board
                                string ChckItmNm = ChckItm.ToString();
                                ChckItmNm = ChckItmNm.Replace("...", " --- ");

                                // Remove extra spaces from checklist item
                                while (ChckItmNm.LastIndexOf("  ") != -1)
                                {
                                    ChckItmNm = ChckItmNm.Replace("  ", " ");
                                }

                                // Parse tokens
                                // Remove spaces in entered hrs
                                BffrTkns = Regex.Split(ChckItmNm.ToString(), " ");

                                Tkns.Clear();
                                HrsPrfxFnd = false;
                                foreach (var BffrTkn in BffrTkns)
                                {
                                    if (HrsPrfxFnd)
                                    {
                                        Mtch = Regex.Matches(BffrTkn, "[0-9.,]");
                                        if (BffrTkn.Length == Mtch.Count)
                                        {
                                            Tkns[Tkns.Count-1] = Tkns[Tkns.Count-1] + BffrTkn;
                                        }
                                        else
                                        {
                                            Tkns.Add(BffrTkn);
                                            if(BffrTkn.Contains("ar:") || BffrTkn.Contains("olp:"))
                                            {
                                                HrsPrfxFnd = true;
                                            }
                                            else
                                            {
                                                HrsPrfxFnd = false;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Tkns.Add(BffrTkn);
                                        if (BffrTkn.Contains("ar:") || BffrTkn.Contains("olp:"))
                                        {
                                            HrsPrfxFnd = true;
                                        }
                                    }

                                }

                                // Select checklist items to process
                                if (PstAllChckLstItms)
                                {
                                    // All checklist items are tasks
                                    TskFnd = true;
                                }
                                else
                                {
                                    // Checklist items containing "ar:" are tasks
                                    foreach (string Tkn in Tkns)
                                    {
                                        if (Tkn.Contains("ar:"))
                                        {
                                            TskFnd = true;
                                        }
                                    }
                                }

                                // If task, get data items
                                if (TskFnd)
                                {
                                    //bool ArFnd;
                                    //bool ArHrsFnd;
                                    bool ErrrAr;
                                    bool ErrrOlp;
                                    //bool OlpFnd;
                                    //bool OlpHrsFnd;
                                    //ArFnd = false;
                                    //ArHrsFnd = false;
                                    Assgnd = "";
                                    ErrrAr = false;
                                    ErrrFnd = false;
                                    ErrrOlp = false;
                                    ErrrTxt = "";
                                    EstmtO = -99;
                                    EstmtL = -99;
                                    EstmtP = -99;
                                    EstmtE = -99;
                                    HrsActl = -99;
                                    HrsRmng = -99;
                                    //OlpFnd = false;
                                    //OlpHrsFnd = false;
                                    PrsnFnd = false;
                                    //RlFnd = false;
                                    Rl = "";
                                    string StrHrs = "";
                                    TskNm = "";
                                    string TskNmFnd = "not started";

                                    // Token processing loop
                                    int iTkn = 0;
                                    do
                                    {
                                        // Initialize token used
                                        bool TknUsd = false;

                                        // Look for role
                                        if (Array.IndexOf(RlLst, Tkns[iTkn].ToUpper()) > -1)
                                        {
                                            //RlFnd = true;
                                            Rl = Tkns[iTkn];
                                            TknUsd = true;

                                            // If accumulating task name, end it
                                            if (TskNmFnd == "started")
                                            {
                                                TskNmFnd = "ended";
                                            }
                                        }

                                        // Look for assigned
                                        if (Tkns[iTkn].Contains("@"))
                                        {
                                            PrsnFnd = true;
                                            TknUsd = true;
                                            Assgnd = Tkns[iTkn];

                                            // Remove dashes from assigned (a common mistake)
                                            while (Assgnd.LastIndexOf("-") != -1)
                                            {
                                                Assgnd = Assgnd.Replace("-", "");
                                            }

                                            // If accumulating task name, end it
                                            if (TskNmFnd == "started")
                                            {
                                                TskNmFnd = "ended";
                                            }
                                        }

                                        // Actual/Remaining hrs
                                        if (Tkns[iTkn].Contains("ar:"))
                                        {
                                            //ArFnd = true;
                                            TknUsd = true;
                                            StrHrs = "";

                                            // If accumulating task name, end it
                                            if (TskNmFnd == "started")
                                            {
                                                TskNmFnd = "ended";
                                            }

                                            // Get string containing numbers
                                            if (Tkns[iTkn].Length != 3)
                                            {
                                                // hrs are in this token
                                                //ArHrsFnd = true;

                                                // Split to get the part with hrs
                                                Tkns1 = Regex.Split(Tkns[iTkn], ":");
                                                StrHrs = Tkns1[1];
                                            }
                                            else
                                            {
                                                // hrs are in following token.
                                                if (iTkn == Tkns.Count - 1)
                                                {
                                                    // At last token so error
                                                    ErrrAr = true;
                                                }
                                                else
                                                {
                                                    // Look for next non-blank token
                                                    {
                                                        iTkn++;
                                                        if (iTkn < Tkns.Count)
                                                        {
                                                            if (Tkns[iTkn] != "")
                                                            {
                                                                StrHrs = Tkns[iTkn];
                                                            }
                                                        }
                                                    } while (StrHrs == "" && iTkn < Tkns.Count) ;
                                                }
                                            }

                                            //  If string found then parse hrs after removing all spaces
                                            if (StrHrs != "")
                                            {
                                                // Clean up hrs string.  Remove spaces and "...." found in some Tutor tasks
                                                Str1 = StrHrs.Replace(" ", "").Replace("....", "");

                                                // Split to get hrs numbers
                                                // For HrsActl ? = 0
                                                // For HrsRmng ? = 5.5h
                                                Tkns2 = Regex.Split(Str1, ",");
                                                if (Tkns2.Length == 2)
                                                {

                                                    if (float.TryParse(Tkns2[0], out Nmbr1))
                                                    {
                                                        HrsActl = float.Parse(Tkns2[0]);
                                                    }
                                                    else
                                                    {
                                                        if (Tkns2[0] == "?" || Tkns2[0] == "??")
                                                        {
                                                            HrsActl = 0f;
                                                        }
                                                        else
                                                        {
                                                            ErrrAr = true;
                                                        }
                                                    }

                                                    if (float.TryParse(Tkns2[1], out Nmbr1))
                                                    {
                                                        HrsRmng = float.Parse(Tkns2[1]);
                                                    }
                                                    else
                                                    {
                                                        if (Tkns2[1] == "?" || Tkns2[1] == "??")
                                                        {
                                                            HrsRmng = 5.5f;
                                                        }
                                                        else
                                                        {
                                                            ErrrAr = true;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    ErrrAr = true;
                                                }
                                            }
                                            else
                                            {
                                                ErrrAr = true;
                                            }
                                        }

                                        // Estimates
                                        if (Tkns[iTkn].Contains("olp:"))
                                        {
                                            //OlpFnd = true;
                                            TknUsd = true;
                                            StrHrs = "";

                                            // If accumulating task name, end it
                                            if (TskNmFnd == "started")
                                            {
                                                TskNmFnd = "ended";
                                            }

                                            // Get string containing numbers
                                            if (Tkns[iTkn].Length != 4)
                                            {
                                                // hrs entries are in this token
                                                //OlpHrsFnd = true;
                                                Tkns1 = Regex.Split(Tkns[iTkn], ":");
                                                StrHrs = Tkns1[1];
                                            }
                                            else
                                            {
                                                // hrs are in following token.
                                                if (iTkn == Tkns.Count - 1)
                                                {
                                                    // At last token so error
                                                    ErrrOlp = true;
                                                }
                                                else
                                                {
                                                    // Look for next non-blank token
                                                    {
                                                        iTkn++;
                                                        if (iTkn < Tkns.Count)
                                                        {
                                                            if (Tkns[iTkn] != "")
                                                            {
                                                                StrHrs = Tkns[iTkn];
                                                            }
                                                        }
                                                    } while (StrHrs == "" && iTkn < Tkns.Count) ;
                                                }
                                            }

                                            // If no error then parse hrs after removing all spaces
                                            if (!ErrrOlp)
                                            {

                                                Tkns2 = Regex.Split(StrHrs.Replace(" ", ""), ",");
                                                if (Tkns2.Length == 3)
                                                {

                                                    if (float.TryParse(Tkns2[0], out Nmbr1))
                                                    {
                                                        EstmtO = float.Parse(Tkns2[0]);
                                                    }
                                                    else
                                                    {
                                                        if (Tkns2[0] == "?" || Tkns2[0] == "??")
                                                        {
                                                            EstmtO = -99f;
                                                        }
                                                        else
                                                        {
                                                            ErrrAr = true;
                                                        }
                                                    }

                                                    if (float.TryParse(Tkns2[1], out Nmbr1))
                                                    {
                                                        EstmtL = float.Parse(Tkns2[1]);
                                                    }
                                                    else
                                                    {
                                                        if (Tkns2[1] == "?" || Tkns2[1] == "??")
                                                        {
                                                            EstmtL = -99f;
                                                        }
                                                        else
                                                        {
                                                            ErrrAr = true;
                                                        }
                                                    }

                                                    if (float.TryParse(Tkns2[2], out Nmbr1))
                                                    {
                                                        EstmtP = float.Parse(Tkns2[2]);
                                                    }
                                                    else
                                                    {
                                                        if (Tkns2[2] == "?" || Tkns2[2] == "??")
                                                        {
                                                            EstmtP = -99f;
                                                        }
                                                        else
                                                        {
                                                            ErrrAr = true;
                                                        }
                                                    }


                                                    //If we have good olp calculate EstmtE
                                                    if (!ErrrOlp && EstmtO != -99f && EstmtL != -99f && EstmtP != -99f)
                                                    {
                                                        EstmtE = (EstmtL + EstmtO + 4 * EstmtP) / 6;
                                                    }
                                                    else
                                                    {
                                                        EstmtE = -99f;
                                                    }

                                                }
                                                else
                                                {
                                                    ErrrOlp = true;
                                                }
                                            }
                                        }

                                        // Look for dashes.  
                                        if (Tkns[iTkn].Contains("---"))
                                        {
                                            TknUsd = true;

                                            // If accumulating task name, end it
                                            if (TskNmFnd == "started")
                                            {
                                                TskNmFnd = "ended";
                                            }
                                        }

                                        if (Tkns[iTkn].Contains("--"))
                                        {
                                            TknUsd = true;

                                            // If accumulating task name, end it
                                            if (TskNmFnd == "started")
                                            {
                                                TskNmFnd = "ended";
                                            }
                                        }

                                        // Look for task name = first token not used to last token not used within line.
                                        if (!TknUsd && !(TskNmFnd == "ended"))
                                        {

                                            if (TskNm == "")
                                            {
                                                TskNmFnd = "started";
                                                TskNm = Tkns[iTkn];
                                            }
                                            else
                                            {
                                                TskNm += " " + Tkns[iTkn];
                                            }
                                        }

                                        iTkn++;

                                    } while (iTkn < Tkns.Count); // End token processing loop

                                    // If errors found then fill error text
                                    // Task name not found
                                    if (TskNm == "")
                                    {
                                        ErrrFnd = true;
                                        TskNm = "Task name not found";
                                        Txt1 = "Task name not found";
                                        if (ErrrTxt == "")
                                        {
                                            ErrrTxt = Txt1;
                                        }
                                        else
                                        {
                                            ErrrTxt += ", " + Txt1;
                                        }
                                    }


                                    // Actual hrs with no assigned
                                    if (!PrsnFnd && HrsActl != -99 && HrsActl != 0)
                                    {
                                        ErrrFnd = true;
                                        Txt1 = "Person not found for task with actual hrs";
                                        if (ErrrTxt == "")
                                        {
                                            ErrrTxt = Txt1;
                                        }
                                        else
                                        {
                                            ErrrTxt += ", " + Txt1;
                                        }
                                    }

                                    // Actual or remaining hrs error
                                    if (ErrrAr)
                                    {
                                        ErrrFnd = true;
                                        Txt1 = "Actual or Remaining Hrs wrong";
                                        if (ErrrTxt == "")
                                        {
                                            ErrrTxt = Txt1;
                                        }
                                        else
                                        {
                                            ErrrTxt += ", " + Txt1;
                                        }
                                    }

                                    // Estimate error
                                    if (ErrrOlp)
                                    {
                                        ErrrFnd = true;
                                        Txt1 = "Estimate error";
                                        if (ErrrTxt == "")
                                        {
                                            ErrrTxt = Txt1;
                                        }
                                        else
                                        {
                                            ErrrTxt += ", " + Txt1;
                                        }
                                    }

                                    // Add assignment row to data table
                                    BrdAssgnmnts.Rows.Add(BrdNm, Lst, CrdNm, CrdPrty, Lbls, WrkPhsNm, TskNm, Rl, Assgnd, HrsActl, HrsRmng, EstmtO, EstmtL, EstmtP, EstmtE, CrdId, ChckLstId, TskId, ErrrFnd, ErrrTxt, ChckItmNm, CrdUrl, CrdLstActvty);
                                } // end if assignment found
                            }  // End foreach checklist item
                        } // End foreach checklist
                    } // End If DbgIncldThsCrd
                } // End foreach card

            }

            // All boards completed.  
            // Write board assignments data table to sheet Tasks
            oShtTsks.Activate();
            iRw1 = 1;
            foreach (DataRow TblRw in BrdAssgnmnts.Rows)
            {
                iRw1 += 1;
                oShtTsks.Cells[iRw1, 1] = TblRw.Field<string>("BrdNm");
                oShtTsks.Cells[iRw1, 2] = TblRw.Field<string>("Lst");
                oShtTsks.Cells[iRw1, 3] = TblRw.Field<string>("CrdNm");
                oShtTsks.Cells[iRw1, 4] = TblRw.Field<string>("CrdPrty");
                oShtTsks.Cells[iRw1, 5] = TblRw.Field<string>("WrkPhsNm");

                // Task name: remove double quotes, leading - and +, leading blanks
                if (TblRw.Field<string>("TskNm").IndexOf("-") == 0 || TblRw.Field<string>("TskNm").IndexOf("+") == 0)
                {
                    Str1 = TblRw.Field<string>("TskNm").Substring(1, TblRw.Field<string>("TskNm").Length - 1);
                }
                else
                {
                    Str1 = TblRw.Field<string>("TskNm");
                }
                while (Str1.IndexOf(" ") == 0)
                {
                    Str1 = Str1.Substring(1, Str1.Length - 1);
                }
                Str1 = Str1.Replace("\"", string.Empty);
                oShtTsks.Cells[iRw1, 6] = Str1;

                oShtTsks.Cells[iRw1, 7] = TblRw.Field<string>("Assgnd");
                oShtTsks.Cells[iRw1, 8] = TblRw.Field<float>("HrsActl");
                oShtTsks.Cells[iRw1, 9] = TblRw.Field<float>("HrsRmnng");
                oShtTsks.Cells[iRw1, 10] = TblRw.Field<string>("Rl");
                oShtTsks.Cells[iRw1, 11] = TblRw.Field<string>("Lbls");

                if (TblRw.Field<float>("EstmtO") != -99f)
                {
                    oShtTsks.Cells[iRw1, 12] = TblRw.Field<float>("EstmtO");
                }
                else
                {
                    oShtTsks.Cells[iRw1, 12] = null;
                }

                if (TblRw.Field<float>("EstmtL") != -99f)
                {
                    oShtTsks.Cells[iRw1, 13] = TblRw.Field<float>("EstmtL");
                }
                else
                {
                    oShtTsks.Cells[iRw1, 13] = null;
                }

                if (TblRw.Field<float>("EstmtP") != -99f)
                {
                    oShtTsks.Cells[iRw1, 14] = TblRw.Field<float>("EstmtP");
                }
                else
                {
                    oShtTsks.Cells[iRw1, 14] = null;
                }

                if (TblRw.Field<float>("EstmtE") != -99f)
                {
                    oShtTsks.Cells[iRw1, 15] = TblRw.Field<float>("EstmtE");
                }
                else
                {
                    oShtTsks.Cells[iRw1, 15] = null;
                }

                oShtTsks.Cells[iRw1, 16] = TblRw.Field<string>("CrdId");
                oShtTsks.Cells[iRw1, 17] = TblRw.Field<string>("ChckLstId");
                oShtTsks.Cells[iRw1, 18] = TblRw.Field<string>("TskId");
                oShtTsks.Cells[iRw1, 19] = TblRw.Field<bool>("ErrrFnd");

                if (TblRw.Field<string>("ErrrTxt") != "")
                {
                    oShtTsks.Cells[iRw1, 20] = TblRw.Field<string>("ErrrTxt");
                    // Color row if error
                    oRng = (Excel.Range)oShtTsks.Range[oShtTsks.Cells[iRw1, 1], oShtTsks.Cells[iRw1, 15]];
                    oRng.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                }
                else
                {
                    oShtTsks.Cells[iRw1, 20] = null;
                }

                // Checkitem name truncated to 100 chars due to MSP task name length limit
                if (TblRw.Field<string>("ChckItmNm").Length <= 250) // was 100 3-17
                {
                    oShtTsks.Cells[iRw1, 21] = TblRw.Field<string>("ChckItmNm");

                }
                else
                {
                    oShtTsks.Cells[iRw1, 21] = TblRw.Field<string>("ChckItmNm").Substring(0, 250);  // was 100 3-17-2017
                }

                // Card URL
                oShtTsks.Cells[iRw1, 22] = TblRw.Field<string>("CrdUrl");

                // Card last activity
                oShtTsks.Cells[iRw1, 23] = TblRw.Field<DateTime>("CrdLstActvty");

                // Sort field.  Sort checklists in order
                //Str1 = TblRw.Field<string>("WrkPhsNm")
                oShtTsks.Cells[iRw1, 24] = TblRw.Field<string>("CrdId") + "|"
                    + TblRw.Field<string>("WrkPhsNm")
                        .Replace("DEFINE", "1DEFINE")
                        .Replace("DESIGN", "2DESIGN")
                        .Replace("DECOMPOSE", "3DECOMPOSE")
                        .Replace("DEVELOP", "4DEVELOP")
                        .Replace("CODE", "4CODE")
                        .Replace("TEST", "5TEST")
                        .Replace("DOCUMENT", "6DOCUMENT")
                    + "|" + TblRw.Field<string>("TskNm") + "|" + TblRw.Field<string>("Assgnd");
            }

            // Sort sheet [Time Records] by story
            Excel.Range oLastAACell;
            Excel.Range oFirstACell;

            oShtTsks.Activate();
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

            //Get complete last Row in Sheet (Not last used just last)     
            int intRows = oSheet.Rows.Count;

            //Get the last cell in col 1
            oLastAACell = (Excel.Range)oSheet.Cells[intRows, 1];

            //Move curser up to the last cell in col 1 that is not blank.  This is the last data row
            oLastAACell = oLastAACell.End[Excel.XlDirection.xlUp];

            // Move cursor to last col in last data row.  This is one corner of the range.
            oLastAACell = (Excel.Range)oSheet.Cells[oLastAACell.Row, 24];

            //Get First Cell of Data (A2)
            oFirstACell = (Excel.Range)oSheet.Cells[2, 1];

            //Get Entire Range of Data
            oRng = (Excel.Range)oSheet.Range[oFirstACell, oLastAACell];
            //oRng.Select();

            //Sort the range by the sort column
            oRng.Sort(oRng.Columns[24, Type.Missing], Excel.XlSortOrder.xlAscending);

            // Write board assignments data table to sheet Export 10K
            oShtExprt10K.Activate();
            iRw1 = 1;
            foreach (DataRow TblRw in BrdAssgnmnts.Rows)
            {
                iRw1 += 1;
                oShtExprt10K.Cells[iRw1, 1] = TblRw.Field<string>("CrdNm");
                oShtExprt10K.Cells[iRw1, 2] = TblRw.Field<string>("WrkPhsNm");

                // Task name: remove double quotes, leading - and +, leading blanks
                if (TblRw.Field<string>("TskNm").IndexOf("-") == 0 || TblRw.Field<string>("TskNm").IndexOf("+") == 0)
                {
                    Str1 = TblRw.Field<string>("TskNm").Substring(1, TblRw.Field<string>("TskNm").Length - 1);
                }
                else
                {
                    Str1 = TblRw.Field<string>("TskNm");
                }
                while (Str1.IndexOf(" ") == 0)
                {
                    Str1 = Str1.Substring(1, Str1.Length - 1);
                }
                Str1 = Str1.Replace("\"", string.Empty);
                oShtExprt10K.Cells[iRw1, 3] = Str1;

                oShtExprt10K.Cells[iRw1, 4] = TblRw.Field<string>("Assgnd");
                oShtExprt10K.Cells[iRw1, 5] = TblRw.Field<float>("HrsActl");
                oShtExprt10K.Cells[iRw1, 6] = TblRw.Field<float>("HrsRmnng");
                oShtExprt10K.Cells[iRw1, 7] = TblRw.Field<string>("Lbls");

                oShtExprt10K.Cells[iRw1, 8] = TblRw.Field<bool>("ErrrFnd");

                if (TblRw.Field<string>("ErrrTxt") != "")
                {
                    oShtExprt10K.Cells[iRw1, 9] = TblRw.Field<string>("ErrrTxt");
                    // Color row if error
                    oRng = (Excel.Range)oShtExprt10K.Range[oShtExprt10K.Cells[iRw1, 1], oShtExprt10K.Cells[iRw1, 15]];
                    oRng.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                }
                else
                {
                    oShtExprt10K.Cells[iRw1, 9] = null;
                }

                // Card URL
                oShtExprt10K.Cells[iRw1, 10] = TblRw.Field<string>("CrdUrl");

            }

            // Save workbook
            oWB.Save();

            // Write time Trello scan done
            Tm = DateTime.Now.ToString("hh:mm:ss");
            Console.Write("\r\nTrello scan done; starting MSP update at " + Tm);
            oShtExec.Cells[5, 2] = Tm;


            // Update Project Online
            // DateTime DtUpdt = new DateTime(2017, 1, 3);
            // string PrjctMsp = "BruceP Test 2";
            if (UpdtMspActls || UpdtMspPrjctd || UpdtMspMsrs)
            {
                oXL.Run("Update_Project_Online", MspExe);
            }

            // Write time completed
            Tm = DateTime.Now.ToString("hh:mm:ss");
            Console.Write("\r\nMSP update completed at " + Tm);
            oShtExec.Cells[9, 2] = Tm;

            // Close xls
            //oXL.Visible = false;
            oWB.Save();
            oWB.Close();
            oXL.Quit();

        }
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
