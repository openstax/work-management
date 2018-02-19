using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;


namespace MspUpdate
{
    public class Configuration
    {
        public List<string> Brds = new List<string>();
        public bool DbgUsr;
        public DateTime IncldCrdsChngdAftr;
        public string MspExe;
        public string MspPrjctNm;
        public bool PrmsOk;
        public bool PstAllChckLstItms;
        public bool PstChckItmNm;
        public string TrlloAppKy;
        public List<string> TrlloLstsIncldd = new List<string>();
        public List<string> TrlloLstsInclddInpt = new List<string>();
        public List<string> TrlloLstsExcldd = new List<string>();
        public List<string> TrlloLstsExclddInpt = new List<string>();
        public List<string> TrlloLstsNtOpn = new List<string>();
        public List<string> TrlloLstsOpn = new List<string>();
        public List<string> TrlloLstsRjctd = new List<string>();
        public string TrlloUsrTkn;
        public DateTime UpdtDt;
        public bool UpdtMspActls;
        public bool UpdtMspMsrs;
        public bool UpdtMspPrjctd;
        public string XlsFlNm;
        public string XlsOutptDrctry;
    }

    public class ParmsFound
    {
        public bool Brds;
        public bool IncldCrdsChngdAftr;
        public bool MspExe;
        public bool MspPrjctNm;
        public bool PstAllChckLstItms;
        public bool PstChckItmNm;
        public bool TrlloAppKy;
        public bool TrlloLstsInclddInpt;
        public bool TrlloLstsExclddInpt;
        public bool TrlloLstsNtOpn;
        public bool TrlloLstsRjctd;
        public bool TrlloUsrTkn;
        public bool UpdtDt;
        public bool UpdtMspActls;
        public bool UpdtMspMsrs;
        public bool UpdtMspPrjctd;
        public bool XlsFlNm;
        public bool XlsOutptDrctry;
    }

    class Program
    {

        struct Result
        {
            public int add;
            public int multiply;
        }
        struct MyStruct
        {
            public List<string> MyList;
            public int MyInt;

            public MyStruct(int myInt)
            {
                MyInt = myInt;
                MyList = new List<string>();
            }
        }
        struct Configuration2
        {
            public List<string> Brds;
            public List<string> TrlloLstsIncldd;
            public bool UpdtMspActls;
            public bool UpdtMspPrjctd;
            public DateTime UpdtDt;
            public string MspPrjctNm;
            public bool PstAllChckLstItms;
            public DateTime IncldCrdsChngdAftr;
            public string XlsFlNm;
        }

        /*
         * public void SetDebugExec()
        {
            DbgExec = true;
        }
        */

        static void Main(string[] args)
        {
            TrelloConnection trelloConnect = new TrelloConnection();
            List<string> Brds = new List<string>();
            var Cnfg = new Configuration();
            string CnfgFlPth = "";
            bool CnslInpt = true;
            bool DbgExec = false;
            DateTime DtNll = new DateTime(1800, 1, 1);
            DateTime DtToUpdt = DateTime.Today;
            // string strDtToUpdt;
            DateTime IncldCrdsChngdAftr = DtNll;
            string MspExe = "";
            //string MspPrjctNm;
            string Prjct = "";
            string PrjctMsp = "";
            var Prms = new Dictionary<string, string>();
            bool PstAllChckLstItms = false;
            bool PstChckItmNm = false;
            string reply = "";
            string Str1 = "";
            string[] StrArry;
            string MsrLblsStr = "";
            string[] Tkns = new string[] { "" };
            string TmStrt;
            List<string> TrlloLstsIncldd = new List<string>(); // Trllo lists included
            bool UpdtMsp = true;
            bool UsrNmFnd = false;
            string XlsFlPth = "";
            string XlsOutptDrctry = "";
            string XlsTmpltPth = "";

            #if DEBUG
                DbgExec = true;
            #endif

            TmStrt = DateTime.Now.ToString("hh:mm:ss");
            Console.Write("Start at " + TmStrt + "\r\n");

            // Get args.  THIS IS BROKEN.
            if (args.Length != 0)
            {
                CnslInpt = false;
                Prjct = args[0].ToUpper();
                UpdtMsp = Convert.ToBoolean(args[1]);
                PstAllChckLstItms = Convert.ToBoolean(args[2]);
                PstChckItmNm = Convert.ToBoolean(args[3]);
                DtToUpdt = Convert.ToDateTime(args[4]);
                IncldCrdsChngdAftr = Convert.ToDateTime(args[5]);
                XlsTmpltPth = args[6];
                //XlsFlNm = args[7];
                MspExe = args[8];
                if (args[9].Contains(","))
                {
                    StrArry = args[9].Split(',');
                    for (int i1 = 0; i1 < StrArry.Length; i1++)
                    {
                        TrlloLstsIncldd.Add(StrArry[i1]);
                    }
                }
                MsrLblsStr = args[10];
            }

            //MessageBox.Show("Entering Main: UpdtMsp=" + Convert.ToString(UpdtMsp) + " PstAllChckLstItms=" + Convert.ToString(PstAllChckLstItms) + " PstChckItmNm=" + Convert.ToString(PstChckItmNm));

            if (CnslInpt)
            {
                // File paths
                int Indx1 = Application.StartupPath.LastIndexOf("\\", 
                    Application.StartupPath.Length, Application.StartupPath.Length, StringComparison.OrdinalIgnoreCase);
                CnfgFlPth = Application.StartupPath.Substring(0, Indx1) + "\\UTS MSP Config.txt";
                XlsTmpltPth = Application.StartupPath + "\\UTS MSP Update Template.xlsm";
                if (DbgExec)
                {
                    CnfgFlPth = "C:\\Users\\Bruce Pike Rice\\Documents\\Repos\\work-management\\JiraInteraction\\bin\\UTS MSP Config.txt";
                    XlsTmpltPth = "C:\\Users\\Bruce Pike Rice\\Documents\\Repos\\work-management\\JiraInteraction\\UTS MSP Update Template.xlsm";
                }

                // Parms entered by user
                Console.WriteLine("MSP Update"); Console.WriteLine("");
                Console.WriteLine("Select Project:"); Console.WriteLine("");
                Console.WriteLine("                [Tutor]: type Tutor");
                Console.WriteLine("                  [CNX]: type CNX");
                Console.WriteLine("                [Books]: type Books");
                Console.WriteLine("[Business Intelligence]: type BIT");
                Console.WriteLine("             [Research]: type Research");
                Console.WriteLine("             [UTS Test]: type UTS Test");
                Console.WriteLine("");
                Console.WriteLine("To quit, type EXIT"); Console.WriteLine("");
                Console.WriteLine("...........................................................");

            }
            bool loop = true;
            while (loop)
            {
                if (CnslInpt)
                {
                    loop = true;
                    reply = Console.ReadLine(); Console.WriteLine("");
                    Prjct = reply.ToUpper();
                }

                loop = false;

                // Select project
                switch (Prjct)
                {
                    case "UTS TEST":
                        Console.WriteLine("> UTS Test selected. Stand by...");

                        // Get configuration
                        Cnfg = Program.Read_Config(Prjct, CnfgFlPth);

                        // Read boards
                        if (Cnfg.PrmsOk)
                        {
                            XlsFlPth = Cnfg.XlsOutptDrctry + Cnfg.XlsFlNm;
                            trelloConnect.CteReadBoard(Prjct, XlsTmpltPth, XlsFlPth, CnfgFlPth, Cnfg, TmStrt);
                        }

                        loop = false;
                        break;

                    case "TUTOR":
                        Console.WriteLine("> Tutor selected. Stand by...");

                        // Get configuration
                        Cnfg = Program.Read_Config(Prjct, CnfgFlPth);

                        // Read boards
                        if (Cnfg.PrmsOk)
                        {
                            XlsFlPth = Cnfg.XlsOutptDrctry + Cnfg.XlsFlNm;
                            trelloConnect.CteReadBoard(Prjct, XlsTmpltPth, XlsFlPth, CnfgFlPth, Cnfg, TmStrt);
                        }
                        loop = false;
                        break;

                    case "CNX":
                        Console.WriteLine("> CNX selected. Stand by....");

                        // Get configuration
                        Cnfg = Program.Read_Config(Prjct, CnfgFlPth);
                        
                        // Read boards
                        if (Cnfg.PrmsOk)
                        {
                            XlsFlPth = Cnfg.XlsOutptDrctry + Cnfg.XlsFlNm;
                            trelloConnect.CteReadBoard(Prjct, XlsTmpltPth, XlsFlPth, CnfgFlPth, Cnfg, TmStrt);
                        }

                        loop = false;
                        break;

                    case "BOOKS":
                        Console.WriteLine("> Books selected. Stand by....");

                        // Get configuration
                        Cnfg = Program.Read_Config(Prjct, CnfgFlPth);

                        // Read boards
                        if (Cnfg.PrmsOk)
                        {
                            XlsFlPth = Cnfg.XlsOutptDrctry + Cnfg.XlsFlNm;
                            trelloConnect.CteReadBoard(Prjct, XlsTmpltPth, XlsFlPth, CnfgFlPth, Cnfg, TmStrt);
                        }

                        loop = false;
                        break;

                    case "RESEARCH":
                        Console.WriteLine("> Research selected. Stand by...");

                        // Get configuration
                        Cnfg = Program.Read_Config(Prjct, CnfgFlPth);
                        
                        // Read boards
                        if (Cnfg.PrmsOk)
                        {
                            XlsFlPth = Cnfg.XlsOutptDrctry + Cnfg.XlsFlNm;
                            trelloConnect.CteReadBoard(Prjct, XlsTmpltPth, XlsFlPth, CnfgFlPth, Cnfg, TmStrt);
                        }

                        loop = false;
                        break;

                    case "BIT":
                        Console.WriteLine("> BIT selected. Stand by...");

                        // Get configuration
                        Cnfg = Program.Read_Config(Prjct, CnfgFlPth);
                        
                        // Read boards
                        if (Cnfg.PrmsOk)
                        {
                            XlsFlPth = Cnfg.XlsOutptDrctry + Cnfg.XlsFlNm;
                            trelloConnect.CteReadBoard(Prjct, XlsTmpltPth, XlsFlPth, CnfgFlPth, Cnfg, TmStrt);
                        }

                        loop = false;
                        break;
                    case "EXIT":
                        Console.WriteLine("Exiting. Stand by...(hit return)");
                        loop = false;
                        break;

                    default:
                        Console.WriteLine(reply + " is not an option. NOTE: these options are case-sensitive, dude. Try again.");

                        break;
                }


            }

            // Pause if console input 
            if (CnslInpt)
            {
                if (Cnfg.PrmsOk)
                {
                    Console.WriteLine("\r\nAll done...hit Enter to exit");
                }
                else
                {
                    Console.WriteLine("\r\nAborting...hit Enter to exit");
                }
                Console.ReadLine();
            }
        }
        private static Configuration Read_Config(string Prjct, string CnfgFlPth)
        {
            List<string> Lst = new List<string>();
            var Cnfg = new Configuration();
            DateTime DtNll = new DateTime(1800, 1, 1);
            int cntr = 0;
            string ln;
            var PrmsFnd = new ParmsFound();
            var Prms = new Dictionary<string, string>();
            string[] Tkns = new string[] { "" };
            string Str1;

            // Initial values
            Cnfg.Brds.Clear();
            Cnfg.DbgUsr = false;
            Cnfg.IncldCrdsChngdAftr = DateTime.Today.AddDays(-3);  
            Cnfg.MspExe = "";
            Cnfg.MspPrjctNm = "";
            Cnfg.PstAllChckLstItms = false;
            Cnfg.PstChckItmNm = false;
            Cnfg.TrlloAppKy = "";
            Cnfg.TrlloLstsExcldd.Clear();
            Cnfg.TrlloLstsExclddInpt.Clear();
            Cnfg.TrlloLstsIncldd.Clear();
            Cnfg.TrlloLstsInclddInpt.Clear();
            Cnfg.TrlloLstsNtOpn.Clear();
            Cnfg.TrlloLstsRjctd.Clear();
            Cnfg.TrlloUsrTkn = "";
            Cnfg.UpdtDt = DateTime.Today;
            Cnfg.UpdtMspActls = false;
            Cnfg.UpdtMspMsrs = false;
            Cnfg.UpdtMspPrjctd = false;
            Cnfg.XlsFlNm = "";
            Cnfg.XlsOutptDrctry = "";

            PrmsFnd.Brds = true;
            PrmsFnd.IncldCrdsChngdAftr = true;
            PrmsFnd.MspExe = true;
            PrmsFnd.MspPrjctNm = true;
            PrmsFnd.PstAllChckLstItms = true;
            PrmsFnd.PstChckItmNm = true;
            PrmsFnd.TrlloAppKy = true;
            PrmsFnd.TrlloLstsExclddInpt = true;
            PrmsFnd.TrlloLstsInclddInpt = true;
            PrmsFnd.TrlloLstsNtOpn = true;
            PrmsFnd.TrlloLstsRjctd = true;
            PrmsFnd.TrlloUsrTkn = true;
            PrmsFnd.UpdtDt = true;
            PrmsFnd.UpdtMspActls = true;
            PrmsFnd.UpdtMspMsrs = true;
            PrmsFnd.UpdtMspPrjctd = true;
            PrmsFnd.XlsFlNm = true;
            PrmsFnd.XlsOutptDrctry = true;

            // Read config file
            // string UsrNm = Environment.UserName;
            System.IO.StreamReader file = new System.IO.StreamReader(CnfgFlPth);
            while ((ln = file.ReadLine()) != null)
            {
                // Fill Prms dictionary with input parms.  
                // Exclude Measure Condition since not needed here and there may be several of these.
                // Will get them during Update Measures and Update Projection.
                if (ln.IndexOf("=") != -1)
                    {
                        Tkns = Regex.Split(ln, "=");
                        if (Tkns[0].IndexOf("Measure Condition") == -1)
                        {
                            Prms.Add(Tkns[0], Tkns[1]);
                            cntr++;
                        }
                    }
            }
            file.Close();

            //Get parms
            try
            {
                Lst.Clear();
                Str1 = Prms[Prjct + ":Boards"];
                Tkns = Regex.Split(Str1, ";");
                foreach (string Tkn in Tkns)
                {
                    if (Tkn != "")
                    {
                        Cnfg.Brds.Add(Tkn);
                    }
                }
            }
            catch
            {
                PrmsFnd.Brds = false;
            }

            if (Cnfg.Brds.Count == 0)
            {
                PrmsFnd.Brds = false;
            }
            try
            {
                Str1 = Prms[Prjct + ":Debug"];
                Cnfg.DbgUsr = Convert.ToBoolean(Str1);
            }
            catch
            {
                Cnfg.DbgUsr = false;
            }

            try
            {
                Str1 = Prms[Prjct + ":Include Cards Changed After"];
                Cnfg.IncldCrdsChngdAftr = Convert.ToDateTime(Str1);
            }
            catch
            {
                PrmsFnd.IncldCrdsChngdAftr = false;
            }

            try
            {
                Cnfg.MspExe = Prms["MS Project Exe"];
            }
            catch
            {
                PrmsFnd.MspExe = false;
            }

            try
            {
                Cnfg.MspPrjctNm = Prms[Prjct + ":MSP Project Name"];
            }
            catch
            {
                PrmsFnd.MspPrjctNm = false;
            }

            try
            {
                Str1 = Prms[Prjct + ":Post Checkitem Name"];
                Cnfg.PstChckItmNm = Convert.ToBoolean(Str1);
            }
            catch
            {
                PrmsFnd.PstChckItmNm = false;
            }

            try
            {
                Str1 = Prms[Prjct + ":Post All Checklist Items"];
                Cnfg.PstAllChckLstItms = Convert.ToBoolean(Str1);
            }
            catch
            {
                PrmsFnd.PstAllChckLstItms = false;
            }

            try
            {
                Cnfg.TrlloAppKy = Prms["Trello AppKey"];
            }
            catch
            {
                PrmsFnd.TrlloAppKy = false;
            }

            try
            {
                Lst.Clear();
                Str1 = Prms[Prjct + ":Trello Lists Included"];
                Tkns = Regex.Split(Str1, ";");
                foreach (string Tkn in Tkns)
                {
                    if (Tkn != "")
                    {
                        Cnfg.TrlloLstsInclddInpt.Add(Tkn);
                    }
                }
            }
            catch
            {
                PrmsFnd.TrlloLstsInclddInpt = false;
            }

            try
            {
                Lst.Clear();
                Str1 = Prms[Prjct + ":Trello Lists Excluded"];
                Tkns = Regex.Split(Str1, ";");
                foreach (string Tkn in Tkns)
                {
                    if (Tkn != "")
                    {
                        Cnfg.TrlloLstsExclddInpt.Add(Tkn);
                    }
                }
            }
            catch
            {
                PrmsFnd.TrlloLstsExclddInpt = false;
            }

            try
            {
                Lst.Clear();
                Str1 = Prms[Prjct + ":Trello Lists Not Open"];
                Tkns = Regex.Split(Str1, ";");
                foreach (string Tkn in Tkns)
                {
                    if (Tkn != "")
                    {
                        Cnfg.TrlloLstsNtOpn.Add(Tkn);
                    }
                }
            }
            catch
            {
                PrmsFnd.TrlloLstsNtOpn = false;
            }

            try
            {
                Lst.Clear();
                Str1 = Prms[Prjct + ":Trello Lists Rejected"];
                Tkns = Regex.Split(Str1, ";");
                foreach (string Tkn in Tkns)
                {
                    if (Tkn != "")
                    {
                        Cnfg.TrlloLstsRjctd.Add(Tkn);
                    }
                }
            }
            catch
            {
                PrmsFnd.TrlloLstsRjctd = false;
            }

            try
            {
                Cnfg.TrlloUsrTkn = Prms["Trello UserToken"];
            }
            catch
            {
                PrmsFnd.TrlloUsrTkn = false;
            }

            try
            {
                Str1 = Prms[Prjct + ":Update Date"];
                Cnfg.UpdtDt = Convert.ToDateTime(Str1);
            }
            catch
            {
                PrmsFnd.UpdtDt = false;
            }

            try
            {
                Str1 = Prms[Prjct + ":Update MSP Actuals"];
                Cnfg.UpdtMspActls = Convert.ToBoolean(Str1);
            }
            catch
            {
                PrmsFnd.UpdtMspActls = false;
            }

            try
            {
                Str1 = Prms[Prjct + ":Update MSP Measures"];
                Cnfg.UpdtMspMsrs = Convert.ToBoolean(Str1);
            }
            catch
            {
                PrmsFnd.UpdtMspMsrs = false;
            }

            try
            {
                Str1 = Prms[Prjct + ":Update MSP Projected"];
                Cnfg.UpdtMspPrjctd = Convert.ToBoolean(Str1);
            }
            catch
            {
                PrmsFnd.UpdtMspPrjctd = false;
            }

            try
            {
                Cnfg.XlsOutptDrctry = Prms["Xls Output Directory"];
            }
            catch
            {
                PrmsFnd.XlsOutptDrctry = false;
            }

            try
            {
                Cnfg.XlsFlNm = Prms[Prjct + ":Xls File Name"];
            }
            catch
            {
                PrmsFnd.XlsFlNm = false;
            }

            Console.WriteLine("Update MSP Actuals = " + Cnfg.UpdtMspActls);
            Console.WriteLine("Update MSP Projected = " + Cnfg.UpdtMspPrjctd);
            Console.WriteLine("Update MSP Measures = " + Cnfg.UpdtMspMsrs);

            Cnfg.PrmsOk = true;
            if (!(PrmsFnd.Brds && PrmsFnd.MspExe && PrmsFnd.MspPrjctNm && PrmsFnd.TrlloAppKy
                && PrmsFnd.TrlloLstsExclddInpt && PrmsFnd.TrlloLstsInclddInpt && PrmsFnd.TrlloLstsNtOpn
                && PrmsFnd.TrlloLstsRjctd && PrmsFnd.TrlloUsrTkn && PrmsFnd.XlsFlNm && PrmsFnd.XlsOutptDrctry))
            {
                Cnfg.PrmsOk = false;
                Console.WriteLine("\n\r");
                if (!PrmsFnd.Brds)
                {
                    Console.WriteLine("ERROR missing parm: Boards");
                }

                if (!PrmsFnd.MspExe)
                {
                    Console.WriteLine("ERROR missing parm: MS Project Exe");
                }

                if (!PrmsFnd.MspPrjctNm)
                {
                    Console.WriteLine("ERROR missing parm: MSP Project Name");
                }

                if (!PrmsFnd.TrlloAppKy)
                {
                    Console.WriteLine("ERROR missing parm: Trello AppKey");
                }

                if (!PrmsFnd.TrlloLstsExclddInpt)
                {
                    Console.WriteLine("ERROR missing parm: Trello Lists Excluded");
                }

                if (!PrmsFnd.TrlloLstsInclddInpt)
                {
                    Console.WriteLine("ERROR missing parm: Trello Lists Included");
                }

                if (!PrmsFnd.TrlloLstsNtOpn)
                {
                    Console.WriteLine("ERROR missing parm: Trello Lists Not Open");
                }

                if (!PrmsFnd.TrlloLstsRjctd)
                {
                    Console.WriteLine("ERROR missing parm: Trello Lists Rejected");
                }

                if (!PrmsFnd.TrlloUsrTkn)
                {
                    Console.WriteLine("ERROR missing parm: Trello UserToken");
                }

                if (!PrmsFnd.XlsFlNm)
                {
                    Console.WriteLine("ERROR missing parm: Xls File Name");
                }

                if (!PrmsFnd.XlsOutptDrctry)
                {
                    Console.WriteLine("ERROR missing parm: Xls Output Directory");
                }
            }

            return Cnfg;
        }
        private static Result Add_Multiply(int a, int b)
        {
            var result = new Result
            {
                add = a * b,
                multiply = a + b
            };
            return result;
        }
    }
}
