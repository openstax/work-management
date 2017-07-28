using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;

// test 2


namespace MspUpdate
{
    public class Configuration
    {
        public List<string> Brds = new List<string>();
        public List<string> TrlloLstsIncldd = new List<string>();
        public List<string> TrlloLstsExcldd = new List<string>();
        public bool UpdtMspActls;
        public bool UpdtMspPrjctd;
        public DateTime UpdtDt;
        public string MspPrjctNm;
        public bool PstAllChckLstItms;
        public DateTime IncldCrdsChngdAftr;
        public string TrlloAppKy;
        public string TrlloUsrTkn;
        public string XlsFlNm;
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
            DateTime DtToUpdt = DateTime.Today;
            // string strDtToUpdt;
            DateTime IncldCrdsChngdAftr = new DateTime(1900, 1, 1);
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
            List<string> TrlloLstsIncldd = new List<string>(); // Trllo lists included
            bool UpdtMsp = true;
            bool UsrNmFnd = false;
            string XlsFlPth = "";
            string XlsOutptDrctry = "";
            string XlsTmpltPth = "";

            #if DEBUG
                DbgExec = true;
            #endif

            Console.Write("Start at " + DateTime.Now.ToString("hh:mm:ss") + "\r\n");

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

            // MessageBox.Show("Entering Main: UpdtMsp=" + Convert.ToString(UpdtMsp) + " PstAllChckLstItms=" + Convert.ToString(PstAllChckLstItms) + " PstChckItmNm=" + Convert.ToString(PstChckItmNm));

            if (CnslInpt)
            {
                // Parms set by username
                string aEnvNm = Environment.UserName;
                switch (Environment.UserName)
                {
                    case "BruceP Rice":
                        UsrNmFnd = true;
                        //CnfgFlPth = "C:\\Users\\Bruce Pike Rice\\Documents\\UTS\\UTS MSP Update\\UTS MSP Config.txt";
                        //XlsTmpltPth = "C:\\Users\\Bruce Pike Rice\\Documents\\Repos\\work-management\\JiraInteraction\\UTS MSP Update Template.xlsm";
                        XlsOutptDrctry = "c:\\temp\\";
                        MspExe = "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\WINPROJ.EXE";
                        break;

                    //case "Norm":
                    //    UsrNmFnd = true;
                    //    XlsTmpltPth = "C:\\Users\\norm\\Source\\Workspaces\\Mayor\\ToolsAutomation\\JiraInteraction\\JiraInteraction\\UTS MSP Update Template.xlsm";
                    //    XlsOutptDrctry = "c:\\temp\\";
                    //    MspExe = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINPROJ.EXE";
                    //    break;

                    default:
                        UsrNmFnd = false;
                        Console.WriteLine("Username not found");
                        break;
                }

                // File paths for debug
                CnfgFlPth = Application.StartupPath;
                XlsTmpltPth = Application.StartupPath;
                if (DbgExec)
                {
                    CnfgFlPth = "C:\\Users\\Bruce Pike Rice\\Documents\\UTS\\UTS MSP Update\\UTS MSP Config.txt";
                    XlsTmpltPth = "C:\\Users\\Bruce Pike Rice\\Documents\\Repos\\work-management\\JiraInteraction\\UTS MSP Update Template.xlsm";
                }
                //CnfgFlPth = "C:\\Users\\Bruce Pike Rice\\Documents\\Bin\\MSP Update\\Deploy";
                //XlsTmpltPth = "C:\\Users\\Bruce Pike Rice\\Documents\\Bin\\MSP Update\\Deploy";

                // Parms entered by user
                if (UsrNmFnd)
                {
                    Console.WriteLine("MSP Update"); Console.WriteLine("");
                    Console.WriteLine("Select Project:"); Console.WriteLine("");
                    //Console.WriteLine("");
                    //Console.WriteLine(" PIVOTAL SYNC --------------------------");
                    //Console.WriteLine("");
                    //Console.WriteLine(" [1 Version of Truth]: type 1VT");
                    //Console.WriteLine("          [All TUTOR]: type All TUTOR");
                    //Console.WriteLine("           [v2 TUTOR]: type V2 TUTOR");
                    //Console.WriteLine("   [count in past 24]: type 24");
                    //Console.WriteLine("   [count 2 days ago]: type dayb4y");
                    //Console.WriteLine("");
                    //Console.WriteLine(" TRELLO SYNC --------------------------");
                    //Console.WriteLine("");
                    Console.WriteLine("             [Tutor]: type Tutor");
                    Console.WriteLine("          [Book Tools]: type Book Tools");
                    Console.WriteLine("            [OS Web]: type OS Web");
                    Console.WriteLine("          [UTS Test]: type UTS Test");
                    //Console.WriteLine("               [1VT]: type TestPO");
                    //Console.WriteLine("      [Norm's Retun]: type NR");

                    //Console.WriteLine("");
                    //Console.WriteLine(" Release Card Data --------------------------");
                    //Console.WriteLine("");
                    //Console.WriteLine("               [RC]: type ReleaseWeb");
                    //Console.WriteLine("");

                    //Console.WriteLine(" CSV SYNC -----------------------------");
                    //Console.WriteLine("");
                    //Console.WriteLine("");
                    //Console.WriteLine("");
                    //Console.WriteLine(" Project Update --------------------------");
                    //Console.WriteLine("");
                    //Console.WriteLine("               [Project update by MSP client]: type Client");
                    Console.WriteLine("");
                    Console.WriteLine("To quit, type EXIT"); Console.WriteLine("");
                    Console.WriteLine("...........................................................");

                }

            }
            bool loop = true;
            while (loop && UsrNmFnd)
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
                        XlsFlPth = XlsOutptDrctry + Cnfg.XlsFlNm;

                        // Read boards
                        trelloConnect.CteReadBoard(Prjct, XlsTmpltPth, XlsFlPth, CnfgFlPth, MspExe, Cnfg);

                        loop = false;
                        break;

                    case "TUTOR":
                        Console.WriteLine("> Tutor selected. Stand by...");

                        // Get configuration
                        Cnfg = Program.Read_Config(Prjct, CnfgFlPth);
                        XlsFlPth = XlsOutptDrctry + Cnfg.XlsFlNm;

                        // Read boards
                        trelloConnect.CteReadBoard(Prjct, XlsTmpltPth, XlsFlPth, CnfgFlPth, MspExe, Cnfg);

                        loop = false;
                        break;

                    case "BOOK TOOLS":
                        Console.WriteLine("> Book Tools selected. Stand by....");

                        // Get configuration
                        Cnfg = Program.Read_Config(Prjct, CnfgFlPth);
                        XlsFlPth = XlsOutptDrctry + Cnfg.XlsFlNm;

                        // Read boards
                        trelloConnect.CteReadBoard(Prjct, XlsTmpltPth, XlsFlPth, CnfgFlPth, MspExe, Cnfg);

                        loop = false;
                        break;

                    case "OS WEB":
                        Console.WriteLine("> OS Web selected. Stand by...");

                        // Get configuration
                        Cnfg = Program.Read_Config(Prjct, CnfgFlPth);
                        XlsFlPth = XlsOutptDrctry + Cnfg.XlsFlNm;

                        // Read boards
                        trelloConnect.CteReadBoard(Prjct, XlsTmpltPth, XlsFlPth, CnfgFlPth, MspExe, Cnfg);

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
                Console.WriteLine("\r\nAll done...hit Enter to exit");
                Console.ReadLine();
            }
        }
        private static Configuration Read_Config(string Prjct, string CnfgFlPth)
        {
            List<string> Lst = new List<string>();
            var Cnfg = new Configuration();
            int cntr = 0;
            string ln;
            var Prms = new Dictionary<string, string>();
            string[] Tkns = new string[] { "" };
            string Str1; 

            // Read config file
            // string UsrNm = Environment.UserName;
            System.IO.StreamReader file = new System.IO.StreamReader(CnfgFlPth);
            while ((ln = file.ReadLine()) != null)
            {
                    if (ln.IndexOf("=") != -1)
                    {
                        Tkns = Regex.Split(ln, "=");
                        Prms.Add(Tkns[0], Tkns[1]);
                        cntr++;
                    }
            }
            file.Close();

            //Get parms
            Cnfg.TrlloAppKy = Prms["Trello AppKey"];
            Cnfg.TrlloUsrTkn = Prms["Trello UserToken"];

            Cnfg.MspPrjctNm = Prms[Prjct + ":MSP Project Name"];

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

            Lst.Clear();
            Str1 = Prms[Prjct + ":Trello Lists Included"];
            Tkns = Regex.Split(Str1, ";");
            foreach (string Tkn in Tkns)
            {
                if (Tkn != "")
                {
                    Cnfg.TrlloLstsIncldd.Add(Tkn);
                }
            }

            Lst.Clear();
            Str1 = Prms[Prjct + ":Trello Lists Excluded"];
            Tkns = Regex.Split(Str1, ";");
            foreach (string Tkn in Tkns)
            {
                if (Tkn != "")
                {
                    Cnfg.TrlloLstsExcldd.Add(Tkn);
                }
            }

            Str1 = Prms[Prjct + ":Update MSP Actuals"];
            Cnfg.UpdtMspActls = Convert.ToBoolean(Str1);

            Str1 = Prms[Prjct + ":Update MSP Projected"];
            Cnfg.UpdtMspPrjctd = Convert.ToBoolean(Str1);

            Str1 = Prms[Prjct + ":Post All Checklist Items"];
            Cnfg.PstAllChckLstItms = Convert.ToBoolean(Str1);

            Str1 = Prms[Prjct + ":Include Cards Changed After"];
            Cnfg.IncldCrdsChngdAftr = Convert.ToDateTime(Str1);

            Cnfg.XlsFlNm = Prms[Prjct + ":Xls File Name"];

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
