using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Net;
using System.IO;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Web;
using System.Web.Script.Serialization;
using System.Drawing;

namespace ProjectWRQuery {
  /// <summary>
  /// Interaction logic for MainWindow.xaml
  /// </summary>
  public partial class MainWindow : System.Windows.Window {
    public MainWindow() {
      string[] args = Environment.GetCommandLineArgs();
      InitializeComponent();
      string[] validUsers = { "usernames" };//update to run
      string strUserName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
      int indDomain = strUserName.LastIndexOf("\\")+1;
      string user = strUserName.Substring(indDomain, strUserName.Length - indDomain);
      int i = args.Count();
      if (true) { //update for validuser check
        BackgroundWorker bw1 = new BackgroundWorker();
        if (args.Count() > 1) {
          string fileName = args[1];
          bw1.DoWork += new DoWorkEventHandler(updateFromFile);
          bw1.RunWorkerAsync(fileName);
          initBtn.IsEnabled = false;
          updateBtn.IsEnabled = false;
          pBar.Visibility = Visibility.Hidden;
        } else { 
        bw1.DoWork += new DoWorkEventHandler(getPList);
        bw1.RunWorkerAsync();
        initBtn.IsEnabled = false;
        updateBtn.IsEnabled = false;
        pBar.Visibility = Visibility.Hidden;
        }
      } else {
        initBtn.IsEnabled = false;
        updateBtn.IsEnabled = false;
        status.Content = "Not an allowed user, contact SCT.";
        pBar.Visibility = Visibility.Hidden;
      }
      
      
      //status.Content = strUserName;
    }

    public void updateFromFile(object sender, DoWorkEventArgs e) {
      string filename = (string) e.Argument;
      updater(filename);
    }

    public List<DevT> getWRplus(string PID) {
      SqlDataReader rdr = null;
      SqlConnection con = null;
      SqlCommand cmd = null;

      List<DevT> dis = new List<DevT>();
      try {
        con = new SqlConnection(String.Format("redacted"));//update ot run
        con.Open();
        #region QueryString
        string qryStr = "Select WR, WRStatus, VAR, VARStatus, VARNeedByDate, ForwardTypes.ForwardTypeName as Priority from (" +
          "SELECT Bug.ProblemID as WR, ProgressStatusTypes.ProgressStatusName as WRStatus, Bug2.ProblemID as VAR, ProgressStatusTypes2.ProgressStatusName as VARStatus, Bug2.TaskPlannedStartDate as VARNeedByDate, Bug2.CrntForwardTypeID as VARPriority " +
          "FROM SWISEDB.dbo.Bug, SWISEDB.dbo.Bug Bug2, SWISEDB.dbo.BugLinks, SWISEDB.dbo.Project, SWISEDB.dbo.ProgressStatusTypes, SWISEDB.dbo.ProgressStatusTypes ProgressStatusTypes2 " +
          "WHERE " +
          "Bug.ProjectID = 17 and Bug2.ProjectID = Project.ProjectID and " +
          "BugLinks.LinkedBugID = Bug.BugID and " +
          "BugLinks.LinkedProjectID = Bug.ProjectID and " +
          "BugLinks.ProjectID = Project.ProjectID and " +
          "Bug2.BugID = BugLinks.BugID and " +
          "Bug2.ProjectID = BugLinks.ProjectID and " +
          "Bug.ProgressStatusID = ProgressStatusTypes.ProgressStatusID and " +
          "Bug.ProjectID = ProgressStatusTypes.ProjectID and " +
          "Bug2.ProgressStatusID = ProgressStatusTypes2.ProgressStatusID and " +
          "Bug2.ProjectID = ProgressStatusTypes2.ProjectID and " +
          "Project.ProjectName like '" + PID + "%' " +
          "UNION " +
          "SELECT Bug.ProblemID as WR, ProgressStatusTypes.ProgressStatusName as WRStatus, Bug2.ProblemID as VAR, ProgressStatusTypes2.ProgressStatusName as VARStatus, Bug2.TaskPlannedStartDate as VARNeedByDate, Bug2.CrntForwardTypeID as VARPriority " +
          "FROM SWISEDB.dbo.Bug, SWISEDB.dbo.Bug Bug2, SWISEDB.dbo.BugLinks, SWISEDB.dbo.Project, SWISEDB.dbo.ProgressStatusTypes, SWISEDB.dbo.ProgressStatusTypes ProgressStatusTypes2 " +
          "WHERE " +
          "Bug.ProjectID = 17 and Bug2.ProjectID = Project.ProjectID and " +
          "BugLinks.BugID = Bug.BugID and " +
          "BugLinks.ProjectID = Bug.ProjectID and " +
          "BugLinks.LinkedProjectID = Project.ProjectID and " +
          "Bug2.BugID = BugLinks.LinkedBugID and " +
          "Bug2.ProjectID = BugLinks.LinkedProjectID and " +
          "Bug.ProgressStatusID = ProgressStatusTypes.ProgressStatusID and " +
          "Bug.ProjectID = ProgressStatusTypes.ProjectID and " +
          "Bug2.ProgressStatusID = ProgressStatusTypes2.ProgressStatusID and " +
          "Bug2.ProjectID = ProgressStatusTypes2.ProjectID and " +
          "Project.ProjectName like '" + PID + "%' ) q" +
          "Left Join ForwardTypes on ForwardTypes.ProjectID = 1450 and ForwardTypes.OrderNo+1 = VARPriority";
        #endregion

        cmd = new SqlCommand(qryStr);
        cmd.Connection = con;
        rdr = cmd.ExecuteReader();
        while (rdr.HasRows) {
          while (rdr.Read()) {
            string dWR = "";
            string dWRstatus = "";
            string dIVnum = "";
            string dIVstatus = "";
            string dNBD = "Date Not Set";//need by date
            string dNBE = "Event Not Set";//need by event
            try {
              dWR = rdr[0].ToString().Substring(3, rdr[0].ToString().Length - 3);
            } catch { }
            try { dWRstatus = rdr[1].ToString(); } catch { }
            try { dIVnum = rdr[2].ToString(); } catch { }
            try { dIVstatus = rdr[3].ToString(); } catch { }
            try { 
              dNBD = rdr[4].ToString();
              if (dNBD == "")
                dNBD = "Date Not Set";
            } catch { dNBD = "Date Not Set"; }
            try { dNBE = rdr[5].ToString(); } catch { dNBE = "Event Not Set"; }
            DevT d = new DevT(dWR, dWRstatus, dIVnum, dIVstatus,dNBD,dNBE);
            dis.Add(d);


          }
          rdr.NextResult();
        }
      } catch (Exception e) {
        string m = e.Message;
        rdr = null;
      } finally {
        if (rdr != null) {
          rdr.Close();
        }
        if (con.State == ConnectionState.Open) {
          con.Close();
        }
      }
      return dis;
    }

    public List<DevT> getWRs(string PID) {
      SqlDataReader rdr = null;
      SqlConnection con = null;
      SqlCommand cmd = null;
      //List<string> WRs = new List<string>();

      List<DevT> dis = new List<DevT>();
      try {
        con = new SqlConnection(String.Format("redacted"));//update to run
        con.Open();
        string qryStr = "SELECT Bug.ProblemID as WR, ProgressStatusTypes.ProgressStatusName as WRStatus, Bug2.ProblemID as VAR, ProgressStatusTypes2.ProgressStatusName as VARStatus  " +
          "FROM SWISEDB.dbo.Bug, SWISEDB.dbo.Bug Bug2, SWISEDB.dbo.BugLinks, SWISEDB.dbo.Project, SWISEDB.dbo.ProgressStatusTypes, SWISEDB.dbo.ProgressStatusTypes ProgressStatusTypes2  " +
          "WHERE  " +
          "Bug.ProjectID = 17 and Bug2.ProjectID = Project.ProjectID and  " +
          "BugLinks.LinkedBugID = Bug.BugID and  " +
          "BugLinks.LinkedProjectID = Bug.ProjectID and  " +
          "BugLinks.ProjectID = Project.ProjectID and  " +
          "Bug2.BugID = BugLinks.BugID and  " +
          "Bug2.ProjectID = BugLinks.ProjectID and  " +
          "Bug.ProgressStatusID = ProgressStatusTypes.ProgressStatusID and " +
          "Bug.ProjectID = ProgressStatusTypes.ProjectID and " +
          "Bug2.ProgressStatusID = ProgressStatusTypes2.ProgressStatusID and "+
          "Bug2.ProjectID = ProgressStatusTypes2.ProjectID and "+
          "Project.ProjectName like '" + PID + "%' " +
          "UNION " +
          "SELECT Bug.ProblemID as WR, ProgressStatusTypes.ProgressStatusName as WRStatus, Bug2.ProblemID as VAR, ProgressStatusTypes2.ProgressStatusName as VARStatus  " +
          "FROM SWISEDB.dbo.Bug, SWISEDB.dbo.Bug Bug2, SWISEDB.dbo.BugLinks, SWISEDB.dbo.Project, SWISEDB.dbo.ProgressStatusTypes, SWISEDB.dbo.ProgressStatusTypes ProgressStatusTypes2 " +
          "WHERE  " +
          "Bug.ProjectID = 17 and Bug2.ProjectID = Project.ProjectID and  " +
          "BugLinks.BugID = Bug.BugID and  " +
          "BugLinks.ProjectID = Bug.ProjectID and " +
          "BugLinks.LinkedProjectID = Project.ProjectID and  " +
          "Bug2.BugID = BugLinks.LinkedBugID and  " +
          "Bug2.ProjectID = BugLinks.LinkedProjectID and  " +
          "Bug.ProgressStatusID = ProgressStatusTypes.ProgressStatusID and " +
          "Bug.ProjectID = ProgressStatusTypes.ProjectID and " +
          "Bug2.ProgressStatusID = ProgressStatusTypes2.ProgressStatusID and " +
          "Bug2.ProjectID = ProgressStatusTypes2.ProjectID and " +
          "Project.ProjectName like '" + PID + "%' ";

        #region OldString
        //string qryStr = "SELECT Bug.ProblemID as WR, Bug2.ProblemID as VAR " +
        //  "FROM SWISEDB.dbo.Bug, SWISEDB.dbo.Bug Bug2, SWISEDB.dbo.BugLinks, SWISEDB.dbo.Project  " +
        //  "WHERE  " +
        //  "Bug.ProjectID = 17 and Bug2.ProjectID = Project.ProjectID and  " +
        //  "BugLinks.LinkedBugID = Bug.BugID and  " +
        //  "BugLinks.LinkedProjectID = Bug.ProjectID and  " +
        //  "BugLinks.ProjectID = Project.ProjectID and  " +
        //  "Bug2.BugID = BugLinks.BugID and  " +
        //  "Bug2.ProjectID = BugLinks.ProjectID and  " +
        //  "Project.ProjectName like '" ID + "%' " +
        //  "UNION " +
        //  "SELECT Bug.ProblemID as WR, Bug2.ProblemID as VAR  " +
        //  "FROM SWISEDB.dbo.Bug, SWISEDB.dbo.Bug Bug2, SWISEDB.dbo.BugLinks, SWISEDB.dbo.Project  " +
        //  "WHERE  " +
        //  "Bug.ProjectID = 17 and Bug2.ProjectID = Project.ProjectID and  " +
        //  "BugLinks.BugID = Bug.BugID and  " +
        //  "BugLinks.ProjectID = Bug.ProjectID and " +
        //  "BugLinks.LinkedProjectID = Project.ProjectID and  " +
        //  "Bug2.BugID = BugLinks.LinkedBugID and  " +
        //  "Bug2.ProjectID = BugLinks.LinkedProjectID and  " +
        //  "Project.ProjectName like '" ID + "%' ";
        #endregion

        cmd = new SqlCommand(qryStr);
        cmd.Connection = con;
        rdr = cmd.ExecuteReader();
        while (rdr.HasRows) {
          while (rdr.Read()) {
            string IVt = "";
            string WRt = "";
            string RDt = "";
            string IVs = "";
            try {
              WRt = rdr[0].ToString().Substring(3, rdr[0].ToString().Length - 3);
            } catch { }
            try { IVt = rdr[1].ToString(); } catch { }
            try { RDt = rdr[2].ToString(); } catch { }
            try { IVs = rdr[3].ToString(); } catch { }
            DevT d = new DevT(WRt, IVt, RDt,IVs,"","");
            dis.Add(d);


          }
          rdr.NextResult();
        }
      } catch (Exception e) {
        string m = e.Message;
        rdr = null;
      } finally {
        if (rdr != null) {
          rdr.Close();
        }
        if (con.State == ConnectionState.Open) {
          con.Close();
        }
      }
      return dis;
    }

    public WR pJSON(string qTxt) {//list<string> qTxts

      JiraResource resource = new JiraResource();
      string bURL = "https://redacted.atlassian.net/rest/api/latest/search?jql="; //base url, update to run
      string JSONdata = null;
      int statusCode = 0;
      Stream s;
      StreamReader r;
      HttpWebResponse webRes;
      HttpWebRequest WebReq = WebRequest.Create(bURL + qTxt) as HttpWebRequest;
      WebReq.ContentType = "application/json";
      WebReq.Method = "GET";
      WebReq.Headers["Authorization"] = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(resource.m_Username + ":" + resource.m_Password));

      try {
        webRes = (HttpWebResponse)WebReq.GetResponse();
        s = webRes.GetResponseStream();
        r = new StreamReader(s);
        JSONdata = r.ReadToEnd();
        statusCode = (int)webRes.StatusCode;
        s.Close();
        r.Close();
      } catch (WebException e) {
        s = e.Response.GetResponseStream();
        r = new StreamReader(s);
        JSONdata = r.ReadToEnd();
        statusCode = (int)((HttpWebResponse)e.Response).StatusCode;
        s.Close();
        r.Close();
      }
      try {
        //JavaScriptSerializer jss = new JavaScriptSerializer();
        //jss.RegisterConverters = unnull;
        WR tem = new JavaScriptSerializer().Deserialize<WR>(JSONdata);
        return tem;
      } catch (Exception e) {
        string m = e.Message;
      }
      return null;
    }

    public List<string> getPID() {

      SqlDataReader rdr = null;
      SqlConnection con = null;
      SqlCommand cmd = null;
      bool fail = false;
      List<string> pids = new List<string>();
      try {
        con = new SqlConnection(String.Format("redacted"));//update to run
        con.Open();
        string qryStr = "select ProjectName from Project where ProjectName not like '*%' order by ProjectName";
        cmd = new SqlCommand(qryStr);
        cmd.Connection = con;
        rdr = cmd.ExecuteReader();
        while (rdr.Read()) {
          pids.Add(rdr[0].ToString());
        }
      } catch {
        rdr = null;
        fail = true;
      } finally {
        if (rdr != null) {
          rdr.Close();
        }
        if (con.State == ConnectionState.Open) {
          con.Close();
        }
      }
      Dispatcher.Invoke((System.Action)delegate() {
        if (fail) {
          status.Content = "Couldn't connect to DevTrack.";
        } else {
          status.Content = "";
          updateBtn.IsEnabled = true;
        }

        //status.Content = "";
      });

      return pids;
    }

    public void getPList(object sender, DoWorkEventArgs e) {
      List<string> rdr = getPID();

      Dispatcher.Invoke((System.Action)delegate() {
        foreach (string pid in rdr) {
          pBox.Items.Add(pid);
        }
        pBox.SelectedItem = "System Certification Team";
        //status.Content = "";
      });
    }

    //public string filterOut(List<string> list, string keep) {
    //  string res = "";
    //  if (list == null) {
    //    list = new List<string>();
    //    list.Add("");
    //  }
    //  try {
    //    string filter = keep.Substring(0, keep.IndexOf(":") - 1);
    //    foreach (string s in list) {
    //      if (s.Contains(filter))
    //        res += s + ", ";
    //    }
    //  } catch {
    //    res += "Error decoding, ";
    //  }
    //  return res;
    //}

    public static string unnull(object Value) {
      return Value == null ? "xxx" : Value.ToString();
    }

    public static string undate(string ugly) {
      //"2015-01-16T16:47:29.000-0600"

      //temp=11/26/2014 12:00:00 AM
      //releasedate=2014-11-26
      string[] a = new string[3];
      try {
        a = ugly.Substring(0, ugly.IndexOf("T")).Split(new Char[] { '-' }).ToArray();
        string[] b = { a[1].TrimStart('0'), a[2].TrimStart('0'), a[0] };
        return string.Join("/", b);
      } catch {
        return a[0];
      }

    }

    public static string getDStatus(string WR, List<DevT> DevTrack) {
      foreach (DevT d in DevTrack) {
        if (d.WR == WR) {
          return d.status;
        }
      }
      return "N/A";
    }
    public static string getIVStatus(string WR, List<DevT> DevTrack) {
      foreach (DevT d in DevTrack) {
        if (d.WR == WR) {
          return d.IVstatus;
        }
      }
      return "N/A";
    }
    public static string getNBDE(string WR, List<DevT> DevTrack) {
      foreach (DevT d in DevTrack) {
        if (d.WR == WR) {
          string temp = d.NeedByEvent+Environment.NewLine+Environment.NewLine+d.NeedByDate;
          return temp;
        }
      }
      return "N/A";
    }

    public static int getJIRAtotal(string jqry) {
      JiraResource resource = new JiraResource();
      string bURL = "https://redacted.atlassian.net/rest/api/latest/search?jql="; //base url, update to run
      string JSONdata = null;
      int statusCode = 0;
      Stream s;
      StreamReader r;
      HttpWebResponse webRes;
      HttpWebRequest WebReq = WebRequest.Create(bURL + jqry+"&startAt=0&maxResults=0") as HttpWebRequest;
      WebReq.ContentType = "application/json";
      WebReq.Method = "GET";
      WebReq.Headers["Authorization"] = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(resource.m_Username + ":" + resource.m_Password));

      try {
        webRes = (HttpWebResponse)WebReq.GetResponse();
        s = webRes.GetResponseStream();
        r = new StreamReader(s);
        JSONdata = r.ReadToEnd();
        statusCode = (int)webRes.StatusCode;
        s.Close();
        r.Close();
      } catch (WebException e) {
        s = e.Response.GetResponseStream();
        r = new StreamReader(s);
        JSONdata = r.ReadToEnd();
        statusCode = (int)((HttpWebResponse)e.Response).StatusCode;
        s.Close();
        r.Close();
      }
      try {
        //"{\"startAt\":0,\"maxResults\":0,\"total\":549,\"issues\":[]}"
        //{"startAt":0,"maxResults":0,"total":549,"issues":[]}
        //pos 36
        int x = JSONdata.LastIndexOf(',');
        int i = Convert.ToInt32(JSONdata.Substring(36, JSONdata.LastIndexOf(',') - 36));
        return i;
      } catch (Exception e) {
        string m = e.Message;
      }
      return 0;
    }
    public static string getIV(string WR, List<DevT> DevTrack) {
      foreach (DevT d in DevTrack) {
        if (d.WR == WR) {
          return d.IV;
        }
      }
      return "N/A";
    }

    static bool ContainsLoop(List<string> list, string value) {
      for (int i = 0; i < list.Count; i++) {
        if (list[i] == value) {
          return true;
        }
      }
      return false;
    }

    public static List<string> getList(Excel.Worksheet xlSheet, int colNum) {
      Excel.Range usedRange = xlSheet.UsedRange.Columns[colNum];
      int rr = usedRange.Rows.Count;
      List<string> slist = new List<string>();
      for (int i = 1; i <= rr; i++) {
        try { slist.Add(xlSheet.Cells[i, colNum].Value.ToString()); }catch{}
      }
      releaseObject(usedRange);
      //releaseObject(xlSheet);
      return slist;
    }

    public int getIndex(List<string> Keys, string value) {//turbo fast
      int index  = -1;
      Parallel.For(0, Keys.Count, i => {
        if (Keys[i] == value) {
          index= i;
        }
      });
      return index;
    }

    //public string getSuite(string JKEY) {
    //  switch (JKEY.Substring(0,JKEY.IndexOf("-"))) {
    //    case "CSM":
    //      return "Platform Server";
    //    case "VM":
    //      return "Platform Visualizaiton";
    //    case "CCM":
    //      return "System Management";
    //    case "SCADAMAIN":
    //      return "SCADA";
    //    case "COMMSMAINT":
    //      return "COMMS";
    //    case "CHRNSMAINT":
    //      return "ISR";
    //    case "ISRMAINT":
    //      return "ISR";
    //    case "GMSMAINT":
    //      return "GMS";
    //    case "EMSMAINT":
    //      return "EMS";
    //    case "DMSMAINT":
    //      return "DMS";
    //    case "GASMAINT":
    //      return "Fluid/Gas";
    //    default:
    //      return "N/A";
    //  }
    //}

    public List<string> DevTWRs(List<DevT> LinkedWRs) {
      List<string> WRVals = new List<string>();
      Parallel.ForEach(LinkedWRs, Issue => {
        WRVals.Add(Issue.WR);
      });
      return WRVals;
    }
    public List<string> ReturnedWRs(List<WR> data) {
      List<string> WRs = new List<string>();
      Parallel.ForEach(data, deets => {
        Parallel.ForEach(deets.issues, issue => {
          try {
            WRs.Add(issue.fields.customfield_10401);
          } catch { }
        });
      });
      return WRs;
    }
    
    public List<string> CompareDJ(List<DevT> DevTrack, List<WR> data){
      List<string> dne = new List<string>();
      List<string> WRs = ReturnedWRs(data);
      foreach (DevT dt in DevTrack) {
        if (getIndex(WRs, dt.WR) < 0) {
          dne.Add(dt.WR);
        }
      }
      //Parallel.ForEach(DevTrack, dt => {
      //  int i = getIndex(WRs, dt.WR);
      //  if (i < 0) {
      //    dne.Add(dt.WR);
      //  }
      //});
      return dne;
    }

    public bool createWRPR(string file, string pid, List<WR> data, List<DevT> DevTrack) {
      bool result = false;
      List<string> cautions = new List<string>();
      Excel.Application excelApp = null;
      Excel.Workbook workbook = null;
      Excel.Sheets sheets = null;
      Excel.Worksheet dataSheet = null;
      Excel.Range xlR = null;
      //int rowS = 3;
      //int colS = 1;
      try {
        excelApp = new Excel.Application();
        workbook = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
        sheets = workbook.Sheets;
        dataSheet = sheets[1];
        dataSheet.Name = DateTime.Now.Month.ToString("00") + "-" + DateTime.Now.Day.ToString("00") + "-" + DateTime.Now.Year.ToString().Substring(2, 2);;
        //int p = 0;
        #region Header format
        string[] headers = { "WR#","WR Status", "Variance", "Project", "DevTeam", "Product", "Version", "Severity", "Project Delivery Priority", "Need by Event/Date", "Date Created", "Linked Variances", "Summary", "Description","Certification Assessment","Certification Required" };
        object[,] xlHeader = new object[1,headers.Length];
        for(int i =0;i<headers.Length;i++){
          xlHeader[0,i] = headers[i];
        }
        xlR = dataSheet.Range["A1","P1"];//xlGetStr
        xlR.Value2 = xlHeader;
        #endregion

        dataSheet.get_Range("A:P", Type.Missing).EntireColumn.VerticalAlignment = XlVAlign.xlVAlignCenter;
        dataSheet.get_Range("A:P", Type.Missing).EntireColumn.WrapText = true;
        dataSheet.get_Range("A:P", Type.Missing).EntireColumn.ColumnWidth = 12;
        dataSheet.get_Range("J:L", Type.Missing).EntireColumn.ColumnWidth = 18;
        dataSheet.get_Range("M:O", Type.Missing).EntireColumn.ColumnWidth = 32;
        dataSheet.get_Range("N:N", Type.Missing).EntireColumn.ColumnWidth = 50;
        dataSheet.get_Range("F:F", Type.Missing).EntireColumn.ColumnWidth = 22;
        int rowCount = 2; //start at row 2
        foreach (WR dd in data) {
          foreach (Jissues info in dd.issues) {
            WRQuery WRQ = new WRQuery(info, pid);
            string[] newLine = { WRQ.WR, getDStatus(WRQ.WR, DevTrack), getIV(WRQ.WR, DevTrack), pid, WRQ.Suite, WRQ.Product, WRQ.CurrentRelease, WRQ.IssueType, WRQ.ProjDeliveryPriority, getNBDE(WRQ.WR, DevTrack), WRQ.Created, WRQ.Variances, "'" + WRQ.Summary, "'" + WRQ.Description };
            object[,] xlNewLine = new object[1, newLine.Length];
            for (int i = 0; i < newLine.Length; i++) {
              xlNewLine[0, i] = newLine[i];
            }
            xlR = dataSheet.Range["A" + rowCount, "N" + rowCount];
            xlR.Value2 = xlNewLine;
            rowCount++;
          }
        }

        xlR = dataSheet.get_Range("A1:P" + (rowCount-1));
        dataSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, xlR, Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "MyTableStyle";
        dataSheet.ListObjects.get_Item("MyTableStyle").TableStyle = "TableStyleLight8";
        dataSheet.get_Range("A1", "P1").EntireRow.RowHeight = 15;

        bool readOnly = false;
        excelApp.DisplayAlerts = false;
        workbook.SaveAs(file, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, readOnly, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing);
        workbook.Close(true, Type.Missing, Type.Missing);
        excelApp.Quit();
        //releaseObject(rangg);
        //releaseObject(merg);
        //releaseObject(mergB);
        result = true;
      } catch (Exception e) {
        string m = e.Message;
        Dispatcher.Invoke((System.Action)delegate() {
          status.Content = "ugh Excel failed.. try again?";
        });
        result = false;
      } finally {
        releaseObject(xlR);
        releaseObject(dataSheet);
        releaseObject(sheets);
        releaseObject(workbook);
        releaseObject(excelApp);
      }
      return result;
    }

    public bool createBlankXL(string file, string pid, List<WR> data, List<DevT> DevTrack) {
      bool result = false;
      List<string> cautions = new List<string>();
      Excel.Application excelApp = null;
      Excel.Workbook workbook = null;
      Excel.Sheets sheets = null;
      Excel.Worksheet dataSheet = null;
      int rowS = 3;
      int colS = 1;
      try {
        excelApp = new Excel.Application();
        workbook = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
        sheets = workbook.Sheets;
        dataSheet = sheets[1];
        dataSheet.Name = "Data";
        int p = 0;
        #region Header format
        dataSheet.Cells[rowS - 2, colS] = pid;
        dataSheet.Cells[rowS - 2, colS + 1] = "Initialized: " + DateTime.Now.Month.ToString("00") + "/" + DateTime.Now.Day.ToString("00") + "/" + DateTime.Now.Year.ToString().Substring(2, 2);
        dataSheet.Cells[rowS - 1, colS + p] = "JIRA ID";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Variances";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "WR";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Issue Type";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Team";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Product";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Release Submitted Against";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Summary";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Proj. Del. Priority";
        //dataSheet.Cells[rowS - 1, colS + p] = "Variance Status";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "IV Status";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Original Statuses";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Status Change?";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Current Statuses";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Original Commitment";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Change?";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Current Commitment";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Original Release Date";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Date Change?";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Current Release Date";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Original Target Release";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Rel. Change?";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Current Target Release";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Created Date";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Requested Date";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Resolution Date";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Required for Certification";
        p++;
        dataSheet.Cells[rowS - 1, colS + p] = "Justification";
        p++;
        #endregion
        //string test = null;
        int i = rowS;
        double j = 100.0;
        double k = j / data.Count();
        List<string> JKEYS = new List<string>();
        JKEYS.Add(" ");
        bool IVTest = true;
        bool IsFilterStatus = false;
        bool ivStatCert = true;
        Dispatcher.Invoke((System.Action)delegate() {
          IVTest = IVCheck.IsChecked.Value;
          IsFilterStatus = StatusFilter.IsChecked.Value;
          if (ivStats.Text != "DevTrack") {
            ivStatCert = false;
          }
        }); 
        #region data formatting
        foreach (WR dd in data) {
          foreach (Jissues info in dd.issues) {
            //foreach (Jissues info in dd.issues) {

            WRQuery WRQ = new WRQuery(info, pid);
            
            if (!ContainsLoop(JKEYS,WRQ.JKey)) {
              int o = 0;
              string IVValue = "";
              if (IVTest&&DevTrack!=null) {
                IVValue = getIV(WRQ.WR, DevTrack);
              } else {
                IVValue = WRQ.Variances;
              }
              dataSheet.Cells[i, colS + o] = WRQ.JKey;//JIRA ID
              JKEYS.Add(WRQ.JKey);
              o++;
              dataSheet.Cells[i, colS + o] = IVValue;             
              o++;
              dataSheet.Cells[i, colS + o] = WRQ.WR;
              o++;
              dataSheet.Cells[i, colS + o] = WRQ.IssueType;
              o++;
              dataSheet.Cells[i, colS + o] = WRQ.Suite;//just team //getSuite(WRQ.JKey);
              o++;
              dataSheet.Cells[i, colS + o] = WRQ.Product;
              o++;
              dataSheet.Cells[i, colS + o] = WRQ.CurrentRelease;
              o++;
              dataSheet.Cells[i, colS + o] = "'"+WRQ.Summary;
              o++;
              dataSheet.Cells[i, colS + o] = WRQ.ProjDeliveryPriority;
              o++;
              if (ivStatCert&&DevTrack!=null) {
                dataSheet.Cells[i, colS + o] = getIVStatus(WRQ.WR, DevTrack);
              } else {
                dataSheet.Cells[i, colS + o] = WRQ.FirstVarianceStatus;
              }              
              o++;
              //if (DevTrack != null) {
              //  dataSheet.Cells[i, colS + o] = "DevT: " + getDStatus(WRQ.WR, DevTrack) + Environment.NewLine + "JIRA as " + WRQ.Resolution + Environment.NewLine + WRQ.Status;
              //} else {
                dataSheet.Cells[i, colS + o] = WRQ.Resolution+Environment.NewLine+WRQ.Status;
              //}
              o++;
              dataSheet.Cells[i, colS + o] = "No Change";
              o++;
              //if (DevTrack != null) {
              //  dataSheet.Cells[i, colS + o] = "DevT: " + getDStatus(WRQ.WR, DevTrack) + Environment.NewLine + "JIRA as " + WRQ.Resolution + Environment.NewLine + WRQ.Status;
              //} else {
                dataSheet.Cells[i, colS + o] = WRQ.Resolution + Environment.NewLine + WRQ.Status;
              //}
              o++;             
              dataSheet.Cells[i, colS + o] = WRQ.Commitment;
              o++;
              dataSheet.Cells[i, colS + o] = "No Change";
              o++;//change?
              dataSheet.Cells[i, colS + o] = WRQ.Commitment;
              o++;//latest commitment
              dataSheet.Cells[i, colS + o] = WRQ.DueDate;
              o++;
              dataSheet.Cells[i, colS + o] = "No Change";
              o++;//change?
              dataSheet.Cells[i, colS + o] = WRQ.DueDate;
              o++;//latest planned due date
              dataSheet.Cells[i, colS + o] = WRQ.TargetRelease;
              o++;
              dataSheet.Cells[i, colS + o] = "No Change";
              o++;//change?
              dataSheet.Cells[i, colS + o] = WRQ.TargetRelease;
              o++;//latest target release
              dataSheet.Cells[i, colS + o] = WRQ.Created;
              o++;
              dataSheet.Cells[i, colS + o] = WRQ.RequestedDate;
              o++;
              dataSheet.Cells[i, colS + o] = WRQ.ResolutionDate;
              i++;

            }
          }
          Dispatcher.Invoke((System.Action)delegate() {
            pBar.Value += k;
          });
        }
        #endregion
        #region Color, width, borders, etc
        int maxRowsA = dataSheet.UsedRange.Rows.Count;
        Excel.Range rangg = dataSheet.get_Range("A2:AA"+maxRowsA.ToString());
        dataSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, rangg, Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "MyTableStyle";
        dataSheet.ListObjects.get_Item("MyTableStyle").TableStyle = "TableStyleMedium1";

        dataSheet.get_Range("A:AA", Type.Missing).EntireColumn.VerticalAlignment = XlVAlign.xlVAlignCenter;
        dataSheet.get_Range("B:Z", Type.Missing).EntireColumn.WrapText = true;
        dataSheet.get_Range("A:F", Type.Missing).EntireColumn.ColumnWidth = 12;
        dataSheet.get_Range("G:G", Type.Missing).EntireColumn.ColumnWidth = 21;
        dataSheet.get_Range("C:C", Type.Missing).EntireColumn.ColumnWidth = 8;
        dataSheet.get_Range("H:H", Type.Missing).EntireColumn.ColumnWidth = 32;
        dataSheet.get_Range("I:Y", Type.Missing).EntireColumn.ColumnWidth = 18;
        dataSheet.get_Range("Z:Z", Type.Missing).EntireColumn.ColumnWidth = 15;
        dataSheet.get_Range("AA:AA", Type.Missing).EntireColumn.ColumnWidth = 32;

        dataSheet.get_Range("H:H", Type.Missing).EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        dataSheet.get_Range("C:C", Type.Missing).EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        dataSheet.get_Range("I:Y", Type.Missing).EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignCenter;

        dataSheet.get_Range("A:AA", Type.Missing).Cells.Font.Size = 8;
        dataSheet.get_Range("A2", "AA2").Cells.Font.Size = 11;
        dataSheet.get_Range("A1", Type.Missing).Cells.Font.Size = 11;
        dataSheet.get_Range("A2", "AA2").Cells.Font.Bold = true;
        dataSheet.get_Range("A1", Type.Missing).Cells.Font.Bold = true;

        //dataSheet.get_Range("A:G", Type.Missing).EntireColumn.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SkyBlue);
        dataSheet.get_Range("A2", "AA2").HorizontalAlignment = XlHAlign.xlHAlignLeft;
        dataSheet.get_Range("A2", "AA2").Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
        dataSheet.get_Range("A2", "I2").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DeepSkyBlue);
        dataSheet.get_Range("J2", "M2").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
        dataSheet.get_Range("N2", "Y2").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.MediumSeaGreen);
        dataSheet.get_Range("Z2", "AA2").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCoral);
        dataSheet.get_Range("A1", "AA1").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
        
        dataSheet.get_Range("A:AA", Type.Missing).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
        dataSheet.get_Range("A1", "AA1").Borders.LineStyle = XlLineStyle.xlLineStyleNone;
        dataSheet.get_Range("I1", "AA1").HorizontalAlignment = XlHAlign.xlHAlignLeft;
        

        #endregion
        //excelApp.Cells.Locked = true;
        //excelApp.get_Range("A:S",Type.Missing).Locked = true;
        Dispatcher.Invoke((System.Action)delegate() {
          status.Content = "Checking for errors...";
        });
        List<string> NotInJIRA = new List<string>();
        if (DevTrack != null) {
          NotInJIRA = CompareDJ(DevTrack, data);
        }
        Excel.Range merg = dataSheet.get_Range("D1", "H1");
        Excel.Range mergB = dataSheet.get_Range("I1", "P1");
        merg.Merge(Type.Missing);
        mergB.Merge(Type.Missing);
        dataSheet.get_Range("A1", "AA1").EntireRow.RowHeight = 50;
        dataSheet.get_Range("A1", Type.Missing).WrapText = true;

        if (cautions.Count > 0) {
          dataSheet.Cells[1, 4] = "Info: " + string.Join(", ", cautions);
        } else {
          dataSheet.Cells[1, 4] = "";
        }
        if (NotInJIRA.Count > 0) {
          string cautCell = dataSheet.get_Range("D1", Type.Missing).Value;
          if (cautCell == "") {
            dataSheet.Cells[1, 4] = "DevT WRs not in JIRA: " + string.Join(", ", NotInJIRA);
          } else {
            dataSheet.Cells[1, 4] = cautCell + Environment.NewLine + "DevT WRs not in JIRA: " + string.Join(", ", NotInJIRA);
          }
        }

        //dataSheet.Cells[1, 4] = "Info: " + string.Join(", ", cautions);
        //if (NotInJIRA.Count > 0) {
        //  dataSheet.Cells[1, 9] = "DevT WRs not in JIRA: " + string.Join(", ", NotInJIRA);
        //}


        Dispatcher.Invoke((System.Action)delegate() {
          status.Content = "Saving file...";
        });
        bool readOnly = true;
        excelApp.DisplayAlerts = false;
        workbook.SaveAs(file, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, readOnly, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing);
        workbook.Close(true, Type.Missing, Type.Missing);
        excelApp.Quit();
        releaseObject(rangg);
        releaseObject(merg);
        releaseObject(mergB);
        result = true;
      } catch (Exception e) {
        string m = e.Message;
        Dispatcher.Invoke((System.Action)delegate() {
          status.Content = "ugh Excel failed.. try again?";
        });
        result = false;
      } finally {
        
        releaseObject(dataSheet);
        releaseObject(sheets);
        releaseObject(workbook);
        releaseObject(excelApp);
      }
      return result;
    }

    public bool dateComp(string worksheetDate, string JIRADate) {
      //string temp = undate(JIRADate + "T");
      worksheetDate += " ";
      if (worksheetDate.Substring(0, worksheetDate.IndexOf(" ")) == undate(JIRADate + "T")) {
        return true;
      } else {
        return false;
      }
    }

    public bool updateMain(Excel.Worksheet frontPage, List<string> JKeys, string IVStatus, string IssueType, string projDelPrio, string JKey, string SWDCommit, string PlanReleaseDate, string TargetRelease, string RequestedDate/*statuses*/,string DWR, List<DevT> LinkedWRs) {
      int commitCell = 14;
      int issueType = 4;
      int projdelprio = 9;
      int releaseDateCell = 17;
      int ivstatusn = 10;
      int targetCell = 20;
      int requestCell = 11;//Original Status
      int acctCell = 26;
      //int indexer = getIndex(DevTWRs(LinkedWRs), DWR);
      int index = getIndex(JKeys, JKey)+1;//JKeys.FindIndex(s => s == JKey); //WTF SLOW
      string acceptance = "";
      try { acceptance = frontPage.Cells[index, acctCell].Value.ToString(); } catch { }
      if (acceptance != "") {
        acceptance = "Was: " + acceptance;
      }
      if (index >0) {//if exists
        frontPage.Cells[index, commitCell+2] = SWDCommit;
        frontPage.Cells[index, issueType] = IssueType;
        frontPage.Cells[index, projdelprio] = projDelPrio;
        frontPage.Cells[index, ivstatusn] = IVStatus;
        string temp = "";
        try{temp=frontPage.Cells[index, commitCell].Value.ToString();}catch{}
        if (temp == SWDCommit) {
          frontPage.Cells[index, commitCell+1] = "No Change";
        } else {
          frontPage.Cells[index, commitCell+1] = "Changed!";
          frontPage.get_Range("A" + index, "I" + index).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
          frontPage.Cells[index, commitCell + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Violet);
          frontPage.Cells[index, acctCell] = acceptance;
        }
        temp = "xxx";
        frontPage.Cells[index, releaseDateCell+2] = PlanReleaseDate;
        try { temp = frontPage.Cells[index, releaseDateCell].Value.ToString(); } catch { }
        if ( dateComp(temp, PlanReleaseDate)) {
          frontPage.Cells[index, releaseDateCell+1] = "No Change";
        } else {
          frontPage.Cells[index, releaseDateCell+1] = "Changed!";
          frontPage.get_Range("A" + index, "I" + index).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
          frontPage.Cells[index, releaseDateCell + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Violet);
          frontPage.Cells[index, acctCell] = acceptance;
        }
        frontPage.Cells[index, targetCell+2] = TargetRelease;
        temp = "xxx";
        try { temp = frontPage.Cells[index, targetCell].Value.ToString(); } catch { }
        if (temp == TargetRelease) {
          frontPage.Cells[index, targetCell+1] = "No Change";
        } else {
          frontPage.Cells[index, targetCell + 1] = "Changed!";
          frontPage.get_Range("A" + index, "I" + index).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
          frontPage.Cells[index, targetCell + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Violet);
          frontPage.Cells[index, acctCell] = acceptance;
        }
        frontPage.Cells[index, requestCell+2] = RequestedDate;
        temp = "xxx";
        try { temp = frontPage.Cells[index, requestCell].Value.ToString(); } catch { }
        if (temp==RequestedDate) {
          frontPage.Cells[index, requestCell+1] = "No Change";
        } else {
          frontPage.Cells[index, requestCell + 1] = "Changed!";
          frontPage.get_Range("A" + index, "I" + index).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
          frontPage.Cells[index, requestCell + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Violet);
          frontPage.Cells[index, acctCell] = acceptance;
        }
        return true;
      }
      return false;
    }

    private void updater(string fileName) {
      bool Jspecial = false;
      bool ivstatusCert = true;
      List<string> cautions = new List<string>();
      Dispatcher.Invoke((System.Action)delegate() {
          pBar.Visibility = Visibility.Visible;
          pBox.IsEnabled = false;
          ITObtn.IsEnabled = false;
          updateBtn.IsEnabled = false;
          pBar.Value = 0;
          initBtn.IsEnabled = false;
          status.Content = "Analyzing Spreadsheet Validity...";
        });

        Excel.Application excelApp = new Excel.Application(); ;
        excelApp.DisplayAlerts = false;
        string fn = fileName;
        Excel.Workbook wrokbak = excelApp.Workbooks.Open(fn, false, false, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        Excel.Sheets sheets = wrokbak.Sheets;
        Excel.Worksheet dataSheet = wrokbak.Sheets[1];
        int maxRow = dataSheet.UsedRange.Rows.Count;
        string PID = "ASDF DNE N/A";
        try { PID = dataSheet.Cells[1, 1].Value.ToString(); } catch { }
        if (PID.Substring(0, 3) == "J*:") {
          Jspecial = true;
          PID = PID.Substring(3, PID.Length - 3);
        }
        string DDate = "Initialized: 01/01/01";
        try { DDate = dataSheet.Cells[1, 2].Value.ToString(); } catch { }
        List<string> DevTcheck = getPID();
        //string tedfr = DDate.Substring(13, 8);
        if (getIndex(DevTcheck, PID) >= 0 || Jspecial) {
          //  return false;
          //}
          int wsC = wrokbak.Sheets.Count;
          Excel.Worksheet nweshat = (Worksheet)sheets.Add(Type.Missing, sheets[1], Type.Missing, Type.Missing);
          if (wsC < 2) {
            nweshat.Name = "Update";
          } else {
            nweshat.Name = "Update" + wsC.ToString();
          }
          int newMaxRow = nweshat.UsedRange.Rows.Count;

          //get current WRs

          string speedval = "Gentle";
          Dispatcher.Invoke((System.Action)delegate() {
            pBar.Visibility = Visibility.Visible;
            pBox.IsEnabled = false;
            ITObtn.IsEnabled = false;
            updateBtn.IsEnabled = false;
            pBar.Value = 0;
            initBtn.IsEnabled = false;
            speedval = dbSpeed.Text;
            if (ivStats.Text != "DevTrack"||Jspecial) {
              ivstatusCert = false;
            }
            if (speedval != "Gentle") {
              status.Content = "Reading Databases...";
            } else {
              status.Content = "Gently Reading Databases...";
            }            
            
          });


          List<DevT> LinkedWRs = new List<DevT>();
          if (!Jspecial) { 
            LinkedWRs = getWRs(PID);
          }
          List<WR> js = new List<WR>();
          List<string> urls = new List<string>();
          if (!Jspecial) {
            urls = makeUrl(LinkedWRs);
          }          
          int speed = 1000;
          if (speedval == "Quick") {
            speed = 260;
          } else if (speedval == "DDOS") {
            speed = 50;
          }
          double k = 0.0;
          if (Jspecial) {

            int total = getJIRAtotal(PID);
            double j = total / 50.0;
            k = 100.0 / j;

            //List<WR> JIRAs = new List<WR>();
            for (int i = 0; i <= total; i += 50) {
              js.Add(pJSON(PID + "&startAt=" + i + "&maxResults=50"));
              Dispatcher.Invoke((System.Action)delegate() {
                pBar.Value += k;
              });
            }
          } else { 
          //get current JIRA
            double j = 100.0;
            k = j / urls.Count();
            foreach (string qry in urls) {
              js.Add(pJSON(qry));
              Dispatcher.Invoke((System.Action)delegate() {
                pBar.Value += k;
              });
              Thread.Sleep(speed);
            }
          }


          //paste into new sheet
          Dispatcher.Invoke((System.Action)delegate() {
            status.Content = "Wrestling Excel...";
            pBar.Value = 0;
          });
          if (js.Count != 0) {
            int rowS = 3;
            int colS = 1;
            #region header for updates
            int p = 0;
            nweshat.Cells[rowS - 2, colS] = "Updated: " + DateTime.Now.Month.ToString("00") + "/" + DateTime.Now.Day.ToString("00") + "/" + DateTime.Now.Year.ToString().Substring(2, 2);
            nweshat.Cells[rowS - 1, colS + p] = "JIRA ID";   //A
            p++;
            nweshat.Cells[rowS - 1, colS + p] = "Variances";   //B
            p++;
            nweshat.Cells[rowS - 1, colS + p] = "WR";   //C
            p++;
            nweshat.Cells[rowS - 1, colS + p] = "Issue Type"; //D
            p++;
            nweshat.Cells[rowS - 1, colS + p] = "Team"; //D
            p++;
            nweshat.Cells[rowS - 1, colS + p] = "Product"; //E
            p++;
            nweshat.Cells[rowS - 1, colS + p] = "Summary";  //F
            p++;
            nweshat.Cells[rowS - 1, colS + p] = "Proj Del Priority";  //G
            //nweshat.Cells[rowS - 1, colS + p] = "Variance Status";  //G
            p++;
            nweshat.Cells[rowS - 1, colS + p] = "IV Status"; //H
            p++;
            nweshat.Cells[rowS - 1, colS + p] = "Statuses"; //H
            p++;
            nweshat.Cells[rowS - 1, colS + p] = "Latest Commitment";  //I +Resolution
            p++;
            nweshat.Cells[rowS - 1, colS + p] = "Planned Release Date";  //J
            p++;
            nweshat.Cells[rowS - 1, colS + p] = "Target Release";   //L
            p++;
            nweshat.Cells[rowS - 1, colS + p] = "Created Date";   //M
            p++;
            nweshat.Cells[rowS - 1, colS + p] = "Requested Date";   //N
            p++;
            nweshat.Cells[rowS - 1, colS + p] = "Resolution Date";   //O
            p++;
            #endregion
            int i = rowS;
            bool IVTest = true;
            bool statusCheck = false;
            Dispatcher.Invoke((System.Action)delegate() {
              IVTest = IVCheck.IsChecked.Value;
              statusCheck = StatusFilter.IsChecked.Value;
            });
            #region data tags for updates
            
            List<string> JKeys = getList(dataSheet, 1);
            foreach (WR dd in js) {
              //List<WRQuery> WRQl = new List<WRQuery>();
              //Parallel.ForEach(dd.issues, info => {
              //  WRQuery WRQ = new WRQuery(info, PID);
              //  WRQl.Add(WRQ);
              //});
              //Parallel.ForEach(js, dd => {
              foreach (Jissues info in dd.issues) {
                WRQuery WRQ = new WRQuery(info, PID);
                int o = 0;
                string IVvalue = "";
                if (IVTest&&!Jspecial) {
                  IVvalue = getIV(WRQ.WR, LinkedWRs);
                } else {
                  IVvalue = WRQ.Variances;
                }
                string VarState = "";
                if (ivstatusCert) {
                  VarState = "DevT: "+getIVStatus(WRQ.WR, LinkedWRs);
                } else {
                  VarState = WRQ.FirstVarianceStatus;
                }

                //string jkty = "";
                //try { jkty = unnull(info.key); } catch { }//JIRA ID    #A
                nweshat.Cells[i, colS + o] = WRQ.JKey;//JIRA ID    #A
                o++;
                nweshat.Cells[i, colS + o] = IVvalue;
                o++;
                nweshat.Cells[i, colS + o] = WRQ.WR;
                o++;
                nweshat.Cells[i, colS + o] = WRQ.IssueType;
                o++;
                nweshat.Cells[i, colS + o] = WRQ.Suite;
                o++; 
                nweshat.Cells[i, colS + o] = WRQ.Product;
                o++;
                nweshat.Cells[i, colS + o] = WRQ.Summary;
                o++;
                nweshat.Cells[i, colS + o] = WRQ.ProjDeliveryPriority;
                o++;
                nweshat.Cells[i, colS + o] = VarState;
                o++;
                //if (Jspecial) {
                  nweshat.Cells[i, colS + o] = WRQ.Resolution + Environment.NewLine + WRQ.Status;
                //} else {
                //  nweshat.Cells[i, colS + o] = "DevT: " + getDStatus(WRQ.WR, LinkedWRs) + Environment.NewLine + "JIRA as " + WRQ.Resolution + Environment.NewLine + WRQ.Status;
                //}                
                o++;
                nweshat.Cells[i, colS + o] = WRQ.Commitment;
                o++;
                nweshat.Cells[i, colS + o] = WRQ.DueDate;
                o++;
                //nweshat.Cells[i, colS + o] = WRQ.CurrentRelease;
                //o++;
                nweshat.Cells[i, colS + o] = WRQ.TargetRelease;
                o++;
                nweshat.Cells[i, colS + o] = WRQ.Created;
                o++;
                nweshat.Cells[i, colS + o] = WRQ.RequestedDate;
                o++;
                nweshat.Cells[i, colS + o] = WRQ.ResolutionDate;
                o++;
                i++;
                bool JRKeyExists = updateMain(dataSheet, JKeys, VarState, WRQ.IssueType, WRQ.ProjDeliveryPriority, WRQ.JKey, WRQ.Commitment, WRQ.DueDate, WRQ.TargetRelease, WRQ.Resolution + Environment.NewLine + WRQ.Status, WRQ.WR, LinkedWRs);
                if (!JRKeyExists&&WRQ.Status!="Closed") {
                  cautions.Add(WRQ.JKey + " is new!");
                  int MaxRowC = dataSheet.UsedRange.Rows.Count + 1;
                  int newColC = 1;

                  dataSheet.Cells[MaxRowC, newColC] = WRQ.JKey;
                  dataSheet.get_Range("A"+MaxRowC, "Y"+MaxRowC).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Violet);
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = IVvalue;
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = WRQ.WR;
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = WRQ.IssueType;
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = WRQ.Suite;//Team //getSuite(WRQ.JKey);
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = WRQ.Product;
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = WRQ.CurrentRelease;
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = WRQ.Summary;
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = WRQ.ProjDeliveryPriority;
                  newColC++;
                    dataSheet.Cells[MaxRowC, newColC] = VarState;
                  newColC++;

                  //if (Jspecial) {
                    dataSheet.Cells[MaxRowC, newColC] = WRQ.Resolution + Environment.NewLine + WRQ.Status;
                    newColC++;
                    dataSheet.Cells[MaxRowC, newColC] = "New Item";
                    newColC++;
                    dataSheet.Cells[MaxRowC, newColC] = WRQ.Resolution + Environment.NewLine + WRQ.Status;
                  //} else {
                  //  string AnDevTtempstat = getDStatus(WRQ.WR, LinkedWRs);
                  //  string AnVarStatus = "notsure";
                  //  if (WRQ.Progress == AnDevTtempstat) {
                  //    AnVarStatus = WRQ.Progress;
                  //  } else {
                  //    AnVarStatus = "JIRA: " + WRQ.Progress + Environment.NewLine + "DevT: " + AnDevTtempstat;
                  //  }
                  //  dataSheet.Cells[MaxRowC, newColC] = AnVarStatus + Environment.NewLine + WRQ.Resolution + Environment.NewLine + WRQ.Status;
                  //  newColC++;
                  //  dataSheet.Cells[MaxRowC, newColC] = "New Item";
                  //  newColC++;
                  //  dataSheet.Cells[MaxRowC, newColC] = AnVarStatus + Environment.NewLine + WRQ.Resolution + Environment.NewLine + WRQ.Status;
                  //}
                  
                  //WR STATUS
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = WRQ.Commitment;
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = "New Item";
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = WRQ.Commitment;
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = WRQ.DueDate;
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = "New Item";
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = WRQ.DueDate;
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = WRQ.TargetRelease;
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = "New Item";
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = WRQ.TargetRelease;
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = WRQ.Created;
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = WRQ.RequestedDate;
                  newColC++;
                  dataSheet.Cells[MaxRowC, newColC] = WRQ.ResolutionDate;

                }

              }
              Dispatcher.Invoke((System.Action)delegate() {
                pBar.Value += k;
              });
              //});
            }
            #endregion
            Dispatcher.Invoke((System.Action)delegate() {
              status.Content = "Checking for errors...";
            });
            List<string> newJKeys = getList(nweshat, 1);
            List<string> newWRcol = getList(nweshat, 3);
            List<int> removedWRs = new List<int>();
            for (int ix = 2; ix < JKeys.Count; ix++) {
              int vv = getIndex(newJKeys, JKeys[ix]);
              if (vv < 0) {//couldn't current JKey in new JKeys
                removedWRs.Add(ix);
              }
            }

            //Parallel.ForEach(JKeys, jkuy => {
            //  int vv = getIndex(newJKeys, jkuy);
            //  if (vv >= 0) {
            //    removedWRs.Add(vv);
            //  }
            //});
            foreach (int lineNum in removedWRs) {
              int unlink = lineNum + 1;
              dataSheet.get_Range("A" + unlink, "Y" + unlink).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Violet);
              //dataSheet.Cells[lineNum + 3, 1]
              dataSheet.Cells[unlink, 13] = "Issue Unlinked";
              dataSheet.Cells[unlink, 16] = "Issue Unlinked";
              dataSheet.Cells[unlink, 19] = "Issue Unlinked";
              dataSheet.Cells[unlink, 23] = "Issue Unlinked";
              string acceptance = "";
              try { acceptance = dataSheet.Cells[unlink, 26].Value.ToString(); } catch { }
              if (acceptance != "") {
                dataSheet.Cells[unlink, 26] = "Was: " + acceptance;
              }
            }
            
            List<string> NotInJIRA = new List<string>();
            foreach (DevT dt in LinkedWRs) {
              if (getIndex(newWRcol, dt.WR) < 0) {
                NotInJIRA.Add(dt.WR);
              }
            }

            List<string> newIssueType = getList(nweshat, 4);
            List<string> newProjDel = getList(nweshat, 8);


            string initDate = dataSheet.get_Range("B1", Type.Missing).Value;
            try { initDate = initDate.Substring(0, initDate.IndexOf(Environment.NewLine)); } catch { }
            dataSheet.Cells[1, 2] = initDate + Environment.NewLine + "Updated: " + DateTime.Now.Month.ToString("00") + "/" + DateTime.Now.Day.ToString("00") + "/" + DateTime.Now.Year.ToString().Substring(2, 2);
            
            if (cautions.Count > 0) {
              dataSheet.Cells[1, 4] = "New Info: " + string.Join(", ", cautions);
            } else {
              dataSheet.Cells[1, 4] = "";
            }
            if (NotInJIRA.Count > 0) {
              string cautCell = dataSheet.get_Range("D1", Type.Missing).Value;
              if (cautCell == "") {
                dataSheet.Cells[1, 4] = "DevT WRs not in JIRA: " + string.Join(", ", NotInJIRA);
              } else {
                dataSheet.Cells[1, 4] = cautCell + Environment.NewLine + "DevT WRs not in JIRA: " + string.Join(", ", NotInJIRA);
              }
            }
            dataSheet.Select();
          }


          Dispatcher.Invoke((System.Action)delegate() {
            status.Content = "Saving Excel...";
          });

          wrokbak.Save();
          wrokbak.Close();
          excelApp.Quit();

          Dispatcher.Invoke((System.Action)delegate() {
            pBar.Visibility = Visibility.Hidden;
            pBar.Value = 0;
            pBox.IsEnabled = true;
            ITObtn.IsEnabled = true;
            //initBtn.IsEnabled = true;
            updateBtn.IsEnabled = true;
            status.Content = "Completed!";
          });
          releaseObject(dataSheet);
          releaseObject(nweshat);
          releaseObject(nweshat);
          releaseObject(sheets);
          releaseObject(wrokbak);
          releaseObject(excelApp);
        } else {
          wrokbak.Close();
          excelApp.Quit();
          releaseObject(dataSheet);
          releaseObject(sheets);
          releaseObject(wrokbak);
          releaseObject(excelApp);
          Dispatcher.Invoke((System.Action)delegate() {
            pBar.Visibility = Visibility.Hidden;
            pBar.Value = 0;
            pBox.IsEnabled = true;
            ITObtn.IsEnabled = true;
            initBtn.IsEnabled = true;
            updateBtn.IsEnabled = true;
            status.Content = "Could not verify Excel file's format.";
          });
        }
      }

    static void releaseObject(object obj) {
      try {
        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        obj = null;
      } catch (Exception ex) {
        obj = null;
        Console.WriteLine("Exception Occured while releasing object " + ex.ToString());
      } finally {
        GC.Collect();
        GC.WaitForPendingFinalizers();
      }
    }

    public void ITOreport(object sender, DoWorkEventArgs e) {
      string dt = DateTime.Now.Day.ToString("00") + DateTime.Now.Month.ToString("00") + DateTime.Now.Year.ToString().Substring(2, 2);
      string jqr = "";
      Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
      dlg.FileName = "xxx_" + dt + "_SWRs";
      dlg.DefaultExt = ".xlsx";
      dlg.Filter = "Excel Workbook (.xlsx)|*.xlsx";
      Nullable<bool> result = dlg.ShowDialog();
      if (result == true) {
        Dispatcher.Invoke((System.Action)delegate() {
          pBar.Visibility = Visibility.Visible;
          pBox.IsEnabled = false;
          ITObtn.IsEnabled = false;
          updateBtn.IsEnabled = false;
          pBar.Value = 0;
          initBtn.IsEnabled = false;
          status.Content = "Reading Databases...";
          jqr = JIRAqry.Text;
        });

        int total = getJIRAtotal(jqr);
        double j = total / 50.0;
        double k = 100.0 / j;


        MessageBoxResult dialogResult = MessageBox.Show("This will report "+total+" items... Okay to continue?", "JIRA-Only Query",MessageBoxButton.OKCancel);
        if (dialogResult == MessageBoxResult.OK) {
          //do something
        


        List<WR> JIRAs = new List<WR>();
        for (int i = 0; i <= total; i += 50) {
          JIRAs.Add(pJSON(jqr + "&startAt=" + i + "&maxResults=50"));
          Dispatcher.Invoke((System.Action)delegate() {
            pBar.Value += k;
          });
        }
        Dispatcher.Invoke((System.Action)delegate() {
          status.Content = "Wrestling Excel...";
          pBar.Value = 0;
        });
        bool worked = false;
        if (JIRAs.Count>0) {
          worked = createBlankXL(dlg.FileName,"J*:"+jqr, JIRAs, null);
        }
        Dispatcher.Invoke((System.Action)delegate() {
          pBar.Visibility = Visibility.Hidden;
          pBar.Value = 0;
          pBox.IsEnabled = true;
          ITObtn.IsEnabled = true;
          //initBtn.IsEnabled = true;
          updateBtn.IsEnabled = true;
          if (worked) {
            status.Content = "Completed!";
          }
        });
        } else if (dialogResult == MessageBoxResult.Cancel) {
          Dispatcher.Invoke((System.Action)delegate() {
            pBar.Visibility = Visibility.Hidden;
            pBar.Value = 0;
            pBox.IsEnabled = true;
            ITObtn.IsEnabled = true;
            updateBtn.IsEnabled = true;
            status.Content = "Cancelled.";
          });
        }
      }
    }

    private void pBox_SelectionChanged(object sender, SelectionChangedEventArgs e) {
      string p = e.AddedItems[0].ToString();
      if (p != "") {
        initBtn.IsEnabled = true;
      } else {
        initBtn.IsEnabled = false;
      }
    }

    public void initialize(object sender, DoWorkEventArgs e) {
      string pid = (string)e.Argument;
      string dt = DateTime.Now.Day.ToString("00") + DateTime.Now.Month.ToString("00") + DateTime.Now.Year.ToString().Substring(2, 2);

      Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
      dlg.FileName = RemoveSpecialCharacters(pid) + "_" + dt + "_WRs";
      dlg.DefaultExt = ".xlsx";
      dlg.Filter = "Excel Workbook (.xlsx)|*.xlsx";
      Nullable<bool> result = dlg.ShowDialog();
      if (result == true) {
        List<DevT> DevTrack = getWRs(pid);
        List<WR> js = new List<WR>();
        //WR js = pJSON(makeUrl(wrList));
        List<string> urls = makeUrl(DevTrack);
        double j = 100.0;
        double k = j / urls.Count();
        string speedval = "Gentle";

        Dispatcher.Invoke((System.Action)delegate() {
          pBar.Visibility = Visibility.Visible;
          pBox.IsEnabled = false;
          ITObtn.IsEnabled = false;
          updateBtn.IsEnabled = false;
          pBar.Value = 0;
          initBtn.IsEnabled = false;
          speedval = dbSpeed.Text;
          if (speedval != "Gentle") {
            status.Content = "Reading Databases...";
          } else {
            status.Content = "Gently Reading Databases...";
          }
        });
        int speed = 1000;
        if (speedval == "Quick") {
          speed = 260;
        } else if (speedval == "DDOS") {
          speed = 0;
        }
        foreach (string qry in urls) {
          js.Add(pJSON(qry));
          Dispatcher.Invoke((System.Action)delegate() {
            pBar.Value += k;
          });
          Thread.Sleep(speed);
        }
        Dispatcher.Invoke((System.Action)delegate() {
          status.Content = "Wrestling Excel...";
          pBar.Value = 0;
        });
        bool worked = false;
        if (js.Count != 0) {
          worked = createBlankXL(dlg.FileName, pid, js, DevTrack);
        }
        Dispatcher.Invoke((System.Action)delegate() {
          pBar.Visibility = Visibility.Hidden;
          pBar.Value = 0;
          pBox.IsEnabled = true;
          ITObtn.IsEnabled = true;
          //initBtn.IsEnabled = true;
          updateBtn.IsEnabled = true;
          if (worked) {
            status.Content = "Completed!";
          }
        });
      }
    }

    private void updateXL(object sender, DoWorkEventArgs e) {
      Microsoft.Win32.OpenFileDialog ofg = new Microsoft.Win32.OpenFileDialog();
      ofg.DefaultExt = ".xlsx";
      ofg.Filter = "WR Spreadsheet (.xlsx)|*.xlsx";
      Nullable<bool> result = ofg.ShowDialog();
      if (result == true) {
        updater(ofg.FileName);
      }
    }
    private void WRPR(object sender, DoWorkEventArgs e) {
      string pid = (string)e.Argument;
      string dt = DateTime.Now.Day.ToString("00") + DateTime.Now.Month.ToString("00") + DateTime.Now.Year.ToString().Substring(2, 2);
      Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
      dlg.FileName = RemoveSpecialCharacters(pid) + "_" + dt + "_WRs";
      dlg.DefaultExt = ".xlsx";
      dlg.Filter = "Excel Workbook (.xlsx)|*.xlsx";
      Nullable<bool> result = dlg.ShowDialog();
      List<DevT> DevTrackList = new List<DevT>();
      if (result == true) {
        DevTrackList = getWRplus(pid);
        List<DevT> OpenDevItems = new List<DevT>();
        List<WR> JIRAitems = new List<WR>();
        Parallel.ForEach(DevTrackList, DTitem => {
          if (DTitem.status != "Confirm Verified" && DTitem.status != "Confirm Duplicate" && DTitem.status != "Confirm Reject") {
            OpenDevItems.Add(DTitem);
          }
        });
        List<string> urls = makeUrl(OpenDevItems);
        double j = 100.0;
        double k = j / urls.Count();
        string speedval = "Gentle";

        Dispatcher.Invoke((System.Action)delegate() {
          pBar.Visibility = Visibility.Visible;
          WRPRbtn.IsEnabled = false;
          pBox.IsEnabled = false;
          ITObtn.IsEnabled = false;
          updateBtn.IsEnabled = false;
          pBar.Value = 0;
          initBtn.IsEnabled = false;
          speedval = dbSpeed.Text;
          if (speedval != "Gentle") {
            status.Content = "Reading Databases...";
          } else {
            status.Content = "Gently Reading Databases...";
          }
        });
        int speed = 1000;
        if (speedval == "Quick") {
          speed = 260;
        } else if (speedval == "DDOS") {
          speed = 0;
        }
        foreach (string qry in urls) {
          JIRAitems.Add(pJSON(qry));
          Dispatcher.Invoke((System.Action)delegate() {
            pBar.Value += k;
          });
          Thread.Sleep(speed);
        }
        Dispatcher.Invoke((System.Action)delegate() {
          status.Content = "Wrestling Excel...";
          pBar.Value = 0;
        });
        bool worked = false;
        if (JIRAitems.Count != 0) {
          worked = createWRPR(dlg.FileName, pid, JIRAitems, OpenDevItems);
        }
        Dispatcher.Invoke((System.Action)delegate() {
          pBar.Visibility = Visibility.Hidden;
          pBar.Value = 0;
          pBox.IsEnabled = true;
          ITObtn.IsEnabled = true;
          updateBtn.IsEnabled = true;
          if (worked) {
            status.Content = "Completed!";
          }
        });

      }//if file name chosen
    }


    private void Button_Click(object sender, RoutedEventArgs e) {
      BackgroundWorker bw1 = new BackgroundWorker();
      bw1.DoWork += new DoWorkEventHandler(initialize);
      bw1.RunWorkerAsync(pBox.Text);
    }

    public List<string> makeUrl(List<DevT> wrList) {
      List<string> urls = new List<string>();
      string url = "";
      int count = wrList.Count;
      for (int i = 0; i <= count - 1; i++) {
        if (url.Length != 0) {
          url += "%20OR%20";
        }
        url += "WR%20~%20\"" + wrList[i].WR + "\"";
        if (i % 30 == 0 && i != 0) {
          urls.Add(url);
          url = "";
        }
      }
      if (url != "") {
        urls.Add(url);
      }
      return urls;
    }

    public static string RemoveSpecialCharacters(string str) {
      StringBuilder sb = new StringBuilder();
      foreach (char c in str) {
        if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '.' || c == '_') {
          sb.Append(c);
        }
      }
      return sb.ToString();
    }

    private void Button_Click_1(object sender, RoutedEventArgs e) {

      BackgroundWorker bw1 = new BackgroundWorker();
      bw1.DoWork += new DoWorkEventHandler(updateXL);
      bw1.RunWorkerAsync();
    }

    private void ITObtn_Click(object sender, RoutedEventArgs e) {
      BackgroundWorker bw1 = new BackgroundWorker();
      bw1.DoWork += new DoWorkEventHandler(ITOreport);
      bw1.RunWorkerAsync();
    }

    private void Button_Click_2(object sender, RoutedEventArgs e) {
      BackgroundWorker bw1 = new BackgroundWorker();
      bw1.DoWork += new DoWorkEventHandler(WRPR);
      bw1.RunWorkerAsync(pBox.Text);
    }

  }
}
