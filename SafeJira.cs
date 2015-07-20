using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ProjectWRQuery {
  public class WRQuery {
    public string JKey;
    public string Product;
    public string Progress;
    public string Fix;
    public string Verification;
    public string Commitment;
    public string Resolution;
    public string ReleaseStart;
    public string ReleaseEnd;
    public string TargetRelease;
    public string WR;
    public string Suite;
    public string RequestedDate;
    public string CurrentRelease;
    public string Created;
    public string Updated;
    public string ResolutionDate;
    public string Variances;
    public string FirstVarianceStatus;
    public string Summary;
    public string FixDescription;
    public string FixName;
    public string ProjDeliveryPriority;
    public string FixReleased;
    public string Priority;
    public string Status;
    public string IssueType;
    public string Project;
    public string DueDate;
    public string Description;
    public WRQuery(Jissues info, string PID) {
      string temp = "";
      try { temp = unnull(info.key); } catch { }//JIRA ID    #A
      JKey = temp;//JIRA ID    #A
      temp = "";
      try { temp = filterOut(info.fields.customfield_11001, PID); } catch { }//Variances   #B
      Variances = temp;
      temp = "Nothing in JIRA";
      try { temp = unnull(info.fields.customfield_12802.value); } catch { }//Variances   #B
      Suite = temp;
      temp = "";
      try { temp = unnull(info.fields.customfield_10401); } catch { }//WR    #C
      WR = temp;
      temp = "";
      try { temp = unnull(info.fields.description); } catch { }//WR    #C
      Description = temp;
      temp = "None";
      try { temp = unnull(info.fields.customfield_13000.value); } catch { }//WR    #C
      ProjDeliveryPriority = temp;
      temp = "";
      try { temp = unnull(info.fields.status.name); } catch { }//WR    #C
      Status = temp;
      temp = "";
      try { temp = unnull(info.fields.issuetype.name); } catch { }//Issue Type    #D
      IssueType = temp;
      temp = "";
      try { temp = unnull(info.fields.customfield_10502.value); } catch { }//Product    #E
      Product = temp;
      temp = "";
      try {temp = unnull(info.fields.summary); } catch { }//Summary		#F
      Summary = temp;
      temp = "";
      try { temp = unnull(info.fields.customfield_10501.value); } catch { }//Summary		#F
      FirstVarianceStatus = temp;
      temp = "";
      try { temp = unnull(info.fields.priority.name); } catch { }//Priority		#G
      Priority = temp;
      temp = "";
      try { temp = unnull(info.fields.customfield_10801.value); } catch { }
      Progress = temp;
      temp = "";
      try { temp = unnull(info.fields.resolution.name); } catch { temp = "Unresolved"; }
      Resolution = temp;  //STATUS     #I
      temp = "";
      try { temp = unnull(info.fields.customfield_12100.value); } catch { }
      Commitment = temp;//SWDev Commitment		#H
      temp = "";
      try { temp = unnull(info.fields.duedate); } catch { }//Planned Release   //J
      DueDate = temp;//Planned Release   //J 
      temp = "";
      try { temp = unnull(info.fields.customfield_10800); } catch { }     //K
      CurrentRelease = temp;
      temp = "";
      try { temp = unnull(info.fields.customfield_10407); } catch { }    ///L
      TargetRelease = temp;
      temp = "";
      try { temp = undate(unnull(info.fields.created)); } catch { }//////DATE			#M
      Created = temp;
      temp = "";
      try { temp = unnull(info.fields.customfield_12101); } catch { }    //N
      RequestedDate = temp;
      temp = "";
      try { temp = undate(unnull(info.fields.resolutiondate)); } catch { }///////DATE		#O
      ResolutionDate = temp;
      temp = "";
      try { temp = unnull(info.fields.fixVersions[0]); } catch { }///////DATE		#O
      FixDescription = temp;
      temp = "";
      try { temp = unnull(info.fields.fixVersions[1]); } catch { }///////DATE		#O
      FixName = temp;
      temp = "";
      try { temp = unnull(info.fields.fixVersions[2]); } catch { }///////DATE		#O
      FixReleased = temp;
      temp = "";
    }
    private static string unnull(object Value) {
      return Value == null ? "xxx" : Value.ToString();
    }
    private static string undate(string ugly) {
      //"2015-01-16T16:47:29.000-0600"

      //temp=11/26/2014 12:00:00 AM
      //releasedate=2014-11-26
      try {
        string[] a = ugly.Substring(0, ugly.IndexOf("T")).Split(new Char[] { '-' }).ToArray();
        string[] b = { a[1], a[2], a[0] };
        return string.Join("/", b);
      } catch {
        return ugly;
      }

    }
    private string filterOut(List<string> list, string keep) {
      bool IVFilter = false;
      string filter = "";
      try {
        if (keep.IndexOf(":") > 0) {
          filter = keep.Substring(0, keep.IndexOf(":") - 1);
        } else {
          filter = keep.Substring(0, keep.IndexOf(" ") - 1);
        }
      }catch{}


      if (IVFilter) {
        string res = "";
        if (list == null) {
          list = new List<string>();
          list.Add("");
        }
        try {
          //string filter = keep.Substring(0, keep.IndexOf(":") - 1);
          foreach (string s in list) {
            if (s.Contains(filter))
              res += s + ", ";
          }
        } catch {
          res += "Error decoding, ";
        }
        return res;
      } else {
        return string.Join(Environment.NewLine, list.ToArray());
      }
    }
  }
}
