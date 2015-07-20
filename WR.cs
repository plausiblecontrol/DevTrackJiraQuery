using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ProjectWRQuery {
  public class WR {
    public List<Jissues> issues { get; set; }
  }
  public class Jissues {
    public Jfields fields { get; set; }
    public string key { get; set; }
  }
  public class Jfields {
    public customField customfield_10502 { get; set; } //product "monarchNET Advanced Tabulars"
    public customField customfield_10501 { get; set; } //status "Closed"
    //public customField customfield_11200 { get; set; } //fix "Fixed"
    public customField customfield_10801 { get; set; } //verification? "Confirm Verified"
    public customField customfield_12100 { get; set; } //commitment "Accepted & Planned"
    public customField customfield_12802 { get; set; }
    public customField customfield_13000 { get; set; } //Project Delivery Priority
    public customField resolution { get; set; } //    "Switched Teams"
    public string customfield_10701 { get; set; } //release start "AdvancedTabulars_v1_8_9_8_wr73344_start"
    public string customfield_10702 { get; set; } //release end "AdvancedTabulars_v1_8_9_8_wr73344_1_end"
    public string customfield_10407 { get; set; } //target release "1.8.10.0"
    public string customfield_10401 { get; set; } //wr# "73344"
    public string customfield_12101 { get; set; }//requested date "2015-01-16T16:47:29.000-0600"
    public string customfield_10800 { get; set; }//current release "6.0.10.0"
    public string created { get; set; } //"2014-04-21T16:27:09.000-0500"
    public string updated { get; set; } //"2014-12-19T09:05:08.000-0600"
    public string resolutiondate { get; set; } //Completed date "2014-07-22T10:39:55.000-0500"
    public string description { get; set; }//LOTS OF STUFF HERE including new lines
    public List<string> customfield_11001 { get; set; } //project variances "PAC1_IV-97"
    //public string description { get; set; } //"Right click menu does not open with expected items on single field editors. The menu does not show any menu items other that edit display or  the command associated with field. This show the same items, Dynamic menu items, help, etc, as on SE displays and AT gridviews."
    public string summary { get; set; } //"(WR-73344) Right click menu does not open with expected items on single field editors."
    public List<versions> fixVersions { get; set; }
    public customField priority { get; set; }
    public customField status { get; set; }
    public customField issuetype { get; set; }
    public customField project { get; set; }
    public string duedate { get; set; }//planned due date "2014-12-10"
  }
  public class versions {
    public string description { get; set; } //suite "Visualization Platform Series 6.2"
    public string name { get; set; } //suitename "VP 6.2"
    public string released { get; set; } //released status "false"
  }
  public class customField {
    public string value { get; set; }
    public string name { get; set; }
  }
}
