using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ProjectWRQuery {
  public class DevT {
    public string WR;
    public string status;
    public string IV;
    public string IVstatus;
    public string NeedByDate;
    public string NeedByEvent;

    public DevT(string WR_Num_As_String, string WR_Status, string InternalVariance, string InternalVarStatus, string dNBD, string dNBE) {
      WR = WR_Num_As_String;
      status = WR_Status;
      IV = InternalVariance;
      IVstatus = InternalVarStatus;
      NeedByDate = dNBD;
      NeedByEvent = dNBE;
    }
    //public string getStatus(string WR_Num_As_String) {

    //}
  }
}
