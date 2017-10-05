using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookRESTAPI_UpdateCalendar
{
    class Constants
    {
        internal const string ClientId = "ENTER_YOUR_CLIENT_ID_HERE";

        internal const string OutlookEndpoint = "https://outlook.office.com/";

        internal const string OutlookRestApiEndpoint = OutlookEndpoint + "api/v2.0/me/";
        
        internal static readonly string[] Scopes = new string[] { OutlookEndpoint + "calendars.readwrite" };
        
        internal const string FindEvent =
            OutlookRestApiEndpoint + "calendarview?startdatetime={0}&enddatetime={1}&subject={2}";

        internal const string MasterEvent = OutlookRestApiEndpoint + "events/{0}";



    }
}
