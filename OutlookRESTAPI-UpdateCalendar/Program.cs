using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookRESTAPI_UpdateCalendar
{
    class Program
    {
        static void Main(string[] args)
        {
            // Find a specific event occurrence by querying for events in a particular time range with a specific meeting subject.
            var start = new DateTime(2017, 10, 04);
            var end = new DateTime(2017, 10, 05);
            var subject = "testmeeting2";
            var geteventsUrl = string.Format(Constants.FindEvent, start.ToString("s"), end.ToString("s"), subject);

            var authToken = AuthHelper.GetAuthToken(Constants.Scopes);

            string response = HttpHelper.SendAsync(HttpMethod.Get, geteventsUrl, authToken);
            if (response == null)
            {
                Console.WriteLine("Failed to find the event");
                return;
            }

            // Get the series master ID from the retrieved occurrence.
            dynamic calEvent = JObject.Parse(response);
            string masterId = calEvent.value[0].SeriesMasterId;
            string getMasterUrl = string.Format(Constants.MasterEvent, masterId);

            // Now update the master event to update all occurrence of this recurring meeting.
            var content = new
            {
                Location = new
                {
                    DisplayName = "Conf Room Baker"
                } 
            };
            string updateResponse = HttpHelper.SendAsync(new HttpMethod("PATCH"), getMasterUrl, authToken, content);

            if (updateResponse == null)
            {
                Console.WriteLine("Failed to update the meeting location");
                return;
            }

            dynamic updatedEvent = JObject.Parse(updateResponse);

            string newLocation = updatedEvent.Location.DisplayName;
            Console.WriteLine($"Successfully updated the meeting location. New location : '{newLocation}' ");
        }
    }
}
