using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Microsoft.Office.Interop.TeamsAuto
{
    public class EventHelper
    {
        public static async Task<List<Event>> GetInComingEventsAsync(string userId, string organizerMail, HttpClient graphHttpClient)
        {
            List<Event> result = new List<Event>();
            HttpResponseMessage response = await graphHttpClient.GetAsync($"{Settings.GraphBaseUri}/users/{userId}/calendar/events");
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.OK)
            {
                throw new FocusException($"List events graph call failed: {responseMsg}");
            }

            GraphDataSet<Event> dataSet = JsonConvert.DeserializeObject<GraphDataSet<Event>>(responseMsg);
            foreach (Event e in dataSet.Value)
            {
                if (e.EventStartTimeOffset > DateTimeOffset.UtcNow && string.Equals(e.OrganizerEmail, organizerMail, StringComparison.OrdinalIgnoreCase))
                {
                    e.UserId = userId;
                    result.Add(e);
                }
            }

            return result;
        }

        public static async Task<byte[]> GetAttachedImgContent(string userId, string eventId, HttpClient graphHttpClient)
        {
            HttpResponseMessage response = await graphHttpClient.GetAsync($"{Settings.GraphBaseUri}/users/{userId}/calendar/events/{eventId}/attachments");
            string responseMsg = await response.Content.ReadAsStringAsync();
            if (response.StatusCode != HttpStatusCode.OK)
            {
                throw new FocusException($"List event attachments graph call failed: {responseMsg}");
            }

            GraphDataSet<EventAttachment> dataSet = JsonConvert.DeserializeObject<GraphDataSet<EventAttachment>>(responseMsg);
            if (dataSet != null && dataSet.Value != null && dataSet.Value.Count > 0)
            {
                foreach (EventAttachment attachment in dataSet.Value)
                {
                    if (attachment.ContentType.IndexOf("image", StringComparison.OrdinalIgnoreCase) > -1)
                    {
                        return Convert.FromBase64String(attachment.ContentBytes);
                    }
                }
            }

            return null;
        }
    }

    public class Event
    {
        [JsonProperty(PropertyName = "start")]
        protected EventTime StartTime;

        [JsonProperty(PropertyName = "organizer")]
        protected ObjectEntity Organizer;

        [JsonProperty(PropertyName = "attendees")]
        protected List<ObjectEntity> Attendees;

        [JsonProperty(PropertyName = "subject")]
        public string DisplayName;

        [JsonProperty(PropertyName = "id")]
        public string Id;

        public string UserId;

        public DateTimeOffset EventStartTimeOffset
        {
            get
            {
                return DateTime.SpecifyKind(StartTime.Time, StartTime.TimeZone.Equals("UTC") ? DateTimeKind.Utc : DateTimeKind.Local);
            }
        }

        public string OrganizerEmail
        {
            get
            {
                return Organizer.Email.MailAddress;
            }
        }

        public List<string> AttendeeEmails
        {
            get
            {
                return Attendees.Select(x => x.Email.MailAddress).ToList();
            }
        }

        public string UniqueId
        {
            get
            {
                return $"{DisplayName}->{EventStartTimeOffset}";
            }
        }
    }

    public class EventTime
    {
        [JsonProperty(PropertyName = "datetime")]
        public DateTime Time;

        [JsonProperty(PropertyName = "timeZone")]
        public string TimeZone;
    }

    public class ObjectEntity
    {
        [JsonProperty(PropertyName = "emailAddress")]
        public EmailEntity Email;
    }

    public class EmailEntity
    {
        [JsonProperty(PropertyName = "address")]
        public string MailAddress;
    }

    public class EventAttachment
    {
        [JsonProperty(PropertyName = "contentType")]
        public string ContentType;

        [JsonProperty(PropertyName = "contentBytes")]
        public string ContentBytes;
    }

}
