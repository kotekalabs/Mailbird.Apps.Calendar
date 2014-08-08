using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Linq;
using System.Windows.Media;
using Mailbird.Apps.Calendar.Engine.Interfaces;
using Mailbird.Apps.Calendar.Engine.Metadata;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;

namespace Mailbird.Apps.Calendar.Engine.CalendarProviders
{
    public class GoogleCalendarProvider : ICalendarProvider
    {
        private CalendarService _calendarService;

        public GoogleCalendarProvider()
        {
            Name = "GoogleCalendarsStorage";
            Authorize();
        }

        public IEnumerable<Metadata.Calendar> Calendars { get; private set; }

        public string Name { get; private set; }

        private void Authorize()
        {
            UserCredential credential;
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    new[] { CalendarService.Scope.Calendar },
                    "MailbirdCalendar",
                    CancellationToken.None).Result;
            }

            // Create the service.
            _calendarService = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "MailbirdCalendar",
            });

            Calendars = GetCalendars();
        }

        public IEnumerable<Metadata.Calendar> GetCalendars()
        {
            try
            {
                var calendarListEntry = _calendarService.CalendarList.List().Execute().Items;
                var list = calendarListEntry.Select(c => new Metadata.Calendar
                {
                    CalendarId = c.Id,
                    Name = c.Summary,
                    Description = c.Description,
                    AccessRights = c.AccessRole == "reader" ? Metadata.Calendar.Access.Read : Metadata.Calendar.Access.Write,
                    CalendarColor = (Color)ColorConverter.ConvertFromString(c.BackgroundColor),
                    Provider = "GoogleCalendarsStorage"
                });
                return list;
            }
            catch (System.Net.Http.HttpRequestException ex)
            {
                return null;

            }
            catch (Google.GoogleApiException ex)
            {
                return null;
            }
        }

        public IEnumerable<Appointment> GetAppointments()
        {
            var appointments = new List<Appointment>();
            foreach (var calendar in Calendars)
            {
                appointments.AddRange(GetAppointments(calendar));
            }
            return appointments;
        }

        public IEnumerable<Appointment> GetAppointments(Metadata.Calendar calendar)
        {
            return GetAppointments(calendar.CalendarId);
        }

        public IEnumerable<Appointment> GetAppointments(string calendarId)
        {
            var calendarEvents = _calendarService.Events.List(calendarId).Execute().Items;

            var list = calendarEvents.Select(a => new Appointment
            {
                Id = a.Id,
                StartTime = (a.Start != null && a.Start.DateTime.HasValue) ? a.Start.DateTime.Value : DateTime.Now,
                EndTime = (a.End != null && a.End.DateTime.HasValue) ? a.End.DateTime.Value : DateTime.Now,
                Subject = a.Summary,
                Description = a.Description,
                Location = a.Location
            });

            return list;
        }

        //adding try catch in below to prevent force close when there is internet connection interruption
        public bool InsertAppointment(Appointment appointment)
        {
            try
            {
                var googleEvent = new Event
                {
                    Start = new EventDateTime
                    {
                        DateTime = appointment.StartTime
                    },
                    End = new EventDateTime { DateTime = appointment.EndTime },
                    Summary = appointment.Subject,
                    Description = appointment.Description,
                    Location = appointment.Location
                };
                _calendarService.Events.Insert(googleEvent, appointment.Calendar.CalendarId).Execute();
                return true;
            }
            catch (System.Net.Http.HttpRequestException ex)
            {
                return false;

            }
            catch (Google.GoogleApiException ex)
            {
                return false;
            }
        }

        public bool UpdateAppointment(Appointment appointment)
        {
            try
            {
                var googleEvent = _calendarService.Events.Get(appointment.Calendar.CalendarId, appointment.Id.ToString()).Execute();
                googleEvent.Start.DateTime = appointment.StartTime;
                googleEvent.End.DateTime = appointment.EndTime;
                googleEvent.Summary = appointment.Subject;
                googleEvent.Description = appointment.Description;
                googleEvent.Location = appointment.Location;
                _calendarService.Events.Update(googleEvent, appointment.Calendar.CalendarId, googleEvent.Id);
                return true;
            }
            catch (System.Net.Http.HttpRequestException ex)
            {
                return false;

            }
            catch (Google.GoogleApiException ex)
            {
                return false;
            }
        }

        public bool RemoveAppointment(Appointment appointment)
        {
            try
            {
                _calendarService.Events.Delete(appointment.Calendar.CalendarId, appointment.Id.ToString());
                return true;
            }
            catch (System.Net.Http.HttpRequestException ex)
            {
                return false;

            }
            catch (Google.GoogleApiException ex)
            {
                return false;
            }
        }
    }
}