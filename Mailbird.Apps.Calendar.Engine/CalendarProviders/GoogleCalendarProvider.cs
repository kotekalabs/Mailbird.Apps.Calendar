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
        private TimeZone _currentTimeZone;
        public GoogleCalendarProvider()
        {
            Name = "GoogleCalendarsStorage";
            _currentTimeZone = TimeZone.CurrentTimeZone;
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
            var calendar = Calendars.FirstOrDefault(x => x.CalendarId.Equals(calendarId));
            var list = calendarEvents.Select(a => new Appointment
            {
                Id = a.Id,
                //change because sometimes get all event from the imported calendar just have date
                StartTime = (a.Start != null && a.Start.DateTime.HasValue) ? a.Start.DateTime.Value : (a.Start.Date != null) ? Convert.ToDateTime(a.Start.Date) : DateTime.Now,
                EndTime = (a.End != null && a.End.DateTime.HasValue) ? a.End.DateTime.Value : (a.End.Date != null) ? Convert.ToDateTime(a.End.Date) : DateTime.Now,
                Subject = a.Summary,
                Description = a.Description,
                Calendar = calendar,
                //fixing binding error
                LabelId = a.ColorId != null ? int.Parse(a.ColorId) : 0,
                AllDayEvent = (a.Start.DateTime.HasValue && a.End.DateTime.HasValue && (Math.Abs(a.End.DateTime.Value.Subtract(a.Start.DateTime.Value).TotalDays) == 0)) ? true : false,
                Location = a.Location
            });

            return list;
        }

        //adding try catch in below to prevent force close when there is internet connection interruption
        public bool InsertAppointment(Appointment appointment)
        {
            try
            {
                Event googleEvent = new Event();
                //changed to be like this because all-day-event in google calendar just have date. if not even we had check the all day event
                // it will be still time limit event.
                if (appointment.AllDayEvent)
                {
                    googleEvent.Start = new EventDateTime
                      {
                          Date = new DateTimeOffset(appointment.StartTime, _currentTimeZone.GetUtcOffset(appointment.StartTime)).DateTime.ToString("yyyy-MM-dd")
                      };
                    googleEvent.End = new EventDateTime
                    {
                        Date = new DateTimeOffset(appointment.EndTime, _currentTimeZone.GetUtcOffset(appointment.EndTime)).DateTime.ToString("yyyy-MM-dd")
                    };

                }
                else
                {
                    googleEvent.Start = new EventDateTime
                      {
                          DateTime = new DateTimeOffset(appointment.StartTime, _currentTimeZone.GetUtcOffset(appointment.StartTime)).DateTime
                      };
                    googleEvent.End = new EventDateTime
                    {
                        DateTime = new DateTimeOffset(appointment.EndTime, _currentTimeZone.GetUtcOffset(appointment.EndTime)).DateTime
                    };

                }
                googleEvent.Summary = appointment.Subject;
                googleEvent.Description = appointment.Description;
                googleEvent.ColorId = !String.IsNullOrEmpty(appointment.LabelId.ToString()) ? appointment.LabelId.ToString() : "0";
                googleEvent.Location = appointment.Location;
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
                if (appointment.AllDayEvent)
                {
                    googleEvent.Start.Date = appointment.StartTime.ToString("yyyy-MM-dd");
                    googleEvent.End.Date = appointment.EndTime.ToString("yyyy-MM-dd");
                }

                googleEvent.Start.DateTime = new DateTimeOffset(appointment.StartTime, _currentTimeZone.GetUtcOffset(appointment.StartTime)).DateTime;
                googleEvent.End.DateTime = new DateTimeOffset(appointment.EndTime, _currentTimeZone.GetUtcOffset(appointment.EndTime)).DateTime;
                googleEvent.ColorId = !String.IsNullOrEmpty(appointment.LabelId.ToString()) ? appointment.LabelId.ToString() : "0";
                googleEvent.Summary = appointment.Subject;
                googleEvent.Description = appointment.Description;
                googleEvent.Location = appointment.Location;
                _calendarService.Events.Update(googleEvent, appointment.Calendar.CalendarId, googleEvent.Id).Execute();
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
                _calendarService.Events.Delete(appointment.Calendar.CalendarId, appointment.Id.ToString()).Execute();
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