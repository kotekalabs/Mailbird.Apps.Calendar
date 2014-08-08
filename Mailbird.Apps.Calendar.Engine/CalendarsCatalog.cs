using System;
using System.Collections.Generic;
using System.Linq;
using Mailbird.Apps.Calendar.Engine.Interfaces;
using Mailbird.Apps.Calendar.Engine.Metadata;

namespace Mailbird.Apps.Calendar.Engine
{
    /// <summary>
    /// CalendarCatalog create all instamces of ICalendarProvider and calendars workout
    /// go throught this clas
    /// </summary>
    public class CalendarsCatalog
    {
        private readonly Dictionary<string, ICalendarProvider> _calendarProviders;

        public Metadata.Calendar DefaultCalendar { get; set; }

        public CalendarsCatalog()
        {
            _calendarProviders = AppDomainAssemblyTypeScanner.TypesOf<ICalendarProvider>().Select(x => (ICalendarProvider)Activator.CreateInstance(x)).ToDictionary(x => x.Name, x => x);
        }

        public IEnumerable<ICalendarProvider> GetProviders
        {
            get { return _calendarProviders.Values; }
        }

        public IEnumerable<Metadata.Calendar> GetCalendars()
        {
            foreach (var calendarProvider in _calendarProviders.Values)
            {
                foreach (var calendar in calendarProvider.GetCalendars())
                {
                    yield return calendar;
                }
            }
        }

        public IEnumerable<Appointment> GetCalendarAppointments()
        {
            return _calendarProviders.Values.SelectMany(calendarProvider => calendarProvider.GetAppointments());
        }

        public bool InsertAppointment(Appointment appointment)
        {
            var provider = appointment.Calendar != null ? appointment.Calendar.Provider : GetProviders.FirstOrDefault().Name.ToString();
            return _calendarProviders[provider].InsertAppointment(appointment);
        }

        public bool UpdateAppointment(Appointment appointment)
        {
            
            var provider = appointment.Calendar != null ? appointment.Calendar.Provider : GetProviders.FirstOrDefault().Name.ToString();
            return _calendarProviders[provider].UpdateAppointment(appointment);
        }

        public bool RemoveAppointment(Appointment appointment)
        {
            var provider = appointment.Calendar != null ? appointment.Calendar.Provider : GetProviders.FirstOrDefault().Name.ToString();
            return _calendarProviders[provider].RemoveAppointment(appointment);
        }
    }
}