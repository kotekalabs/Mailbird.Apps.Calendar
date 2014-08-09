﻿using System.Collections.ObjectModel;
﻿using DevExpress.Mvvm.POCO;
﻿using DevExpress.Xpf.Scheduler;
﻿using Mailbird.Apps.Calendar.Engine;
 ﻿using System;
using System.Collections.Generic;
﻿using System.Linq;
﻿using DevExpress.Xpf.Core.Native;
﻿using Mailbird.Apps.Calendar.Engine.Interfaces;
﻿using Mailbird.Apps.Calendar.Engine.Metadata;
﻿using Mailbird.Apps.Calendar.Infrastructure;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.Threading;

namespace Mailbird.Apps.Calendar.ViewModels
{
    public class MainWindowViewModel : ViewModelBase
    {
        #region PrivateProps

        private readonly Dictionary<object, Appointment> _appointments = new Dictionary<object, Appointment>();

        private readonly CalendarsCatalog _calendarsCatalog = new CalendarsCatalog();

        private readonly ObservableCollection<TreeData> _treeData = new ObservableCollection<TreeData>();

        private readonly ObservableCollection<Engine.Metadata.Calendar> _calendarCollection = new ObservableCollection<Engine.Metadata.Calendar>();

        private ObservableCollection<Appointment> _appointmentCollection = new ObservableCollection<Appointment>();
        #endregion PrivateProps

        #region PublicProps

        public FlyoutViewModel FlyoutViewModel { get; private set; }

        public ObservableCollection<Appointment> AppointmentCollection
        {
            get
            {
                return _appointmentCollection;
            }

            set
            {
                _appointmentCollection = value;
                RaisePropertyChanged(() => AppointmentCollection);
            }
        }

        public ObservableCollection<TreeData> TreeData
        {
            get { return _treeData; }
        }

        private DispatcherTimer _timer;

        private CancellationTokenSource _cts;

        #endregion PublicProps

        public MainWindowViewModel()
        {

            foreach (var provider in _calendarsCatalog.GetProviders)
            {
                AddElementToTree(provider);
                foreach (var calendar in provider.GetCalendars())
                {
                    AddElementToTree(calendar);
                    _calendarCollection.Add(calendar);
                }
            }

            _cts = new CancellationTokenSource();
            SyncTaskAsync();

            FlyoutViewModel = new FlyoutViewModel
            {
                AddAppointmentAction = AddAppointment,
                UpdateAppointmentAction = UpdateAppointment,
                RemoveAppointmentAction = RemoveAppointment
            };
            //Make asynchronous
            var calendars = Task.Factory.StartNew(() => _calendarsCatalog.GetCalendars());
            calendars.Result.ToArray();

            
            _timer = new DispatcherTimer();
            _timer.Interval = TimeSpan.FromMinutes(1);
            _timer.Tick += _timer_Tick;
            _timer.Start();
        }

        void _timer_Tick(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("Fired Timer");
            SyncTaskAsync();
        }

        private void SyncTaskAsync()
        {
            //using CTS for making the thread can be cancelled
            //using FromCurrentSynchronizationContext for make the thread can be update the UI without Dispatcher.
           var task = Task.Factory.StartNew(()=>
            {
                
                var appointmentList = _calendarsCatalog.GetCalendarAppointments().ToList();
                

                foreach (var a in appointmentList)
                {
                    // Make sure we don't get any duplicates
                    if (!_appointments.ContainsKey(a.Id))
                        _appointments.Add(a.Id, a);
                }

                return appointmentList;
            });
           task.ContinueWith(x =>
           {
               if (x.IsFaulted || x.IsCanceled)
                   return;
               AppointmentCollection = new ObservableCollection<Appointment>(x.Result);
             
           },_cts.Token, TaskContinuationOptions.LongRunning,TaskScheduler.FromCurrentSynchronizationContext());

        }
        public void AddAppointment(Appointment appointment)
        {
            AppointmentCollection.Add(appointment);
            if (appointment.Id == null || _appointments.ContainsKey(appointment.Id))
                appointment.Id = Guid.NewGuid();
            if (appointment.Calendar == null)
                appointment.Calendar = _calendarsCatalog.DefaultCalendar != null ? _calendarsCatalog.DefaultCalendar : _calendarCollection.FirstOrDefault(x => x.AccessRights == Engine.Metadata.Calendar.Access.Write);
            _appointments.Add(appointment.Id, appointment);
            _calendarsCatalog.InsertAppointment(appointment);
        }

        public void UpdateAppointment(object appointmentId, Appointment appointment)
        {
            var appointmentToUpdate = _appointments[appointmentId];
            AppointmentCollection.Remove(appointmentToUpdate);
            AppointmentCollection.Add(appointment);
            _appointments[appointmentId] = appointment;
            if (appointment.Calendar == null)
                appointment.Calendar = _calendarsCatalog.DefaultCalendar != null ? _calendarsCatalog.DefaultCalendar : _calendarCollection.FirstOrDefault(x => x.AccessRights == Engine.Metadata.Calendar.Access.Write);
            _calendarsCatalog.UpdateAppointment(appointment);
        }

        public void RemoveAppointment(object appintmentId)
        {
            //double check prevent Calendar to be null  -> Will Force Close when Drag & Move Events
            var appointment = _appointments[appintmentId]; // make like this for more debuggable code
            if (appointment.Calendar == null)
                appointment.Calendar = _calendarsCatalog.DefaultCalendar != null ? _calendarsCatalog.DefaultCalendar : _calendarCollection.FirstOrDefault(x => x.AccessRights == Engine.Metadata.Calendar.Access.Write);
            AppointmentCollection.Remove(_appointments[appintmentId]);
            _calendarsCatalog.RemoveAppointment(_appointments[appintmentId]);
            _appointments.Remove(appintmentId);
        }

        public void AppointmentOnViewChanged(Appointment appointment)
        {
            var app = AppointmentCollection.First(f => f.Id.ToString() == appointment.Id.ToString());
            appointment.ReminderInfo = app.ReminderInfo;
            appointment.Calendar = app.Calendar;
            UpdateAppointment(appointment.Id, appointment);
        }

        private void AddElementToTree(object element)
        {
            Dispatcher.CurrentDispatcher.BeginInvoke(new Action(() =>
            {
                if (element is ICalendarProvider)
                {
                    TreeData.Add(new TreeData
                    {
                        DataType = TreeDataType.Provider,
                        Data = element,
                        Name = (element as ICalendarProvider).Name,
                        ParentID = "0"
                    });
                }
                if (element is Mailbird.Apps.Calendar.Engine.Metadata.Calendar)
                {
                    TreeData.Add(new TreeData
                    {
                        DataType = TreeDataType.Calendar,
                        Data = element,
                        Name = (element as Mailbird.Apps.Calendar.Engine.Metadata.Calendar).Name,
                        ParentID = (element as Mailbird.Apps.Calendar.Engine.Metadata.Calendar).Provider
                    });
                }
            }));
        }

        public void OpenInnerFlyout(SchedulerControl scheduler)
        {
            FlyoutViewModel.SelectedStartDateTime = scheduler.SelectedInterval.Start;
            FlyoutViewModel.SelectedEndDateTime = scheduler.SelectedInterval.End;
            FlyoutViewModel.IsOpen = true;
        }

        public void CloseInnerFlyout()
        {
            if (FlyoutViewModel.IsEdited)
            {
                FlyoutViewModel.OkCommandeExecute();
            }
            else
            {
                FlyoutViewModel.IsOpen = false;
            }
        }
    }
}
