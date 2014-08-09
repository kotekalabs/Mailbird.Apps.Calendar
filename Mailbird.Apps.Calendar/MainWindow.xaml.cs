using DevExpress.Xpf.Scheduler;
using System.Windows;
using DevExpress.XtraScheduler;
using Mailbird.Apps.Calendar.ViewModels;
using Appointment = Mailbird.Apps.Calendar.Engine.Metadata.Appointment;

namespace Mailbird.Apps.Calendar
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            Scheduler.DayView.MoreButtonDownStyle = (Style)FindResource("MoreButtonDownStyle");
            Scheduler.DayView.MoreButtonUpStyle = (Style)FindResource("MoreButtonUpStyle");
        }

        public MainWindowViewModel ViewModel
        {
            get { return (MainWindowViewModel) DataContext; }
        }
        
        private void Scheduler_OnEditAppointmentFormShowing(object sender, EditAppointmentFormEventArgs e)
        {
            e.Cancel = true;
            ViewModel.OpenInnerFlyout(Scheduler);
            FlyoutControl.Focus();
        }

        private void UIElement_GotFocus(object sender, RoutedEventArgs e)
        {
            if (!FlyoutControl.IsOpen) return;
            ViewModel.CloseInnerFlyout();
        }

        private void SchedulerStorage_AppointmentsChanged(object sender, DevExpress.XtraScheduler.PersistentObjectsEventArgs e)
        {
            foreach (DevExpress.XtraScheduler.Appointment obj in e.Objects)
            {
                var app = new Appointment
                {
                    Id = obj.Id,
                    AllDayEvent = obj.AllDay,
                    Calendar = (Engine.Metadata.Calendar)obj.CustomFields["cfCalendar"],
                    Description = obj.Description,
                    EndTime = obj.End,
                    LabelId = obj.LabelId,
                    Location = obj.Location,
                    ResourceId = obj.ResourceId,
                    StartTime = obj.Start,
                    StatusId = obj.StatusId,
                    Subject = obj.Subject
                };
                ViewModel.AppointmentOnViewChanged(app);
            }
        }

        private void SchedulerStorage_OnAppointmentsDeleted(object sender, PersistentObjectsEventArgs e)
        {
            foreach (DevExpress.XtraScheduler.Appointment obj in e.Objects)
            {
                ViewModel.RemoveAppointment(obj.Id);
            }
        }
    }
}
