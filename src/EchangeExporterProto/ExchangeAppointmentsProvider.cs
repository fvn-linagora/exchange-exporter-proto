using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using Bender.Collections;
using Newtonsoft.Json;

using Microsoft.Exchange.WebServices.Data;
using AppointmentType = Microsoft.Exchange.WebServices.Data.AppointmentType;
using Attendee = Microsoft.Exchange.WebServices.Data.Attendee;
using EWSAppointment = Microsoft.Exchange.WebServices.Data.Appointment;
using EWSAppointmentType = Microsoft.Exchange.WebServices.Data.AppointmentType;
using Appointment = Messages.Appointment;

namespace EchangeExporterProto
{
    public interface IAppointmentsProvider
    {
        IEnumerable<Appointment> FindByMailbox(string primaryEmailAddress);
    }

    public class ExchangeAppointmentsProvider : IAppointmentsProvider
    {
        private static readonly ICollection<EWSAppointmentType> singleAndRecurringMasterAppointmentTypes =
            new List<EWSAppointmentType> {EWSAppointmentType.RecurringMaster, EWSAppointmentType.Single};

        private readonly Func<string, ExchangeService> impersonateServiceProvider;
        private readonly TractableJsonSerializer serializer;

        public ExchangeAppointmentsProvider(TractableJsonSerializer serializer, Func<string, ExchangeService> impersonateServiceProvider)
        {
            this.serializer = serializer;
            this.impersonateServiceProvider = impersonateServiceProvider;
        }

        public IEnumerable<Appointment> FindByMailbox(string primaryEmailAddress)
        {
            return FindAllMeetingsForMailbox(primaryEmailAddress);
        }

        private IEnumerable<Appointment> FindAllMeetingsForMailbox(string primaryEmailAddress)
        {
            PropertySet includeMostProps = BuildAppointmentPropertySet();
            ExchangeService service = impersonateServiceProvider(primaryEmailAddress);

            var findAllAppointments = new Func<CalendarFolder, IEnumerable<EWSAppointment>>
                (calendar => FindAllAppointments(service, calendar.Id));
            var repairMaster = new Func<ExchangeService, EWSAppointment, EWSAppointment>
                (RepairReccurenceMasterAttendees).Partial(service);
            var fetchAppointmentDetails = new Func<EWSAppointment, EWSAppointment>(
                app => EWSAppointment.Bind(service, app.Id, includeMostProps));

            var mailboxAppointments = GetAllCalendars(service)
                .SelectMany(findAllAppointments);

            var singleAndReccurringMasterAppointments = mailboxAppointments
                .Where(app => singleAndRecurringMasterAppointmentTypes.Contains(app.AppointmentType));

            var singleAndReccurringMasterAppointmentsWithContext = singleAndReccurringMasterAppointments
                .Select(fetchAppointmentDetails)
                .Select(repairMaster);

            return singleAndReccurringMasterAppointmentsWithContext
                .Select(Convert);
        }


        private Appointment Convert(EWSAppointment appointment)
        {
            // TODO: REWRITE me as a(n external) mapper class
            // TODO: there should be a genuine mapping eventually here, as models may diverge !
            var serializedPayload = serializer.ToJson(appointment);
            Appointment result = JsonConvert.DeserializeObject<Appointment>(serializedPayload);
            // restore (lost) MimeContent
            result.MimeContent = new Messages.MimeContent {
                CharacterSet = appointment.MimeContent.CharacterSet,
                Content = appointment.MimeContent.Content,
            };
            return result;
        }

        private EWSAppointment RepairReccurenceMasterAttendees(ExchangeService service, EWSAppointment appointment)
        {
            if (appointment.AppointmentType != AppointmentType.RecurringMaster)
                return appointment;
            if (appointment.ModifiedOccurrences == null)
                return appointment;

            var appointmentsExceptionsWithAttendees = appointment.ModifiedOccurrences.Select(exception =>
                EWSAppointment.Bind(service, exception.ItemId, new PropertySet(
                    BasePropertySet.IdOnly,
                    AppointmentSchema.RequiredAttendees,
                    AppointmentSchema.OptionalAttendees,
                    AppointmentSchema.Resources)))
                .ToList();

            ReplaceAppointmentsAttendees(appointment, appointmentsExceptionsWithAttendees, app => app.RequiredAttendees);
            ReplaceAppointmentsAttendees(appointment, appointmentsExceptionsWithAttendees, app => app.OptionalAttendees);
            ReplaceAppointmentsAttendees(appointment, appointmentsExceptionsWithAttendees, app => app.Resources);
//            ReplaceAttendees(appointmentsExceptionsWithAttendees.SelectMany(app => app.RequiredAttendees),
//                appointment.RequiredAttendees);
//            ReplaceAttendees(appointmentsExceptionsWithAttendees.SelectMany(app => app.OptionalAttendees),
//                appointment.OptionalAttendees);
//            ReplaceAttendees(appointmentsExceptionsWithAttendees.SelectMany(app => app.Resources),
//                appointment.Resources);

            return appointment;
        }

        private void ReplaceAppointmentsAttendees(EWSAppointment master, IEnumerable<EWSAppointment> exceptions,
            Expression<Func<EWSAppointment, AttendeeCollection>> attendeePropMapper)
        {
            var propExpr = (MemberExpression) attendeePropMapper.Body;
            if (!(propExpr.Member is PropertyInfo)) return;

            var attendeesGetter = attendeePropMapper.Compile();

            var masterAttendees = attendeesGetter(master);
            masterAttendees.Clear();
            var exceptionsAttendees = exceptions.SelectMany(attendeesGetter);

            exceptionsAttendees
                .ToLookup(GetAttendeesKey).Select(e => e.First())
                .ForEach(masterAttendees.Add);
        }

        private static void ReplaceAttendees(IEnumerable<Attendee> exceptionsAttendees, AttendeeCollection masterAttendees)
        {
            masterAttendees.Clear();
            exceptionsAttendees
                .ToLookup(GetAttendeesKey).Select(e => e.First())
                .ForEach(masterAttendees.Add);
        }

        private static string GetAttendeesKey(Attendee att) => (att?.ToString() ?? "").Trim().ToLower();

        private static IEnumerable<EWSAppointment> FindAllAppointments(ExchangeService service, FolderId calendarId)
        {
            var appIdsView = new ItemView(int.MaxValue)
            {
                PropertySet = new PropertySet(BasePropertySet.IdOnly, AppointmentSchema.AppointmentType)
            };

            var result = PagedItemsSearch.PageSearchItems<Microsoft.Exchange.WebServices.Data.Appointment>(service, calendarId, 500,
                appIdsView.PropertySet, AppointmentSchema.DateTimeCreated);

            return result;
        }

        public static IEnumerable<CalendarFolder> GetAllCalendars(ExchangeService service)
        {
            int folderViewSize = int.MaxValue; // Kids, dont do this at home !
            var view = new FolderView(folderViewSize)
            {
                PropertySet =
                    new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName, FolderSchema.FolderClass),
                Traversal = FolderTraversal.Deep
            };
            FindFoldersResults findFolderResults = service.FindFolders(WellKnownFolderName.MsgFolderRoot, view);

            return findFolderResults.OfType<CalendarFolder>();
        }

        private static PropertySet BuildAppointmentPropertySet()
        {
            return new PropertySet(
                BasePropertySet.FirstClassProperties,
                AppointmentSchema.AppointmentType,
                ItemSchema.Body,
                AppointmentSchema.RequiredAttendees,
                AppointmentSchema.OptionalAttendees,
                AppointmentSchema.Resources,
                ItemSchema.Categories,
                ItemSchema.Culture,
                ItemSchema.DateTimeCreated,
                ItemSchema.DateTimeReceived,
                ItemSchema.DateTimeSent,
                ItemSchema.DisplayTo,
                ItemSchema.DisplayCc,
                AppointmentSchema.Duration,
                AppointmentSchema.End,
                AppointmentSchema.Start,
                AppointmentSchema.StartTimeZone,
                ItemSchema.Subject,
                AppointmentSchema.TimeZone,
                ItemSchema.MimeContent,
                AppointmentSchema.ModifiedOccurrences,
                AppointmentSchema.DeletedOccurrences,
                AppointmentSchema.AppointmentType,
                AppointmentSchema.IsResponseRequested,
                AppointmentSchema.When
            );
        }
    }
}