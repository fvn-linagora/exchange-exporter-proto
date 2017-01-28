using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

using Microsoft.Exchange.WebServices.Data;
using AppointmentType = Microsoft.Exchange.WebServices.Data.AppointmentType;
using Attendee = Microsoft.Exchange.WebServices.Data.Attendee;
using EWSAppointment = Microsoft.Exchange.WebServices.Data.Appointment;
using EWSAppointmentType = Microsoft.Exchange.WebServices.Data.AppointmentType;
using EWSMailboxType = Microsoft.Exchange.WebServices.Data.MailboxType;
using Messages;
using Appointment = Messages.Appointment;

namespace EchangeExporterProto
{
    public interface IAppointmentsProvider
    {
        IEnumerable<Appointment> FindByMailbox(string primaryEmailAddress);
        IEnumerable<Appointment> FindByMailbox(string primaryEmailAddress, DateTime? sinceDate);
    }

    public class ExchangeAppointmentsProvider : IAppointmentsProvider
    {
        private static readonly ICollection<EWSAppointmentType> singleAndRecurringMasterAppointmentTypes =
            new List<EWSAppointmentType> {EWSAppointmentType.RecurringMaster, EWSAppointmentType.Single};

        private static readonly ISet<EWSMailboxType> expandableMailboxTypes = new HashSet<EWSMailboxType> {
            EWSMailboxType.PublicGroup, EWSMailboxType.PublicFolder, EWSMailboxType.ContactGroup
        };

        private readonly Func<string, ExchangeService> impersonateServiceProvider;
        private readonly TractableJsonSerializer serializer;

        public ExchangeAppointmentsProvider(TractableJsonSerializer serializer, Func<string, ExchangeService> impersonateServiceProvider)
        {
            this.serializer = serializer;
            this.impersonateServiceProvider = impersonateServiceProvider;
        }
        public IEnumerable<Appointment> FindByMailbox(string primaryEmailAddress)
        {
            return FindByMailbox(primaryEmailAddress, null);
        }

        public IEnumerable<Appointment> FindByMailbox(string primaryEmailAddress, DateTime? sinceDate)
        {
            return FindAllMeetingsForMailbox(primaryEmailAddress, sinceDate);
        }

        private IEnumerable<Appointment> FindAllMeetingsForMailbox(string primaryEmailAddress, DateTime? sinceDate)
        {
            PropertySet includeMostProps = BuildAppointmentPropertySet();
            ExchangeService service = impersonateServiceProvider(primaryEmailAddress);

            var findAllAppointments = new Func<CalendarFolder, IEnumerable<EWSAppointment>>
                (calendar => FindAllAppointments(service, calendar.Id, sinceDate));
            var repairMaster = new Func<ExchangeService, EWSAppointment, AppointmentWithParticipation>
                (RepairReccurenceMasterAttendees).Partial(service);
            var fetchAppointmentDetails = new Func<EWSAppointment, EWSAppointment>(
                app => FetchAppointmentDetails(app, includeMostProps, service));
            var expandDistributionLists = new Func<ExchangeService, EWSAppointment, EWSAppointment>(ExpandGroups).Partial(service);

            var mailboxAppointments = GetAllCalendars(service)
                .SelectMany(findAllAppointments);

            var singleAndReccurringMasterAppointments = mailboxAppointments
                .Where(app => singleAndRecurringMasterAppointmentTypes.Contains(app.AppointmentType));

            var singleAndReccurringMasterAppointmentsWithContext = singleAndReccurringMasterAppointments
                .Select(fetchAppointmentDetails)
                .Select(expandDistributionLists)
                .Select(repairMaster);

            return singleAndReccurringMasterAppointmentsWithContext
                .Where(awp => awp != default(AppointmentWithParticipation))
                .Select(ToAppointmentDTO);
        }

        private static EWSAppointment FetchAppointmentDetails(EWSAppointment app, PropertySet includeMostProps, ExchangeService service)
        {
            try
            {
                return EWSAppointment.Bind(service, app.Id, includeMostProps);
            }
            catch (ServiceResponseException ex) when (ex.ErrorCode == ServiceError.ErrorRecurrenceHasNoOccurrence)
            {
                return default(EWSAppointment);
            }
        }

        private Appointment ToAppointmentDTO(AppointmentWithParticipation appointment)
        {
            var appointmentWithAttendeeStatus = Convert(appointment.Appointment);

            if (appointment.Appointment.AppointmentType != AppointmentType.RecurringMaster)
                return appointmentWithAttendeeStatus;

            var mapOfExceptionAttendees = appointment.ExceptionsAttendees
                .ToDictionary(k => k.Key.UniqueId, v => v.Value);

            foreach (var exc in appointmentWithAttendeeStatus.ModifiedOccurrences ?? Enumerable.Empty<ModifiedOccurrence>())
            {
                var exceptionAttendeesParticipation = mapOfExceptionAttendees[exc.ItemId.UniqueId];

                exc.Attendees = new Messages.ExceptionAttendees
                {
                    Optional = new HashSet<OptionalAttendee>(exceptionAttendeesParticipation.Optional
                        .Select(MapToOptionalAttendees)),
                    Required = new HashSet<RequiredAttendee>(exceptionAttendeesParticipation.Required
                        .Select(MapToRequiredAttendees)),
                    Resources = new HashSet<Resource>(exceptionAttendeesParticipation.Resources
                        .Select(MapToResources)),
                };

            }

            return appointmentWithAttendeeStatus;
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

        private static AppointmentWithParticipation RepairReccurenceMasterAttendees(ExchangeService service, EWSAppointment appointment)
        {
            if (appointment == default(EWSAppointment))
                return default(AppointmentWithParticipation); // avoid breaking iterator

            if (appointment.AppointmentType != AppointmentType.RecurringMaster)
                return new AppointmentWithParticipation(appointment);
            if (appointment.ModifiedOccurrences == null)
                return new AppointmentWithParticipation(appointment);

            var appointmentsExceptionsWithAttendees = appointment.ModifiedOccurrences.Select(exception =>
                EWSAppointment.Bind(service, exception.ItemId, new PropertySet(
                    BasePropertySet.IdOnly,
                    AppointmentSchema.RequiredAttendees,
                    AppointmentSchema.OptionalAttendees,
                    AppointmentSchema.Resources)))
                .ToList();

            return new AppointmentWithParticipation(appointment) {
                ExceptionsAttendees = appointmentsExceptionsWithAttendees
                    .Select(e => new {
                        ItemId = e.Id,
                        ExceptionsAttendees = new ExceptionAttendees {
                            Optional = e.OptionalAttendees,
                            Required = e.RequiredAttendees,
                            Resources = e.Resources
                        }
                    })
                    .ToDictionary(k => k.ItemId, v => v.ExceptionsAttendees),
            };
        }

        private static EWSAppointment ExpandGroups(ExchangeService service, EWSAppointment appointment)
        {
            if (appointment == default(EWSAppointment))
                return default(EWSAppointment); // avoid breaking iterator

            var expandDL = new Func<ExchangeService, string, IEnumerable<Attendee>> (ExpandDistributionLists).Partial(service);

            // Expand DL in attendees collections
            appointment.RequiredAttendees
                .Where(HasExpandableMailbox)
                .Select(x => x.Address)
                .SelectMany(expandDL)
                .ToList()
                .ForEach(appointment.RequiredAttendees.Add);
            appointment.OptionalAttendees
                .Where(HasExpandableMailbox)
                .Select(x => x.Address)
                .SelectMany(expandDL)
                .ToList()
                .ForEach(appointment.OptionalAttendees.Add);
            appointment.Resources
                .Where(HasExpandableMailbox)
                .Select(x => x.Address)
                .SelectMany(expandDL)
                .ToList()
                .ForEach(appointment.Resources.Add);

            // Return updated apppointment
            return appointment;
        }

        private static IEnumerable<Attendee> ExpandDistributionLists(ExchangeService service, string mailbox)
        {
            if (mailbox != null)
            {
                foreach (EmailAddress address in service.ExpandGroup(mailbox).Members)
                    if (address.MailboxType == EWSMailboxType.PublicGroup && !String.IsNullOrWhiteSpace(address.Address))
                        ExpandDistributionLists(service, address.Address);
                    else if (address != null && !String.IsNullOrWhiteSpace(address.Address))
                        yield return new Attendee(address);
            }
        }

        private static bool HasExpandableMailbox(Attendee attendee) => attendee.MailboxType.HasValue && expandableMailboxTypes.Contains(attendee.MailboxType.Value);


        private static IEnumerable<EWSAppointment> FindAllAppointments(ExchangeService service, FolderId calendarId, DateTime? sinceDate)
        {
            var appIdsView = new ItemView(int.MaxValue)
            {
                PropertySet = new PropertySet(BasePropertySet.IdOnly, AppointmentSchema.AppointmentType, ItemSchema.LastModifiedTime)
            };

            var filter = sinceDate.HasValue ? BuildModifiedSinceFilter(sinceDate.Value) : null;

            var result = PagedItemsSearch.PageSearchItems<Microsoft.Exchange.WebServices.Data.Appointment>(service, calendarId, 500,
                appIdsView.PropertySet, AppointmentSchema.DateTimeCreated, filter);

            return result;
        }
        private static SearchFilter.SearchFilterCollection BuildModifiedSinceFilter(DateTime sinceDate)
        {
            var searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And);
            searchFilter.Add(new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.LastModifiedTime, sinceDate));
            searchFilter.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Appointment"));
            return searchFilter;
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

        private static RequiredAttendee MapToRequiredAttendees(Attendee att)
        {
            return new RequiredAttendee
            {
                Address = att.Address,
                Name = att.Name,
                MailboxType = (Messages.MailboxType)(int?)att.MailboxType,
                ResponseType = (Messages.MeetingResponseType)(int?)att.ResponseType,
                RoutingType = att.RoutingType,
            };
        }

        private static OptionalAttendee MapToOptionalAttendees(Attendee att)
        {
            return new OptionalAttendee
            {
                Address = att.Address,
                Name = att.Name,
                MailboxType = (Messages.MailboxType)(int?)att.MailboxType,
                ResponseType = (Messages.MeetingResponseType)(int?)att.ResponseType,
                RoutingType = att.RoutingType,
            };
        }

        private static Resource MapToResources(Attendee att)
        {
            return new Resource
            {
                Address = att.Address,
                Name = att.Name,
                MailboxType = (Messages.MailboxType)(int?)att.MailboxType,
                ResponseType = (Messages.MeetingResponseType)(int?)att.ResponseType,
                RoutingType = att.RoutingType,
            };
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