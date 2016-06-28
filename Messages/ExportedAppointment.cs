using System;
using System.Collections.Generic;

namespace Messages
{
    public class Body
    {
        public int bodyType { get; set; }
    }

    public class ConversationId
    {
        public string uniqueId { get; set; }
    }

    public class ItemId
    {
        public string changeKey { get; set; }
        public string uniqueId { get; set; }
    }

    public class FirstOccurrence
    {
        public string end { get; set; }
        public ItemId itemId { get; set; }
        public string originalStart { get; set; }
        public string start { get; set; }
    }

    public class Id
    {
        public string changeKey { get; set; }
        public string uniqueId { get; set; }
    }

    public class Organizer
    {
        public string address { get; set; }
        public int mailboxType { get; set; }
        public string name { get; set; }
        public string routingType { get; set; }
    }

    public class ParentFolderId
    {
        public string changeKey { get; set; }
        public string uniqueId { get; set; }
    }

    public class Recurrence
    {
        public string SourceType { get; set; }
        public int dayOfMonth { get; set; }
        public bool hasEnd { get; set; }
        public int month { get; set; }
        public string startDate { get; set; }
    }

    public class DaylightTransitionEnd
    {
        public string SourceType { get; set; }
        public int Day { get; set; }
        public int DayOfWeek { get; set; }
        public bool IsFixedDateRule { get; set; }
        public int Month { get; set; }
        public string TimeOfDay { get; set; }
        public int Week { get; set; }
    }

    public class DaylightTransitionStart
    {
        public string SourceType { get; set; }
        public int Day { get; set; }
        public int DayOfWeek { get; set; }
        public bool IsFixedDateRule { get; set; }
        public int Month { get; set; }
        public string TimeOfDay { get; set; }
        public int Week { get; set; }
    }

    public class Value
    {
        public string DateEnd { get; set; }
        public string DateStart { get; set; }
        public string DaylightDelta { get; set; }
        public DaylightTransitionEnd DaylightTransitionEnd { get; set; }
        public DaylightTransitionStart DaylightTransitionStart { get; set; }
    }

    public class AdjustmentRules
    {
        public string SourceType  { get; set; }
        public List<Value> Values { get; set; }
    }

    public class StartTimeZone
    {
        public AdjustmentRules AdjustmentRules { get; set; }
        public string BaseUtcOffset { get; set; }
        public string DaylightName { get; set; }
        public string DisplayName { get; set; }
        public string Id { get; set; }
        public string StandardName { get; set; }
        public bool SupportsDaylightSavingTime { get; set; }
    }

    public enum MeetingResponseType
    {
        Unknown = 0,
        Organizer = 1,
        Tentative = 2,
        Accept = 3,
        Decline = 4,
        NoResponseReceived = 5,
    }

    public sealed class Attendee
    {
        public string Name { get; set; }
        public string Address { get; set; }
        public string RoutingType { get; set; }
        public DateTime? LastResponseTime { get; private set; }
        public MeetingResponseType? MailboxType { get; set; }
    }

    public class Appointment
    {
        //public bool allowNewTimeProposal { get; set; }
        //public int allowedResponseActions { get; set; }
        //public int appointmentSequenceNumber { get; set; }
        //public int appointmentState { get; set; }
        //public int appointmentType { get; set; }
        //public List<object> attachments { get; set; }
        //public Body body { get; set; }
        //public List<object> categories { get; set; }
        //public int conferenceType { get; set; }
        //public ConversationId conversationId { get; set; }
        //public string culture { get; set; }
        //public string dateTimeCreated { get; set; }
        //public string dateTimeReceived { get; set; }
        //public string dateTimeSent { get; set; }
        public string duration { get; set; }
        //public int effectiveRights { get; set; }
        public string end { get; set; }
        //public List<object> extendedProperties { get; set; }
        //public FirstOccurrence firstOccurrence { get; set; }
        //public bool hasAttachments { get; set; }
        //public string iCalDateTimeStamp { get; set; }
        public string iCalUid { get; set; }
        //public Id id { get; set; }
        //public int importance { get; set; }
        //public bool isAllDayEvent { get; set; }
        //public bool isAssociated { get; set; }
        //public bool isAttachment { get; set; }
        //public bool isCancelled { get; set; }
        //public bool isDirty { get; set; }
        //public bool isDraft { get; set; }
        //public bool isFromMe { get; set; }
        //public bool isMeeting { get; set; }
        //public bool isNew { get; set; }
        //public bool isRecurring { get; set; }
        //public bool isReminderSet { get; set; }
        //public bool isResend { get; set; }
        //public bool isResponseRequested { get; set; }
        //public bool isSubmitted { get; set; }
        //public bool isUnmodified { get; set; }
        //public string itemClass { get; set; }
        //public string lastModifiedName { get; set; }
        //public string lastModifiedTime { get; set; }
        //public int legacyFreeBusyStatus { get; set; }
        //public bool meetingRequestWasSent { get; set; }
        //public int myResponseType { get; set; }
        //public List<object> optionalAttendees { get; set; }
        public Organizer organizer { get; set; }
        //public ParentFolderId parentFolderId { get; set; }
        //public Recurrence recurrence { get; set; }
        //public string reminderDueBy { get; set; }
        //public int reminderMinutesBeforeStart { get; set; }
        public List<Attendee> requiredAttendees { get; set; }
        //public List<object> resources { get; set; }
        //public int sensitivity { get; set; }
        //public int size { get; set; }
        public string start { get; set; }
        //public StartTimeZone startTimeZone { get; set; }
        public string subject { get; set; }
        //public string timeZone { get; set; }
        //public string webClientEditFormQueryString { get; set; }
        //public string webClientReadFormQueryString { get; set; }
    }

    public class ExportedAppointment
    {
        // public string id { get; set; }
        public string mailbox { get; set; }
        public Appointment appointment { get; set; }
    }
}
