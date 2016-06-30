﻿using System;
using System.Collections.Generic;

namespace Messages
{
    public class Attendee
    {
        public string Name { get; set; }
        public string Address { get; set; }
        public string RoutingType { get; set; }
        public int? MailboxType { get; set; }
    }

    public class Organizer : Attendee {}

    public class RequiredAttendee : Attendee
    {
        public int? ResponseType { get; set; }
    }

    public class OptionalAttendee : Attendee
    {
        public int? ResponseType { get; set; }
    }

    public class Recurrence
    {
        public List<int> DaysOfTheWeek { get; set; }
        public int FirstDayOfWeek { get; set; }
        public int Interval { get; set; }
        public string StartDate { get; set; }
        public bool HasEnd { get; set; }
        public string EndDate { get; set; }
    }

    public struct ItemId
    {
        public string uniqueId { get; set; }
        public string changeKey { get; set; }
    }

    public class ModifiedOccurrence
    {
        public ItemId ItemId { get; set; }
        public List<Attendee> Attendees { get; set; }
        public string Start { get; set; }
        public string End { get; set; }
        public string OriginalStart { get; set; }
    }

    public class DeletedOccurrence
    {
        public string OriginalStart { get; set; }
    }

    public class DaylightTransitionStart
    {
        public string TimeOfDay { get; set; }
        public int Month { get; set; }
        public int Week { get; set; }
        public int Day { get; set; }
        public int DayOfWeek { get; set; }
        public bool IsFixedDateRule { get; set; }
    }

    public class DaylightTransitionEnd
    {
        public string TimeOfDay { get; set; }
        public int Month { get; set; }
        public int Week { get; set; }
        public int Day { get; set; }
        public int DayOfWeek { get; set; }
        public bool IsFixedDateRule { get; set; }
    }

    //public class AdjustmentRulesValue
    //{
    //    public string DateStart { get; set; }
    //    public string DateEnd { get; set; }
    //    public string DaylightDelta { get; set; }
    //    public DaylightTransitionStart DaylightTransitionStart { get; set; }
    //    public DaylightTransitionEnd DaylightTransitionEnd { get; set; }
    //}

    //public class AdjustmentRules
    //{
    //    public List<AdjustmentRulesValue> Values { get; set; }
    //}

    public class StartTimeZone
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string StandardName { get; set; }
        public string DaylightName { get; set; }
        public string BaseUtcOffset { get; set; }
        // public AdjustmentRules AdjustmentRules { get; set; }
        public bool SupportsDaylightSavingTime { get; set; }
    }

    public class Id
    {
        public string uniqueId { get; set; }
        public string changeKey { get; set; }
    }

    public class Body
    {
        public int bodyType { get; set; }
        public string text { get; set; }
    }

    public class ConversationId
    {
        public string UniqueId { get; set; }
    }

    public enum AppointmentType
    {
        Single,
        Occurrence,
        Exception,
        RecurringMaster
    }

    public class Appointment
    {
        public string Start { get; set; }
        public string End { get; set; }
        public bool IsAllDayEvent { get; set; }
        public int LegacyFreeBusyStatus { get; set; }
        public string Location { get; set; }
        public bool IsMeeting { get; set; }
        public bool MeetingRequestWasSent { get; set; }
        public bool IsResponseRequested { get; set; }
        public AppointmentType AppointmentType { get; set; }
        public int MyResponseType { get; set; }
        public Organizer Organizer { get; set; }
        public List<RequiredAttendee> RequiredAttendees { get; set; }
        public List<OptionalAttendee > OptionalAttendees { get; set; }
        // public List<object> Resources { get; set; }
        public string Duration { get; set; }
        public string TimeZone { get; set; }
        public int AppointmentSequenceNumber { get; set; }
        public int AppointmentState { get; set; }
        public Recurrence Recurrence { get; set; }
        //public FirstOccurrence FirstOccurrence { get; set; }
        //public LastOccurrence LastOccurrence { get; set; }
        public List<ModifiedOccurrence> ModifiedOccurrences { get; set; }
        public List<DeletedOccurrence> DeletedOccurrences { get; set; }
        public StartTimeZone StartTimeZone { get; set; }
        public int ConferenceType { get; set; }
        public bool AllowNewTimeProposal { get; set; }
        public string ICalUid { get; set; }
        public string ICalDateTimeStamp { get; set; }
        public bool IsAttachment { get; set; }
        public bool IsNew { get; set; }
        public Id Id { get; set; }
        public Id ParentFolderId { get; set; }
        public int Sensitivity { get; set; }
        public List<object> Attachments { get; set; }
        public string DateTimeReceived { get; set; }
        public int Size { get; set; }
        // public List<object> Categories { get; set; }
        public string Culture { get; set; }
        public int Importance { get; set; }
        public bool IsSubmitted { get; set; }
        public bool IsAssociated { get; set; }
        public bool IsDraft { get; set; }
        public bool IsFromMe { get; set; }
        public bool IsResend { get; set; }
        public bool IsUnmodified { get; set; }
        public string DateTimeSent { get; set; }
        public string DateTimeCreated { get; set; }
        public int AllowedResponseActions { get; set; }
        public string ReminderDueBy { get; set; }
        public bool IsReminderSet { get; set; }
        public int ReminderMinutesBeforeStart { get; set; }
        public string DisplayTo { get; set; }
        public bool HasAttachments { get; set; }
        public Body Body { get; set; }
        public string ItemClass { get; set; }
        public string Subject { get; set; }
        public string WebClientReadFormQueryString { get; set; }
        public string WebClientEditFormQueryString { get; set; }
        // public List<object> ExtendedProperties { get; set; }
        public int EffectiveRights { get; set; }
        public string LastModifiedName { get; set; }
        public string LastModifiedTime { get; set; }
        public ConversationId ConversationId { get; set; }
        public bool IsDirty { get; set; }
    }
}
