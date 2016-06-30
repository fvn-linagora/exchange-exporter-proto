using System;
using System.Collections.Generic;
using System.Linq;

namespace EchangeDumpedMessagesListener
{
    public static class AttendeeCollectionExtensions
    {
        public static IEnumerable<DDay.iCal.Attendee> MapToICalAttendees(this IEnumerable<Messages.Attendee> attendees) {

            return attendees
                .Where(IsAttendeesAddressSet)
                .Select(a => new {
                    Uri = new Uri("mailto:" + a.Address),
                    DisplayName = a.Name,
                })
                .Select(a => new DDay.iCal.Attendee(a.Uri) {
                    CommonName = a.DisplayName
                });
        }

        private static bool IsAttendeesAddressSet(Messages.Attendee a)
        {
            return a.RoutingType.Equals("SMTP", StringComparison.OrdinalIgnoreCase) && !String.IsNullOrWhiteSpace(a.Address);
        }
    }
}
