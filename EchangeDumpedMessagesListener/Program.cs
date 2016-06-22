using EasyNetQ;
using System;
using System.Collections.Generic;
using System.Linq;
using Messages;

namespace EchangeDumpedMessagesListener
{
    class Program
    {
        private static readonly string MQCONNETIONSTRING = "host=10.69.0.117";

        static void Main(string[] args)
        {
            using (var bus = RabbitHutch.CreateBus(MQCONNETIONSTRING))
            {
                bus.Subscribe<NewAppointmentDumped>("test", HandleNewAppointment);

                Console.WriteLine("Listening for messages. Hit <return> to quit.");
                Console.ReadLine();
            }
        }

        static void HandleNewAppointment(NewAppointmentDumped app)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Got message: {0}", app.Appointment.subject);
            Console.ResetColor();
        }
    }
}
