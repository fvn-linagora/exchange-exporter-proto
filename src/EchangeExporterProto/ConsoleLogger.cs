using System;

namespace EchangeExporterProto
{
    interface ILog
    {
        void Info(string message);
        void Error(string message);
        void Warn(string message);
    }
    class ConsoleLogger : ILog
    {
        public void Error(string message)
        {
            ColoredConsoleWrite(ConsoleColor.Red, message);
        }

        public void Info(string message)
        {
            ColoredConsoleWrite(ConsoleColor.White, message);
        }

        public void Warn(string message)
        {
            ColoredConsoleWrite(ConsoleColor.Yellow, message);
        }
        public static void ColoredConsoleWrite(ConsoleColor color, string text)
        {
            ConsoleColor originalColor = Console.ForegroundColor;
            Console.ForegroundColor = color;
            Console.Write(text);
            Console.ForegroundColor = originalColor;
        }
    }
}
