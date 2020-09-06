
using System;

namespace BehaviorSelectorConverter
{
    public static class ConsoleUtility
    {
        // Info.

        public static void Info(string message)
        {
            Console.ForegroundColor = ConsoleColor.DarkGray;

            Console.WriteLine(message);

            Console.ResetColor();
        }

        public static void Info(string format, params object[] args)
        {
            Info(string.Format(format, args));
        }

        // Warning.

        public static void Warning(string message)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;

            Console.WriteLine(message);

            Console.ResetColor();
        }

        public static void Warning(string format, params object[] args)
        {
            Warning(string.Format(format, args));
        }

        // Error.

        public static void Error(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;

            Console.WriteLine(message);

            Console.ResetColor();
        }

        public static void Error(string format, params object[] args)
        {
            Error(string.Format(format, args));
        }

        // Progress.

        public static void Progress(string message)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;

            Console.WriteLine();

            Console.WriteLine(message);

            Console.WriteLine();

            Console.ResetColor();
        }

        // Task.

        public static void Task(string message)
        {
            Console.ForegroundColor = ConsoleColor.DarkGreen;

            Console.WriteLine(message);

            Console.ResetColor();
        }

        public static void Task(string format, params object[] args)
        {
            Task(string.Format(format, args));
        }
    }
}
