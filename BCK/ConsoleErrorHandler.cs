using System;

namespace BCK
{
    class ConsoleErrorHandler : IErrorHandler
    {
        public ConsoleErrorHandler()
        {

        }

        public void HandleError(Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
