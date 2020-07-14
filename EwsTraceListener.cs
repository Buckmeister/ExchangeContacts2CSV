using Microsoft.Exchange.WebServices.Data;
using System;

namespace ExchangeContacts2CSV
{
    class EwsTraceListener : ITraceListener
    {
        #region ITraceListener Members
        public void Trace(string traceType, string traceMessage)
        {
            Console.WriteLine("EWS Trace Typ: " + traceType);
            Console.WriteLine(traceMessage);
        }
        #endregion
        
    }
}
