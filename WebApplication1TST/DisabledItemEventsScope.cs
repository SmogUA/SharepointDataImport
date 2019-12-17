using System;
using Microsoft.SharePoint;

namespace DataImport
{
    public class DisabledItemEventsScope : SPItemEventReceiver, IDisposable
    {
        private bool oldValue;

        public void Dispose()
        {
            EventFiringEnabled = oldValue;
        }
        public DisabledItemEventsScope()
        {
            oldValue = EventFiringEnabled;
            EventFiringEnabled = false;
        }
    }
}
