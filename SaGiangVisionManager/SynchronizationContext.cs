using System;

using System.Collections.Generic;
using System.Text;

namespace Cognex.DataMan.SDK
{
    public delegate void SendOrPostCallback(object state);
    public class SynchronizationContext
    {
        private static System.Windows.Forms.Control _ctrl = new System.Windows.Forms.Control();

        public SynchronizationContext()
        {
            IntPtr x = _ctrl.Handle;
        }

        public void Post(SendOrPostCallback d, object state)
        {
            _ctrl.BeginInvoke(d, state);
        }

        public void Send(SendOrPostCallback d, object state)
        {
            if (_ctrl.InvokeRequired)
                _ctrl.EndInvoke(_ctrl.BeginInvoke(d, state));
            else
                _ctrl.Invoke(d, state);
        }
    }

    public static class WindowsFormsSynchronizationContext
    {
        private static SynchronizationContext _syncContext = new SynchronizationContext();

        public static SynchronizationContext Current
        {
            get
            {
                return _syncContext;
            }
        }
    }
}
