using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SubconAddOn.Services
{
    public class NamedMutex : IDisposable
    {
        private Mutex _mutex;
        private bool _acquired;

        public string Key { get; }

        public NamedMutex(string key)
        {
            Key = $"Global\\SAPB1_ADDON_{key}";
            try
            {
                _mutex = new Mutex(false, Key);
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to create named mutex.", ex);
            }
        }

        // Try acquire the mutex (non-blocking)
        public bool TryAcquire()
        {
            if (_mutex == null) return false;
            _acquired = _mutex.WaitOne(0); // returns immediately
            return _acquired;
        }

        // Acquire with timeout
        public bool Wait(int millisecondsTimeout)
        {
            if (_mutex == null) return false;
            _acquired = _mutex.WaitOne(millisecondsTimeout);
            return _acquired;
        }

        // Release manually
        public void Release()
        {
            if (_acquired && _mutex != null)
            {
                try
                {
                    _mutex.ReleaseMutex();
                    _acquired = false;
                }
                catch { }
            }
        }

        public void Dispose()
        {
            Release();
            _mutex?.Dispose();
            _mutex = null;
        }
    }

}
