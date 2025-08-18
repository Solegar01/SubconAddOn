using SAPbobsCOM;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace SubconAddOn.Helpers
{
    public class SqlLockHelper : IDisposable
    {
        private readonly Company _company;
        private string _currentLockKey;
        private string _lockedBy;
        private bool _disposed;

        public SqlLockHelper(Company company)
        {
            _company = company ?? throw new ArgumentNullException(nameof(company));
        }

        public async Task<bool> AcquireLockAsync(
            string lockKey,
            string lockedBy,
            int retryDelayMs = 500,
            int maxRetries = 10,
            CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(lockKey)) throw new ArgumentNullException(nameof(lockKey));
            if (string.IsNullOrWhiteSpace(lockedBy)) lockedBy = Environment.UserName;

            _currentLockKey = lockKey;
            _lockedBy = lockedBy;

            int retries = 0;

            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();

                try
                {
                    var rs = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);

                    string sql = $@"
                        INSERT INTO ADD_ON_LOCK (LockKey, LockedBy, LockTime)
                        SELECT '{lockKey}', '{lockedBy}', GETDATE()
                        WHERE NOT EXISTS (SELECT 1 FROM ADD_ON_LOCK WHERE LockKey = '{lockKey}')";

                    rs.DoQuery(sql);

                    // If inserted, we have the lock
                    return true;
                }
                catch (Exception ex)
                {
                    // Likely a PK violation — means locked by someone else
                    retries++;
                    if (retries >= maxRetries)
                        return false;

                    await Task.Delay(retryDelayMs, cancellationToken).ConfigureAwait(false);
                }
            }
        }

        public void ReleaseLock()
        {
            if (_currentLockKey == null) return;

            try
            {
                var rs = (Recordset)_company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string sql = $"DELETE FROM ADD_ON_LOCK WHERE LockKey = '{_currentLockKey}' AND LockedBy = '{_lockedBy}'";
                rs.DoQuery(sql);
            }
            finally
            {
                _currentLockKey = null;
                _lockedBy = null;
            }
        }

        public void Dispose()
        {
            if (_disposed) return;

            if (_currentLockKey != null)
            {
                try { ReleaseLock(); } catch { }
            }

            _disposed = true;
        }
    }
}
