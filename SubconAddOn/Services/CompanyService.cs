using SAPbobsCOM;
using SAPbouiCOM.Framework;
using System;

namespace SubconAddOn.Services
{
    public static class CompanyService
    {
        private static Company _company;

        /// <summary>
        /// Mengambil objek Company aktif dari UI API (GetDICompany)
        /// </summary>
        public static Company GetCompany()
        {
            if (_company == null || !_company.Connected)
            {
                // Ambil koneksi DI-API dari UI-API
                _company = (Company)Application.SBO_Application.Company.GetDICompany();

                if (!_company.Connected)
                {
                    throw new Exception("Failed to get connected DI-API company object.");
                }
            }

            return _company;
        }

        /// <summary>
        /// Menutup koneksi DI-API jika masih terhubung.
        /// </summary>
        public static void Disconnect()
        {
            if (_company != null && _company.Connected)
            {
                try
                {
                    _company.Disconnect();
                }
                catch (Exception ex)
                {
                    // Logging optional
                    System.Diagnostics.Debug.WriteLine("Disconnect error: " + ex.Message);
                }
            }

            _company = null;
        }
    }
}
