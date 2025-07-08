using System;
using System.Collections.Generic;
using SAPbouiCOM.Framework;

namespace SubconAddOn
{
    internal static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                // 1️⃣  Bootstrap UI‐API
                Application oApp = args.Length < 1
                    ? new Application()            // run from SAP (normal)
                    : new Application(args[0]);     // debug via SboGuiApi.Connect

                // 2️⃣  Hand off ke controller
                AddonController.Start();

                // 3️⃣  Dispatch loop
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Add‑on failed to start:\n" + ex.Message,
                    "Headless Add‑on",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}
