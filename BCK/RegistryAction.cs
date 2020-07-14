using Microsoft.Win32;
using System;
using System.Security.AccessControl;

namespace BCK
{
    static class RegistryAction
    {
        static RegistryKey rkApp = null;
        static readonly string strAppSubkey = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;

        private static void initRegKey(RegistryRights regRights)
        {
            RegistryKey rkHKCU = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64);
            RegistryKey rkHKCUSoftware = rkHKCU.OpenSubKey("SOFTWARE");

            try
            {
                rkApp = rkHKCUSoftware.OpenSubKey(strAppSubkey,
                    RegistryKeyPermissionCheck.ReadWriteSubTree,
                    regRights);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while opening HKCU\\SOFTWARE Subkey '" +
                    strAppSubkey + "' - Message: " + ex.Message);
                rkApp = null;
            }
            
        }

        private static void closeRegKey()
        {
            if (rkApp != null)
            {
                rkApp.Close();
            }
        }
        
        public static string GetValue(string vName)
        {
            initRegKey(RegistryRights.ReadKey);
            if (rkApp != null)
            {
                string regValue = rkApp.GetValue(vName).ToString();
                closeRegKey();
                return regValue;
            }
            return string.Empty;
        }

        public static bool SetValue(string vName, string vValue)
        {
            initRegKey(RegistryRights.WriteKey);
            if (rkApp != null)
            {
                rkApp.SetValue(vName, vValue, RegistryValueKind.String);
                closeRegKey();
                return true;
            }
            return false;
        }
    }
}
