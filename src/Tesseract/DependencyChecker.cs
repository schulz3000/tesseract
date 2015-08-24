using Microsoft.Win32;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace Tesseract
{
    static class DependencyChecker
    {
        [DllImport("msi.dll")]
        static extern INSTALLSTATE MsiQueryProductState(string product);
        public static bool IsVCRedist2012Installed()
        {
            return IsVCRedist2012InstalledByMsi() || IsVCRedist2012InstalledByRegistry() || IsVCRedist2012InstalledByFile();
        }

        static bool IsVCRedist2012InstalledByMsi()
        {
            var productcode = Is64BitProzess ? "{CF2BEA3C-26EA-32F8-AA9B-331F7E34BA97}" : "{BD95A8CD-1D9F-35AD-981A-3E7925026EBB}";
            var productState = MsiQueryProductState(productcode);
            return productState == INSTALLSTATE.INSTALLSTATE_DEFAULT || productState == INSTALLSTATE.INSTALLSTATE_LOCAL;
        }

        static bool IsVCRedist2012InstalledByRegistry()
        {
            var regKey = Registry.LocalMachine.OpenSubKey(@"Software\Wow6432Node\Microsoft\VisualStudio\11.0\VC\Runtimes\" + (Is64BitProzess ? "x64" : "x86"));
            var regInstallKey = regKey.GetValue("Installed");

            return regInstallKey != null && ((int)regInstallKey) == 1;
        }

        static bool IsVCRedist2012InstalledByFile()
        {
            return File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), @"system32\msvcr110.dll"));
        }

        static bool Is64BitProzess { get { return IntPtr.Size == 8; } }

        enum INSTALLSTATE
        {
            INSTALLSTATE_NOTUSED = -7,  // component disabled
            INSTALLSTATE_BADCONFIG = -6,  // configuration data corrupt
            INSTALLSTATE_INCOMPLETE = -5,  // installation suspended or in progress
            INSTALLSTATE_SOURCEABSENT = -4,  // run from source, source is unavailable
            INSTALLSTATE_MOREDATA = -3,  // return buffer overflow
            INSTALLSTATE_INVALIDARG = -2,  // invalid function argument
            INSTALLSTATE_UNKNOWN = -1,  // unrecognized product or feature
            INSTALLSTATE_BROKEN = 0,  // broken
            INSTALLSTATE_ADVERTISED = 1,  // advertised feature
            INSTALLSTATE_REMOVED = 1,  // component being removed (action state, not settable)
            INSTALLSTATE_ABSENT = 2,  // uninstalled (or action state absent but clients remain)
            INSTALLSTATE_LOCAL = 3,  // installed on local drive
            INSTALLSTATE_SOURCE = 4,  // run from source, CD or net
            INSTALLSTATE_DEFAULT = 5,  // use default, local or source
        }
    }
}
