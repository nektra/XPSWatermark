using System;
using System.Collections.Generic;
using System.Text;
using Nektra.Deviare2;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;

namespace IEPrintWatermark
{
    class Program
    {
        static NktSpyMgr spyMgr;
        static NktProcess procIE;

        static void Main(string[] args)
        {
            spyMgr = new NktSpyMgr();
            if (spyMgr.Initialize() < 0)
            {
                Console.WriteLine("ERROR: Cannot initialize Deviare engine");
                return;
            }
            spyMgr.OnCreateProcessCall += new DNktSpyMgrEvents_OnCreateProcessCallEventHandler(spyMgr_OnCreateProcessCall);
            spyMgr.OnLoadLibraryCall += new DNktSpyMgrEvents_OnLoadLibraryCallEventHandler(spyMgr_OnLoadLibraryCall);

            KillRunningInternetExplorerInstances();

            if (LaunchAndHookInternetExplorer() == false)
            {
                Console.WriteLine("ERROR: Unable to launch Microsoft Internet Explorer");
                return;
            }

            Console.Write("Close IE or press any key to quit...");
            while (procIE.get_IsActive(100) != false)
            {
                if (Console.KeyAvailable != false)
                {
                    Console.ReadKey(true);
                    break;
                }
            }
            Console.WriteLine("");
        }

        static void spyMgr_OnLoadLibraryCall(NktProcess proc, string dllName, object moduleHandle)
        {
            System.Diagnostics.Trace.WriteLine("IEPrintWatermark [LoadLibraryCall]: " + dllName);
            if (dllName.ToLower().EndsWith("xpsservices.dll") != false)
            {
                HookXpsInterfaces(proc);
            }
        }

        static void spyMgr_OnCreateProcessCall(NktProcess proc, int childPid, int mainThreadId, bool is64BitProcess, bool canHookNow)
        {
            NktProcess childProc = spyMgr.ProcessFromPID(childPid);
            if (childProc != null && childProc.Name.ToLower().EndsWith("iexplore.exe") != false)
            {
                spyMgr.LoadAgent(childProc);
            }
        }

        private static void KillRunningInternetExplorerInstances()
        {
            foreach (NktProcess proc in spyMgr.Processes())
            {
                if (proc.Name.ToLower().EndsWith("iexplore.exe") != false)
                    proc.Terminate(-1);
            }
        }

        private static bool LaunchAndHookInternetExplorer()
        {
            object continueEvent;
            string sExeName;

            sExeName = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            sExeName = "\"" + sExeName + "\\Internet Explorer\\iexplore.exe\" http://www.google.com";
            procIE = spyMgr.CreateProcess(sExeName, true, out continueEvent);
            if (procIE == null)
                return false;

            spyMgr.LoadAgent(procIE);

            spyMgr.ResumeProcess(procIE, continueEvent);
            return true;
        }

        private static void HookXpsInterfaces(NktProcess proc)
        {
            NktProcessMemory procMem;
            string dllName;
            int pointerSize, retVal;
            object callParams;
            IntPtr remoteBuffer, ptrVal;

            System.Diagnostics.Trace.WriteLine("IEPrintWatermark [HookXpsInterfaces]: Start (a)");
            dllName = GetAppPath() + "IEPrintWatermarkHelper";
            if (proc.PlatformBits == 64)
                dllName += "64";
            dllName += ".dll";

            pointerSize = 0;
            procMem = null;
            remoteBuffer = IntPtr.Zero;
            ptrVal = IntPtr.Zero;
            try
            {
                System.Diagnostics.Trace.WriteLine("IEPrintWatermark [HookXpsInterfaces]: Start (b)");
                //allocate memory for retrieving results
                pointerSize = proc.PlatformBits / 8;
                procMem = proc.Memory();
                System.Diagnostics.Trace.WriteLine("IEPrintWatermark [HookXpsInterfaces]: Start (c)");
                remoteBuffer = procMem.AllocMem(new IntPtr(pointerSize), false);
                System.Diagnostics.Trace.WriteLine("IEPrintWatermark [HookXpsInterfaces]: Start (d)");
                //load helper dll and retrieve the pointer we need
                spyMgr.LoadCustomDll(proc, dllName, true, true);
                System.Diagnostics.Trace.WriteLine("IEPrintWatermark [HookXpsInterfaces]: LoadCustomDll 0x" + LastCallError().ToString("X"));
                if (pointerSize == 4)
                    callParams = new int[1] { remoteBuffer.ToInt32() };
                else
                    callParams = new long[1] { remoteBuffer.ToInt64() };
                retVal = spyMgr.CallCustomApi(proc, dllName, "GetXpsAddresses", ref callParams, true);
                System.Diagnostics.Trace.WriteLine("IEPrintWatermark [HookXpsInterfaces]: CallCustomApi 0x" + LastCallError().ToString("X"));
                spyMgr.UnloadCustomDll(proc, dllName, true);
            }
            catch (System.Exception)
            {
                retVal = -1;
            }
            System.Diagnostics.Trace.WriteLine("IEPrintWatermark [HookXpsInterfaces]: retVal 0x" + retVal.ToString("X"));
            //get IXpsOMPageReference::CollectPartResources's address
            if (retVal >= 0)
            {
                try
                {
                    if (pointerSize == 4)
                        ptrVal = new IntPtr(Convert.ToInt32(procMem.Read(remoteBuffer, eNktDboFundamentalType.ftSignedDoubleWord)));
                    else
                        ptrVal = new IntPtr(Convert.ToInt64(procMem.Read(remoteBuffer, eNktDboFundamentalType.ftSignedQuadWord)));
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Trace.WriteLine("IEPrintWatermark [X]: " + ex.ToString());
                    ptrVal = IntPtr.Zero;
                }
            }
            //free memory
            try
            {
                if (procMem != null && remoteBuffer != IntPtr.Zero)
                    procMem.FreeMem(remoteBuffer);
            }
            catch (System.Exception)
            { }
            System.Diagnostics.Trace.WriteLine("IEPrintWatermark [HookXpsInterfaces]: ptrVal 0x" + ptrVal.ToInt32().ToString("X"));

            //if we have an address, create a hook for it
            if (ptrVal != IntPtr.Zero)
            {
                NktHook hk;

                hk = spyMgr.CreateHookForAddress(ptrVal, "XpsServices.dll!IXpsOMPageReference::SetPage", (int)eNktHookFlags.flgOnlyPreCall);
                hk.AddCustomHandler(GetAppPath() + "IEPrintWatermarkHelperCS.dll", (int)eNktHookCustomHandlerFlags.flgChDontCallIfLoaderLocked, "");
                hk.Hook(true);
                hk.Attach(proc, true);
            }
        }

        private static string GetAppPath()
        {
            string s = System.Reflection.Assembly.GetExecutingAssembly().Location;
            s = System.IO.Path.GetDirectoryName(s);
            if (s.EndsWith(@"\") == false)
                s = s + @"\";
            return s;
        }

        [DllImport("ole32.dll")]
        static extern int GetErrorInfo(uint dwReserved, out IErrorInfo pperrinfo);

        [Guid("1CF2B120-547D-101B-8E65-08002B2BD119"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
         ComImport]
        public interface IErrorInfo
        {
            [MethodImplAttribute(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int GetGuid(out Guid pGuid);

            [MethodImplAttribute(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int GetSource([MarshalAs(UnmanagedType.BStr)] out string pBstrSource);

            [MethodImplAttribute(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int GetDescription([MarshalAs(UnmanagedType.BStr)] out string pbstrDescription);

            [MethodImplAttribute(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int GetHelpFile([MarshalAs(UnmanagedType.BStr)] out string pBstrHelpFile);

            [MethodImplAttribute(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int GetHelpContext(out uint pdwHelpContext);
        }

        private static int LastCallError()
        {
            IErrorInfo errInfo = null;
            string sDesc = null;
            int res;

            if (GetErrorInfo(0, out errInfo) < 0 || errInfo == null)
                return -1;
            try
            {
                if (errInfo.GetDescription(out sDesc) < 0 || sDesc == null)
                    return -1;
            }
            catch (System.Exception ex)
            {
                return -2;
            }
            if (Int32.TryParse(sDesc, out res) == false)
                return -3;
            return res;
        }
    }
}
