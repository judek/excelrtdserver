using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Timers;
using System.Runtime.InteropServices;

namespace ExcelRTDServer
{

    //To register on system use the commands
    
    //For  x86
    // %SystemRoot%\Microsoft.Net\Framework\v4.0.30319\RegAsm.exe ExcelRTDServer.dll /codebase

    //For x64
    //%SystemRoot%\Microsoft.Net\Framework64\v4.0.30319\RegAsm.exe ExcelRTDServer.dll /codebase
    
    //Note that only one of the commands can be registered at the same time.  If you want to register both
    //It is probably best to name the assemblies differently 
    
    //Then to see it run in Excel use the following function from within an Excel cell:
    // =RTD("Jude.Sample.RtdServer", , "topic")

    

    [Guid("2A03620E-2254-4061-9BC2-023F4EF62B15"), ProgId("Jude.Sample.RtdServer")]
    public class RtdServer : IRtdServer
    {
        private IRTDUpdateEvent m_callback;
        private Timer m_NotifyPeriodTimer;
        private int m_topicId;

        private void TimerEventHandler(Object source, ElapsedEventArgs e)
        {
            m_NotifyPeriodTimer.Stop();//Stopping the timer is important since we don’t want to call UpdateNotify repeatedly
            m_callback.UpdateNotify();
        }
        
        public int ServerStart(IRTDUpdateEvent callback)
        {
            m_callback = callback;
            m_NotifyPeriodTimer = new Timer(2000);
            m_NotifyPeriodTimer.Elapsed += TimerEventHandler;
            return 1;
        }

        private string GetTime()
        {
            return DateTime.Now.ToString("hh:mm:ss:ff");
        }

        public object ConnectData(int topicId, ref Array strings, ref bool newValues)
        {
            m_topicId = topicId;
            m_NotifyPeriodTimer.Start();
            return GetTime();
        }

        public Array RefreshData(ref int topicCount)
        {
            object[,] data = new object[2, 1];
            data[0, 0] = m_topicId;
            data[1, 0] = GetTime();

            topicCount = 1;

            m_NotifyPeriodTimer.Start();
            return data;
        }

        public void DisconnectData(int topicId)
        {
            m_NotifyPeriodTimer.Stop();
        }

        public int Heartbeat()
        {
            return DateTime.Now.Second;
        }

        public void ServerTerminate()
        {
            //throw new NotImplementedException();
        }
    }
}
