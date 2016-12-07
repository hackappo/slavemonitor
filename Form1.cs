using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Automation;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;

namespace slaveMonitor
{
    public enum ErrorType
    {
        NoError = 0,
        SecurityCheckError = 1,
        ConnectionError = 2,
        SlaveAgentNotLaunched = 3,
        SlaveAgentTerminated = 4
    }

    public partial class SlaveMonitor : Form
    {
        const string ExpectedProcessName = "jp2launcher";
        const string SlaveMonitorProcessName = "slaveMonitor";
        const string Command = "javaws";
        const string ScriptName = @"-Xnosplash .\cloudbees-agent.jnlp";
        private TimeSpan pollInverval = TimeSpan.FromMinutes(10);
        private TimeSpan inBetweenDialogsInterval = TimeSpan.FromSeconds(5);
        private Thread workerThread;
        

        public SlaveMonitor()
        {
            InitializeComponent();
            Show();
            KillPrevSlaveMonitor();

            workerThread = new Thread(new System.Threading.ThreadStart(Loop));
            workerThread.Start();
        }

        public bool IsHostProcessAlive()
        {
            label1.Text = "Checking...";
            int? handleCount = 0;
            var processList = Process.GetProcesses();
            var targetProcess =
                (processList.Where(process => process.ProcessName.Equals(ExpectedProcessName))).FirstOrDefault();
            if (targetProcess != null)
            {
                handleCount = targetProcess?.HandleCount;
                if (handleCount != 22)   //Failing case had 20 User Handles
                    throw new ApplicationException();
                else
                    return true;
            }
            else
            {
                return false;
            }

        }

        public void LaunchJenkinsApplet()
        {
            label1.Text = "Launching Jenkins slave agent...";
            var processInfo = new ProcessStartInfo();
            processInfo.FileName = Command;
            processInfo.Arguments = ScriptName;
            Debug.WriteLine("Running command: {0} {1}", processInfo.FileName, processInfo.Arguments);
            var p = Process.Start(processInfo);
            Debug.WriteLine(p.HandleCount);
        }

        public void KillAgent()
        {
            var processList = Process.GetProcesses();
            var targetProcess =
                (processList.Where(process => process.ProcessName.Equals(ExpectedProcessName))).FirstOrDefault();
            targetProcess?.Kill();
        }

        public void KillPrevSlaveMonitor()
        {
            //Kill all previous slaveMonitor processes but current one
            var currentProcessId = Process.GetCurrentProcess().Id;
            var processList = Process.GetProcesses();
            var processes = (processList.Where(process => process.ProcessName.Equals(SlaveMonitorProcessName) && process.Id != currentProcessId));

            foreach (var process in processes)
            {
                process.Kill();
            }
        }

        public ErrorType CheckErrors()
        {
            Debug.WriteLine("Looking for Security dialog");

            var propCondition = new AndCondition(new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window) ,new PropertyCondition(AutomationElement.NameProperty, "Security Warning"));
            var servletWnd = AutomationElement.RootElement.FindFirst(TreeScope.Children, propCondition);
            if (servletWnd != null)
                return ErrorType.SecurityCheckError;

            Debug.WriteLine("Looking Slave agent Window");
            propCondition = new AndCondition(new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window), new PropertyCondition(AutomationElement.NameProperty, "Jenkins slave agent"));
            var slaveWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, propCondition);

            if (slaveWindow == null)
                return ErrorType.SlaveAgentNotLaunched;

            Debug.WriteLine("Looking for connection error Window");
            propCondition = new AndCondition(new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window), new PropertyCondition(AutomationElement.NameProperty, "Error"));
            var errorDialog = slaveWindow.FindFirst(TreeScope.Children, propCondition);

            if(errorDialog != null)
                return ErrorType.ConnectionError;

            return ErrorType.NoError;
        }

        public void Loop()
        {
            while (true)
            {
                try
                {
                    if (!IsHostProcessAlive())
                    {
                        Debug.WriteLine("Process not found, running applet");
                        LaunchJenkinsApplet();
                        System.Threading.Thread.Sleep(inBetweenDialogsInterval);
                        var executionError = CheckErrors();
                        if (executionError != ErrorType.NoError)
                        {
                            Debug.WriteLine($"ErrorType {executionError}: Human intervention required!!");
                            if(executionError != ErrorType.SlaveAgentNotLaunched)
                                KillAgent();
                            break;
                        }
                        else
                        {
                            label1.Text = "He's awake!!!";
                        }
                    }

                    System.Threading.Thread.Sleep(pollInverval);
                }
                catch (ApplicationException)
                {
                    throw;
                }
            }

            this.Close();
        }
    }
}
