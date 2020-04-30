using System;
using Ranorex;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Text.RegularExpressions;

namespace MyRanorexTest
{
    class Program
    {
        static void Main(string[] args)
        {
            InitResolver();
            RanorexInit();
            run();
        }
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void InitResolver()
        {
            Ranorex.Core.Resolver.AssemblyLoader.Initialize();
        }
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void RanorexInit()
        {
            TestingBootstrapper.SetupCore();
        }
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static int run()
        {
            int error = 0;
            //Start calculator and wait for UI to be loaded
            try
            {
                System.Diagnostics.Process pr = System.Diagnostics.Process.Start("calc.exe");
                Thread.Sleep(2000);
                //Get process name
                string processName = GetActualCalculatorProcessName();

                //Find Calculator | Windows 10
                if (IsWindows10())
                {
                    WindowsApp calculator = Host.Local.FindSingle<WindowsApp>("winapp[@processname='" + processName + "']");

                    Button button = calculator.FindSingle<Ranorex.Button>(".//button[@automationid='num2Button']");
                    button.Click();
                    Report.Info(button.Text);

                    button = calculator.FindSingle<Ranorex.Button>(".//button[@automationid='plusButton']");
                    button.Click();
                    //Report.Info(button.Text);
                    Report.Log(ReportLevel.Info, button.Text);

                    button = calculator.FindSingle<Ranorex.Button>(".//button[@automationid='num3Button']");
                    button.Click();
                    Report.Info(button.Text);

                    button = calculator.FindSingle<Ranorex.Button>(".//button[@automationid='minusButton']");
                    button.Click();
                    Report.Info(button.Text);

                    button = calculator.FindSingle<Ranorex.Button>(".//button[@automationid='num3Button']");
                    button.Click();
                    Report.Info(button.Text);

                    button = calculator.FindSingle<Ranorex.Button>(".//button[@automationid='equalButton']");
                    button.Click();
                    Report.Info(button.Text);
                    new Regex(Regex.Escape("Open flashing door."));
                    //if (button.Text=="2")
                    Report.Success("Test OK");
                    Thread.Sleep(2000);
                    //Close calculator
                    calculator.As<Form>().Close();
                }
                //Find Calculator | Windows 8.X or older
                else
                {
                    Form calculator = Host.Local.FindSingle<Form>("form[@processname='" + processName + "']");

                    calculator.EnsureVisible();

                    Button button = calculator.FindSingle<Ranorex.Button>(".//button[@controlid='132']");
                    button.Click();

                    button = calculator.FindSingle<Ranorex.Button>(".//button[@controlid='92']");
                    button.Click();

                    button = calculator.FindSingle<Ranorex.Button>(".//button[@controlid='133']");
                    button.Click();

                    button = calculator.FindSingle<Ranorex.Button>(".//button[@controlid='121']");
                    button.Click();

                    //Close calculator
                    calculator.Close();
                }
            }
            catch (RanorexException e)
            {
                Console.WriteLine(e.ToString());
                error = -1;
            }

            return error;
        }
        private static string GetActualCalculatorProcessName()
        {
            string processName = String.Empty;
            var processes = System.Diagnostics.Process.GetProcesses();

            foreach (var item in processes)
            {
                if (item.ProcessName.ToLowerInvariant().Contains("calc"))
                {
                    processName = item.ProcessName;
                    break;
                }
            }

            return processName;
        }

        private static bool IsWindows10()
        {
            var reg = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion");

            string productName = (string)reg.GetValue("ProductName");

            return productName.StartsWith("Windows 10");
        }

    }
}
