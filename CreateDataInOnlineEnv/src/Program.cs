using Microsoft.Crm.QA.Utf;
using System;
using System.Collections;
using System.Globalization;
using System.Reflection;
using UnitTest;
namespace CommonScenarioTests
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Utility.InitializeUtf(assembly.GetName().Name);

            ProgramSwitches switches = ProcessArguments(args);

            if (switches.ShowHelp)
            {
                // ShowUsage();
            }
            else
            {
                TestSuite suite = (TestSuite)Utility.BuildTestSuiteMatching(
                        assembly,
                        Environment.CommandLine,
                        (string[])switches.NonSwitchArgs.ToArray(typeof(string)),
                        switches.ExactMatch);
                suite.Execute();
            }
        }

        private static ProgramSwitches ProcessArguments(string[] args)
        {
            ProgramSwitches switches = new ProgramSwitches();

            foreach (string argument in args)
            {
                if (!String.IsNullOrEmpty(argument))
                {
                    if ((argument.Length > 1) && ((argument[0] == '/') || (argument[0] == '-')))
                    {
                        string[] subArguments = argument.Split(new char[] { ':' }, 2);
                        // Note: subArguments length is assured to be greater than 1.
                        switch (subArguments[0].Substring(1).ToLower(CultureInfo.InvariantCulture))
                        {
                            case "?":
                                switches.ShowHelp = true;
                                break;
                            case "f1perf":
                                switches.F1Profiling = true;
                                if (subArguments.Length > 1)
                                {
                                    switches.PerformanceProcess = subArguments[1];
                                }
                                if (switches.OfficeProfiling) // only one of the performance profiler can be enabled
                                {
                                    switches.ShowHelp = true;
                                }
                                break;
                            case "offperf":
                                switches.OfficeProfiling = true;
                                if (subArguments.Length > 1)
                                {
                                    switches.PerformanceProcess = subArguments[1];
                                }
                                if (switches.F1Profiling) // only one of the performance profiler can be enabled
                                {
                                    switches.ShowHelp = true;
                                }
                                break;
                            case "crmplatformprofiling":
                                switches.CrmPlatformProfiling = true;
                                break;
                            case "starttime":
                                Console.WriteLine("Current Time: " + DateTime.Now.ToLongTimeString());
                                if (subArguments.Length > 1)
                                {
                                    switches.StartTime = Convert.ToDateTime(subArguments[1], CultureInfo.CurrentCulture.DateTimeFormat);
                                }
                                break;
                            case "exactmatch":
                                switches.ExactMatch = true;
                                break;
                            default:
                                // Unrecognized switch, do nothing
                                break;
                        }
                    }
                    else
                    {
                        switches.NonSwitchArgs.Add(argument);
                    }
                }
            }

            return switches;
        }

        private class ProgramSwitches
        {
            /// <summary>
            /// Show help.
            /// </summary>
            /// <remarks>By default show help is false.</remarks>
            public bool ShowHelp = false;

            /// <summary>
            /// Indicates if F1 profiling is enabled.
            /// </summary>
            /// <remarks>By default performance tracing is off.</remarks>
            public bool F1Profiling = false;

            /// <summary>
            /// Indicates if Office profiling is enabled.
            /// </summary>
            /// <remarks>By default performance tracing is off.</remarks>
            public bool OfficeProfiling = false;

            /// <summary>
            /// Indicates if CRM Platform profiling is enabled.
            /// </summary>
            /// <remarks>By default CRM Platform profiling is off.</remarks>
            public bool CrmPlatformProfiling = false;

            /// <summary>
            /// Indicate the process that should be profiled for performance.
            /// </summary>
            public string PerformanceProcess = String.Empty;

            /// <summary>
            /// Indicate start time of the run.
            /// </summary>
            public DateTime StartTime = DateTime.Now;

            /// <summary>
            /// Non switch arguments.
            /// </summary>
            /// <remarks>By default it will be empty collection.</remarks>
            public ArrayList NonSwitchArgs = new ArrayList();

            /// <summary>
            /// Determines if an exact name match is required to run test cases
            /// </summary>
            public bool ExactMatch { get; set; }
        }
    }
}
