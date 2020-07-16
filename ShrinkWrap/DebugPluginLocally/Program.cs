﻿using System;
using Inventor;
using System.IO;

namespace DebugPluginLocally
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var inv = new InventorConnector())
            {
                InventorServer server = inv.GetInventorServer();

                try
                {
                    Console.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff"));
                    Console.WriteLine("Running locally...");
                    // run the plugin
                    //inv.GetInventorApplication().SilentOperation = true;
                    DebugSamplePlugin(server);
                }
                catch (Exception e)
                {
                    string message = $"Exception: {e.Message}";
                    if (e.InnerException != null)
                        message += $"{System.Environment.NewLine}    Inner exception: {e.InnerException.Message}";

                    Console.WriteLine(message);
                }
                finally
                {
                    //inv.GetInventorApplication().SilentOperation = false;
                    if (System.Diagnostics.Debugger.IsAttached)
                    {
                        Console.WriteLine(DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff"));
                        Console.WriteLine("Press any key to exit. All documents will be closed.");
                        Console.ReadKey();
                    }
                }
            }
        }

        /// <summary>
        /// Opens box.ipt and runs SamplePlugin
        /// </summary>
        /// <param name="app"></param>
        private static void DebugSamplePlugin(InventorServer app)
        {
            // create an instance of ShrinkWrapPlugin
            ShrinkWrapPlugin.SampleAutomation plugin = new ShrinkWrapPlugin.SampleAutomation(app);

            // run the plugin
            plugin.RunWithArguments(null, null);
        }
    }
}
