using System.Collections.Generic;
using Autodesk.Forge.DesignAutomation.Model;

namespace Interaction
{
    /// <summary>
    /// Customizable part of Publisher class.
    /// </summary>
    internal partial class Publisher
    {
        /// <summary>
        /// Constants.
        /// </summary>
        private static class Constants
        {
            private const int EngineVersion = 24;
            public static readonly string Engine = $"Autodesk.Inventor+{EngineVersion}";

            public const string Description = "Shrinkwraps model and converts to FBX";

            internal static class Bundle
            {
                public static readonly string Id = "ShrinkWrap";
                public const string Label = "prod";

                public static readonly AppBundle Definition = new AppBundle
                {
                    Engine = Engine,
                    Id = Id,
                    Description = Description
                };
            }

            internal static class Activity
            {
                public static readonly string Id = Bundle.Id;
                public const string Label = Bundle.Label;
            }

            internal static class Parameters
            {
                public const string inputFile = nameof(inputFile);
                public const string inputJson = nameof(inputJson);
                public const string outputIpt = nameof(outputIpt);
                public const string outputObjZip = nameof(outputObjZip);
                public const string outputFbx = nameof(outputFbx);
            }
        }


        /// <summary>
        /// Get command line for activity.
        /// </summary>
        private static List<string> GetActivityCommandLine()
        {
            return new List<string> {
              $"$(engine.path)\\InventorCoreConsole.exe /al $(appbundles[{Constants.Activity.Id}].path)"
            };
            //return new List<string> {
            //  $"$(engine.path)\\InventorCoreConsole.exe /al $(appbundles[{Constants.Activity.Id}].path)",
            //  $"$(appbundles[ShrinkWrap].path)\\ShrinkWrapPlugin.bundle\\Contents\\converter.bat"
            //};
        }

        /// <summary>
        /// Get activity parameters.
        /// </summary>
        private static Dictionary<string, Parameter> GetActivityParams()
        {
            return new Dictionary<string, Parameter>
                    {
                        {
                            Constants.Parameters.inputFile,
                            new Parameter
                            {
                                Verb = Verb.Get,
                                LocalName = "input",
                                Zip = true,
                                Description = "Input assembly to simplify"
                            }
                        },
                        {
                            Constants.Parameters.inputJson,
                            new Parameter
                            {
                                Verb = Verb.Get,
                                LocalName = "params.json",
                            }
                        },
                        {
                            Constants.Parameters.outputIpt,
                            new Parameter
                            {
                                Verb = Verb.Put,
                                LocalName = "output.ipt",
                            }
                        },
                        {
                            Constants.Parameters.outputObjZip,
                            new Parameter
                            {
                                Verb = Verb.Put,
                                Required = false,
                                Zip = true,
                                LocalName = "outputObjZip",
                            }
                        },
                        {
                            Constants.Parameters.outputFbx,
                            new Parameter
                            {
                                Verb = Verb.Put,
                                Required = false,
                                LocalName = "output.fbx",
                            }
                        }
                    };
        }

        /// <summary>
        /// Get arguments for workitem.
        /// </summary>
        private static Dictionary<string, IArgument> GetWorkItemArgs()
        {
            string json = System.IO.File.ReadAllText("params.json").Replace("\r\n", "");
            const string myOssBucket = "https://developer.api.autodesk.com/oss/v2/buckets/pugvwjyn5udkpb0g2zv54zicddguubqz_simplify";
            const string accessToken = "eyJhbGciOiJIUzI1NiIsImtpZCI6Imp3dF9zeW1tZXRyaWNfa2V5In0.eyJzY29wZSI6WyJjb2RlOmFsbCIsImRhdGE6d3JpdGUiLCJkYXRhOnJlYWQiLCJidWNrZXQ6Y3JlYXRlIiwiYnVja2V0OmRlbGV0ZSIsImJ1Y2tldDpyZWFkIl0sImNsaWVudF9pZCI6InB1R1Z3SlluNXVES1BCMEcyenY1NHpJQ0RkZ3VVYnF6IiwiYXVkIjoiaHR0cHM6Ly9hdXRvZGVzay5jb20vYXVkL2p3dGV4cDYwIiwianRpIjoiN3dwOUpGd2VnbWt5V3pyMHpsMkdoRlNNSWF3QnAyVWszYWFRYjNSSmJXRmhJVE4zZkQ2QUVieHF2Q0cyRHl1TSIsImV4cCI6MTU4OTQ2NzEzM30.vqqu-Q45MeLtt8KFzrEfNAnLOPpNYX9tjdX0dH1BinE";

            // TODO: update the URLs below with real values
            return new Dictionary<string, IArgument>
                    {
                        {
                            Constants.Parameters.inputFile,
                            new XrefTreeArgument
                            {
                                Url = $"{myOssBucket}/objects/BastianToyotaGoodYear.zip",
                                Headers = new Dictionary<string, string>()
                                {
                                    { "Authorization", "Bearer " + accessToken }
                                }
                            }
                        },
                        {
                            Constants.Parameters.inputJson,
                            new XrefTreeArgument
                            {
                                Url = "data:application/json," + json
                            }
                        },
                        {
                            Constants.Parameters.outputIpt,
                            new XrefTreeArgument
                            {
                                Verb = Verb.Put,
                                Optional = true,
                                Url = $"{myOssBucket}/objects/output.ipt",
                                Headers = new Dictionary<string, string>()
                                {
                                    { "Authorization", "Bearer " + accessToken }
                                }
                            }
                        },
                        {
                            Constants.Parameters.outputObjZip,
                            new XrefTreeArgument
                            {
                                Verb = Verb.Put,
                                Optional = true,
                                Url = $"{myOssBucket}/objects/output.obj.zip",
                                Headers = new Dictionary<string, string>()
                                {
                                    { "Authorization", "Bearer " + accessToken }
                                }
                            }
                        },
                        {
                            Constants.Parameters.outputFbx,
                            new XrefTreeArgument
                            {
                                Verb = Verb.Put,
                                Optional = true,
                                 Url = $"{myOssBucket}/objects/output.fbx",
                                 Headers = new Dictionary<string, string>()
                                {
                                    { "Authorization", "Bearer " + accessToken }
                                }
                            }
                        }
                    };
        }
    }
}
