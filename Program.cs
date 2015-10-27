using System;
using System.Collections.Generic;
using Microsoft.TeamFoundation.TestManagement.Client;
using NDesk.Options;
using System.Configuration;

namespace TfsSyncUtility
{
	public class Program
	{
		public static Dictionary<Int64, string> TestMethodsList;

		private static void Main(string[] args)
		{
            //*****Variables*****
            //Path to TFS server
		    string serverName = @ConfigurationManager.AppSettings["tfsServerPath"];
			//******************

            #region Options
			string project = null;
			string path = null;
			bool force = false;
			bool rewrite = false;
			var options = new OptionSet
			{
				{ "p|path=", v => path = v },
				{ "n|name=", v => project = v },
				{ "r", v => rewrite = v != null },
				{ "f", v => force = v != null },
			};

			try
			{
				options.Parse(args);
				if (project == null && path == null)
				{
					Console.WriteLine("This app should run with options.");
					Environment.ExitCode = 2;
				}
			}
			catch (OptionException e)
			{
				Console.WriteLine("Unable to parse keys" + e.Message);
			}
			#endregion
			
            try
			{
				var file = path;
				var tfs = new TfsWrapper(serverName, project);
				var teamProject = tfs.ConnectToTestManagementTfs();
				var parser = new ParsingAssembly(path);
				TestMethodsList = parser.GetTestMethods();
				foreach (int id in TestMethodsList.Keys)
				{
					string fullTestName;
					ITestCase testCase = teamProject.TestCases.Find(id);
					var implementation = (ITmiTestImplementation)testCase.Implementation;
					//if not automated
					if (!tfs.IsImplemented(testCase))
					{
						TestMethodsList.TryGetValue(id, out fullTestName);
						tfs.SetAutomation(testCase, fullTestName, file);
					}
					//if already automated
					else
					{
						if (rewrite && force)
						{
							TestMethodsList.TryGetValue(id, out fullTestName);
							if (fullTestName != implementation.TestName)
							{
								tfs.RemoveAutomation(testCase);
								tfs.SetAutomation(testCase, fullTestName, file);
                                Console.WriteLine("Force changed");
                                Console.WriteLine("------------");
							}
						}
						else if (rewrite)
						{
							TestMethodsList.TryGetValue(id, out fullTestName);
							if (fullTestName != implementation.TestName)
							{
								Console.WriteLine("Automation on case Id={0} with current association {1} will be changed to {2}.", id,
									implementation.TestName, fullTestName);
								Console.WriteLine("Press Y to perform altering or press Enter to skip");
								while (true)
								{
									ConsoleKeyInfo result = Console.ReadKey();
									if (result.Key == ConsoleKey.Y)
									{
										tfs.RemoveAutomation(testCase);
										tfs.SetAutomation(testCase, fullTestName, file);
										Console.WriteLine("Changed");
										Console.WriteLine("------------");
										break;
									}
									if (result.Key == ConsoleKey.Enter)
									{
										Console.WriteLine("Skiped");
										Console.WriteLine("------------");
										break;
									}
								}
							}
						}
					}
				}
			}
			catch (Exception e)
			{
				Console.Write("Smth bad happened: " + e.Message);
				Environment.ExitCode = 1;
			}
            // comment this region in case of using in TFS build template
            #region finally
            finally
            {
                while (true)
                {
                    Console.WriteLine("Press ESC to exit");
                    var result = Console.ReadKey();
                    if (result.Key == ConsoleKey.Escape)
                    {
                        break;
                    }
                }
            }
            #endregion
        }
	}
}
