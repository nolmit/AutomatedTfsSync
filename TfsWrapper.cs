using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;

namespace TfsSyncUtility
{
	public class TfsWrapper
    {  
		private readonly string _projectName;

		private readonly string _serverName;

		public TfsWrapper(string serverName, string projectName)
		{
			this._projectName = projectName;
			this._serverName = serverName;
		}

		public ITestManagementTeamProject ConnectToTestManagementTfs()
		{
			var tfsUri = new Uri(_serverName);
			ITestManagementTeamProject teamProject;
			try
			{
				TfsTeamProjectCollection projectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(tfsUri);
				teamProject =
					projectCollection.GetService<ITestManagementService>().GetTeamProject(_projectName);
                Console.WriteLine("Connected to {0}/{1} ->TestManagementService", tfsUri , _projectName);
			}
			catch (Exception)
			{
                Console.WriteLine("Unable to connect the TFS server {0}/{1} ->TestManagementService", tfsUri, _projectName);
				throw;
			}
			return teamProject;
		}

		public IBuildDefinition GetBuildDefinitionsTfs(string buidDefinitionName)
		{
			Uri tfsUri = new Uri(_serverName);
			IBuildDefinition testBuild;
			try
			{
				TfsTeamProjectCollection projectCollection = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(tfsUri);
				var teamProject =
					projectCollection.GetService<IBuildServer>();
                testBuild = teamProject.QueryBuildDefinitions(_projectName).First(a => a.Name.Contains(buidDefinitionName));
                Console.WriteLine("Connected to {0}/{1} ->BuildServer", tfsUri, _projectName);
                Console.WriteLine("Build definition name: {0}", buidDefinitionName);
			}
			catch (Exception)
			{
				Console.WriteLine("Unable to connect the TFS server {0}/{1} -> BuildServer", tfsUri, _projectName);
				throw;
			}
			return testBuild;
		}

		public bool IsImplemented(ITestCase test)
		{
			var isImplemented = false;
			var implementation = (ITmiTestImplementation)test.Implementation;
			if (implementation != null)
			{
				isImplemented = true;
				Console.WriteLine(test.Id + " is already implemented");
			}
			return isImplemented;
		}

		public void SetAutomation(ITestCase testCase, string automationTestName, string filename)
		{
			try
			{
				var arr = filename.Split(new[] { "\\" }, StringSplitOptions.None);
				var crypto = new SHA1CryptoServiceProvider();
				var bytes = new byte[16];
				Array.Copy(crypto.ComputeHash(Encoding.Unicode.GetBytes(automationTestName)), bytes, bytes.Length);
				var automationGuid = new Guid(bytes);

				testCase.Implementation = testCase.Project.CreateTmiTestImplementation(
					automationTestName,
					"UI Test",
					arr.Last(),
					automationGuid);
				testCase.Save();
				Console.WriteLine("Automation is associated to {0}", testCase.Title);
			}
			catch (Exception)
			{
				Console.WriteLine("Unable to set automation to : {0}", testCase.Title);
				throw;
			}
		}

		public void RemoveAutomation(ITestCase testCase)
		{
            //standard tfs fields values
			const string automationField = "Microsoft.VSTS.TCM.AutomationStatus";
			const string automationFieldValue = "Not Automated";
			try
			{
				testCase.Implementation = null;
				testCase.CustomFields[automationField].Value = automationFieldValue;
				testCase.Save();
			}
			catch (Exception)
			{
				Console.WriteLine("Unable to remove automation from : {0}", testCase.Title);
				throw;
			}
		}
	}
}
