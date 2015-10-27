using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using Mono.Cecil;

namespace TfsSyncUtility
{
	public class ParsingAssembly
	{
        //*****Variables********
        readonly string unitTestAttribute = ConfigurationManager.AppSettings["unitTestAttribute"];
        readonly string tcAttribute = ConfigurationManager.AppSettings["customTcAttribute"];
       
        private readonly string _testAssembly;

		public ParsingAssembly(string testAssembly)
		{
			this._testAssembly = testAssembly;
		}

		public Dictionary<Int64, string> GetTestMethods()
		{
			var idNamePair = new Dictionary<Int64, string>();
			var assembly = AssemblyDefinition.ReadAssembly(_testAssembly);
			var allMethods = assembly.MainModule.Types
					.SelectMany(t => t.Methods);
			foreach (var methodDefinition in allMethods)
			{
				var hasUnitTestAttribute = methodDefinition.CustomAttributes.Any(a => a.AttributeType.FullName == unitTestAttribute);
				var hasTcAttribute = methodDefinition.CustomAttributes.Any(a => a.AttributeType.FullName == tcAttribute);
				if (hasUnitTestAttribute && hasTcAttribute)
				{
					var tcId = 0;
					string menthodName = methodDefinition.DeclaringType.FullName + "." + methodDefinition.Name;
					CustomAttribute firstOrDefault = methodDefinition.CustomAttributes.FirstOrDefault(a => a.AttributeType.FullName == tcAttribute);
					if (firstOrDefault != null)
						tcId = Convert.ToInt32(firstOrDefault.ConstructorArguments[0].Value);
					try
					{
						if (!idNamePair.ContainsKey(tcId))
							idNamePair.Add(tcId, menthodName);
					}
					catch (Exception e)
					{
						Console.WriteLine(e.Message);
						throw;
					}
				}
			}

			return idNamePair;
		}
	}
}
