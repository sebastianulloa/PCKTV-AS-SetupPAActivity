/*
****************************************************************************
*  Copyright (c) 2022,  Skyline Communications NV  All Rights Reserved.    *
****************************************************************************

By using this script, you expressly agree with the usage terms and
conditions set out below.
This script and all related materials are protected by copyrights and
other intellectual property rights that exclusively belong
to Skyline Communications.

A user license granted for this script is strictly for personal use only.
This script may not be used in any way by anyone without the prior
written consent of Skyline Communications. Any sublicensing of this
script is forbidden.

Any modifications to this script by the user are only allowed for
personal use and within the intended purpose of the script,
and will remain the sole responsibility of the user.
Skyline Communications will not be responsible for any damages or
malfunctions whatsoever of the script resulting from a modification
or adaptation by the user.

The content of this script is confidential information.
The user hereby agrees to keep this confidential information strictly
secret and confidential and not to disclose or reveal it, in whole
or in part, directly or indirectly to any person, entity, organization
or administration without the prior written consent of
Skyline Communications.

Any inquiries can be addressed to:

	Skyline Communications NV
	Ambachtenstraat 33
	B-8870 Izegem
	Belgium
	Tel.	: +32 51 31 35 69
	Fax.	: +32 51 31 01 29
	E-mail	: info@skyline.be
	Web		: www.skyline.be
	Contact	: Ben Vandenberghe

****************************************************************************
Revision History:

DATE		VERSION		AUTHOR			COMMENTS

dd/mm/2022	1.0.0.1		XXX, Skyline	Initial version
****************************************************************************
*/

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Skyline.DataMiner.Automation;
using Skyline.DataMiner.Net.Apps.DataMinerObjectModel;
using Skyline.DataMiner.Net.Messages;
using Skyline.DataMiner.Net.Messages.SLDataGateway;
using Skyline.DataMiner.Net.Profiles;
using Skyline.DataMiner.Net.Sections;
using Parameter = Skyline.DataMiner.Net.Profiles.Parameter;

/// <summary>
/// DataMiner Script Class.
/// </summary>
public class Script
{
	/// <summary>
	/// The Script entry point.
	/// </summary>
	/// <param name="engine">Link with SLAutomation process.</param>
	public void Run(Engine engine)
	{
		try
		{
			var name = engine.GetScriptParam("Activity Name").Value;
			var domDefinitionName = engine.GetScriptParam("DOM Definition").Value;
			var createScriptMessage = new SaveAutomationScriptMessage
			{
				IsUpdate = false,
				Definition = new GetScriptInfoResponseMessage
				{
					CheckSets = false,
					Options = ScriptOptions.AllowUndef,
					Name = name,
					Description = "Automatically generated automation script for PA Task: " + name,
					Type = AutomationScriptType.Automation,
					Dummies = new[]
					{
						new AutomationProtocolInfo
						{
							Description = "FunctionDVE",
							ProtocolId = 1,
							ProtocolName = "Skyline Process Automation.PA Script Task",
							ProtocolVersion = "Production",
						},
					},
					Parameters = new[]
					{
						new AutomationParameterInfo
						{
							Description = "Info",
							ParameterId = 1,
						},
						new AutomationParameterInfo
						{
							Description = "ProcessInfo",
							ParameterId = 2,
						},
						new AutomationParameterInfo
						{
							Description = "ProfileInstance",
							ParameterId = 3,
						},
					},
					Exes = new[]
					{
						new AutomationExeInfo
						{
							Id = 2,
							CSharpDebugMode = false,
							CSharpDllRefs = @"C:\Skyline DataMiner\Files\Newtonsoft.Json.dll;C:\Skyline DataMiner\Files\SLSRMLibrary.dll;C:\Skyline DataMiner\ProtocolScripts\ProcessAutomation.dll",
							PreCompile = false,
							Type = AutomationExeType.CSharpCode,
							ValueOffset = 0,
							Value = defaultcsharp,
						},
					},
					Memories = new AutomationMemoryInfo[0],
				},
			};

			engine.SendSLNetMessage(createScriptMessage);

			// create/check parameters based on DOM
			// From given DOM Definition, create profile parameters to include in profile definition
			DomHelper domHelper = new DomHelper(engine.SendSLNetMessages, "process_automation");

			var domDefinitionFilter = DomDefinitionExposers.Name.Equal(domDefinitionName);
			var definition = domHelper.DomDefinitions.Read(domDefinitionFilter).FirstOrDefault();
			if (definition == null)
			{
				engine.Log("Setup PA Activity|Given DOM definition doesn't exist: " + domDefinitionName);
				return;
			}

			List<SectionDefinition> sections = new List<SectionDefinition>();
			foreach (var sectionLink in definition.SectionDefinitionLinks)
			{
				var sectionFilter = SectionDefinitionExposers.ID.Equal(sectionLink.SectionDefinitionID);
				var section = domHelper.SectionDefinitions.Read(sectionFilter).First();
				sections.Add(section);
			}

			var fields = GetAllFieldNames(sections);

			var currentProfileParameters = new GetProfileManagerParameterMessage();
			var profileParameters = engine.SendSLNetMessage(currentProfileParameters)[0] as ProfileManagerResponseMessage;

			var existingParameters = profileParameters.ManagerObjects.Select(x => x.Item2 as Parameter).ToList();

			var fieldParameters = GetParametersFromFields(fields);

			List<Parameter> parametersToCreate = new List<Parameter>();
			List<Parameter> parametersForProfile = new List<Parameter>();
			CheckExistingParameters(existingParameters, fieldParameters, parametersToCreate, parametersForProfile);

			var profileparamSetter = new SetProfileManagerParameterMessage(parametersToCreate);
			var paramSetterResponse = engine.SendSLNetMessage(profileparamSetter).First() as ProfileManagerResponseMessage;
			if (!paramSetterResponse.Success)
			{
				engine.Log("Setup PA Activity|Failed to set up new parameters for the profile.");
				return;
			}

			var profilesRequest = new GetProfileDefinitionMessage();
			var availableProfiles = engine.SendSLNetMessage(profilesRequest)[0] as ProfileManagerResponseMessage;

			var scriptTaskProfile = GetPaScriptTaskProfile(engine, availableProfiles);

			// check if profile def exists, if so, update with new parameters
			var existingProfileDefinitions = availableProfiles.ManagerObjects.Select(x => x.Item2 as ProfileDefinition).Where(x => x != null && x.Name.Equals(name));
			if (!existingProfileDefinitions.Any())
			{
				CreateProfileDefinition(engine, name, parametersForProfile, scriptTaskProfile);
			}
			else
			{
				// update
				var createdProfile = existingProfileDefinitions.First();
				foreach (var profileParam in parametersForProfile)
				{
					if (!createdProfile.ParameterSettings.ContainsKey(profileParam.ID))
					{
						createdProfile.Parameters.Add(profileParam);
					}
				}

				SetProfileDefinitionMessage updateProfileMessage = new SetProfileDefinitionMessage(createdProfile);
				var response = engine.SendSLNetMessage(updateProfileMessage)[0] as ProfileManagerResponseMessage;
			}

			// Commenting out profile instance code because that doesn't seem to work well with process automation
			// var profileDefId = profileDefResponse.UpdatedProfileDefinitions.First().ID;

			// var profileInstance = new ProfileInstance
			// {
			// 	Name = name,
			// 	AppliesToID = profileDefId,
			// };

			// var newProfileInstanceMessage = new SetProfileInstanceMessage(profileInstance);
			// engine.SendSLNetMessage(newProfileInstanceMessage);
		}
		catch (Exception ex)
		{
			engine.Log("Setup PA Activity|Error creating activity: " + ex);
		}
	}

	private static void CreateProfileDefinition(Engine engine, string name, List<Parameter> parametersForProfile, ProfileDefinition scriptTaskProfile)
	{
		var profile = new ProfileDefinition
		{
			Name = name,
			Scripts = new[] { new ScriptEntry { Name = "PA Script", Script = name } },
		};

		profile.BasedOnIDs.Add(scriptTaskProfile.ID);

		foreach (var profileParam in parametersForProfile)
		{
			profile.Parameters.Add(profileParam);
		}

		var newProfileMessage = new SetProfileDefinitionMessage(profile);
		engine.SendSLNetMessage(newProfileMessage);
	}

	private static void CheckExistingParameters(List<Parameter> existingParameters, List<Parameter> fieldParameters, List<Parameter> parametersToCreate, List<Parameter> parametersForProfile)
	{
		foreach (var parameter in fieldParameters)
		{
			if (!existingParameters.Any(x => x.Name.Equals(parameter.Name)))
			{
				parametersToCreate.Add(parameter);
				parametersForProfile.Add(parameter);
			}
			else
			{
				parametersForProfile.Add(existingParameters.First(x => x.Name == parameter.Name));
			}
		}
	}

	private static List<Parameter> GetParametersFromFields(List<string> fields)
	{
		var profileDefinitionParams = new List<Parameter>();
		foreach (var field in fields)
		{
			var profileParam = new Parameter(Guid.NewGuid())
			{
				Name = field,
				Type = Parameter.ParameterType.Text,
				IsOptional = true,
			};

			profileDefinitionParams.Add(profileParam);
		}

		return profileDefinitionParams;
	}

	private static ProfileDefinition GetPaScriptTaskProfile(Engine engine, ProfileManagerResponseMessage availableProfiles)
	{
		var baseProfiles = availableProfiles.ManagerObjects.Select(x => x.Item2 as ProfileDefinition).Where(x => x != null && x.Name.Equals("PA Script Task"));
		if (!baseProfiles.Any())
		{
			engine.Log("Missing PA Script Task profile.");
			engine.ExitFail("Missing PA Script Task profile.");
		}

		var scripTaskProfile = baseProfiles.First();
		return scripTaskProfile;
	}

	private static List<string> GetAllFieldNames(List<SectionDefinition> sections)
	{
		List<string> fieldsList = new List<string>();
		foreach (var section in sections)
		{
			var fields = section.GetAllFieldDescriptors();
			foreach (var field in fields)
			{
				fieldsList.Add(field.Name);
			}
		}

		return fieldsList;
	}

	#region defaultcsharp
	private const string defaultcsharp = @"
/*
****************************************************************************
*  Copyright (c) 2022,  Skyline Communications NV  All Rights Reserved.    *
****************************************************************************

By using this script, you expressly agree with the usage terms and
conditions set out below.
This script and all related materials are protected by copyrights and
other intellectual property rights that exclusively belong
to Skyline Communications.

A user license granted for this script is strictly for personal use only.
This script may not be used in any way by anyone without the prior
written consent of Skyline Communications. Any sublicensing of this
script is forbidden.

Any modifications to this script by the user are only allowed for
personal use and within the intended purpose of the script,
and will remain the sole responsibility of the user.
Skyline Communications will not be responsible for any damages or
malfunctions whatsoever of the script resulting from a modification
or adaptation by the user.

The content of this script is confidential information.
The user hereby agrees to keep this confidential information strictly
secret and confidential and not to disclose or reveal it, in whole
or in part, directly or indirectly to any person, entity, organization
or administration without the prior written consent of
Skyline Communications.

Any inquiries can be addressed to:

	Skyline Communications NV
	Ambachtenstraat 33
	B-8870 Izegem
	Belgium
	Tel.	: +32 51 31 35 69
	Fax.	: +32 51 31 01 29
	E-mail	: info@skyline.be
	Web		: www.skyline.be
	Contact	: Ben Vandenberghe

****************************************************************************
Revision History:

DATE		VERSION		AUTHOR			COMMENTS

dd/mm/2022	1.0.0.1		XXX, Skyline	Initial version
****************************************************************************
*/

using System;
using System.Diagnostics;
using Skyline.DataMiner.Automation;
using Skyline.DataMiner.DataMinerSolutions.ProcessAutomation.Helpers.Logging;
using Skyline.DataMiner.DataMinerSolutions.ProcessAutomation.Manager;

/// <summary>
/// DataMiner Script Class.
/// </summary>
public class Script
{
	/// <summary>
	/// The Script entry point.
	/// </summary>
	/// <param name=""engine"">Link with SLAutomation process.</param>
	public void Run(Engine engine)
	{
		var helper = new PaProfileLoadDomHelper(engine);
		try
		{
			
		}
		catch (Exception ex)
		{
			engine.Log(""Error: "" + ex);
		}
	}
}
";
	#endregion
}