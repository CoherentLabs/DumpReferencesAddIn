/*
Copyright (C) 2013 Coherent Labs AD

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated 
documentation files (the "Software"), to deal in the Software without restriction, including without limitation the 
rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit 
persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions 
of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED 
TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL 
THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, 
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
*/
using System;
using System.Collections.Generic;
using Extensibility;
using EnvDTE;
using EnvDTE80;

using Microsoft.VisualStudio.VCProjectEngine;
using Microsoft.VisualStudio.CommandBars;

namespace DumpReferencesAddIn
{
    /// <summary>The object for implementing an Add-in.</summary>
	/// <seealso class='IDTExtensibility2' />
    public class Connect : IDTExtensibility2, IDTCommandTarget
	{
		/// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
		public Connect()
		{
		}

        struct RefData
        {
            public string Parent;
            public string ReferencedProject;
            public string ReferencedGUID;
        }

        private void EnumSFProjects(Project sf, List<Project> list)
        {
            for (var i = 1; i <= sf.ProjectItems.Count; i++)
            {
                var subProject = sf.ProjectItems.Item(i).SubProject;
                if (subProject == null)
                {
                    continue;
                }

                if (subProject.Kind == TypeProject)
                {
                    list.Add(subProject);
                }
                else if (subProject.Kind == TypeSolutionFolder)
                {
                    EnumSFProjects(subProject, list);
                }
            }
        }

        private void DumpReferences()
        {
            Window window = _applicationObject.Windows.Item(Constants.vsWindowKindOutput);
          
            OutputWindow outputWindow = (OutputWindow)window.Object;
            OutputWindowPane owp;
            owp = outputWindow.OutputWindowPanes.Add("Dump References");

            owp.OutputString("Initializing references dump on " + _applicationObject.Solution.FullName + "\n");

            owp.OutputString("================ All project names ================ \n");
            Dictionary<string, string> allProjectGUIDs = new Dictionary<string, string>();
            List<Project> allProjs = new List<Project>();
            foreach (Project proj in _applicationObject.Solution.Projects)
            {
                if (proj.Kind == TypeProject)
                {
                    allProjs.Add(proj);
                }
                else if(proj.Kind == TypeSolutionFolder)
                {
                    EnumSFProjects(proj, allProjs);                    
                }
            }

            foreach (Project pr in allProjs)
            {
                VCProject vcpr = (VCProject)pr.Object;
                owp.OutputString(vcpr.Name + "\n");
                allProjectGUIDs.Add(vcpr.Name, vcpr.ProjectGUID);
            }

            HashSet<string> uniqueReferences = new HashSet<string>();

            List<RefData> allReferences = new List<RefData>();
            owp.OutputString("================ Full list of expanded dependencies ================ \n");
            foreach (Project proj in allProjs)
            {
                owp.OutputString("--- " + proj.Name + " ---\n");
                if (proj.Properties != null)
                {
                    foreach (Property prop in proj.Properties)
                    {
                        if (prop.Name == "VCReferences")
                        {
                            VCReferences references = (VCReferences)prop.Object;
                            foreach (VCReference reference in references)
                            {
                                owp.OutputString(reference.Name + "\n");
                                uniqueReferences.Add(reference.Name);
                                RefData item = new RefData();
                                item.Parent = proj.Name;
                                item.ReferencedProject = reference.Name;

                                VCProjectReference projRef = reference as VCProjectReference;
                                if (projRef != null)
                                {
                                    item.ReferencedGUID = projRef.ReferencedProjectIdentifier;
                                }
                                allReferences.Add(item);
                            }
                        }
                    }
                }
            }

            owp.OutputString("================ List of expanded unique dependencies ================ \n");
            foreach (var reference in uniqueReferences)
            {
                owp.OutputString(reference + "\n");
            }

            owp.OutputString("================ Possible issues ================ \n");
            foreach (RefData reference in allReferences)
            {
                if (reference.ReferencedGUID != "")
                {
                    string projectGUIDInSolution;
                    if (allProjectGUIDs.TryGetValue(reference.ReferencedProject, out projectGUIDInSolution))
                    {
                        if (projectGUIDInSolution != reference.ReferencedGUID)
                        {
                            owp.OutputString("Reference " + reference.ReferencedProject + " referenced by " + reference.Parent + " has a different GUID than the loaded in the solution! \n");
                            owp.OutputString(reference.ReferencedProject + " in solution: " + projectGUIDInSolution + " in reference: " + reference.ReferencedGUID + "\n");
                        }
                    }
                    else
                    {
                        owp.OutputString("Reference " + reference.ReferencedProject + " referenced by " + reference.Parent + " is MISSING FROM THE SOLUTION! \n");
                    }
                }
                else
                {
                    owp.OutputString("Reference " + reference.ReferencedProject + " referenced by " + reference.Parent + " has an UNKNOWN GUID! \n");
                }
            }
        }
                
		/// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
		/// <param term='application'>Root object of the host application.</param>
		/// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
		/// <param term='addInInst'>Object representing this Add-in.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
			_applicationObject = (DTE2)application;
			_addInInstance = (AddIn)addInInst;

            Commands2 commands = (Commands2)_applicationObject.Commands;

            try
            {
                _command = _applicationObject.Commands.Item(_addInInstance.ProgID + "." + COMMAND_NAME, -1);
            }
            catch
            {
            }
            // Add the command if it does not exist
            if (_command == null)
            {
                object[] contextUIGuids = new object[] { };
                _command = commands.AddNamedCommand(_addInInstance, COMMAND_NAME, "Dump references", "Dump the references of all projects", true, 59, ref contextUIGuids,
                                                (int)(vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled));

                CommandBar standardToolBar = ((CommandBars)_applicationObject.CommandBars)["Solution"];
                CommandBarButton dumpRefsBtn = (CommandBarButton)_command.AddControl(standardToolBar, standardToolBar.Controls.Count);

                // Change some button properties
                dumpRefsBtn.Caption = "Dump References";
                dumpRefsBtn.Style = MsoButtonStyle.msoButtonIcon;
                dumpRefsBtn.BeginGroup = true;
                dumpRefsBtn.Visible = true;
            }
		}

        public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
        {
            if (neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
            {
                if (commandName == _addInInstance.ProgID + "." + COMMAND_NAME)
                {
                    status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                    return;
                }
            }
        }

        public void Exec(string commandName, vsCommandExecOption executeOption,
                 ref object varIn, ref object varOut, ref bool handled)
        {
            handled = false;
            if (executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault)
            {
                if (commandName == _addInInstance.ProgID + "." + COMMAND_NAME)
                {
                    DumpReferences();
                    handled = true;
                    return;
                }
            }
        }

		/// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
		/// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        { }

		/// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />		
		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref Array custom)
		{
		}

		/// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnBeginShutdown(ref Array custom)
		{
		}
		
		private DTE2 _applicationObject;
		private AddIn _addInInstance;
        private Command _command;
        private string COMMAND_NAME = "DumpReference";

        const string TypeProject = "{8BC9CEB8-8B4A-11D0-8D11-00A0C91BC942}";
        const string TypeSolutionFolder = "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}";
	}
}