// **********************************************************************************
// The MIT License (MIT)
// 
// Copyright (c) 2014 Rob Prouse
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy of
// this software and associated documentation files (the "Software"), to deal in
// the Software without restriction, including without limitation the rights to
// use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
// the Software, and to permit persons to whom the Software is furnished to do so,
// subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
// FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
// COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
// IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
// CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// 
// **********************************************************************************

#region Using Directives

using System;
using System.Collections;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using EnvDTE;
using Thread = System.Threading.Thread;

#endregion

namespace Alteridem.ProjectSync.Automation
{
    public class VisualStudio
    {
        #region Private Member Variables

        private const string VS_PROGID = "VisualStudio.DTE";
        private const uint S_OK = 0;
        private const int MAX_RETRIES = 1000;
        private const int DELAY_TIMEOUT = 10;
        private const uint RPC_E_SERVERCALL_RETRYLATER = 0x8001010A;

        private DTE _dte;

        #endregion

        #region Native P/Invoke imports

        [DllImport("ole32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        static extern uint GetRunningObjectTable(uint res, out IRunningObjectTable objectTable);

        [DllImport("ole32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        static extern uint CreateBindCtx(uint res, out IBindCtx ctx);

        #endregion

        #region Public Methods

        /// <summary>
        /// Connect to a running instance of Visual Studio
        /// </summary>
        /// <param name="solutionPath"></param>
        /// <returns></returns>
        public bool Connect(string solutionPath)
        {
            _dte = Retry(() => GetCurrentDte(solutionPath));
            return _dte != null;
        }

        /// <summary>
        /// Used for testing to connect to the currently running instance of
        /// Visual Studio 2013
        /// </summary>
        /// <returns></returns>
        public bool ConnectToFirstRunningInstance()
        {
            _dte = Retry(() => (DTE)Marshal.GetActiveObject("VisualStudio.DTE.12.0"));
            if (_dte != null) return true;

            _dte = Retry(() => (DTE)Marshal.GetActiveObject("VisualStudio.DTE.11.0"));
            if (_dte != null) return true;

            _dte = Retry(() => (DTE)Marshal.GetActiveObject("VisualStudio.DTE.10.0"));
            if (_dte != null) return true;

            _dte = Retry(() => (DTE)Marshal.GetActiveObject("VisualStudio.DTE.9.0"));
            if (_dte != null) return true;

            return false;
        }

        /// <summary>
        /// Launch Visual Studio and open the given solution
        /// </summary>
        /// <param name="solutionPath"></param>
        /// <returns></returns>
        public bool Launch(string solutionPath)
        {
            try
            {
                _dte = new DTE();
                _dte.MainWindow.Visible = true;
                _dte.Solution.Open(solutionPath);
            }
            catch (Exception x)
            {
                Console.WriteLine("Failed to launch visual studio or load the solution file {0}.  Exception: {1}", solutionPath, x.Message);
            }
            return true;
        }

        /// <summary>
        /// Disconnect from Visual Studio and optionally close it
        /// </summary>
        /// <param name="closeDte"></param>
        /// <returns></returns>
        public bool Disconnect(bool closeDte)
        {
            try
            {
                if (_dte == null)
                {
                    return true;
                }
                if (closeDte)
                {
                    _dte.Quit();
                }
                Marshal.ReleaseComObject(_dte);
            }
            catch (Exception x)
            {
                Console.WriteLine("Failed to connect to a running instance of Visual Studio. Exception was thrown: {0}", x.Message);
            }
            return _dte == null;
        }

        public Project GetProject(string projectName)
        {
            Console.WriteLine("Looking for {0} project...", projectName);
            Debug.WriteLine(string.Format("Looking for {0} project...", projectName));

            if (_dte == null)
            {
                throw new InvalidOperationException("Call Connect() or Launch() first");
            }

            Projects projects = Retry(() => _dte.Solution.Projects);

            if (projects == null)
            {
                return null;
            }

            Project[] projArray = Retry(() =>
            {
                IEnumerable ie = projects;
                return ie.Cast<Project>().ToArray();
            });

            int retries = MAX_RETRIES;
            int i = 0;
            // Console.WriteLine( "Looping through collection of projects." );
            while (projArray != null && i < projArray.Length)
            {
                try
                {
                    Project proj = projArray[i];

                    // Console.WriteLine( "Found project: {0}", proj.Name );
                    if (proj.Name.ToUpper() == projectName.ToUpper() && proj.Kind != "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}") // Constants.vsProjectKindSolutionItems
                    {
                        return proj;
                    }
                    Project proj2 = GetProject(proj, projectName);
                    if (proj2 != null)
                    {
                        return proj2;
                    }
                    i++;
                }
                catch (COMException cx)
                {
                    if ((uint)cx.ErrorCode == RPC_E_SERVERCALL_RETRYLATER)
                    {
                        Thread.Sleep(DELAY_TIMEOUT);
                        Console.WriteLine("retry counter#2    attempt #{0}", --retries);
                    }
                    else
                    {
                        Console.WriteLine("Exception: other");
                        throw;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("an unexpected exception happened.  We are eating this exception:  {0}", e.Message);
                }
            }
            return null;
        }

        public Project GetProject(Project parentProj, string projectName)
        {
            try
            {
                Debug.WriteLine("GetProject: {0}, {1}", parentProj.Name, projectName);
                Console.WriteLine("GetProject: {0}, {1}", parentProj.Name, projectName);
            }
            catch
            {
                Debug.WriteLine("Exception fetching project name");
                Console.WriteLine("Exception fetching project name.");
            }

            int retries = MAX_RETRIES;
            ProjectItem[] items = SafeGetProjectItems(parentProj);
            if (_dte == null)
            {
                throw new InvalidOperationException("Call Connect() or Launch() first");
            }

            if (items != null)
            {
                return Retry(() =>
                {
                    int itemsIndex = 0;
                    while (itemsIndex < items.Length)
                    {
                        ProjectItem item = items[itemsIndex];
                        Project proj = item.SubProject;
                        if (proj != null)
                        {
                            if (proj.Name.ToUpper() == projectName.ToUpper()
                                && proj.Kind != "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}")
                            // Constants.vsProjectKindSolutionItems
                            {
                                return proj;
                            }

                            return Retry(() => GetProject(proj, projectName));
                        }
                        itemsIndex++;
                    }
                    return null;
                });
            }
            return null;
        }

        public void EnsureInProject(Project project, string fileName)
        {
            ProjectItem pi = Retry(() => GetProjectItem(project, fileName));
            if (pi == null)
                throw new ApplicationException(string.Format("Failed to ensure the file {0} in project {1}", fileName, project.Name));
        }

        public bool RemoveFromProject(Project project, string fileName)
        {
            if (DoesExistInProject(project, fileName))
            {
                ProjectItem item = Retry(() => GetProjectItem(project, fileName));
                if (item == null) return false;

                // remove it from the project
                item.Remove();
                return true;
            }
            return false;
        }

        /// <summary>
        /// Call this method to see if the filename passed in exists in the project.
        /// </summary>
        /// <param name="project">The visual studio project to check.</param>
        /// <param name="fileName">The filename of the file searching for.</param>
        /// <returns><b>true</b> if the file exists in the project, <b>false</b> otherwise.</returns>
        public bool DoesExistInProject(Project project, string fileName)
        {
            var item = Retry(() => GetProjectItem(project, fileName));
            return item != null;
        }

        #endregion

        #region Non-Public methods

        private ProjectItem[] SafeGetProjectItems(Project project)
        {
            return Retry(() =>
            {
                ProjectItems items = project.ProjectItems;
                if (items != null)
                {
                    IEnumerable ie = items;
                    return ie.Cast<ProjectItem>().ToArray();
                }
                return null;
            });
        }

        private ProjectItem[] SafeGetProjectItems(ProjectItem projectItem)
        {
            ProjectItem[] retArray = Retry(() =>
            {
                ProjectItems items = projectItem.ProjectItems;
                IEnumerable ie = items;
                return ie.Cast<ProjectItem>().ToArray();
            });
            return retArray;
        }


        private ProjectItem GetProjectItem(Project project, string name)
        {
            ProjectItem[] items = SafeGetProjectItems(project);
            if (items != null)
            {
                return Retry(() =>
                {
                    int itemIndex = 0;
                    while (itemIndex < items.Length)
                    {
                        ProjectItem projectItem = items[itemIndex];
                        Debug.WriteLine("Comparing " + projectItem.Name);
                        if (string.Compare(projectItem.Name, name, true, CultureInfo.InvariantCulture) == 0)
                        {
                            return projectItem;
                        }
                        projectItem = GetChildProjectItem(name, projectItem);
                        if (projectItem != null) return projectItem;
                        itemIndex++;
                    }
                    return null;
                });
            }
            return null;
        }

        private ProjectItem GetChildProjectItem(string name, ProjectItem projectItem)
        {
            ProjectItem[] childItems = SafeGetProjectItems(projectItem);
            if (childItems != null)
            {
                ProjectItem child = Retry(() =>
                {
                    int index = 0;

                    while (index < childItems.Length)
                    {
                        ProjectItem childItem = childItems[index];

                        for (short i = 0; i < childItem.FileCount; ++i)
                        {
                            Debug.WriteLine("\tComparing " + childItem.FileNames[i]);
                            if (childItem.FileNames[i].EndsWith(name, true, CultureInfo.InvariantCulture))
                            {
                                return childItem;
                            }
                        }
                        index++;
                    }
                    return null;
                });
                if (child != null) return projectItem;
            }
            return null;
        }

        private DTE GetCurrentDte(string solutionPath)
        {
            IRunningObjectTable runningObjectTable;
            uint hres = GetRunningObjectTable(0, out runningObjectTable);
            if (hres == S_OK)
            {
                IEnumMoniker enumMoniker;
                runningObjectTable.EnumRunning(out enumMoniker);
                if (enumMoniker != null)
                {
                    const int MAX_MONIKERS = 5000;
                    var monArr = new IMoniker[MAX_MONIKERS];
                    IntPtr pCount = Marshal.AllocHGlobal(IntPtr.Size);
                    enumMoniker.Next(MAX_MONIKERS, monArr, pCount);
                    //create a binding to moniker
                    int count = Marshal.ReadInt32(pCount);
                    IBindCtx ctx;
                    CreateBindCtx(0, out ctx);

                    for (int i = 0; i < count; i++)
                    {
                        DTE dte = Retry(() =>
                        {
                            string name;
                            monArr[i].GetDisplayName(ctx, null, out name);
                            if (name.Contains(VS_PROGID))
                            {
                                object tmp;
                                runningObjectTable.GetObject(monArr[i], out tmp);
                                var dteTmp = (DTE)tmp;
                                Console.WriteLine(dteTmp.Solution.FileName);
                                var cmpr = new CaseInsensitiveComparer();
                                if (cmpr.Compare(dteTmp.Solution.FileName, solutionPath) == 0)
                                {
                                    Console.WriteLine("Attached to existing visual studio process.");
                                    return dteTmp;
                                }
                            }
                            return null;
                        });
                        if (dte != null) return dte;
                    }
                }
            }
            return null;
        }

        private static T Retry<T>(Func<T> func, int retries = MAX_RETRIES) where T : class
        {
            while (retries > 0)
            {
                try
                {
                    return func();
                }
                catch (COMException cx)
                {
                    if ((uint)cx.ErrorCode == RPC_E_SERVERCALL_RETRYLATER)
                    {
                        Thread.Sleep(DELAY_TIMEOUT);
                        retries--;
                    }
                    else
                    {
                        throw;
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
