using alteridem.net.Extensions;
using Alteridem.ProjectSync.Model;

namespace Alteridem.ProjectSync
{
    class Program
    {
        static void Main(string[] args)
        {
            var solution = new Solution(@"C:\src\NUnit\nunit-framework\NUnitFramework.sln");

            var projectgroup = new ProjectGroup("nunit.framework");
            solution.ProjectGroups.Add(projectgroup);

            var project = new Project("nunit-framework-2.0");
            project.Ignore.Add("async.cs");
            project.Ignore.Add("lambda.cs");

            projectgroup.Projects.Add(project);
            projectgroup.Projects.Add(new Project("nunit-framework-3.5"));
            projectgroup.Projects.Add(new Project("nunit-framework-4.0"));
            projectgroup.Projects.Add(new Project("nunit-framework-4.5"));

            solution.SerializeToFile("nunit.xml");
        }
    }
}
