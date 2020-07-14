using System.Deployment.Application;
using System.Reflection;


namespace BCK
{
    class DeploymentAction
    {
        public static string GetRunningVersion()
        {
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                return ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString(4);
            }
            else
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }
    }
}
