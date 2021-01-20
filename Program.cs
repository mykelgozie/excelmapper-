using System;
using Topshelf;

namespace ErrorService
{
    class Program
    {
        static void Main(string[] args)
        {
            var exitCode = HostFactory.Run(X => {

                X.Service<ServiceClass>(S =>
                {
                    S.ConstructUsing(service => new ServiceClass());
                    S.WhenStarted(service => service.Start());
                    S.WhenStopped(service => service.Stop());

                });

                X.RunAsLocalSystem();
                X.SetServiceName(" Excel Mapper ");
                X.SetDisplayName("Excel Mapper ");
                X.SetDescription(" Mapper data to excel ");
            });

            int exitCodeValue = (int)Convert.ChangeType(exitCode, exitCode.GetTypeCode());
            Environment.ExitCode = exitCodeValue;
        }
    }
}
