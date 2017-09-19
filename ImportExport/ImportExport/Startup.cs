using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ImportExport.Startup))]
namespace ImportExport
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
