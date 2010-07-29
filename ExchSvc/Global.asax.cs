using System;
using System.ServiceModel.Activation;
using System.Web;
using System.Web.Routing;

namespace TALHO
{
    public class Global : HttpApplication
    {
        void Application_Start(object sender, EventArgs e)
        {
            RegisterRoutes();
        }

        protected void Page_Load(object sender, EventArgs e)
        {

        }


        private void RegisterRoutes()
        {
            RouteTable.Routes.Add(new ServiceRoute("ExchSvc", new WebServiceHostFactory(), typeof(ExchSvc)));
            RouteTable.Routes.Add(new ServiceRoute("DstrSvc", new WebServiceHostFactory(), typeof(DstrSvc)));
        }
    }
}
