using ADPF.API.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using System.Web.Http.Cors;

namespace ADPF.API
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            // Web API configuration and services

            // Web API routes
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );
        }

        public static HttpConfiguration Register()
        {
            // Web API configuration and services
            var config = new HttpConfiguration();
            //config.Formatters.Remove(config.Formatters.XmlFormatter);
            EnableCorsAttribute cors = new EnableCorsAttribute("*", "*", "GET,POST");
            config.EnableCors(cors);
            // Web API routes
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );

            //config.Routes.MapHttpRoute(//added by DWD
            //    name: "V2",
            //    routeTemplate: "api/V2/{controller}/{id}",
            //    defaults: new { id = RouteParameter.Optional }
            //);

            config.Filters.Add(new NotImplExceptionFilterAttribute());
            //config.MessageHandlers.Add(new LocalizationMessageHandler());

            //config.EnableCors();
            //CorsHttpConfigurationExtensions


            return config;
        }
    }

}