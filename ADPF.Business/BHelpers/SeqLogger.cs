using ADPF.Utilities.Helpers;
using Serilog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADPF.Business.BHelpers
{
    public static class SeqLogger
    {
        static ILogger _loggerobj;
        public static ILogger LoggerProp
        {
            get
            {
                if (_loggerobj == null)
                {
                    _loggerobj = new LoggerConfiguration()
                           .WriteTo.Seq(ConfigurationManager.AppSettings["SEQUrl"].ToString())
                           .CreateLogger();
                }

                return _loggerobj;
            }
        }

        public static void LogError(Exception ex, string Parameter, string ControllerName, string ActionName, params object[] AdditionalValue)
        {
            LoggerProp.Error("Controller : {ControllerName} - Action : {ActionName} - Request : {Request} - ErrorMsg : {ErrorMessage} - Application : {Application} AdditionlValue : {AdditionalValue}",
                ControllerName, ActionName, Parameter, Logger.GetExceptionDetails(ex), "ADPF.API", AdditionalValue);
        }
        public static bool LogErrorMsg(string MsgTemplate, string ExceptionMsg, params object[] Values)
        {
            try
            {
                //  SwaggerConfiguration();
                LoggerProp.Error(new FieldAccessException(ExceptionMsg), MsgTemplate, Values);
                // _loggerobj.CloseAndFlush();
                return true;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

        }
    }
}
