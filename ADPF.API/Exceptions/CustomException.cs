using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using System.Web.Http.Filters;

namespace ADPF.API.Exceptions
{
    public class CustomException :Exception
    {
        public string ErrorCode { get; set; }

        public CustomException(string message, string code) : base(message)
        {
            this.ErrorCode = code;
        }
    }

    public class NotImplExceptionFilterAttribute : ExceptionFilterAttribute
    {
        public override void OnException(HttpActionExecutedContext context)
        {
            if (context.Exception is CustomException)
            {
                CustomException exception = (CustomException)context.Exception;
                HttpError error = new HttpError();
                error.Add("Message", exception.Message);
                error.Add("ExceptionMessage", exception.Message);
                error.Add("ExceptionCode", exception.ErrorCode);
                error.Add("ExceptionType", exception.Source);
                error.Add("StackTrace", exception.StackTrace);
                context.Response = context.Request.CreateErrorResponse(HttpStatusCode.InternalServerError, error);

            }
        }
    }

    public class ErrorResult

    {
        public string ErrorMessage { get; set; }

    }
}