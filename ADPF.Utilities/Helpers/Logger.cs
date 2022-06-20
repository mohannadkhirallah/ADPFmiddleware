﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADPF.Utilities.Helpers
{
    public static class Logger
    {
        public static string GetExceptionDetails(Exception exception)
        {
            StringBuilder sbExceptionMessage = new StringBuilder();

            sbExceptionMessage.Append("Exception Type" + Environment.NewLine);

            // Get the exception type

            sbExceptionMessage.Append(exception.GetType().Name);

            // Environment.NewLine writes new line character - \n

            sbExceptionMessage.Append(Environment.NewLine + Environment.NewLine);

            sbExceptionMessage.Append("Message" + Environment.NewLine);

            // Get the exception message

            sbExceptionMessage.Append(exception.Message + Environment.NewLine + Environment.NewLine);

            sbExceptionMessage.Append("Stack Trace" + Environment.NewLine);

            // Get the exception stack trace

            sbExceptionMessage.Append(exception.StackTrace + Environment.NewLine + Environment.NewLine);

            // Retrieve inner exception if any

            Exception innerException = exception.InnerException;

            // If inner exception exists

            while (innerException != null)
            {
                sbExceptionMessage.Append("Exception Type" + Environment.NewLine);

                sbExceptionMessage.Append(innerException.GetType().Name);

                sbExceptionMessage.Append(Environment.NewLine + Environment.NewLine);

                sbExceptionMessage.Append("Message" + Environment.NewLine);

                sbExceptionMessage.Append(innerException.Message + Environment.NewLine + Environment.NewLine);

                sbExceptionMessage.Append("Stack Trace" + Environment.NewLine);

                sbExceptionMessage.Append(innerException.StackTrace + Environment.NewLine + Environment.NewLine);

                // Retrieve inner exception if any

                innerException = innerException.InnerException;
            }

            return sbExceptionMessage.ToString();
        }

    }
}
