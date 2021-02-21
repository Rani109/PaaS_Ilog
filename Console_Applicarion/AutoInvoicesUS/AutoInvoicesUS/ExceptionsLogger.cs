using System;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using Dapper;

namespace AutoInvoicesUS
{
    public static class ExceptionsLogger
    {
        public static int? AddExceptionToDatabase(Exception ex, DateTime now)
        {
            if (ex == null)
                return null;

            int? exceptionID = null;
            try
            {
                exceptionID = SqlHelper.ExecuteScalarSP<int?>("sp_add_exception", new
                {
                    EXCEPTION_TYPE = ex.GetType().Name,
                    EXCEPTION_MESSAGE = ex.Message,
                    EXCEPTION_STACK_TRACE = ex.StackTrace,
                    EXCEPTION_STACK_FRAME = GetExceptionStackFrame(ex),
                    EXCEPTION_SOURCE = ex.Source,
                    EXCEPTION_METHOD = GetExceptionMethod(ex),
                    INNER_EXCEPTION_MESSAGE = (ex.InnerException != null ? ex.InnerException.Message : null),
                    INNER_EXCEPTION_METHOD = (ex.InnerException != null ? GetExceptionMethod(ex.InnerException) : null),
                    EXCEPTION_DATE = now,
                    USER_CODE = 0,
                    USER_NAME = ".",
                    IP = (string)null,
                    REQUEST_URL = (string)null,
                    PAGE_NAME = "AutoInvoicesUS.exe",
                    REFERER = "AutoInvoicesUS.exe",
                    FILTER_SEARCH_TEXT = (string)null,
                    POSTBACK_CONTROL = (string)null
                });
            }
            catch
            {
                exceptionID = null;
            }

            return exceptionID;
        }

        public static string GetExceptionStackFrame(Exception ex)
        {
            StackTrace st = new StackTrace(ex, true);
            StackFrame sf = null;
            StringBuilder sb = new StringBuilder();
            int lineNumber = 0;
            MethodBase method = null;
            Type reflectedType = null;
            for (int i = 0; i < st.FrameCount; i++)
            {
                sf = st.GetFrame(i);
                if (sf != null)
                {
                    method = sf.GetMethod();
                    if (method != null)
                    {
                        reflectedType = method.ReflectedType;
                        if (reflectedType != null)
                        {
                            sb.Append(reflectedType.FullName);
                            sb.Append(".");
                            sb.Append(method.Name);
                            sb.Append((method.MemberType == MemberTypes.Method ? "()" : string.Empty));
                            lineNumber = sf.GetFileLineNumber();
                            if (lineNumber > 0)
                            {
                                sb.Append(" at line ");
                                sb.Append(lineNumber);
                            }
                            sb.Append(Environment.NewLine);
                        }
                    }
                }
            }
            return sb.ToString();
        }

        public static string GetExceptionMethod(Exception ex)
        {
            MethodBase methodBase = ex.TargetSite;
            if (methodBase == null)
                return null;

            try
            {
                return
                    methodBase.ReflectedType.Name + "." +
                    methodBase.Name +
                    (methodBase.MemberType == MemberTypes.Method ? "()" : string.Empty);
            }
            catch
            {
                return methodBase.Name;
            }
        }
    }
}