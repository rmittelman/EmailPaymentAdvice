﻿using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]
namespace Aimm.Logging
{
    public static class LogIt
    {
        static log4net.ILog _log = null;
        public static log4net.ILog Log
        {
            get
            {
                if (_log == null)
                    _log = log4net.LogManager.GetLogger("");

                return _log;
            }
        }
        
        public static string LogMethod(params object[] parameterValues)
        {
            var stackTrace = new StackTrace();
            StackFrame stackFrame = stackTrace.GetFrame(1);
            string message = "";
            ParameterInfo[] parameters = stackFrame.GetMethod().GetParameters();
            var parameterString = new StringBuilder();
            if (parameters.Length == parameterValues.Length)
            {
                for (int i = 0; i < parameterValues.Length; i++)
                    parameterString.AppendFormat("{0}: {1}, ", parameters[i].Name, parameterValues[i] ?? "");

                if (parameterString.Length > 0)
                    parameterString.Remove(parameterString.Length - 2, 2);

                message = $"-- {stackFrame.GetMethod().ReflectedType.Name}.{stackFrame.GetMethod().Name} ({parameterString.ToString()})";
            }
            else
            {
                message = $"-- {stackFrame.GetMethod().ReflectedType.Name}.{stackFrame.GetMethod().Name}";
            }

            Log.Debug(message);
            return message;
        }


        public static string LogDebug(string message, [CallerMemberName] string caller = null)
        {
            Log.Debug($"[{caller}] {message}");
            return message;
        }

        public static string LogInfo(string message, [CallerMemberName] string caller = null)
        {
            Log.Info($"[{caller}] {message}");
            return message;
        }

        public static string LogWarn(string message, [CallerMemberName] string caller = null)
        {
            Log.Warn($"[{caller}] {message}");
            return message;
        }

        public static string LogError(string message, [CallerMemberName] string caller = null)
        {
            Log.Error($"[{caller}] {message}");
            return message;
        }

        public static string GetAndLogMessage(Exception ex, [CallerFilePath] string filePath = null, [CallerMemberName] string caller = null)
        {
            string message = $"Exception in {filePath}:{caller}";
            string detailedMessage = message + ":\n" + MaskPassword(ex.Message);

            Log.Error(detailedMessage, ex);

            return detailedMessage;
        }

        static string MaskPassword(string text)
        {
            return Regex.Replace(text, "password[ =].*?[;$]", "password=...;", RegexOptions.IgnoreCase);
        }
    }
}
