using Siemens.Engineering;
using Siemens.Engineering.HW;
using Siemens.Engineering.MC.Drives;
using Siemens.Engineering.MC.Drives.Enums;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace SDRhelper
{
    public class StartdriveHelper
    {
        #region Drive parameter handling

        #region Reading drive parameters for further operation

        #region Offline
        /// <summary>
        /// Returns a drive parameter for further operation or internal library use.
        /// <para>This method accesses parameters with indices</para>
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="parameter">Parameter number to be accessed</param>
        /// <param name="index">Parameter index to be accessed. -1 for parameters without indices</param>
        /// <returns>
        /// DriveParameter Object if parameter is accessible/exists.
        /// <para>Null if parameter is not accessible/exists.</para>
        /// </returns>
        public static DriveParameter GetParameter(DriveObject actDriveObject, int parameter, int index)
        {
            return actDriveObject.Parameters.Find(parameter, index);
        }

        /// <summary>
        /// Returns a drive parameter for further operation or internal library use.
        /// <para>This method accesses parameters without indices</para>
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="parameter">Parameter number to be accessed</param>
        /// <returns>
        /// DriveParameter Object if parameter is accessible/exists.
        /// <para>Null if parameter is not accessible/exists.</para>
        /// </returns>
        public static DriveParameter GetParameter(DriveObject actDriveObject, int parameter)
        {
            return actDriveObject.Parameters.Find(parameter, -1);
        }

        /// <summary>
        /// Returns a drive parameter for further operation or internal library use.
        /// <para>This method accesses parameters with indices and/or bits</para>
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="parameter">
        /// Parameter to be accessed without leading character
        /// <para>e.g. "63", "899.0", "1082[0]"</para>
        /// </param>
        /// <returns>
        /// DriveParameter Object if parameter is accessible/exists.
        /// <para>Null if parameter is not accessible/exists.</para>
        /// </returns>
        public static DriveParameter GetParameter(DriveObject actDriveObject, string parameter)
        {
            DriveParameter param = null;

            if (!parameter.ToLower().Contains("p") && !parameter.ToLower().Contains("r"))
            {
                //No 'p' in the parameter string; Add 'p' and try to access the parameter
                string paramString = "p" + parameter;
                param = GetParameterByString(actDriveObject, paramString);

                if (param == null)
                {
                    //Parameter not found; Add 'r' instead and try to access the parameter
                    paramString = "r" + parameter;
                    param = GetParameterByString(actDriveObject, paramString);
                }
            }
            else
            {
                //Parameter contains 'p' or 'r'; try to access the parameter
                param = GetParameterByString(actDriveObject, parameter);
            }
            return param;
        }

        /// <summary>
        /// Returns a drive parameter for further operation or internal library use.
        /// <para>This method accesses parameters with indices and/or bits</para>
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="parameter">
        /// Parameter to be accessed with leading character
        /// <para>e.g. "r63", "r899.0", "p1082[0]"</para>
        /// </param>
        /// <returns>
        /// DriveParameter Object if parameter is accessible/exists.
        /// <para>Null if parameter is not accessible/exists.</para>
        /// </returns>
        private static DriveParameter GetParameterByString(DriveObject actDriveObject, string parameter)
        {
            DriveParameter param = null;
            if (!parameter.Contains('.'))
            {
                //No '.' in the parameter string; try to access the parameter directly
                param = actDriveObject.Parameters.Find(parameter);
            }
            else
            {
                //Found '.' in the parameter string; split string and try to access the parameter
                string[] parameters = parameter.Split('.');
                param = actDriveObject.Parameters.Find(parameters[0]);
                if (param != null)
                {
                    //Accessing the desired bit in the parameter
                    param = param.Bits.Find(parameter);
                }
            }
            if (param == null && parameter.Contains('['))
            {
                //Parameter with multiple indices; split the parameter string and try to access the parameter
                StringBuilder builder = new StringBuilder();
                builder.Append(parameter.Split('[')[0]);
                if (parameter.Split(']').Count() == 2)
                {
                    builder.Append(parameter.Split(']')[1]);
                }
                param = GetParameter(actDriveObject, builder.ToString());
            }
            return param;
        }
        #endregion

        #region Online

        /// <summary>
        /// Returns a drive parameter for further operation or internal library use.
        /// <para>This method accesses parameters with indices</para>
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <param name="parameter">Parameter number to be accessed</param>
        /// <param name="index">Parameter index to be accessed. -1 for parameters without indices</param>
        /// <returns>
        /// DriveParameter Object if parameter is accessible/exists.
        /// <para>Null if parameter is not accessible/exists.</para>
        /// </returns>
        public static DriveParameter GetParameter(OnlineDriveObject actDriveObject, int parameter, int index)
        {
            return actDriveObject.Parameters.Find(parameter, index);
        }

        /// <summary>
        /// Returns a drive parameter for further operation or internal library use.
        /// <para>This method accesses parameters without indices</para>
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <param name="parameter">Parameter number to be accessed</param>
        /// <returns>
        /// DriveParameter Object if parameter is accessible/exists.
        /// <para>Null if parameter is not accessible/exists.</para>
        /// </returns>
        public static DriveParameter GetParameter(OnlineDriveObject actDriveObject, int parameter)
        {
            return actDriveObject.Parameters.Find(parameter, -1);
        }

        /// <summary>
        /// Returns a drive parameter for further operation or internal library use.
        /// <para>This method accesses parameters with indices and/or bits</para>
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <param name="parameter">
        /// Parameter to be accessed without leading character
        /// <para>e.g. "63", "899.0", "1082[0]"</para>
        /// </param>
        /// <returns>
        /// DriveParameter Object if parameter is accessible/exists.
        /// <para>Null if parameter is not accessible/exists.</para>
        /// </returns>
        public static DriveParameter GetParameter(OnlineDriveObject actDriveObject, string parameter)
        {
            DriveParameter param = null;

            if (!parameter.ToLower().Contains("p") && !parameter.ToLower().Contains("r"))
            {
                //No 'p' in the parameter string; Add 'p' and try to access the parameter
                string paramString = "p" + parameter;
                param = GetParameterByString(actDriveObject, paramString);

                if (param == null)
                {
                    //Parameter not found; Add 'r' instead and try to access the parameter
                    paramString = "r" + parameter;
                    param = GetParameterByString(actDriveObject, paramString);
                }
            }
            else
            {
                //Parameter contains 'p' or 'r'; try to access the parameter
                param = GetParameterByString(actDriveObject, parameter);
            }
            return param;
        }

        /// <summary>
        /// Returns a drive parameter for further operation or internal library use.
        /// <para>This method accesses parameters with indices and/or bits</para>
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <param name="parameter">
        /// Parameter to be accessed with leading character
        /// <para>e.g. "r63", "r899.0", "p1082[0]"</para>
        /// </param>
        /// <returns>
        /// DriveParameter Object if parameter is accessible/exists.
        /// <para>Null if parameter is not accessible/exists.</para>
        /// </returns>
        private static DriveParameter GetParameterByString(OnlineDriveObject actDriveObject, string parameter)
        {
            DriveParameter param = null;
            if (!parameter.Contains('.'))
            {
                //No '.' in the parameter string; try to access the parameter directly
                param = actDriveObject.Parameters.Find(parameter);
            }
            else
            {
                //Found '.' in the parameter string; split string and try to access the parameter
                string[] parameters = parameter.Split('.');
                param = actDriveObject.Parameters.Find(parameters[0]);
                if (param != null)
                {
                    //Accessing the desired bit in the parameter
                    param = param.Bits.Find(parameter);
                }
            }
            if (param == null && parameter.Contains('['))
            {
                //Parameter with multiple indices; split the parameter string and try to access the parameter
                StringBuilder builder = new StringBuilder();
                builder.Append(parameter.Split('[')[0]);
                if (parameter.Split(']').Count() == 2)
                {
                    builder.Append(parameter.Split(']')[1]);
                }
                param = GetParameter(actDriveObject, builder.ToString());
            }
            return param;
        }

        #endregion

        #endregion

        #region Reading drive parameters for diagnostic

        #region Offline

        /// <summary>
        /// Returns the value of a drive parameter for diagnostic use.
        /// <para>This method accesses parameters with indices</para>
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="parameter">Parameter number to be accessed</param>
        /// <param name="index">Parameter index to be accessed. -1 for parameters without indices</param>
        /// <returns>
        /// Parameter value as string if parameter is accessible/exists.
        /// <para>Null if parameter is not accessible/exists.</para>
        /// </returns>
        public static string ReadParameterValue(DriveObject actDriveObject, int parameter, int index)
        {
            string value = null;
            DriveParameter param = null;
            param = GetParameter(actDriveObject, parameter, index);

            if (param != null)
            {
                if (param.Value.GetType() == param.GetType())
                {
                    DriveParameter param2 = param.Value as DriveParameter;
                    value = param2.Name;
                }
                else
                {
                    value = param.Value.ToString();
                }
            }
            return value;
        }

        /// <summary>
        /// Returns the value of a drive parameter for diagnostic use.
        /// <para>This method accesses parameters without indices</para>
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="parameter">Parameter number to be accessed</param>
        /// <returns>
        /// Parameter value as string if parameter is accessible/exists.
        /// <para>Null if parameter is not accessible/exists.</para>
        /// </returns>
        public static string ReadParameterValue(DriveObject actDriveObject, int parameter)
        {
            string value = null;
            DriveParameter param = null;
            param = GetParameter(actDriveObject, parameter);

            if (param != null)
            {
                if (param.Value.GetType() == param.GetType())
                {
                    DriveParameter param2 = param.Value as DriveParameter;
                    value = param2.Name;
                }
                else
                {
                    value = param.Value.ToString();
                }
            }
            return value;
        }

        /// <summary>
        /// Returns the value of a drive parameter for diagnostic use.
        /// <para>This method accesses parameters with indices and/or bits</para>
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="parameter">
        /// Parameter to be accessed without leading character
        /// <para>e.g. "63", "899.0", "1082[0]"</para>
        /// </param>
        /// <returns>
        /// DriveParameter Object if parameter is accessible/exists.
        /// <para>Null if parameter is not accessible/exists.</para>
        /// </returns>
        public static string ReadParameterValue(DriveObject actDriveObject, string parameter)
        {
            string value = null;
            DriveParameter param = null;
            param = GetParameter(actDriveObject, parameter);

            if (param != null)
            {
                if (param.Value.GetType() == param.GetType())
                {
                    DriveParameter param2 = param.Value as DriveParameter;
                    value = param2.Name;
                }
                else
                {
                    value = param.Value.ToString();
                }
            }
            return value;
        }

        #endregion

        #region Online


        /// <summary>
        /// Returns the value of a drive parameter for diagnostic use.
        /// <para>This method accesses parameters with indices</para>
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <param name="parameter">Parameter number to be accessed</param>
        /// <param name="index">Parameter index to be accessed. -1 for parameters without indices</param>
        /// <returns>
        /// Parameter value as string if parameter is accessible/exists.
        /// <para>Null if parameter is not accessible/exists.</para>
        /// </returns>
        public static string ReadParameterValue(OnlineDriveObject actDriveObject, int parameter, int index)
        {
            string value = null;
            DriveParameter param = null;
            param = GetParameter(actDriveObject, parameter, index);

            if (param != null)
            {
                if (param.Value.GetType() == param.GetType())
                {
                    DriveParameter param2 = param.Value as DriveParameter;
                    value = param2.Name;
                }
                else
                {
                    value = param.Value.ToString();
                }
            }
            return value;
        }

        /// <summary>
        /// Returns the value of a drive parameter for diagnostic use.
        /// <para>This method accesses parameters without indices</para>
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <param name="parameter">Parameter number to be accessed</param>
        /// <returns>
        /// Parameter value as string if parameter is accessible/exists.
        /// <para>Null if parameter is not accessible/exists.</para>
        /// </returns>
        public static string ReadParameterValue(OnlineDriveObject actDriveObject, int parameter)
        {
            string value = null;
            DriveParameter param = null;
            param = GetParameter(actDriveObject, parameter);

            if (param != null)
            {
                if (param.Value.GetType() == param.GetType())
                {
                    DriveParameter param2 = param.Value as DriveParameter;
                    value = param2.Name;
                }
                else
                {
                    value = param.Value.ToString();
                }
            }
            return value;
        }

        /// <summary>
        /// Returns the value of a drive parameter for diagnostic use.
        /// <para>This method accesses parameters with indices and/or bits</para>
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <param name="parameter">
        /// Parameter to be accessed without leading character
        /// <para>e.g. "63", "899.0", "1082[0]"</para>
        /// </param>
        /// <returns>
        /// DriveParameter Object if parameter is accessible/exists.
        /// <para>Null if parameter is not accessible/exists.</para>
        /// </returns>
        public static string ReadParameterValue(OnlineDriveObject actDriveObject, string parameter)
        {
            string value = null;
            DriveParameter param = null;
            param = GetParameter(actDriveObject, parameter);

            if (param != null)
            {
                if (param.Value.GetType() == param.GetType())
                {
                    DriveParameter param2 = param.Value as DriveParameter;
                    value = param2.Name;
                }
                else
                {
                    value = param.Value.ToString();
                }
            }
            return value;
        }

        #endregion

        #endregion

        #region Setting integer parameter values

        #region Offline

        /// <summary>
        /// Sets parameter to an integer value
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="parameter">Number of parameter to be written</param>
        /// <param name="index">Index of parameter to be written</param>
        /// <param name="value">Parameter value as integer to be written</param>
        /// <returns>true if the value was written successfully
        /// <para>false if the value was not written</para></returns>
        public static bool SetParameter(DriveObject actDriveObject, int parameter, int index, int value)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter, index);
            bool success = false;

            //Parameter exists
            if (param != null)
            {
                if (param.Name.ToLower().StartsWith("p"))
                {
                    //Set parameter value
                    try
                    {
                        param.Value = value;
                    }
                    catch (Exception e)
                    { }

                    //Read back parameter value
                    if (Convert.ToInt32(param.Value) == value)
                    {
                        success = true;
                    }
                }
            }
            return success;
        }

        /// <summary>
        /// Sets parameter to an integer value
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="parameter">Number of parameter to be written</param>
        /// <param name="value">Parameter value as integer to be written</param>
        /// <returns>true if the value was written successfully
        /// <para>false if the value was not written</para></returns>
        public static bool SetParameter(DriveObject actDriveObject, int parameter, int value)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter);
            bool success = false;

            //Parameter exists
            if (param != null)
            {
                if (param.Name.ToLower().StartsWith("p"))
                {
                    //Set parameter value
                    try
                    {
                        param.Value = value;
                    }
                    catch (Exception e)
                    { }

                    //Read back parameter value
                    if (Convert.ToInt32(param.Value) == value)
                    {
                        success = true;
                    }
                }
            }
            return success;
        }

        /// <summary>
        /// Sets parameter to an integer value
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="parameter">string representation of parameter to be written
        /// <para>e.g. "2000" or "1082[0]". No leading "p" required.</para></param>
        /// <param name="value">Parameter value as integer to be written</param>
        /// <returns>true if the value was written successfully
        /// <para>false if the value was not written</para></returns>
        public static bool SetParameter(DriveObject actDriveObject, string parameter, int value)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter);
            bool success = false;

            //Parameter exists
            if (param != null)
            {
                if (param.Name.ToLower().StartsWith("p"))
                {
                    //Set parameter value
                    try
                    {
                        param.Value = value;
                    }
                    catch (Exception e)
                    { }

                    //Read back parameter value
                    if (Convert.ToInt32(param.Value) == value)
                    {
                        success = true;
                    }
                }
            }
            return success;
        }

        #endregion

        #region Online

        /// <summary>
        /// Sets parameter to an integer value
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <param name="parameter">Number of parameter to be written</param>
        /// <param name="index">Index of parameter to be written</param>
        /// <param name="value">Parameter value as integer to be written</param>
        /// <returns>true if the value was written successfully
        /// <para>false if the value was not written</para></returns>
        public static bool SetParameter(OnlineDriveObject actDriveObject, int parameter, int index, int value)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter, index);
            bool success = false;

            //Parameter exists
            if (param != null)
            {
                if (param.Name.ToLower().StartsWith("p"))
                {
                    //Set parameter value
                    try
                    {
                        param.Value = value;
                    }
                    catch (Exception e)
                    { }

                    //Read back parameter value
                    if (Convert.ToInt32(param.Value) == value)
                    {
                        success = true;
                    }
                }
            }
            return success;
        }

        /// <summary>
        /// Sets parameter to an integer value
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <param name="parameter">Number of parameter to be written</param>
        /// <param name="value">Parameter value as integer to be written</param>
        /// <returns>true if the value was written successfully
        /// <para>false if the value was not written</para></returns>
        public static bool SetParameter(OnlineDriveObject actDriveObject, int parameter, int value)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter);
            bool success = false;

            //Parameter exists
            if (param != null)
            {
                if (param.Name.ToLower().StartsWith("p"))
                {
                    //Set parameter value
                    try
                    {
                        param.Value = value;
                    }
                    catch (Exception e)
                    { }

                    //Read back parameter value
                    if (Convert.ToInt32(param.Value) == value)
                    {
                        success = true;
                    }
                }
            }
            return success;
        }

        /// <summary>
        /// Sets parameter to an integer value
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <param name="parameter">string representation of parameter to be written
        /// <para>e.g. "2000" or "1082[0]". No leading "p" required.</para></param>
        /// <param name="value">Parameter value as integer to be written</param>
        /// <returns>true if the value was written successfully
        /// <para>false if the value was not written</para></returns>
        public static bool SetParameter(OnlineDriveObject actDriveObject, string parameter, int value)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter);
            bool success = false;

            //Parameter exists
            if (param != null)
            {
                if (param.Name.ToLower().StartsWith("p"))
                {
                    //Set parameter value
                    try
                    {
                        param.Value = value;
                    }
                    catch (Exception e)
                    { }

                    //Read back parameter value
                    if (Convert.ToInt32(param.Value) == value)
                    {
                        success = true;
                    }
                }
            }
            return success;
        }

        #endregion

        #endregion

        #region Setting double parameter values

        #region Offline

        /// <summary>
        /// Sets parameter to a double value
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="parameter">Number of parameter to be written</param>
        /// <param name="index">Index of parameter to be written</param>
        /// <param name="value">Parameter value as double to be written</param>
        /// <returns>true if the value was written successfully
        /// <para>false if the value was not written</para></returns>
        public static bool SetParameter(DriveObject actDriveObject, int parameter, int index, double value)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter, index);
            bool success = false;

            //Parameter exists
            if (param != null)
            {
                if (param.Name.ToLower().StartsWith("p"))
                {
                    //Set parameter value
                    try
                    {
                        param.Value = value;
                    }
                    catch (Exception e)
                    { }

                    //Double conversion in inaccurate, compare strings
                    //Change decimal separator to '.'
                    string stringVal = value.ToString().Replace(",", ".");
                    //Read back parameter value
                    string stringReadVal = param.Value.ToString().Replace(",",".");
                    
                    //Compare if desired value was set
                    if (stringReadVal == stringVal)
                    {
                        success = true;
                    }
                }
            }
            return success;
        }

        /// <summary>
        /// Sets parameter to a double value
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="parameter">Number of parameter to be written</param>
        /// <param name="value">Parameter value as double to be written</param>
        /// <returns>true if the value was written successfully
        /// <para>false if the value was not written</para></returns>
        public static bool SetParameter(DriveObject actDriveObject, int parameter, double value)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter);
            bool success = false;

            //Parameter exists
            if (param != null)
            {
                if (param.Name.ToLower().StartsWith("p"))
                {
                    //Set parameter value
                    try
                    {
                        param.Value = value;
                    }
                    catch (Exception e)
                    { }

                    //Double conversion in inaccurate, compare strings
                    //Change decimal separator to '.'
                    string stringVal = value.ToString().Replace(",", ".");
                    //Read back parameter value
                    string stringReadVal = param.Value.ToString().Replace(",", ".");

                    //Compare if desired value was set
                    if (stringReadVal == stringVal)
                    {
                        success = true;
                    }
                }
            }
            return success;
        }

        /// <summary>
        /// Sets parameter to a double value
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="parameter">string representation of parameter to be written
        /// <para>e.g. "2000" or "1082[0]". No leading "p" required.</para></param>
        /// <param name="value">Parameter value as double to be written</param>
        /// <returns>true if the value was written successfully
        /// <para>false if the value was not written</para></returns>
        public static bool SetParameter(DriveObject actDriveObject, string parameter, double value)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter);
            bool success = false;

            //Parameter exists
            if (param != null)
            {
                if (param.Name.ToLower().StartsWith("p"))
                {
                    //Set parameter value
                    try
                    {
                        param.Value = value;
                    }
                    catch (Exception e)
                    { }

                    //Double conversion in inaccurate, compare strings
                    //Change decimal separator to '.'
                    string stringVal = value.ToString().Replace(",", ".");
                    //Read back parameter value
                    string stringReadVal = param.Value.ToString().Replace(",", ".");

                    //Compare if desired value was set
                    if (stringReadVal == stringVal)
                    {
                        success = true;
                    }
                }
            }
            return success;
        }

        #endregion

        #region Online


        /// <summary>
        /// Sets parameter to a double value
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <param name="parameter">Number of parameter to be written</param>
        /// <param name="index">Index of parameter to be written</param>
        /// <param name="value">Parameter value as double to be written</param>
        /// <returns>true if the value was written successfully
        /// <para>false if the value was not written</para></returns>
        public static bool SetParameter(OnlineDriveObject actDriveObject, int parameter, int index, double value)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter, index);
            bool success = false;

            //Parameter exists
            if (param != null)
            {
                if (param.Name.ToLower().StartsWith("p"))
                {
                    //Set parameter value
                    try
                    {
                        param.Value = value;
                    }
                    catch (Exception e)
                    { }

                    //Double conversion in inaccurate, compare strings
                    //Change decimal separator to '.'
                    string stringVal = value.ToString().Replace(",", ".");
                    //Read back parameter value
                    string stringReadVal = param.Value.ToString().Replace(",", ".");

                    //Compare if desired value was set
                    if (stringReadVal == stringVal)
                    {
                        success = true;
                    }
                }
            }
            return success;
        }

        /// <summary>
        /// Sets parameter to a double value
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <param name="parameter">Number of parameter to be written</param>
        /// <param name="value">Parameter value as double to be written</param>
        /// <returns>true if the value was written successfully
        /// <para>false if the value was not written</para></returns>
        public static bool SetParameter(OnlineDriveObject actDriveObject, int parameter, double value)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter);
            bool success = false;

            //Parameter exists
            if (param != null)
            {
                if (param.Name.ToLower().StartsWith("p"))
                {
                    //Set parameter value
                    try
                    {
                        param.Value = value;
                    }
                    catch (Exception e)
                    { }

                    //Double conversion in inaccurate, compare strings
                    //Change decimal separator to '.'
                    string stringVal = value.ToString().Replace(",", ".");
                    //Read back parameter value
                    string stringReadVal = param.Value.ToString().Replace(",", ".");

                    //Compare if desired value was set
                    if (stringReadVal == stringVal)
                    {
                        success = true;
                    }
                }
            }
            return success;
        }

        /// <summary>
        /// Sets parameter to a double value
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <param name="parameter">string representation of parameter to be written
        /// <para>e.g. "2000" or "1082[0]". No leading "p" required.</para></param>
        /// <param name="value">Parameter value as double to be written</param>
        /// <returns>true if the value was written successfully
        /// <para>false if the value was not written</para></returns>
        public static bool SetParameter(OnlineDriveObject actDriveObject, string parameter, double value)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter);
            bool success = false;

            //Parameter exists
            if (param != null)
            {
                if (param.Name.ToLower().StartsWith("p"))
                {
                    //Set parameter value
                    try
                    {
                        param.Value = value;
                    }
                    catch (Exception e)
                    { }

                    //Double conversion in inaccurate, compare strings
                    //Change decimal separator to '.'
                    string stringVal = value.ToString().Replace(",", ".");
                    //Read back parameter value
                    string stringReadVal = param.Value.ToString().Replace(",", ".");

                    //Compare if desired value was set
                    if (stringReadVal == stringVal)
                    {
                        success = true;
                    }
                }
            }
            return success;
        }

        #endregion

        #endregion

        #region Setting string parameter values

        #region Offline

        /// <summary>
        /// Sets parameter to a value represented as string
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="parameter">Number of parameter to be written</param>
        /// <param name="index">Index of parameter to be written</param>
        /// <param name="value">Parameter value as string to be written</param>
        /// <returns>true if the value was written successfully
        /// <para>false if the value was not written</para></returns>
        public static bool SetParameter(DriveObject actDriveObject, int parameter, int index, string value)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter, index);
            bool success = false;

            //Parameter exists
            if (param != null)
            {
                if (param.Name.ToLower().StartsWith("p"))
                {
                    //Set parameter value
                    try
                    {
                        param.Value = value;
                    }
                    catch (Exception e)
                    { }

                    //Change decimal separator to '.'
                    string stringVal = value.ToString().Replace(",", ".");
                    //Read back parameter value
                    string stringReadVal = param.Value.ToString().Replace(",", ".");

                    //When writing e.g. "100.00", the read value is "100"
                    //Usung Regex to remove any amount of decimal places consisting of '0' (10.00 --> 10)
                    stringVal = Regex.Replace(stringVal, @"\.0+$", "");
                    //Using Regex to remove any trailing '0' (10.10 --> 10.1)
                    stringVal = Regex.Replace(stringVal, @"(?<=[0-9]+\.[0-9]*[1-9])0", "");

                    //Compare if desired value was set
                    if (stringReadVal == stringVal)
                    {
                        success = true;
                    }
                }
            }
            return success;
        }

        /// <summary>
        /// Sets parameter to a value represented as string
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="parameter">Number of parameter to be written</param>
        /// <param name="value">Parameter value as string to be written</param>
        /// <returns>true if the value was written successfully
        /// <para>false if the value was not written</para></returns>
        public static bool SetParameter(DriveObject actDriveObject, int parameter, string value)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter);
            bool success = false;

            //Parameter exists
            if (param != null)
            {
                if (param.Name.ToLower().StartsWith("p"))
                {
                    //Set parameter value
                    try
                    {
                        param.Value = value;
                    }
                    catch (Exception e)
                    { }

                    //Change decimal separator to '.'
                    string stringVal = value.ToString().Replace(",", ".");
                    //Read back parameter value
                    string stringReadVal = param.Value.ToString().Replace(",", ".");

                    //When writing e.g. "100.00", the read value is "100"
                    //Usung Regex to remove any amount of decimal places consisting of '0' (10.00 --> 10)
                    stringVal = Regex.Replace(stringVal, @"\.0+$", "");
                    //Using Regex to remove any trailing '0' (10.10 --> 10.1)
                    stringVal = Regex.Replace(stringVal, @"(?<=[0-9]+\.[0-9]*[1-9])0", "");

                    //Compare if desired value was set
                    if (stringReadVal == stringVal)
                    {
                        success = true;
                    }
                }
            }
            return success;
        }

        /// <summary>
        /// Sets parameter to a value represented as string
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="parameter">string representation of parameter to be written
        /// <para>e.g. "2000" or "1082[0]". No leading "p" required.</para></param>
        /// <param name="value">Parameter value as string to be written</param>
        /// <returns>true if the value was written successfully
        /// <para>false if the value was not written</para></returns>
        public static bool SetParameter(DriveObject actDriveObject, string parameter, string value)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter);
            bool success = false;

            //Parameter exists
            if (param != null)
            {
                if (param.Name.ToLower().StartsWith("p"))
                {
                    //Set parameter value
                    try
                    {
                        param.Value = value;
                    }
                    catch (Exception e)
                    { }

                    //Change decimal separator to '.'
                    string stringVal = value.ToString().Replace(",", ".");
                    //Read back parameter value
                    string stringReadVal = param.Value.ToString().Replace(",", ".");

                    //When writing e.g. "100.00", the read value is "100"
                    //Usung Regex to remove any amount of decimal places consisting of '0' (10.00 --> 10)
                    stringVal = Regex.Replace(stringVal, @"\.0+$", "");
                    //Using Regex to remove any trailing '0' (10.10 --> 10.1)
                    stringVal = Regex.Replace(stringVal, @"(?<=[0-9]+\.[0-9]*[1-9])0", "");

                    //Compare if desired value was set
                    if (stringReadVal == stringVal)
                    {
                        success = true;
                    }
                }
            }
            return success;
        }

        #endregion

        #region Online


        /// <summary>
        /// Sets parameter to a value represented as string
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <param name="parameter">Number of parameter to be written</param>
        /// <param name="index">Index of parameter to be written</param>
        /// <param name="value">Parameter value as string to be written</param>
        /// <returns>true if the value was written successfully
        /// <para>false if the value was not written</para></returns>
        public static bool SetParameter(OnlineDriveObject actDriveObject, int parameter, int index, string value)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter, index);
            bool success = false;

            //Parameter exists
            if (param != null)
            {
                if (param.Name.ToLower().StartsWith("p"))
                {
                    //Set parameter value
                    try
                    {
                        param.Value = value;
                    }
                    catch (Exception e)
                    { }

                    //Change decimal separator to '.'
                    string stringVal = value.ToString().Replace(",", ".");
                    //Read back parameter value
                    string stringReadVal = param.Value.ToString().Replace(",", ".");

                    //When writing e.g. "100.00", the read value is "100"
                    //Usung Regex to remove any amount of decimal places consisting of '0' (10.00 --> 10)
                    stringVal = Regex.Replace(stringVal, @"\.0+$", "");
                    //Using Regex to remove any trailing '0' (10.10 --> 10.1)
                    stringVal = Regex.Replace(stringVal, @"(?<=[0-9]+\.[0-9]*[1-9])0", "");

                    //Compare if desired value was set
                    if (stringReadVal == stringVal)
                    {
                        success = true;
                    }
                }
            }
            return success;
        }

        /// <summary>
        /// Sets parameter to a value represented as string
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <param name="parameter">Number of parameter to be written</param>
        /// <param name="value">Parameter value as string to be written</param>
        /// <returns>true if the value was written successfully
        /// <para>false if the value was not written</para></returns>
        public static bool SetParameter(OnlineDriveObject actDriveObject, int parameter, string value)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter);
            bool success = false;

            //Parameter exists
            if (param != null)
            {
                if (param.Name.ToLower().StartsWith("p"))
                {
                    //Set parameter value
                    try
                    {
                        param.Value = value;
                    }
                    catch (Exception e)
                    { }

                    //Change decimal separator to '.'
                    string stringVal = value.ToString().Replace(",", ".");
                    //Read back parameter value
                    string stringReadVal = param.Value.ToString().Replace(",", ".");

                    //When writing e.g. "100.00", the read value is "100"
                    //Usung Regex to remove any amount of decimal places consisting of '0' (10.00 --> 10)
                    stringVal = Regex.Replace(stringVal, @"\.0+$", "");
                    //Using Regex to remove any trailing '0' (10.10 --> 10.1)
                    stringVal = Regex.Replace(stringVal, @"(?<=[0-9]+\.[0-9]*[1-9])0", "");

                    //Compare if desired value was set
                    if (stringReadVal == stringVal)
                    {
                        success = true;
                    }
                }
            }
            return success;
        }

        /// <summary>
        /// Sets parameter to a value represented as string
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <param name="parameter">string representation of parameter to be written
        /// <para>e.g. "2000" or "1082[0]". No leading "p" required.</para></param>
        /// <param name="value">Parameter value as string to be written</param>
        /// <returns>true if the value was written successfully
        /// <para>false if the value was not written</para></returns>
        public static bool SetParameter(OnlineDriveObject actDriveObject, string parameter, string value)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter);
            bool success = false;

            //Parameter exists
            if (param != null)
            {
                if (param.Name.ToLower().StartsWith("p"))
                {
                    //Set parameter value
                    try
                    {
                        param.Value = value;
                    }
                    catch (Exception e)
                    { }

                    //Change decimal separator to '.'
                    string stringVal = value.ToString().Replace(",", ".");
                    //Read back parameter value
                    string stringReadVal = param.Value.ToString().Replace(",", ".");

                    //When writing e.g. "100.00", the read value is "100"
                    //Usung Regex to remove any amount of decimal places consisting of '0' (10.00 --> 10)
                    stringVal = Regex.Replace(stringVal, @"\.0+$", "");
                    //Using Regex to remove any trailing '0' (10.10 --> 10.1)
                    stringVal = Regex.Replace(stringVal, @"(?<=[0-9]+\.[0-9]*[1-9])0", "");

                    //Compare if desired value was set
                    if (stringReadVal == stringVal)
                    {
                        success = true;
                    }
                }
            }
            return success;
        }

        #endregion

        #endregion

        #region Connecting BiCo parameters

        #region Offline

        /// <summary>
        /// Connects BiCo parameters on the same drive object
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="parameter">Parameter which represents a BiCo input e.g. "840[0]"</param>
        /// <param name="setToParameter">Parameter which represents a BiCo output to be connected to the BiCo input e.g. "r2090.0"</param>
        /// <returns>true if the parameters were connected successfully
        /// <para>false if the parameters were not connected</para></returns>
        public static bool ConnectParameter(DriveObject actDriveObject, string parameter, string setToParameter)
        {
            DriveParameter inputParam = GetParameter(actDriveObject, parameter);
            DriveParameter connectedParam = GetParameter(actDriveObject, setToParameter);

            //Both parameters can be accessed
            if (inputParam != null && connectedParam != null)
            {
                //Connect parameters
                try
                {
                    inputParam.Value = connectedParam;
                }
                catch (Exception e)
                { }

                DriveParameter checkParam = null;
                //Read back connection on input parameter
                try
                {
                    checkParam = (DriveParameter)GetParameter(actDriveObject, parameter).Value;
                }
                catch (Exception e)
                { }
                
                if (checkParam != null)
                {
                    string chekParamString = checkParam.Name;
                    //Checking if the connected parameter matches the input string
                    //Chekcing for equality is not possible; input string does not require leading 'p' or 'r'
                    if (chekParamString.Contains(setToParameter))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Connects BiCo parameters across drive objects on the same device
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the BiCo input is found</param>
        /// <param name="parameter">Parameter which represents a BiCo input e.g. "840[0]"</param>
        /// <param name="connectedDriveObject">DriveObject on which the BiCo output is found</param>
        /// <param name="setToParameter">Parameter which represents a BiCo output to be connected to the BiCo input e.g. "r2090.0"</param>
        /// <returns>true if the parameters were connected successfully
        /// <para>false if the parameters were not connected</para></returns>
        public static bool ConnectParameter(DriveObject actDriveObject, string parameter, DriveObject connectedDriveObject, string setToParameter)
        {
            DriveParameter inputParam = GetParameter(actDriveObject, parameter);
            DriveParameter connectedParam = GetParameter(connectedDriveObject, setToParameter);

            //Both parameters can be accessed
            if (inputParam != null && connectedParam != null)
            {
                //Connect parameters
                try
                {
                    inputParam.Value = connectedParam;
                }
                catch (Exception e)
                { }

                DriveParameter checkParam = null;
                //Read back connection on input parameter
                try
                {
                    checkParam = (DriveParameter)GetParameter(actDriveObject, parameter).Value;
                }
                catch (Exception e)
                { }

                if (checkParam != null)
                {
                    string chekParamString = checkParam.Name;
                    //Checking if the connected parameter matches the input string
                    //Chekcing for equality is not possible; input string does not require leading 'p' or 'r'
                    if (chekParamString.Contains(setToParameter))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        #endregion

        #region Online

        /// <summary>
        /// Connects BiCo parameters on the same drive object
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <param name="parameter">Parameter which represents a BiCo input e.g. "840[0]"</param>
        /// <param name="setToParameter">Parameter which represents a BiCo output to be connected to the BiCo input e.g. "r2090.0"</param>
        /// <returns>true if the parameters were connected successfully
        /// <para>false if the parameters were not connected</para></returns>
        public static bool ConnectParameter(OnlineDriveObject actDriveObject, string parameter, string setToParameter)
        {
            DriveParameter inputParam = GetParameter(actDriveObject, parameter);
            DriveParameter connectedParam = GetParameter(actDriveObject, setToParameter);

            //Both parameters can be accessed
            if (inputParam != null && connectedParam != null)
            {
                //Connect parameters
                try
                {
                    inputParam.Value = connectedParam;
                }
                catch (Exception e)
                { }

                DriveParameter checkParam = null;
                //Read back connection on input parameter
                try
                {
                    checkParam = (DriveParameter)GetParameter(actDriveObject, parameter).Value;
                }
                catch (Exception e)
                { }

                if (checkParam != null)
                {
                    string chekParamString = checkParam.Name;
                    //Checking if the connected parameter matches the input string
                    //Chekcing for equality is not possible; input string does not require leading 'p' or 'r'
                    if (chekParamString.Contains(setToParameter))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Connects BiCo parameters across drive objects on the same device
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the BiCo input is found</param>
        /// <param name="parameter">Parameter which represents a BiCo input e.g. "840[0]"</param>
        /// <param name="connectedDriveObject">OnlineDriveObject on which the BiCo output is found</param>
        /// <param name="setToParameter">Parameter which represents a BiCo output to be connected to the BiCo input e.g. "r2090.0"</param>
        /// <returns>true if the parameters were connected successfully
        /// <para>false if the parameters were not connected</para></returns>
        public static bool ConnectParameter(OnlineDriveObject actDriveObject, string parameter, OnlineDriveObject connectedDriveObject, string setToParameter)
        {
            DriveParameter inputParam = GetParameter(actDriveObject, parameter);
            DriveParameter connectedParam = GetParameter(connectedDriveObject, setToParameter);

            //Both parameters can be accessed
            if (inputParam != null && connectedParam != null)
            {
                //Connect parameters
                try
                {
                    inputParam.Value = connectedParam;
                }
                catch (Exception e)
                { }

                DriveParameter checkParam = null;
                //Read back connection on input parameter
                try
                {
                    checkParam = (DriveParameter)GetParameter(actDriveObject, parameter).Value;
                }
                catch (Exception e)
                { }

                if (checkParam != null)
                {
                    string chekParamString = checkParam.Name;
                    //Checking if the connected parameter matches the input string
                    //Chekcing for equality is not possible; input string does not require leading 'p' or 'r'
                    if (chekParamString.Contains(setToParameter))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        #endregion

        #endregion

        #region Parameter limit handling

        #region Offline

        /// <summary>
        /// Opens a TIA project
        /// </summary>
        /// <param name="actDriveObject">Path of TIA project to be opened</param>
        /// <param name="parameter">Parameter where the opened project is stored</param>
        /// <param name="selectLimit">true: Return upper parameter limit
        /// <para>false: Return lower parameter limit</para></param>
        /// <returns>Parameter limt as string if parameter limit exists accessible/exists.
        /// <para>null if there is no limit for this parameter</para></returns>
        public static string GetParameterLimit(DriveObject actDriveObject, string parameter, bool selectLimit)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter);
            string boundary = null;

            if (param != null)
            {
                if (selectLimit == true)
                {
                    boundary = param.MaxValue.ToString();
                }
                else
                {
                    boundary = param.MinValue.ToString();
                }
            }
            return boundary;
        }

        #endregion

        #region Online

        /// <summary>
        /// Opens a TIA project
        /// </summary>
        /// <param name="actDriveObject">Path of TIA project to be opened</param>
        /// <param name="parameter">Parameter where the opened project is stored</param>
        /// <param name="selectLimit">true: Return upper parameter limit
        /// <para>false: Return lower parameter limit</para></param>
        /// <returns>Parameter limt as string if parameter limit exists accessible/exists.
        /// <para>null if there is no limit for this parameter</para></returns>
        public static string GetParameterLimit(OnlineDriveObject actDriveObject, string parameter, bool selectLimit)
        {
            DriveParameter param = GetParameter(actDriveObject, parameter);
            string boundary = null;

            if (param != null)
            {
                if (selectLimit == true)
                {
                    boundary = param.MaxValue.ToString();
                }
                else
                {
                    boundary = param.MinValue.ToString();
                }
            }
            return boundary;
        }

        #endregion

        #endregion

        #endregion

        #region Telegram handling

        /// <summary>
        /// Sets the main telegram
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="telegramNumber">Number of telegram to set</param>
        /// <returns>
        /// true if telegram was set
        /// <para>false if telegram was not set</para></returns>
        public static bool SetMainTelegramNumber(DriveObject actDriveObject, int telegramNumber)
        {
            //Get Telegram Composition
            TelegramComposition myTelegramComp = actDriveObject.Telegrams;
            Telegram myTelegram = null;

            //Check if there is any telegram available on the axis
            if (myTelegramComp.Find(TelegramType.MainTelegram) != null)
            {
                myTelegram = myTelegramComp.Find(TelegramType.MainTelegram);

                if (myTelegram.CanChangeTelegram(telegramNumber))
                {
                    myTelegram.TelegramNumber = telegramNumber;
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Connects the main telegram to MC-Servo OB.
        /// <para>A connected controller with existing MC-Servo OB is required</para>
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <returns>
        /// true if MC-Servo was connected
        /// <para>false if MC Servo was not connected</para></returns>
        public static bool SetMainTelegramMCServo(DriveObject actDriveObject)
        {
            const int MCServo = 32768;
            //Get Telegram Composition
            TelegramComposition myTelegramComp = actDriveObject.Telegrams;

            foreach (Telegram myTelegram in myTelegramComp)
            {
                if (myTelegram.Type == TelegramType.MainTelegram)
                {
                    foreach (Address addr in myTelegram.Addresses)
                    {
                        if (addr.IoType == AddressIoType.Input ||
                            addr.IoType == AddressIoType.Output)
                        {
                            addr.SetAttribute("ProcessImage", MCServo);
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Returns starting input address of main telegram
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <returns>
        /// starting input address of the main telegram
        /// <para>-1  if there was an error during method call</para></returns>
        public static int GetMainTelegramAddressIn(DriveObject actDriveObject)
        {
            //Get Telegram Composition
            TelegramComposition myTelegramComp = actDriveObject.Telegrams;
            int address = -1;

            foreach (Telegram myTelegram in myTelegramComp)
            {
                if (myTelegram.Type == TelegramType.MainTelegram)
                {
                    foreach (Address addr in myTelegram.Addresses)
                    {
                        //Get input address
                        if (addr.IoType == AddressIoType.Input)
                        {
                            address = addr.StartAddress;
                        }
                    }
                }
            }
            return address;
        }

        /// <summary>
        /// Returns starting output address of main telegram
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <returns>
        /// starting output address of the main telegram
        /// <para>-1  if there was an error during method call</para></returns>
        public static int GetMainTelegramAddressOut(DriveObject actDriveObject)
        {
            //get Telegram Composition
            TelegramComposition myTelegramComp = actDriveObject.Telegrams;
            int address = 0;

            foreach (Telegram myTelegram in myTelegramComp)
            {
                if (myTelegram.Type == TelegramType.MainTelegram)
                {
                    foreach (Address addr in myTelegram.Addresses)
                    {
                        //Get output address
                        if (addr.IoType == AddressIoType.Output)
                        {
                            address = addr.StartAddress;
                        }
                    }
                }
            }
            return address;
        }

        /// <summary>
        /// Adds an additional free telegram
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="sendLength">Amount of words to send</param>
        /// <param name="receiveLength">Amount of words to receive</param>
        /// <returns>
        /// true if telegram was set
        /// <para>false if telegram was not set</para></returns>
        public static bool AddAdditionalTelegram(DriveObject actDriveObject, int sendLength, int receiveLength)
        {
            //Get Telegram Composition
            TelegramComposition myTelegramComp = actDriveObject.Telegrams;

            if (myTelegramComp.CanInsertAdditionalTelegram(sendLength, receiveLength))
            {
                myTelegramComp.InsertAdditionalTelegram(sendLength, receiveLength);
                return true;
            }
            return false;
        }

        /// <summary>
        /// Deletes the additional telegrams
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <returns>
        /// true if telegram was deleted
        /// <para>false if telegram was not deleted</para></returns>
        public static bool DeleteAdditionalTelegrams(DriveObject actDriveObject)
        {
            bool success = false;

            //get Telegram Composition
            TelegramComposition myTelegramComp = actDriveObject.Telegrams;

            foreach (Telegram telegram in myTelegramComp)
            {
                //Find AdditionalTelegram
                if (telegram.Type == TelegramType.AdditionalTelegram)
                {
                    myTelegramComp.EraseTelegram(TelegramType.AdditionalTelegram);
                    success = true;

                    break;
                }
            }
            return success;
        }

        /// <summary>
        /// Adds a torque telegram 750
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <returns>
        /// true if telegram was set
        /// <para>false if telegram was not set</para></returns>
        public static bool AddTorqueTelegram(DriveObject actDriveObject)
        {
            //get Telegram Composition
            TelegramComposition myTelegramComp = actDriveObject.Telegrams;
            bool success = false;

            //Torque Telegram has the number 750
            if (myTelegramComp.CanInsertTorqueTelegram(750))
            {
                myTelegramComp.InsertTorqueTelegram(750);
                success = true;
            }
            return success;
        }

        /// <summary>
        /// Deletes the torque telegrams
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <returns>
        /// true if telegram was deleted
        /// <para>false if telegram was not deleted</para></returns>
        public static bool DeleteTorqueTelegram(DriveObject actDriveObject)
        {
            bool success = false;

            //get Telegram Composition
            TelegramComposition myTelegramComp = actDriveObject.Telegrams;

            foreach (Telegram telegram in myTelegramComp)
            {
                //Find TorqueTelegram
                if (telegram.Type == TelegramType.TorqueTelegram)
                {
                    myTelegramComp.EraseTelegram(TelegramType.TorqueTelegram);
                    success = true;

                    break;
                }
            }
            return success;
        }

        /// <summary>
        /// Adds a safety telegram
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="telegramNumber">Number of telegram to set</param>
        /// <returns>
        /// true if telegram was set
        /// <para>false if telegram was not set</para></returns>
        public static bool AddSafetyTelegram(DriveObject actDriveObject, int telegramNumber)
        {
            //get Telegram Composition
            TelegramComposition myTelegramComp = actDriveObject.Telegrams;
            bool success = false;

            if (myTelegramComp.CanInsertSafetyTelegram(telegramNumber))
            {
                myTelegramComp.InsertSafetyTelegram(telegramNumber);
                success = true;
            }
            return success;
        }

        /// <summary>
        /// Deletes the safety telegram
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <returns>
        /// true if telegram was deleted
        /// <para>false if telegram was not deleted</para></returns>
        public static bool DeleteSafetyTelegram(DriveObject actDriveObject)
        {
            bool success = false;

            //get Telegram Composition
            TelegramComposition myTelegramComp = actDriveObject.Telegrams;

            foreach (Telegram telegram in myTelegramComp)
            {
                //Find SafetyTelegram
                if (telegram.Type == TelegramType.SafetyTelegram)
                {
                    myTelegramComp.EraseTelegram(TelegramType.SafetyTelegram);
                    success = true;

                    break;
                }
            }
            return success;
        }

        /// <summary>
        /// Adds Safety Info/Control Channel
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="telegramNumber">Number of telegram to set</param>
        /// <returns>
        /// true if telegram was set
        /// <para>false if telegram was not set</para></returns>
        public static bool AddSafetyInfoControlChannel(DriveObject actDriveObject, int telegramNumber)
        {
            bool success = false;
            //get Telegram Composition
            TelegramComposition myTelegramComp = actDriveObject.Telegrams;

            if (myTelegramComp.CanInsertSupplementaryTelegram(telegramNumber))
            {
                myTelegramComp.InsertSupplementaryTelegram(telegramNumber);
                success = true;
            }
            return success;
        }

        /// <summary>
        /// Deletes the Safety Info/Control Channel
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <returns>
        /// true if telegram was deleted
        /// <para>false if telegram was not deleted</para></returns>
        public static bool DeleteSafetyInfoControlChannel(DriveObject actDriveObject)
        {
            bool success = false;
            bool delete = false;

            //get Telegram Composition
            TelegramComposition myTelegramComp = actDriveObject.Telegrams;

            foreach (Telegram telegram in myTelegramComp)
            {
                //Find Safety Info/Control Channel
                if (telegram.Type == TelegramType.SupplementaryTelegram)
                {
                    delete = true;
                    break;
                }
            }
            if (delete)
            {
                myTelegramComp.EraseTelegram(TelegramType.SupplementaryTelegram);
                success = true;
            }
            return success;
        }

        /// <summary>
        /// Changes the main telegram to free telegram
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="sendLength">Send length in words</param>
        /// <param name="receiveLength">Receive length in words</param>
        /// <param name="keepAddr">true: Keep telegram address
        /// <para>false: Ignore previous telegram address</para></param>
        /// <returns>
        /// true if telegram was changed with success
        /// <para>false if telegram or length was not changed</para></returns>
        public static bool SetFreeTelegram(DriveObject actDriveObject, int sendLength, int receiveLength, bool keepAddr)
        {
            //get Telegram Composition
            TelegramComposition myTelegramComp = actDriveObject.Telegrams;
            bool success = false;

            foreach (Telegram telegram in myTelegramComp)
            {
                if(telegram.Type == TelegramType.MainTelegram)
                {
                    //Change telegram to free telegram
                    if (telegram.CanChangeTelegram(999) &&
                        telegram.CanChangeSize(AddressIoType.Input, sendLength, keepAddr) &&
                        telegram.CanChangeSize(AddressIoType.Output, (receiveLength), keepAddr))
                    {
                        telegram.TelegramNumber = 999;
                        telegram.ChangeSize(AddressIoType.Input, sendLength, keepAddr);
                        telegram.ChangeSize(AddressIoType.Output, (receiveLength), keepAddr);

                        success = true;
                    }
                }
            }
            return success;
        }

        /// <summary>
        /// Adds a Telegram Extension
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="sendExtenstion">Extension length (send direction) in words</param>
        /// <param name="receiveExtension">Extension length (receive direction) in words</param>
        /// <param name="keepAddr">true: Keep telegram address
        /// <para>false: Ignore previous telegram address</para></param>
        /// <returns>
        /// true if Extension was added with success
        /// <para>false if Extension could not be added</para></returns>
        public static bool AddMainTelegramExtension(DriveObject actDriveObject, int sendExtenstion, int receiveExtension, bool keepAddr)
        {
            TelegramComposition myTelegramComp = actDriveObject.Telegrams;
            bool success = false;

            foreach (Telegram telegram in myTelegramComp)
            {
                if (telegram.Type == TelegramType.MainTelegram)
                {
                    //Total send length is default length + extension
                    int sendUpdate = telegram.GetSize(AddressIoType.Input) + sendExtenstion;
                    //Total receive length is default length + extension
                    int rcvUpdate = telegram.GetSize(AddressIoType.Output) + receiveExtension;

                    if (telegram.CanChangeSize(AddressIoType.Input, sendUpdate, keepAddr) &&
                        telegram.CanChangeSize(AddressIoType.Output, rcvUpdate, keepAddr))
                    {
                        telegram.ChangeSize(AddressIoType.Input, sendUpdate, keepAddr);
                        telegram.ChangeSize(AddressIoType.Output, rcvUpdate, keepAddr);

                        success = true;
                    }
                }
            }
            return success;
        }

        #endregion

        #region Drive object handling

        #region Offline

        /// <summary>
        /// Returns a DriveObject of Contol Unit for CU3x0-2 devices
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <returns>
        /// DriveObject Object of Control Unit if Control Unit was found.
        /// <para>Null if Control Unit was not found.</para>
        /// </returns>
        public static DriveObject GetControlUnit(DriveObject actDriveObject)
        {
            DriveObject ControlUnit = null;
            try
            {
                //get the DeviceItem of the actual Drive Object
                DeviceItem actDeviceItem =
                    (DeviceItem)actDriveObject.Parent.Parent;

                //get the Device of the actual DeviceItem
                if (actDeviceItem.TypeIdentifier == "System:Rack")
                {
                    Device drive_unit = (Device)actDeviceItem.Parent;

                    if ((drive_unit.TypeIdentifier.ToString().
                        Contains("System:Device.S120")) ||
                        (drive_unit.TypeIdentifier.ToString().
                        Contains("System:Device.S150")) ||
                        (drive_unit.TypeIdentifier.ToString().
                        Contains("System:Device.G130")) ||
                        (drive_unit.TypeIdentifier.ToString().
                        Contains("System:Device.G150"))
                        )
                    {
                        //Iterate thorough DeviceItems to find Control Unit
                        foreach (DeviceItem deviceItems in drive_unit.DeviceItems)
                        {
                            if ((deviceItems.TypeIdentifier ==
                                 "System:Rack.S120") ||
                                 (deviceItems.TypeIdentifier ==
                                 "System:Rack.S150") ||
                                 (deviceItems.TypeIdentifier ==
                                 "System:Rack.G130") ||
                                 (deviceItems.TypeIdentifier ==
                                 "System:Rack.G150")
                               )
                            {
                                ControlUnit =
                                deviceItems.GetService<DriveObjectContainer>().
                                DriveObjects[0];
                                return ControlUnit;
                            }
                        }
                    }
                }
                return null;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        /// <summary>
        /// Returns a DriveObject of Infeed Axis for CU3x0-2 devices
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <returns>
        /// DriveObject Object of Infeed Axis if Infeed was found.
        /// <para>Null if Infeed was not found.</para>
        /// </returns>
        public static DriveObject GetInfeedAxis(DriveObject actDriveObject)
        {
            DriveObject InfeedAxis = null;
            string DriveObjectType = null;

            try
            {
                //get the DeviceItem of the actual Drive Object
                DeviceItem actDeviceItem = (DeviceItem)actDriveObject.Parent.Parent;

                if (actDeviceItem.TypeIdentifier == "System:Rack")
                {
                    //get the Device of the actual DeviceItem
                    Device drive_unit = (Device)actDeviceItem.Parent;

                    //In case of S120/S150/G150 Devices get the infeed axis
                    if ((drive_unit.TypeIdentifier.ToString().
                        Contains("System:Device.S120")) ||
                        (drive_unit.TypeIdentifier.ToString().
                        Contains("System:Device.S150")) ||
                        (drive_unit.TypeIdentifier.ToString().
                        Contains("System:Device.G150")))
                    {
                        foreach (DeviceItem deviceItems in drive_unit.DeviceItems)
                        {
                            if (deviceItems.TypeIdentifier == "System:Rack")
                            {
                                InfeedAxis =
                                deviceItems.GetService<DriveObjectContainer>().
                                DriveObjects[0];

                                //in r107 the type of the drive object is saved
                                DriveObjectType = ReadParameterValue(InfeedAxis, 107);

                                if ((DriveObjectType == "10") ||
                                    (DriveObjectType == "20") ||
                                    (DriveObjectType == "21") ||
                                    (DriveObjectType == "30") ||
                                    (DriveObjectType == "40") ||
                                    (DriveObjectType == "41") ||
                                    (DriveObjectType == "42"))
                                {
                                    return InfeedAxis;
                                }
                            }
                        }
                    }
                }
                return null;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        /// <summary>
        /// Searches for another axis on CU3x0-2 devices
        /// <para>using a given axis name</para>
        /// </summary>
        /// <param name="actDriveObject">DriveObject on which the method is executed</param>
        /// <param name="nameOfS120Axis">Name of the axis to be returned</param>
        /// <returns>
        /// DriveObject Object if given axis name was found.
        /// <para>Null if axis name was not found.</para>
        /// </returns>
        public static DriveObject GetDriveAxisByName(DriveObject actDriveObject, String nameOfS120Axis)
        {
            DriveObject S120DriveAxis = null;
            try
            {
                //get the DeviceItem of the actual Drive Object
                DeviceItem actDeviceItem = (DeviceItem)actDriveObject.Parent.Parent;

                if (actDeviceItem.TypeIdentifier == "System:Rack")
                {
                    //get the Device of the actual DeviceItem
                    Device drive_unit = (Device)actDeviceItem.Parent;

                    //In case of S120 Devices get S120 drive axis
                    if (drive_unit.TypeIdentifier.ToString().
                        Contains("System:Device.S120"))
                    {
                        foreach (DeviceItem deviceItems
                            in drive_unit.DeviceItems)
                        {
                            if (deviceItems.Name == nameOfS120Axis)
                            {
                                S120DriveAxis =
                                deviceItems.GetService<DriveObjectContainer>().
                                DriveObjects[0];

                                return S120DriveAxis;
                            }
                        }
                    }
                }
                return null;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        #endregion

        #region Online

        /// <summary>
        /// Returns a OnlineDriveObject of Contol Unit for CU3x0-2 devices
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <returns>
        /// OnlineDriveObject Object of Control Unit if Control Unit was found.
        /// <para>Null if Control Unit was not found.</para>
        /// </returns>
        public static OnlineDriveObject GetControlUnit(OnlineDriveObject actDriveObject)
        {
            OnlineDriveObject ControlUnit = null;
            try
            {
                //get the DeviceItem of the actual Drive Object
                DeviceItem actDeviceItem =
                    (DeviceItem)actDriveObject.Parent.Parent;

                //get the Device of the actual DeviceItem
                if (actDeviceItem.TypeIdentifier == "System:Rack")
                {
                    Device drive_unit = (Device)actDeviceItem.Parent;

                    if ((drive_unit.TypeIdentifier.ToString().
                        Contains("System:Device.S120")) ||
                        (drive_unit.TypeIdentifier.ToString().
                        Contains("System:Device.S150")) ||
                        (drive_unit.TypeIdentifier.ToString().
                        Contains("System:Device.G130")) ||
                        (drive_unit.TypeIdentifier.ToString().
                        Contains("System:Device.G150"))
                        )
                    {
                        //Iterate thorough DeviceItems to find Control Unit
                        foreach (DeviceItem deviceItems in drive_unit.DeviceItems)
                        {
                            if ((deviceItems.TypeIdentifier ==
                                 "System:Rack.S120") ||
                                 (deviceItems.TypeIdentifier ==
                                 "System:Rack.S150") ||
                                 (deviceItems.TypeIdentifier ==
                                 "System:Rack.G130") ||
                                 (deviceItems.TypeIdentifier ==
                                 "System:Rack.G150")
                               )
                            {
                                ControlUnit =
                                deviceItems.GetService<OnlineDriveObjectContainer>().
                                OnlineDriveObjects[0];
                                return ControlUnit;
                            }
                        }
                    }
                }
                return null;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        /// <summary>
        /// Returns a OnlineDriveObject of Infeed Axis for CU3x0-2 devices
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <returns>
        /// OnlineDriveObject Object of Infeed Axis if Infeed was found.
        /// <para>Null if Infeed was not found.</para>
        /// </returns>
        public static OnlineDriveObject GetInfeedAxis(OnlineDriveObject actDriveObject)
        {
            OnlineDriveObject InfeedAxis = null;
            string DriveObjectType = null;

            try
            {
                //get the DeviceItem of the actual Drive Object
                DeviceItem actDeviceItem = (DeviceItem)actDriveObject.Parent.Parent;

                if (actDeviceItem.TypeIdentifier == "System:Rack")
                {
                    //get the Device of the actual DeviceItem
                    Device drive_unit = (Device)actDeviceItem.Parent;

                    //In case of S120/S150/G150 Devices get the infeed axis
                    if ((drive_unit.TypeIdentifier.ToString().
                        Contains("System:Device.S120")) ||
                        (drive_unit.TypeIdentifier.ToString().
                        Contains("System:Device.S150")) ||
                        (drive_unit.TypeIdentifier.ToString().
                        Contains("System:Device.G150")))
                    {
                        foreach (DeviceItem deviceItems in drive_unit.DeviceItems)
                        {
                            if (deviceItems.TypeIdentifier == "System:Rack")
                            {
                                InfeedAxis =
                                deviceItems.GetService<OnlineDriveObjectContainer>().
                                OnlineDriveObjects[0];

                                //in r107 the type of the drive object is saved
                                DriveObjectType = ReadParameterValue(InfeedAxis, 107);

                                if ((DriveObjectType == "10") ||
                                    (DriveObjectType == "20") ||
                                    (DriveObjectType == "21") ||
                                    (DriveObjectType == "30") ||
                                    (DriveObjectType == "40") ||
                                    (DriveObjectType == "41") ||
                                    (DriveObjectType == "42"))
                                {
                                    return InfeedAxis;
                                }
                            }
                        }
                    }
                }
                return null;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        /// <summary>
        /// Searches for another axis on CU3x0-2 devices
        /// <para>using a given axis name</para>
        /// </summary>
        /// <param name="actDriveObject">OnlineDriveObject on which the method is executed</param>
        /// <param name="nameOfS120Axis">Name of the axis to be returned</param>
        /// <returns>
        /// OnlineDriveObject Object if given axis name was found.
        /// <para>Null if axis name was not found.</para>
        /// </returns>
        public static OnlineDriveObject GetDriveAxisByName(OnlineDriveObject actDriveObject, String nameOfS120Axis)
        {
            OnlineDriveObject S120DriveAxis = null;
            try
            {
                //get the DeviceItem of the actual Drive Object
                DeviceItem actDeviceItem = (DeviceItem)actDriveObject.Parent.Parent;

                if (actDeviceItem.TypeIdentifier == "System:Rack")
                {
                    //get the Device of the actual DeviceItem
                    Device drive_unit = (Device)actDeviceItem.Parent;

                    //In case of S120 Devices get S120 drive axis
                    if (drive_unit.TypeIdentifier.ToString().
                        Contains("System:Device.S120"))
                    {
                        foreach (DeviceItem deviceItems
                            in drive_unit.DeviceItems)
                        {
                            if (deviceItems.Name == nameOfS120Axis)
                            {
                                S120DriveAxis =
                                deviceItems.GetService<OnlineDriveObjectContainer>().
                                OnlineDriveObjects[0];

                                return S120DriveAxis;
                            }
                        }
                    }
                }
                return null;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        #endregion

        #endregion

        #region Project handling

        /// <summary>
        /// Opens a TIA project
        /// </summary>
        /// <param name="filePath">Path of TIA project to be opened</param>
        /// <param name="project">Parameter where the opened project is stored</param>
        /// <param name="tiaPortal">Instance of TIA Portal</param>
        /// <returns>true if TIA project was opened successfully
        /// <para>false if TIA project was not opened</para></returns>
        public static bool OpenProject(string filePath, TiaPortal tiaPortal)
        {
            Project project = tiaPortal.Projects.First();
            try
            {
                //Open project if no project opened or if the new project has another name
                if (project == null)
                {
                    project = tiaPortal.Projects.Open(new FileInfo(filePath));
                }
                else if (project.Path != new FileInfo(filePath))
                {
                    project.Close();
                    project = tiaPortal.Projects.Open(new FileInfo(filePath));
                }
            }
            catch (Exception e)
            {
                return false;
            }
            if (project != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        #endregion

        #region Enumerating devices

        /// <summary>
        /// Searches the project for devices and adds them to a list
        /// </summary>
        /// <param name="project">Project which is searced for devices</param>
        /// <returns>List of all devices that meet the search criteria
        /// <para>Search criteria is defined in the method <see cref="EnumerateDevices"/></para>
        /// <para>Default search criteria: SINAMICS devices</para></returns>
        public static IList<Device> EnumerateDevicesInProject(Project project)
        {
            IList<Device> deviceList = new List<Device>();
            DeviceComposition deviceComposition = project.Devices;

            EnumerateDevices(deviceList, deviceComposition);
            EnumerateDevicesInGroups(deviceList, project);

            return deviceList;
        }

        /// <summary>
        /// Searches groups for devices and adds them to a list
        /// </summary>
        /// <param name="deviceList">List to which found devices are added</param>
        /// <param name="project">Project which is searced for devices</param>
        private static void EnumerateDevicesInGroups(IList<Device> deviceList, Project project)
        {
            foreach (DeviceUserGroup deviceUserGroup in project.DeviceGroups)
            {
                EnumerateDeviceUserGroup(deviceList, deviceUserGroup);
            }
        }

        /// <summary>
        /// Searches groups recursively for devices and adds them to a list
        /// </summary>
        /// <param name="deviceList">List to which found devices are added</param>
        /// <param name="deviceUserGroup">Group which is searced for devices</param>
        private static void EnumerateDeviceUserGroup(IList<Device> deviceList, DeviceUserGroup deviceUserGroup)
        {
            EnumerateDevices(deviceList, deviceUserGroup.Devices);
            foreach (DeviceUserGroup subDeviceUserGroup in deviceUserGroup.Groups)
            {
                //Recursion to handle nested groups
                EnumerateDeviceUserGroup(deviceList, subDeviceUserGroup);
            }
        }

        /// <summary>
        /// Searches device composition for devices and adds them to a list
        /// <para>Criteria which device is added to a list is defined in this method</para>
        /// </summary>
        /// <param name="deviceList">List to which found devices are added</param>
        /// <param name="deviceComposition">Device composition which is searched for devices</param>
        private static void EnumerateDevices(IList<Device> deviceList, DeviceComposition deviceComposition)
        {
            foreach (Device device in deviceComposition)
            {
                //Only SINAMICS type devices are added to the device list
                if (device.TypeIdentifier.ToString().Contains("System:Device.G110M") ||
                    device.TypeIdentifier.ToString().Contains("System:Device.G120") ||
                    device.TypeIdentifier.ToString().Contains("System:Device.G120C") ||
                    device.TypeIdentifier.ToString().Contains("System:Device.G115D") ||
                    device.TypeIdentifier.ToString().Contains("System:Device.G120D") ||
                    device.TypeIdentifier.ToString().Contains("System:Device.G120P") ||
                    device.TypeIdentifier.ToString().Contains("System:Device.G130") ||
                    device.TypeIdentifier.ToString().Contains("System:Device.G150") ||
                    device.TypeIdentifier.ToString().Contains("System:Device.S120") ||
                    device.TypeIdentifier.ToString().Contains("System:Device.S120-DCU") ||
                    device.TypeIdentifier.ToString().Contains("System:Device.S150") ||
                    device.TypeIdentifier.ToString().Contains("System:Device.S210"))
                {
                    //Add device to list
                    deviceList.Add(device);
                }
            }
        }

        #region Offline

        /// <summary>
        /// Searches a list of devices for drive objects and adds the drive objects to another list
        /// </summary>
        /// <param name="deviceList">List of devices which is searched for device objects</param>
        /// <returns>List of all device objekts that meet the search criteria
        /// <para>Search criteria is defined in the method <see cref="EnumerateDriveObjectsInDevice"/></para>
        /// <para>Default search criteria: SINAMICS drive objects</para></returns>
        public static IList<DriveObject> EnumerateDriveObjectsInDeviceList(IList<Device> deviceList)
        {
            IList<DriveObject> driveObjectList = new List<DriveObject>();

            foreach (Device device in deviceList)
            {
                EnumerateDriveObjectsInDevice(driveObjectList, device);
            }

            return driveObjectList;
        }

        /// <summary>
        /// Searches a device for drive objects and adds the drive objects to another list
        /// </summary>
        /// <param name="driveObjectList">List to which found device objects are added</param>
        /// <param name="device">Device which is searched for device objects</param>
        private static void EnumerateDriveObjectsInDevice(IList<DriveObject> driveObjectList, Device device)
        {
            //Searching for drive objects of CU3x0- and S210-devices
            if (device.TypeIdentifier.ToString().Contains("System:Device.S120") ||
                device.TypeIdentifier.ToString().Contains("System:Device.S150") ||
                device.TypeIdentifier.ToString().Contains("System:Device.G130") ||
                device.TypeIdentifier.ToString().Contains("System:Device.G150") ||
                device.TypeIdentifier.ToString().Contains("System:Device.S120-DCU") ||
                device.TypeIdentifier.ToString().Contains("System:Device.S210"))
            {
                foreach (DeviceItem deviceItem in device.DeviceItems)
                {
                    if (deviceItem.TypeIdentifier.ToString().Contains("System:Rack"))
                    {
                        driveObjectList.Add(deviceItem.GetService<DriveObjectContainer>().DriveObjects[0]);
                    }
                }
            }

            //Searching for drive objects of Sinamics G-devices
            else if (device.TypeIdentifier.ToString().Contains("System:Device.G110M") ||
                device.TypeIdentifier.ToString().Contains("System:Device.G120C") ||
                device.TypeIdentifier.ToString().Contains("System:Device.G115D") ||
                device.TypeIdentifier.ToString().Contains("System:Device.G120D") ||
                device.TypeIdentifier.ToString().Contains("System:Device.G120"))
            {
                foreach (DeviceItem deviceItem in device.DeviceItems)
                {
                    if (deviceItem.Classification == DeviceItemClassifications.HM)
                    {
                        driveObjectList.Add(deviceItem.GetService<DriveObjectContainer>().DriveObjects[0]);
                    }
                }
            }
        }

        #endregion

        #region Online

        /// <summary>
        /// Searches a list of devices for online drive objects and adds the online drive objects to another list
        /// </summary>
        /// <param name="deviceList">List of devices which is searched for online device objects</param>
        /// <returns>List of all online device objekts that meet the search criteria
        /// <para>Search criteria is defined in the method <see cref="EnumerateDriveObjectsInDevice"/></para>
        /// <para>Default search criteria: SINAMICS online drive objects</para></returns>
        public static IList<OnlineDriveObject> EnumerateOnlineDriveObjectsInDeviceList(IList<Device> deviceList)
        {
            IList<OnlineDriveObject> onlineDriveObjectList = new List<OnlineDriveObject>();

            foreach (Device device in deviceList)
            {
                EnumerateOnlineDriveObjectsInDevice(onlineDriveObjectList, device);
            }

            return onlineDriveObjectList;
        }

        /// <summary>
        /// Searches a device for online drive objects and adds the online drive objects to another list
        /// </summary>
        /// <param name="driveObjectList">List to which found online device objects are added</param>
        /// <param name="device">Device which is searched for online device objects</param>
        private static void EnumerateOnlineDriveObjectsInDevice(IList<OnlineDriveObject> onlineDriveObjectList, Device device)
        {
            //Searching for online drive objects of CU3x0- and S210-devices
            if (device.TypeIdentifier.ToString().Contains("System:Device.S120") ||
                device.TypeIdentifier.ToString().Contains("System:Device.S150") ||
                device.TypeIdentifier.ToString().Contains("System:Device.G130") ||
                device.TypeIdentifier.ToString().Contains("System:Device.G150") ||
                device.TypeIdentifier.ToString().Contains("System:Device.S120-DCU") ||
                device.TypeIdentifier.ToString().Contains("System:Device.S210"))
            {
                foreach (DeviceItem deviceItem in device.DeviceItems)
                {
                    if (deviceItem.TypeIdentifier.ToString().Contains("System:Rack"))
                    {
                        onlineDriveObjectList.Add(deviceItem.GetService<OnlineDriveObjectContainer>().OnlineDriveObjects[0]);
                    }
                }
            }

            //Searching for online drive objects of Sinamics G-devices
            else if (device.TypeIdentifier.ToString().Contains("System:Device.G110M") ||
                device.TypeIdentifier.ToString().Contains("System:Device.G120C") ||
                device.TypeIdentifier.ToString().Contains("System:Device.G115D") ||
                device.TypeIdentifier.ToString().Contains("System:Device.G120D") ||
                device.TypeIdentifier.ToString().Contains("System:Device.G120"))
            {
                foreach (DeviceItem deviceItem in device.DeviceItems)
                {
                    if (deviceItem.Classification == DeviceItemClassifications.HM)
                    {
                        onlineDriveObjectList.Add(deviceItem.GetService<OnlineDriveObjectContainer>().OnlineDriveObjects[0]);
                    }
                }
            }
        }

        #endregion

        #endregion
    }
}