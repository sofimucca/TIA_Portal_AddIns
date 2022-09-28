using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.HW;
using System.Windows.Forms;
using Siemens.Engineering.MC.Drives;
using Siemens.Engineering.MC.Drives.DFI;
using Siemens.Engineering.MC.Drives.Enums;
using System.Collections.Generic;
using System.Net.Configuration;
using System;
using Siemens.Engineering.SW.Blocks;
using SDRhelper;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.Hmi.Screen;
using static System.Net.Mime.MediaTypeNames;
using System.IO;
using System.Security.Cryptography;
using System.Net;
using System.Linq;

namespace TIA_Add_In_Example_Project
{
    public class AddIn : ContextMenuAddIn
    {

        public string logtext = "";
        
        /// <summary>
        ///The global TIA Portal Object 
        ///<para>It will be used in the TIA Add-In.</para>
        /// </summary>
        TiaPortal _tiaportal;

        /// <summary>
        /// The display name of the Add-In.
        /// </summary>
        private const string s_DisplayNameOfAddIn = "Add-In custom ETI";

        /// <summary>
        /// The constructor of the AddIn.
        /// Creates an object of the class AddIn
        /// Called from AddInProvider, when the first
        /// right-click is performed in TIA
        /// Motherclass' constructor of ContextMenuAddin
        /// will be executed, too. 
        /// </summary>
        /// <param name="tiaportal">
        /// Represents the actual used TIA Portal process.
        /// </param>
        public AddIn(TiaPortal tiaportal) : base(s_DisplayNameOfAddIn)
        {
            /*
             * The acutal TIA Portal process is saved in the 
             * global TIA Portal variable _tiaportal
             * tiaportal comes as input Parameter from the
             * AddInProvider
            */
            _tiaportal = tiaportal;
        }

        /// <summary>
        /// The method is supplemented to include the Add-In
        /// in the Context Menu of TIA Portal.
        /// Called when a right-click is performed in TIA
        /// and a mouse-over is performed on the name of the Add-In.
        /// </summary>
        /// <typeparam name="addInRootSubmenu">
        /// The Add-In will be displayed in 
        /// the Context Menu of TIA Portal.
        /// </typeparam>
        /// <example>
        /// ActionItems like Buttons/Checkboxes/Radiobuttons
        /// are possible. In this example, only Buttons will be created 
        /// which will start the Add-In program code.
        /// </example>
        protected override void BuildContextMenuItems(ContextMenuAddInRoot
            addInRootSubmenu)
        {
            /* Method addInRootSubmenu.Items.AddActionItem
             * Will Create a Pushbutton with the text 'Start Add-In Code'
             * 1st input parameter of AddActionItem is the text of the 
             *          button
             * 2nd input parameter of AddActionItem is the clickDelegate, 
             *          which will be executed in case the button 'Start 
             *          Add-In Code' will be clicked/pressed.
             * 3rd input parameter of AddActionItem is the 
             *          updateStatusDelegate, which will be executed in 
             *          case there is a mouseover the button 'Start 
             *          Add-In Code'.
             * in <placeholder> the type of AddActionItem will be 
             *          specified, because AddActionItem is generic 
             * AddActionItem<DeviceItem> will create a button that will be 
             *          displayed if a rightclick on a DeviceItem will be 
             *          performed in TIA Portal
             * AddActionItem<Project> will create a button that will be 
             *          displayed if a rightclick on the project name 
             *          will be performed in TIA Portal
            */
            //    addInRootSubmenu.Items.AddActionItem<Project>(
            //        ("Start Add-In"), ListDevice, OnCanSomething);
            //    addInRootSubmenu.Items.AddActionItem<Project>(
            //        "Not Available here", OnClickProject, 
            //        OnStatusUpdateProject);
            //}

            addInRootSubmenu.Items.AddActionItem<Project>(
            ("Explore project"), ExploreProject);
            addInRootSubmenu.Items.AddActionItem<DeviceItem>(
            ("Set Telegram"), SetFreeTelegram);
            addInRootSubmenu.Items.AddActionItem<DeviceItem>(
            ("Set Safety Telegram"), SetSafetyTelegram);
            addInRootSubmenu.Items.AddActionItem<DeviceItem>(
            ("Get drives data"), GetDrivesData);
            addInRootSubmenu.Items.AddActionItem<DeviceItem>(
            ("Read drives telegrams"), SetSafetyTelegram);
            addInRootSubmenu.Items.AddActionItem<Device>(
            ("Make text list"), GetPageTextList);


        }

        //Put the following function definitions before or after the
        //function definition of "private static void IterateThroughDevices(Project project)":



        /// <summary>
        /// The method contains the program code of the TIA Add-In.
        /// Called when the button 'Start Add-In Code' will be pressed.
        /// </summary>
        /// <typeparam name="menuSelectionProvider">
        /// here, the same type as in addInRootSubmenu.Items.AddActionItem
        /// must be used -> here it is <DeviceItem>
        /// </typeparam>
        /// 


        private void WriteLogFile(string fileName, string text)
        {
            string path = Path.GetDirectoryName(_tiaportal.Projects[0].Path.ToString()) + @"\" + fileName;

            File.WriteAllText(path, text);

            MessageBox.Show("Log file:" + path);

        }


        private void ExploreProject(MenuSelectionProvider<Project>
    menuSelectionProvider)
        {

            Project project = _tiaportal.Projects[0];

            try
            {

                foreach (Device device in project.Devices)
                {

                    if (device.DeviceItems != null)
                    {
                        this.logtext += ExploreDevice(device.DeviceItems);
                    }

                    this.logtext += device.Name + ";\t" + device.TypeIdentifier + "\n";

                }

            }
            catch (Exception)
            {

                this.logtext += "Fail!\n";

            }

            WriteLogFile("devices.txt", this.logtext);

        }


        private string ExploreDevice(DeviceItemComposition devices)
        {

            foreach (DeviceItem device in devices)
            {

                this.logtext += device.Name + ";\t" + device.TypeIdentifier + "\n";

                
                    if (device.DeviceItems != null)
                    {
                        this.logtext += ExploreDevice(device.DeviceItems);
                    }

            }

            return this.logtext + "\n";

        }


        #region Drive


        private void ChangeTelegramAddress(Telegram telegram, int startAddress)
        {

            foreach (Address addr in telegram.Addresses)
            {
                //Get output address
                if (addr.IoType == AddressIoType.Input || addr.IoType == AddressIoType.Output)
                {

                    addr.StartAddress = startAddress;
                }
            }
        }

        private void getAttributes(IList<EngineeringAttributeInfo> collection, string fileName)
        {
            string infos = " ";
            foreach (EngineeringAttributeInfo info in collection)
            {
                infos += info.Name + " " + info.AccessMode + " " + info.CreateRelevance +"\n";
            }

            WriteLogFile(fileName, infos);

        }

        private bool SetSafetyTelegramParameters(DriveObject actDriveObject, int safeMonitoringTime=450)
        {

            bool success = false;

            try
            {
                Telegram safetyTelegram = actDriveObject.Telegrams.Find(TelegramType.SafetyTelegram);
                safetyTelegram.SetAttribute("Failsafe_ManualAssignmentFMonitoringtime", true);
                safetyTelegram.SetAttribute("Failsafe_FMonitoringtime", safeMonitoringTime);
                //safetyTelegram.SetAttribute("Failsafe_FSourceAddress", startAddress);
                //safetyTelegram.SetAttribute("Failsafe_FDestinationAddress", startAddress);

                success = true;
            }
            catch (Exception)
            {

                throw;
            }


            return success;

        }


        private bool GetTelegramAddress(TelegramComposition telegramComp)

        {
            bool result;


            if (telegramComp == null)
            {
                result = false;
            }
            else
            {
                foreach (Telegram myTelegram in telegramComp)
                {
                    this.logtext += "Telegram type:  \t" + myTelegram.Type.ToString() + "\n";


                    foreach (Address addr in myTelegram.Addresses)
                    {
                        //Get output address
                        if (addr.IoType == AddressIoType.Input)
                        {

                            this.logtext += "address input:\t" + addr.StartAddress + "\n";
                        }
                        else if (addr.IoType == AddressIoType.Output)
                        {

                            this.logtext += "address output:\t" + addr.StartAddress + "\n";

                        }
                    }



                }

                result = true;

            }

            return result;

        }

        private void GetDrivesData(MenuSelectionProvider<DeviceItem>
menuSelectionProvider)
        {

            IEnumerable<DeviceItem> selection = menuSelectionProvider.GetSelection<DeviceItem>();

            this.logtext += "Drives name;\tTelegram input;\tTelegram output;\tSafety telegram input;\tSafety telegram outputs;\n";
            this.logtext += "\n\n";

            foreach (DeviceItem deviceItem in selection)
            {
                DriveObject driveObject = deviceItem.GetService<DriveObjectContainer>().DriveObjects[0];

                if (driveObject != null)
                {
                    try
                    {

                        string[] telegrams = {deviceItem.Name,"\t\t","\t\t","\t\t","\t\t"};
                        var index= -1;

                        foreach (Telegram myTelegram in driveObject.Telegrams)
                        {
                            switch (myTelegram.Type)
                            {
                                case TelegramType.MainTelegram:
                                    index = 1;
                                    break;
                                case TelegramType.SupplementaryTelegram:
                                    index = -1;
                                    break;
                                case TelegramType.AdditionalTelegram:
                                    index = -1;
                                    break;
                                case TelegramType.SafetyTelegram:
                                    index = 3;
                                    break;
                                case TelegramType.TorqueTelegram:
                                    index = -1;
                                    break;
                                default:
                                    index = -1;
                                    break;
                            }

                            if (index>0)
                            {
                                foreach (Address addr in myTelegram.Addresses)
                                {
                                    //Get output address
                                    if (addr.IoType == AddressIoType.Input)
                                    {

                                        telegrams[index] += addr.StartAddress;
                                    }
                                    else if (addr.IoType == AddressIoType.Output)
                                    {

                                        telegrams[index+1] += addr.StartAddress;

                                    }
                                }
                            }

                        }
                        
                        foreach (string word in telegrams)
                        {
                            this.logtext += word +";";
                        }
                        this.logtext += "\n";
                    }
                    catch (Exception)
                    {
                        this.logtext += deviceItem.Name + "\tfail: check if address is already occupied.\n";
                        MessageBox.Show("FAIL! See log file for more information.");
                        return;

                    }
                }
                else
                {
                    this.logtext += deviceItem.Name + "\tis not a drive object!\n";
                }



            }

            WriteLogFile("drivesData.csv", this.logtext);
            MessageBox.Show("Procedure completed with success, see log file for more information.");




        }

        private void SetFreeTelegram(MenuSelectionProvider<DeviceItem>
menuSelectionProvider)
        {

            Form addressForm = new Form();

            if (addressForm.ShowDialog() != DialogResult.OK)
                return;

            int startingAddress = addressForm.adressStart;
            int size = addressForm.addresSize;
            int incr = addressForm.addressIncr;

            IEnumerable<DeviceItem> selection = menuSelectionProvider.GetSelection<DeviceItem>();

            this.logtext += "Log File\t\t" + DateTime.Now.ToString();
            this.logtext += "\n\n";

            foreach (DeviceItem deviceItem in selection)
            {
                DriveObject driveObject = deviceItem.GetService<DriveObjectContainer>().DriveObjects[0];

                if (driveObject != null)
                {
                    try
                    {

                        SDRhelper.StartdriveHelper.SetFreeTelegram(driveObject, size, size, false);
                        ChangeTelegramAddress(driveObject.Telegrams.Find(TelegramType.MainTelegram), startingAddress);

                        startingAddress += incr;

                        this.logtext += "\n\n" + deviceItem.Name + "\n";
                        GetTelegramAddress(driveObject.Telegrams);
                        this.logtext += "\n";

                        this.logtext += "\n-------------------------------------------------------\n";

                    }
                    catch (Exception)
                    {
                        this.logtext += deviceItem.Name + "\tfail: check if address is already occupied.\n";
                        MessageBox.Show("FAIL! See log file for more information.");
                        return;

                    }
                }
                else
                {
                    this.logtext += deviceItem.Name + "\tis not a drive object!\n";
                }


                
            }

            WriteLogFile("setTelegramLog.txt", this.logtext);
            MessageBox.Show("Procedure completed with success, see log file for more information.");
           



        }

        private void SetSafetyTelegram(MenuSelectionProvider<DeviceItem>
menuSelectionProvider)
        {
            const int safetyTelegram = 30;

            Form addressForm = new Form();
            addressForm.BackColor = System.Drawing.Color.Yellow;
            addressForm.Text = "Set Safety Telegram";
            addressForm.labelSize.Text = "Fail Safe Monitoring Time";
            addressForm.numericUpDownSize.Value = 450;

            if (addressForm.ShowDialog() != DialogResult.OK)
                return;

            int startingAddress = addressForm.adressStart;
            int safeMonitoringTime = addressForm.addresSize;
            int incr = addressForm.addressIncr;

            string profiSafeTags = "";

            IEnumerable<DeviceItem> selection = menuSelectionProvider.GetSelection<DeviceItem>();

            this.logtext += "Log File\t\t" + DateTime.Now.ToString();
            this.logtext += "\n\n";


            foreach (DeviceItem deviceItem in selection)
            {
                DriveObject driveObject = deviceItem.GetService<DriveObjectContainer>().DriveObjects[0];

                if (driveObject != null)
                {
                    try
                    {

                        //getAttributes(driveObject.Telegrams.Find(TelegramType.SafetyTelegram).GetAttributeInfos(), "safetyAttributes.txt");
                        SDRhelper.StartdriveHelper.AddSafetyTelegram(driveObject, safetyTelegram);
                        SetSafetyTelegramParameters(driveObject, safeMonitoringTime);
                        ChangeTelegramAddress(driveObject.Telegrams.Find(TelegramType.SafetyTelegram), startingAddress);

                        profiSafeTags += "PROFIsafeStatDrive_" + deviceItem.Name + ";\t%I" + startingAddress + ".0\n";
                        profiSafeTags += "PROFIsafeCtrlDrive_" + deviceItem.Name + ";\t%Q" + startingAddress + ".0\n";

                        this.logtext += deviceItem.Name + "\t" + startingAddress + "\n";

                        startingAddress += incr;

                    }
                    catch (Exception)
                    {
                        this.logtext += deviceItem.Name + "\tfail: check if  address is already occupied.\n";
                        MessageBox.Show("FAIL! See log file for more information.");
                        return;

                    }


                }
                else
                {
                    this.logtext += deviceItem.Name + "\tis not a drive object!\n";

                }


                
            }

            WriteLogFile("setTelegramLog.txt", this.logtext);
            WriteLogFile("profiSafeTags.csv", profiSafeTags);

           
           MessageBox.Show("Procedure completed with success, see log file for more information.");
 
        }



        #endregion


        #region hmi


        private static HmiTarget GetTheFirstHmiTarget(Project project)
        {
            if (project == null)
            {
                Console.WriteLine("Project cannot be null");
                throw new ArgumentNullException("project");
            }
            foreach (Device device in project.Devices)
            //This example looks for devices located directly in the project.
            //Devices which are stored in a subfolder of the project will not be affected by this
            //example.
            {
                foreach (DeviceItem deviceItem in device.DeviceItems)
                {
                    DeviceItem deviceItemToGetService = deviceItem as DeviceItem;
                    SoftwareContainer container =
                    deviceItemToGetService.GetService<SoftwareContainer>();
                    if (container != null)
                    {
                        HmiTarget hmi = container.Software as HmiTarget;
                        if (hmi != null)
                        {
                            return hmi;
                        }
                    }
                }
            }
            return null;
        }


        private static string GetPage(string screenName)
        {

            string[] words = screenName.Split('_');

            string pageNumber = words[0].Remove(0, 1);

            string pageText = "";

            for (int i = 1; i < words.Length; i++)
            {
                pageText += words[i] + " ";
            }

            string page = pageNumber + ";\t" + pageText;


            return page;
        }


        private  void GetScreensFromFolder(ScreenUserFolderComposition folders)
        {

            foreach (ScreenFolder folder in folders)
            {
                foreach (Siemens.Engineering.Hmi.Screen.Screen screen in folder.Screens)
                {
                    this.logtext += $"{GetPage(screen.Name)}\n";
                }

                if (folder.Folders != null)
                {
                    GetScreensFromFolder(folder.Folders);
                }

            }
        }


        public void GetPageTextList(MenuSelectionProvider<Device> menuSelectionProvider)

        {


            foreach (Device device in menuSelectionProvider.GetSelection())

            {
                foreach (DeviceItem deviceItem in device.DeviceItems)
                {
                    DeviceItem deviceItemToGetService = deviceItem;

                    SoftwareContainer container = deviceItemToGetService.GetService<SoftwareContainer>();

                    if (container != null)
                    {
                        HmiTarget hmi = container.Software as HmiTarget;

                        if (hmi != null)
                        {

                            this.logtext = "";

                            foreach (Siemens.Engineering.Hmi.Screen.Screen screen in hmi.ScreenFolder.Screens)
                            {
                                this.logtext += $"{GetPage(screen.Name)}\n";
                            }

                            if (hmi.ScreenFolder.Folders != null)
                            {
                                GetScreensFromFolder(hmi.ScreenFolder.Folders);

                            }

                            string path = Path.GetDirectoryName(_tiaportal.Projects[0].Path.ToString()) + @"\" + hmi.Name + ".csv";

                            MessageBox.Show(path);

                            File.WriteAllText(path, this.logtext);


                        }


                    }
                }
            }

        }



        #endregion




        /// <summary>
        /// Called when there is a mousover the button at a DeviceItem.
        /// It will be used to enable the button.
        /// </summary>
        /// <typeparam name="menuSelectionProvider">
        /// here, the same type as in addInRootSubmenu.Items.AddActionItem
        /// must be used -> here it is <DeviceItem>
        /// </typeparam>
        //private MenuStatus OnCanSomething(MenuSelectionProvider
        //    <DeviceItem> menuSelectionProvider)
        //{
        //    //enable the button
        //    return MenuStatus.Enabled;
        //}

        /// <summary>
        /// The method contains the program code of the TIA Add-In.
        /// Called when the button will be pressed on project level.
        /// </summary>
        /// <typeparam name="menuSelectionProvider">
        /// here, the same type as in addInRootSubmenu.Items.AddActionItem
        /// must be used -> here it is <Project>
        /// </typeparam>
        //private void OnClickProject(MenuSelectionProvider<Project> 
        //    menuSelectionProvider)
        //{
        //    //Do Nothing on Project level
        //}

        ///// <summary>
        ///// Called when there is a mousover the button at the Project 
        ///// Level. It will be used to disable the button because no 
        ///// action should be performed on project level.
        ///// </summary>
        ///// <typeparam name="menuSelectionProvider">
        ///// here, the same type as in addInRootSubmenu.Items.AddActionItem
        ///// must be used -> here it is <Project>
        ///// </typeparam>
        //private MenuStatus OnStatusUpdateProject(MenuSelectionProvider
        //    <Project> menuSelectionProvider)
        //{
        //    //disable the button
        //    return MenuStatus.Disabled;
        //}
    }
}