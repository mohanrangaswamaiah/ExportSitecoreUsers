
namespace SitecoreCustom.Web.Extensions.Commands
{
    using Sitecore.Shell.Framework.Commands;
    using System;
    using Sitecore.Web.UI.Sheer;
    using System.IO;
    using Sitecore.SecurityModel;
    using System.Text;
	using Sitecore.Diagnostics;
 
    public class ExportSitecoreUsers : Command
    {
        string filepath = string.Empty;  
        public override void Execute(CommandContext context)
        {
            try
            {
                CreateFile();
                if (!string.IsNullOrWhiteSpace(filepath))
                    SheerResponse.Download(filepath);

            }
            catch (Exception ex)
            {
				Log.Error(ex.Message, ex);
                SheerResponse.Alert("Error Occurred: Please try again");
            }

        }

        /// <summary>
        /// Method to create csv file with list of users.
        /// </summary>         
        private void CreateFile()
        {

            string fileExtension = ".csv";
            string path = Sitecore.Configuration.Settings.DataFolder + "Export User";

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            string fileName = "Sitecore_User";
            fileName = fileName + "-" + DateTime.Now.ToString("MM-dd-yyyy_HH-mm-ss");
            _filepath = path + @"\" + fileName + fileExtension;

          
             using (new SecurityDisabler())
            {
                var users = Sitecore.Security.Accounts.UserManager.GetUsers();

                StringBuilder records = new StringBuilder();
                records.Append(string.Format("{0},{1},{2},{3},{4},{5},{6},{7}",
                                                         "User Name",
                                                         "Domain",
                                                         "Fully Qualified Name",
                                                         "Full Name",
                                                         "Email",
                                                         "Comment",
                                                         "Language",
                                                         "Locked"));  

                foreach (var user in users)
                {
                    try
                    {
                        if (user != null)
                        {
                            records.Append(Environment.NewLine);
                            records.Append(string.Format("{0},{1},{2},{3},{4},{5},{6},{7}",
                                                       user.LocalName,
                                                       user.Domain.Name,
                                                       user.DisplayName,
                                                       user.Profile.FullName,
                                                       user.Profile.Email,
                                                       user.Profile.Comment,
                                                       user.Profile.ClientLanguage,
                                                       user.Profile.State));
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Log(ex.Message, EnumUtil.LoggingType.Error, EnumUtil.ExceptionTier.None, ex);
                    }
                }



                using (var writer = new StreamWriter(_filepath, false))
                {                   

                    writer.WriteLine(records.ToString());
                    writer.Close();
                }
            }

        }


    }
}
