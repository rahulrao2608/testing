namespace EmailDownloadService
{
    partial class ProjectInstaller
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.SPILisaEmailDownloader = new System.ServiceProcess.ServiceProcessInstaller();
            this.EmailDownloadService = new System.ServiceProcess.ServiceInstaller();
            // 
            // SPILisaEmailDownloader
            // 
            this.SPILisaEmailDownloader.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
            this.SPILisaEmailDownloader.Password = null;
            this.SPILisaEmailDownloader.Username = null;
            // 
            // EmailDownloadService
            // 
            this.EmailDownloadService.ServiceName = "EmailDownloadService";
            this.EmailDownloadService.StartType = System.ServiceProcess.ServiceStartMode.Automatic;
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.SPILisaEmailDownloader,
            this.EmailDownloadService});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller SPILisaEmailDownloader;
        private System.ServiceProcess.ServiceInstaller EmailDownloadService;
    }
}