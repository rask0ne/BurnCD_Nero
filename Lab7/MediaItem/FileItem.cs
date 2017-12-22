using System;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using IMAPI2.Interop;

namespace IMAPI2.MediaItem
{

    class FileItem : IMediaItem
    {
        private Int64 m_fileLength = 0;

        public FileItem(string path)
        {
            if (!File.Exists(path))
            {
                throw new FileNotFoundException("The file added to FileItem was not found!", path);
            }

            filePath = path;


            FileInfo fileInfo = new FileInfo(filePath);
            displayName = fileInfo.Name;
            m_fileLength = fileInfo.Length;

            //
            // Get the File icon
            //
            SHFILEINFO shinfo = new SHFILEINFO();
            IntPtr hImg = Win32.SHGetFileInfo(filePath, 0, ref shinfo,
                (uint)Marshal.SizeOf(shinfo), Win32.SHGFI_ICON | Win32.SHGFI_SMALLICON);

            if (shinfo.hIcon != null)
            {
                //The icon is returned in the hIcon member of the shinfo struct
                System.Drawing.IconConverter imageConverter = new System.Drawing.IconConverter();
                System.Drawing.Icon icon = System.Drawing.Icon.FromHandle(shinfo.hIcon);
                try
                {
                    fileIconImage = (System.Drawing.Image)
                        imageConverter.ConvertTo(icon, typeof(System.Drawing.Image));
                }
                catch (NotSupportedException)
                {
                }

                Win32.DestroyIcon(shinfo.hIcon);
            }
        }

        public Int64 SizeOnDisc
        {
            get
            {
                ulong cluster = 0;
                string fileName = Path;
                string driveLetter = System.IO.Path
                    .GetPathRoot(fileName)
                    .TrimEnd('\\');

                string queryString = string.Format("SELECT BlockSize, NumberOfBlocks " +
                                                   "  FROM Win32_Volume " +
                                                   "  WHERE DriveLetter = '{0}'", driveLetter);

                using (ManagementObjectSearcher searcher = new
                    ManagementObjectSearcher(queryString))
                {
                    foreach (ManagementObject item in searcher.Get())
                    {
                        cluster = (ulong)item["BlockSize"];
                        break;
                    }
                }
                var clusterSize = Convert.ToInt64(cluster);
                var size = m_fileLength;
                long bytes = ((size + clusterSize - 1) / clusterSize) * clusterSize;

                return bytes;
            }
        }

        public string Path
        {
            get
            {
                return filePath;
            }
        }
        private string filePath;

        public System.Drawing.Image FileIconImage
        {
            get
            {
                return fileIconImage;
            }
        }
        private System.Drawing.Image fileIconImage = null;

        public override string ToString()
        {
            return displayName;
        }
        private string displayName;

        public bool AddToFileSystem(IFsiDirectoryItem rootItem)
        {
            IStream stream = null;

            try
            {
                Win32.SHCreateStreamOnFile(filePath, Win32.STGM_READ | Win32.STGM_SHARE_DENY_WRITE, ref stream);

                if (stream != null)
                {
                    rootItem.AddFile(displayName, stream);
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error adding file",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (stream != null)
                {
                    Marshal.FinalReleaseComObject(stream);
                }
            }

            return false;
        }
    }
}