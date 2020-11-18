using ReplaceBookMark.DriveHelper;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DriverHelper
{
    public class CommonHelper
    {
        private MySqlHelper mySqlHelper;
        public CommonHelper()
        {
            mySqlHelper = new MySqlHelper();
        }
        #region 文件扫描方法

        /// <summary>
        /// 文件信息集合
        /// </summary>
        private List<FileInfo> fileInfos = new List<FileInfo>();


        public List<FileInfo> GetFileInfos(string path)
        {
            fileInfos.Clear();//清空文件名集合

            try
            {
                getDirectory(path);
                return fileInfos;
            }
            catch
            {
                return fileInfos;
            }
        }

        /// <summary>
        /// 获得指定路径下所有文件名
        /// </summary>
        /// <param name="path">文件写入流</param>
        /// <param name="indent">输出时的缩进量</param>
        public  void getFileName(string path)
        {
            DirectoryInfo root = new DirectoryInfo(path);
            foreach (FileInfo f in root.GetFiles())
            {
                fileInfos.Add(f);
            }
        }

        /// <summary>
        /// 获得指定路径下所有子目录名
        /// </summary>
        /// <param name="sw">文件写入流</param>
        /// <param name="path">文件夹路径</param>
        /// <param name="indent">输出时的缩进量</param>
        public void getDirectory(string path)
        {
            getFileName(path);
            DirectoryInfo root = new DirectoryInfo(path);
            foreach (DirectoryInfo d in root.GetDirectories())
            {
                getDirectory(d.FullName);
            }
        }
        #endregion


        #region 开始替换书签
        public void startReplace(string groupID,string DirPath)
        {
            string oldDirName = new DirectoryInfo(DirPath).Name;
            string newDirName = "ReplaceResult";

            string mapSQL = $"SELECT OldBookMarkName,NewBookMarkName FROM bookmarkmap WHERE GroupID = '{groupID}'";
            DataTable mapDt = mySqlHelper.ExecuteDataTable(mapSQL, null);
            List<FileInfo> fileInfos = GetFileInfos(DirPath);
            AsposeHelper asposeHelper = new AsposeHelper();

            foreach (FileInfo fileInfo in fileInfos)
            {
                try
                {
                    if (!fileInfo.Extension.ToLower().Contains("doc") || fileInfo.Name.Contains("~$"))
                    {
                        continue;
                    }
                    asposeHelper.OpenDocument(fileInfo.FullName);

                    foreach (DataRow dr in mapDt.Rows)
                    {
                        asposeHelper.RepalceBookMark(dr["OldBookMarkName"].ToString(), dr["NewBookMarkName"].ToString());
                    }

                    asposeHelper.SaveDocument(fileInfo.FullName.Replace(oldDirName, newDirName));
                }
                catch (Exception e)
                {
                    LogHelper.DoErrorLog($"报错信息：{e.ToString()}\r\n 文件信息：{fileInfo.FullName}");
                }
            }
        }

        #endregion

    }
}