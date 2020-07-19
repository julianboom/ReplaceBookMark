using Aspose.Words;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DriverHelper
{
    class AsposeHelper
    {
        #region 公共函数
        private DocumentBuilder oDocBuilder; //   a   reference   to   Word   Builder   
        private Document oDoc; //   a   reference   to   the   document
        private int Docversion = 2003; // a editon to the saved document

        public void OpenDocument(string fileName)
        {
            if (File.Exists(fileName))
            {
                oDoc = new Document(fileName);
            }
            else
            {
                oDoc = new Document();
            }
        }

        public void CreateDocument()
        {
            oDoc = new Document();
        }
        
        public void BuildDocument()
        {
            oDocBuilder = new DocumentBuilder(oDoc);
        }

        public void SaveDocument(string fileName)
        {
            if (this.Docversion >= 2007)
            {
                oDoc.Save(fileName, SaveFormat.Docx);

            }
            else
            {
                oDoc.Save(fileName, SaveFormat.Doc);
            }
        }
        #endregion

        /// <summary>
        /// 通过书签名定位书签并调整书签名
        /// </summary>
        /// <param name="oldBookMarkName">原始书签名</param>
        /// <param name="newBookmarkName">新书签名</param>
        public void RepalceBookMark(string oldBookMarkName,string newBookmarkName)
        {
            // Use the indexer of the Bookmarks collection to obtain the desired bookmark.
            Bookmark bookmark = oDoc.Range.Bookmarks[oldBookMarkName];

            // Set the name of the bookmark.
            if(bookmark != null)
            {
                bookmark.Name = newBookmarkName;
            }
        }
    }
}
