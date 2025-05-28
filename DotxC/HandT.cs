using Microsoft.Office.Interop.Word;
using System.Collections.Generic; 

namespace DotxC
{
    public static class HandT
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="tpath">模板位置</param>
        /// <param name="spath">保存位置</param>
        /// <param name="pairs">键值对</param>
        public static void Replace(string tpath, string spath, Dictionary<string, string> pairs)
        {
            // 创建一个新的Word文档对象
            Application wordApp = new Application();
            Document wordDoc = wordApp.Documents.Add();
            // 打开Word模板文件
            object oTemplate = tpath;
            wordDoc = wordApp.Documents.Add(ref oTemplate);
            // 将模板文件中的内容复制到新的Word文档对象中
            wordDoc.Content.InsertParagraphAfter();
            wordDoc.Content.InsertFile(oTemplate.ToString());
            // 将新的内容替换为您需要的内容
            foreach (var pair in pairs)
            {
                FindAndReplace(wordDoc, pair.Key, pair.Value);
            }
            // 保存新的Word文档
            object oFilename = spath;
            wordDoc.SaveAs2(ref oFilename);
            wordApp.Quit();
        }

        static void FindAndReplace(Document doc, string findText, string replaceWithText)
        {
            // 将新的内容替换为您需要的内容
            foreach (Range range in doc.StoryRanges)
            {
                range.Find.Execute(FindText: findText, ReplaceWith: replaceWithText,
                    MatchWholeWord: true, Replace: WdReplace.wdReplaceAll);
            }
        }
    }
}