using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2013.PowerPoint.Roaming;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Text;
namespace QuestDataConverter
{
    abstract class BaseConverter
    {
        /// <summary>
        /// ヘッダー情報を差し込む場所
        /// </summary>
        public string questLabel = "";

        /// <summary>
        /// コンバート時に1つ前の行のセル情報を使う時があるので用意しておく
        /// </summary>
        protected List<XLCellValue> previousCells = new();

        /// <summary>
        /// CSVの１行目に必要になるヘッダー用文字列を作成する
        /// </summary>
        /// <returns>ヘッダー用文字列リスト</returns>
        public abstract List<string> SetHeader();

        /// <summary>
        /// 解析内容をもとにCSVを出力する
        /// </summary>
        /// <param name="output">出力するテキスト</param>
        /// <param name="filePath">出力先</param>
        public static void WriteCSV(List<string> output, string filePath)
        {
            using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                foreach (var line in output)
                {
                    sw.WriteLine(line);
                }
            }
        }
        protected abstract void ConvertSub(XLCellValue[] pickupCellData, List<int> pickUpList, List<string> output);

        /// <summary>
        /// 必要なパラメータ列のデータのみをピックアップする
        /// </summary>
        /// <param name="data">ワークシートの情報</param>
        /// <param name="pickUpList">必要なデータ列番号のリスト</param>
        /// <returns>CSVとして出力したい文字列</returns>
        public List<string> Convert(IXLRange data, List<int> pickUpList)
        {
            List<string> output = new();
            output = SetHeader();

            //最初の３行はデータ行ではないから無視して、それ以降の各行ごとに処理をする
            foreach (var rowData in data.Rows().Skip(3))
            {
                //A列がコメントアウト「//」だった場合、その行はデータとして扱わずスキップする
                if (rowData.Cell(1).Value.ToString() == "//")
                    continue;
                List<XLCellValue> Cells = new();
                ///各行の不要な行を除いたデータを作る
                foreach (var cellData in rowData.Cells())
                {
                    //cellDataの列番号が抽出したい列のどれかなら抽出する
                    if (pickUpList.Any(param => param == cellData.Address.ColumnNumber))
                        Cells.Add(cellData.Value);
                }

                //行のなかの必要なセルが集まったデータを解釈する
                ConvertSub(Cells.ToArray(), pickUpList, output);
                //この行のデータを次の行で使うことがあるので保持しておく
                previousCells = Cells;
            }
            return output;
        }

        public string GetParentDirectory(string dir)
        {
            DirectoryInfo directoryInfoParent = new DirectoryInfo(dir);
            return directoryInfoParent.Parent.FullName + "\\";
        }
    }
}