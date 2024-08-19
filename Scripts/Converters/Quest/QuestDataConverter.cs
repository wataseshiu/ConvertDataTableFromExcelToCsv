using ClosedXML.Excel;
using System;
namespace QuestDataConverter
{
    class QuestDataConverter : BaseConverter, IConverter
    {
        List<int> pickupColumnNumberList = new();
        private string filePath = @"QuestData.csv";

        public void ConvertCsv(IXLRange data, string fileDirectory)
        {
            ///データとして抽出したい列のEnum番号リストを作成
            pickupColumnNumberList.Add((int)QuestDataColumnNumber.data3_1);
            pickupColumnNumberList.Add((int)QuestDataColumnNumber.data3_2);
            pickupColumnNumberList.Add((int)QuestDataColumnNumber.data3_3);

            var textData = Convert(data, pickupColumnNumberList);
            filePath = GetParentDirectory(fileDirectory) + filePath;

            //整形した形のままCSVを作成する
            WriteCSV(textData, filePath);
        }

        public override List<string> SetHeader()
        {
            List<string> header = new() { "data3_1,data3_2,data3_3" };
            return header;
        }

        /// <summary>
        /// 出力する必要のあるデータが入ったセル群を処理する
        /// </summary>
        /// <param name="pickupCellData">行の中の必要なセルだけが入ったデータ配列</param>
        /// <param name="pickUpList">このコンバートで必要な行の番号</param>
        /// <param name="output">出力するテキスト</param>
        protected override void ConvertSub(XLCellValue[] pickupCellData, List<int> pickUpList, List<string> output)
        {
            int count = 0;
            string temp = "";
            //QuestDataでは純粋にコンマ区切りで羅列する
            foreach (var data in pickUpList)
            {
                temp += pickupCellData[count++].ToString() + ",";
            }
            temp = temp.Remove(temp.Length - 1);
            output.Add(temp);
        }
    }
}