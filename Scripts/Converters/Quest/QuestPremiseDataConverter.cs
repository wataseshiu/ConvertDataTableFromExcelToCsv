using ClosedXML.Excel;
namespace QuestDataConverter
{
    class QuestPremiseDataConverter : BaseConverter, IConverter
    {
        List<int> pickupColumnNumberList = new();
        private string filePath = @"QuestPremiseData.csv";

        public void ConvertCsv(IXLRange data, string fileDirectory)
        {
            ///データとして抽出したい列のEnum番号リストを作成
            pickupColumnNumberList.Add((int)QuestPremiseDataColumnNumber.data5_1);
            pickupColumnNumberList.Add((int)QuestPremiseDataColumnNumber.data5_2);
            pickupColumnNumberList.Add((int)QuestPremiseDataColumnNumber.data5_3);

            var textData = Convert(data, pickupColumnNumberList);
            filePath = GetParentDirectory(fileDirectory) + filePath;

            //整形した形のままCSVを作成する
            WriteCSV(textData, filePath);
        }

        public override List<string> SetHeader()
        {
            List<string> header = new() { "data5_1,data5_2,data5_3" };
            return header;
        }

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