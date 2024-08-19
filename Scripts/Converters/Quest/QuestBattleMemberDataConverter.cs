using ClosedXML.Excel;
namespace QuestDataConverter
{
    class QuestBattleMemberDataConverter : BaseConverter, IConverter
    {
        List<int> pickupColumnNumberList = new List<int>();
        private string filePath = @"QuestBattleMemberData.csv";

        public void ConvertCsv(IXLRange data, string fileDirectory)
        {
            ///データとして抽出したい列のEnum番号リストを作成
            pickupColumnNumberList.Add((int)QuestBattleMemberDataColumnNumber.data1);
            pickupColumnNumberList.Add((int)QuestBattleMemberDataColumnNumber.data2);
            pickupColumnNumberList.Add((int)QuestBattleMemberDataColumnNumber.data3);

            var textData = Convert(data, pickupColumnNumberList);
            filePath = GetParentDirectory(fileDirectory) + filePath;

            //整形した形のままCSVを作成する
            WriteCSV(textData, filePath);
        }

        public override List<string> SetHeader()
        {
            List<string> header = new() { "questLabel,battleMember" };
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
            //QuestLabelは後で使うのでとっておく
            questLabel = pickupCellData.First().ToString();
            //NPC1人目のデータの時は空欄であろうが出力する
            output.Add(questLabel + "," + pickupCellData[1].ToString());

            foreach (var data in pickupCellData.Skip(2))
            {
                //それ以外の時は空欄の場合は出力しない
                if (data.ToString() != "")
                {
                    output.Add(questLabel + "," + data.ToString());
                }
            }
        }
    }
}