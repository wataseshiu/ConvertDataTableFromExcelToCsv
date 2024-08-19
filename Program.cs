using System.Diagnostics;
using ClosedXML.Excel;
namespace QuestDataConverter
{
    partial class Program
    {
        /// <summary>
        /// 各種コンバートで使うパラメータをまとめる
        /// </summary>
        struct ConvertParameter
        {
            //コマンドライン引数から渡される文字列と比較して、どのConvertParameterを使用するかを決める
            public string convertType;
            //コンバート対象となるExcelのファイルを指定する
            public string filePath;
            //コンバート対象となるExcelファイルのシートindex
            public int workSheetIndex;
            //使用するConverterのインターフェース
            public IConverter converter;

            public ConvertParameter(string convertType, string filePath, int workSheetIndex, IConverter converter)
            {
                this.convertType = convertType;
                this.filePath = filePath;
                this.workSheetIndex = workSheetIndex;
                this.converter = converter;
            }
        }

        static void Main(string[] args)
        {
            List<ConvertParameter> convertParameters = new List<ConvertParameter>{
            new ConvertParameter("QuestData", args[1] + @"クエストデータ.xlsm", 1, new QuestDataConverter()),
            new ConvertParameter("QuestBattleMemberData", args[1] + @"クエストデータ.xlsm", 1, new QuestBattleMemberDataConverter()),
            new ConvertParameter("QuestPremiseData", args[1] + @"クエストデータ.xlsm", 6, new QuestPremiseDataConverter()),
            new ConvertParameter("QuestGroupData", args[1] + @"クエストデータ.xlsm", 2, new QuestGroupDataConverter()),
            new ConvertParameter("QuestCategoryData", args[1] + @"クエストデータ.xlsm", 3, new QuestCategoryDataConverter()),
            new ConvertParameter("QuestRewardData", args[1] + @"クエストデータ.xlsm", 4, new QuestRewardDataConverter()),
            //~~省略~~//
            };

            //Lotteryとログインボーナスはargsが1個多い、コンバートエラー回避のためにargsの数を見てから追加する
            if (args.Length >= 3)
            {
                convertParameters.Add(new ConvertParameter("LotteryTableData", args[1] + Path.GetFileName(args[2]), 1, new LotteryTableDataConverter(Path.GetFileName(args[2]))));
                convertParameters.Add(new ConvertParameter("LoginBonusShowcaseData", args[1] + Path.GetFileName(args[2]), 1, new LoginBonusShowcaseDataConverter(Path.GetFileName(args[2]))));
            }

            //コマンドライン引数から取得するconvertTypeから、今回使用するConvertParameterを取得する
            ConvertParameter convertParameter = convertParameters.Single(param => param.convertType == args[0]);

            IXLRange tbl;

            ///ファイルストリームを使って読み込み専用でエクセルファイルを開くと、Excelでファイルを開いていてもエラーが起きなくて済む
            using (var fs = new FileStream(convertParameter.filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var wb = new XLWorkbook(fs))
                {
                    ///ワークシートの読み込み
                    tbl = wb.Worksheet(convertParameter.workSheetIndex).RangeUsed();
                    ///データコンバート
                    convertParameter.converter.ConvertCsv(tbl, args[1]);
                }
            }
        }
    }
}