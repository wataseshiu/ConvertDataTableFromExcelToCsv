using ClosedXML.Excel;

interface IConverter
{
    /// <summary>
    /// IConverter経由でProgram.csから呼び出されるCSVコンバート
    /// </summary>
    /// <param name="data">ワークシートの情報</param>
    /// <param name="fileDirectory">エクセルデータの配置ディレクトリ</param>
    void ConvertCsv(IXLRange data, string fileDirectory);
}