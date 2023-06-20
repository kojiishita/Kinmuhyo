using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using System.Security;

namespace Kinmuhyo
{
    internal class Program
    {
        private static List<Jisseki> _jissekiList = new();
        private static List<Genka> _genkaList = new();

        private static string _kinmuhyoFolder = null!;
        private static string _genkahyoFolder = null!;
        private static string _outputFile = null!;
        private static int _nen;
        private static int _jissekiHaneiDay;

        static void Main(string[] args)
        {
            var configuration = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json")
                .Build();

            // 勤務表フォルダ
            _kinmuhyoFolder = configuration["KinmuhyoFolder"];
            
            // 原価表フォルダ
            _genkahyoFolder = configuration["GenkaFolder"];

            // 出力ファイル
            _outputFile = configuration["OutputFile"];

            // 集計年
            _nen = int.Parse(configuration["Nen"]);

            // 実績反映日
            _jissekiHaneiDay = int.Parse(configuration["JissekiHaneiDay"]);

            // 勤務表フォルダのチェック
            if (Directory.Exists(_kinmuhyoFolder))
            {
                // 勤務表フォルダの*.xlsmファイルをすべて処理する
                foreach (var file in Directory.EnumerateFiles(_kinmuhyoFolder, "*.xlsm", SearchOption.TopDirectoryOnly))
                {
                    try
                    {
                        ReadKinmuhyo(file);
                    }
                    catch (SecurityException)
                    {
                        Console.WriteLine($"は処理できません。({file})");
                    }
                }

                // 結果をコンソール出力
                foreach (var jisseki in _jissekiList)
                {
                    Console.WriteLine($"{jisseki.ShainBango},{jisseki.ShainName},{jisseki.Nengetstu},{jisseki.KyakusakiCode},{jisseki.KyakusakiName},{jisseki.Jikan}");
                }
            }
            else
            {
                Console.WriteLine($"パスが存在しないため処理をスキップします。({_kinmuhyoFolder}) ");
            }

            // 原価表フォルダのチェック
            if (Directory.Exists(_genkahyoFolder))
            {
                // 原価表フォルダの売上原価表*ファイルをすべて処理する
                foreach (var file in Directory.EnumerateFiles(_genkahyoFolder, "売上原価表*", SearchOption.TopDirectoryOnly))
                {
                    try
                    {
                        ReadGenkahyo(file);
                    }
                    catch (SecurityException)
                    {
                        Console.WriteLine($"は処理できません。({file})");
                    }
                }

                // 結果をコンソール出力
                foreach (var genka in _genkaList)
                {
                    Console.WriteLine($"{genka.Nengetstu},{genka.Bunrui},{genka.Busho},{genka.KyakusakiName},{genka.Keiyaku},{genka.Yotei},{genka.Jisseki}");
                }
            }
            else
            {
                Console.WriteLine($"パスが存在しないため処理をスキップします。({_genkahyoFolder}) ");
            }

            // 結果をExcelに出力
            WriteResult();
        }

        private static void ReadKinmuhyo(string path)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var xlsxFile = new FileInfo(path);
            using (var package = new ExcelPackage(xlsxFile, "0901"))
            {
                var worksheet = package.Workbook.Worksheets["プロジェクト月計"];
                if (worksheet != null)
                {
                    worksheet.Cells[2, 2].Style.Numberformat.Format = "yyyy/mm";
                    var nengetsu = worksheet.Cells[2, 2].Text;
                    var name = worksheet.Cells[2, 3].Text;
                    var shainBango = worksheet.Cells[2, 4].Text;

                    Console.WriteLine($"プロジェクト月計シートを処理します({nengetsu},{name},{shainBango})");

                    for (int i = 4; i <= 100; i++)
                    {
                        if (worksheet.Cells[i, 2].Text == string.Empty || worksheet.Cells[i, 2].Text == "0")
                        {
                            break;
                        }

                        if (worksheet.Cells[i, 4].Value.ToString() != "0")
                        {
                            Console.WriteLine($"{name}\t{worksheet.Cells[i, 2].Value}\t{worksheet.Cells[i, 4].Value}");
                            _jissekiList.Add(new Jisseki()
                            {
                                ShainBango = int.Parse(shainBango),
                                ShainName = name,
                                Nengetstu = DateTime.ParseExact(nengetsu, "yyyy/MM", null),
                                KyakusakiCode = worksheet.Cells[i, 5].Value.ToString()!,
                                KyakusakiName = worksheet.Cells[i, 2].Value.ToString()!,
                                Jikan = decimal.Parse(worksheet.Cells[i, 4].Value.ToString()!),
                            });
                        }
                    }
                }
            }
        }

        private static void ReadGenkahyo(string path)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var xlsxFile = new FileInfo(path);
            using (var package = new ExcelPackage(xlsxFile))
            {
                for (var tsuki = 1; tsuki <= 12; tsuki++)
                {
                    GetGenkahyo(package, tsuki);
                }
            }
        }

        private static void GetGenkahyo(ExcelPackage pkg, int tsuki)
        {
            var sheetName = GetTsukiSheetName(tsuki);

            // 指定月のシートを取得
            ExcelWorksheet? ws;
            ws = pkg.Workbook.Worksheets[sheetName.hankakuSheet];
            ws ??= pkg.Workbook.Worksheets[sheetName.zenkakuSheet];

            // 使用されているセル範囲を取得する
            var usedRange = ws.Cells[ws.Dimension.Address];

            // 必要な項目を取得
            var curBunrui = string.Empty;
            var curBusho = string.Empty;
            for (var row = 2; row <= usedRange.Rows; row++)
            {
                // 分類
                var bunrui = ws.Cells[row, 1].Text;
                if (!string.IsNullOrEmpty(bunrui) && curBunrui != bunrui)
                {
                    curBunrui = bunrui;
                }

                // 部署
                var busho = ws.Cells[row, 3].Text;
                if (!string.IsNullOrEmpty(busho) && curBusho != busho)
                {
                    curBusho = busho;
                }

                // 客先名
                var kyakusakimei = ws.Cells[row, 4].Text;

                // 契約
                var keiyaku = ws.Cells[row, 5].Text;

                // 予定
                var yotei = ws.Cells[row, 9].Value;

                // 実績
                var jisseki = ws.Cells[row, 10].Value;

                // 客先名と契約が設定されているなら処理する
                if (!string.IsNullOrEmpty(ws.Cells[row, 4].Text) && !string.IsNullOrEmpty(ws.Cells[row, 5].Text))
                {
                    // 年月
                    DateTime nengetsu;
                    if (tsuki >= 1 && tsuki <= 6)
                    {
                        nengetsu = new DateTime(_nen + 1, tsuki, 1);
                    }
                    else
                    {
                        nengetsu = new DateTime(_nen, tsuki, 1);
                    }

                    _genkaList.Add(new Genka()
                    {
                        Nengetstu = nengetsu,
                        Bunrui = curBunrui,
                        Busho = curBusho,
                        KyakusakiName = kyakusakimei,
                        Keiyaku = keiyaku,
                        Yotei = yotei != null ? decimal.Parse(yotei.ToString()!) : 0m,
                        Jisseki = jisseki != null ? decimal.Parse(jisseki.ToString()!) : 0m,
                    });
                }
            }
        }

        private static void WriteResult()
        {
            using var package = new ExcelPackage(_outputFile);
            WriteResultKinmuhyo(package);
            WriteResultGenka(package);
            package.SaveAs(new FileInfo(_outputFile));
        }

        private static void WriteResultKinmuhyo(ExcelPackage pkg)
        {
            // 単価設定を取得
            var tanka = GetTanka(pkg);

            // 顧客設定を取得
            var kokyaku = GetKokyaku(pkg);

            // 社員別実績時間シートを取得する
            var ws = pkg.Workbook.Worksheets["社員別実績時間明細"];

            // 使用されているセル範囲を取得する
            var usedRange = ws.Cells[ws.Dimension.Address];

            // セル範囲をクリアする
            usedRange.Clear();

            // ヘッダを設定
            int row = 1;
            ws.Cells[row, 1].Value = "社員番号";
            ws.Cells[row, 2].Value = "社員名";
            ws.Cells[row, 3].Value = "年月";
            ws.Cells[row, 4].Value = "客先コード";
            ws.Cells[row, 5].Value = "客先名";
            ws.Cells[row, 6].Value = "作業時間";
            ws.Cells[row, 7].Value = "作業時間合計";
            ws.Cells[row, 8].Value = "換算単価";
            ws.Cells[row, 9].Value = "直接費";
            row++;

            // 社員番号ごとの時間合計を得る
            var sumJikan =
                _jissekiList.GroupBy(group => group.ShainBango).Select(e => new { e.Key, Total = e.Sum(group => group.Jikan) });

            // 結果を設定
            foreach (var jisseki in _jissekiList)
            {
                var totalJikan =
                    sumJikan.Where(e => e.Key == jisseki.ShainBango).Select(e => e.Total).First();
                var setteiTanka =
                    tanka.Where(e => e.ShainBango == jisseki.ShainBango).Select(e => e.SetteiTanka).First();

                ws.Cells[row, 1].Value = jisseki.ShainBango;
                ws.Cells[row, 2].Value = jisseki.ShainName;
                ws.Cells[row, 3].Value = jisseki.Nengetstu;
                var cellStyle = ws.Cells[row, 3].Style;
                cellStyle.Numberformat.Format = "yyyy/mm";
                ws.Cells[row, 4].Value = jisseki.KyakusakiCode;
                ws.Cells[row, 5].Value = jisseki.KyakusakiName;
                ws.Cells[row, 6].Value = jisseki.Jikan;
                ws.Cells[row, 7].Value = totalJikan;
                var kansanTanka = 0m;
                if (totalJikan > 150)
                {
                    // 設定単価*合計時間/150
                    kansanTanka = setteiTanka * totalJikan / 150m;
                }
                else
                {
                    // 設定単価
                    kansanTanka = totalJikan;
                }
                ws.Cells[row, 8].Value = kansanTanka;

                // 換算単価*時間/合計時間
                ws.Cells[row, 9].Value = kansanTanka * jisseki.Jikan / totalJikan;

                row++;
            }
        }

        private static void WriteResultGenka(ExcelPackage pkg)
        {
            // 顧客設定を取得
            var kokyaku = GetKokyaku(pkg);

            // 顧客別ジョブ採算シートを取得する
            var ws = pkg.Workbook.Worksheets["顧客別ジョブ採算明細"];

            // 使用されているセル範囲を取得する
            var usedRange = ws.Cells[ws.Dimension.Address];

            // セル範囲をクリアする
            usedRange.Clear();

            // ヘッダを設定
            int row = 1;
            ws.Cells[row, 1].Value = "部署";
            ws.Cells[row, 2].Value = "年月";
            ws.Cells[row, 3].Value = "客先コード";
            ws.Cells[row, 4].Value = "客先名";
            ws.Cells[row, 5].Value = "契約";
            ws.Cells[row, 6].Value = "売上";
            ws.Cells[row, 7].Value = "仕入";
            ws.Cells[row, 8].Value = "直接原価";
            row++;

            // 部署、年月、客先名、契約、分類で集計
            var sumGenka =
                _genkaList.Where(e => e.Bunrui == "売上" | e.Bunrui == "仕入")
                .GroupBy(group => new { group.Busho, group.Nengetstu, group.KyakusakiName, group.Keiyaku })
                .Select(e => new 
                { 
                    e.Key.Busho,
                    e.Key.Nengetstu,
                    e.Key.KyakusakiName,
                    e.Key.Keiyaku,
                    UriageYotei = e.Sum(x => x.Bunrui == "売上" ? x.Yotei : 0m),
                    UriageJisseki = e.Sum(x => x.Bunrui == "売上" ? x.Jisseki : 0m),
                    ShiireYotei = e.Sum(x => x.Bunrui == "仕入" ? x.Yotei : 0m),
                    ShiireJisseki = e.Sum(x => x.Bunrui == "仕入" ? x.Jisseki : 0m),
                });

            // 結果を設定
            foreach (var genka in sumGenka)
            {
                ws.Cells[row, 1].Value = genka.Busho;
                ws.Cells[row, 2].Value = genka.Nengetstu;
                var cellStyle = ws.Cells[row, 2].Style;
                cellStyle.Numberformat.Format = "yyyy/mm";
                var kyakusakiCode = kokyaku.FirstOrDefault(e => e.Name == genka.KyakusakiName)?.Code;
                ws.Cells[row, 3].Value = kyakusakiCode;
                ws.Cells[row, 4].Value = genka.KyakusakiName;
                ws.Cells[row, 5].Value = genka.Keiyaku;

                // 原価表の年月+設定ファイルの実績反映日が現在日付以上なら実績を反映
                decimal kingaku;
                if (DateTime.Today.Date <= new DateTime(genka.Nengetstu.Year, genka.Nengetstu.Month, _jissekiHaneiDay))
                {
                    ws.Cells[row, 6].Value = genka.UriageJisseki;
                    ws.Cells[row, 7].Value = genka.ShiireJisseki;
                }
                else
                {
                    ws.Cells[row, 6].Value = genka.UriageYotei;
                    ws.Cells[row, 7].Value = genka.ShiireYotei;
                }

                // 次行
                row++;
            }
        }

        /// <summary>
        /// 顧客を取得します
        /// </summary>
        /// <param name="pkg">ExcelPackage</param>
        /// <returns>単価</returns>
        private static IEnumerable<Kokyaku> GetKokyaku(ExcelPackage pkg)
        {
            // シートを取得
            var ws = pkg.Workbook.Worksheets["顧客設定"];

            // 使用されているセル範囲を取得する
            var usedRange = ws.Cells[ws.Dimension.Address];

            // 結果を取得
            var result = new List<Kokyaku>();
            for (var row = 2; row <= usedRange.Rows; row++)
            {
                // コードが空なら終わる
                if (string.IsNullOrEmpty(ws.Cells[row, 2].Text))
                {
                    break;
                }

                result.Add(new Kokyaku()
                {
                    Code = ws.Cells[row, 2].Text,
                    Name = ws.Cells[row, 1].Text,
                });
            }

            return result;
        }

        /// <summary>
        /// 単価を取得します
        /// </summary>
        /// <param name="pkg">ExcelPackage</param>
        /// <returns>単価</returns>
        private static IEnumerable<Tanka> GetTanka(ExcelPackage pkg)
        {
            // シートを取得
            var ws = pkg.Workbook.Worksheets["単価設定"];

            // 使用されているセル範囲を取得する
            var usedRange = ws.Cells[ws.Dimension.Address];

            // 結果を取得
            var result = new List<Tanka>();
            for (var row = 2; row <= usedRange.Rows; row++)
            {
                // 社員番号が空なら終わる
                if (string.IsNullOrEmpty(ws.Cells[row, 1].Text))
                {
                    break;
                }

                result.Add(new Tanka()
                {
                    ShainBango = int.Parse(ws.Cells[row, 1].Text),
                    ShainName = ws.Cells[row, 2].Text,
                    SetteiTanka = decimal.Parse(ws.Cells[row, 3].Text),
                });
            }

            return result;
        }

        /// <summary>
        /// シート名から数値型の月を得る
        /// </summary>
        static (string hankakuSheet, string zenkakuSheet) GetTsukiSheetName(int tsuki)
        {
            return tsuki switch
            {
                7 => ("7月", "７月"),
                8 => ("8月", "８月"),
                9 => ("9月", "９月"),
                10 => ("10月", "１０月"),
                11 => ("11月", "１１月"),
                12 => ("12月", "１２月"),
                1 => ("1月", "１月"),
                2 => ("2月", "２月"),
                3 => ("3月", "３月"),
                4 => ("4月", "４月"),
                5 => ("5月", "５月"),
                6 => ("6月", "６月"),
                _ => throw new NotImplementedException(),
            };
        }
    }
}