using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;
using OfficeOpenXml;

public class FundoInvestimento
{
    public string CNPJ_FUNDO { get; set; }
    public DateTime DT_COMPTC { get; set; }
    public string TP_FUNDO { get; set; }
    public int NR_COTST { get; set; }
    public float VL_QUOTA { get; set; }
    public decimal VL_PATRIM_LIQ { get; set; }
    public decimal CAPTC_DIA { get; set; }
    public decimal RESG_DIA { get; set; } 

    public override string ToString()
    {
        return $"CNPJ: {CNPJ_FUNDO}, Data: {DT_COMPTC}, Tipo: {TP_FUNDO}, Valor da Cota: {VL_QUOTA}, Patrimonio Liquido: {VL_PATRIM_LIQ}, Captação: {CAPTC_DIA}, Resgates: {RESG_DIA}, Número de Cotistas: {NR_COTST}";
    }
}

public class Program
{
    private static readonly HttpClient client = new HttpClient();

    public static async Task Main(string[] args)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        do
        {
            string cnpj;
            do
            {
                Console.Write("Digite o CNPJ do fundo (ex: 00.017.024/0001-53): ");
                cnpj = Console.ReadLine();

                if (!CnpjValido(cnpj))
                {
                    Console.WriteLine("CNPJ inválido, verifique o formato e insira novamente.");
                }
            } while (!CnpjValido(cnpj));

            DateTime dataInicial;
            do
            {
                Console.Write("Digite a data inicial (ex: 01072023): ");
                if (!DateTime.TryParseExact(Console.ReadLine(), "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dataInicial))
                {
                    Console.WriteLine("Data inválida, verifique o formato e insira novamente.");
                }
            } while (dataInicial == default);

            DateTime dataFinal;
            do
            {
                Console.Write("Digite a data final (ex: 31072023): ");
                if (!DateTime.TryParseExact(Console.ReadLine(), "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dataFinal))
                {
                    Console.WriteLine("Data inválida, verifique o formato e insira novamente.");
                }
            } while (dataFinal == default);

            var fundos = new List<FundoInvestimento>();

            for (int ano = dataInicial.Year; ano <= dataFinal.Year; ano++)
            {
                for (int mes = dataInicial.Month; mes <= dataFinal.Month; mes++)
                {
                    var fundosDoMes = await BaixarDadosFundos(ano.ToString(), mes.ToString("D2"));
                    fundos.AddRange(fundosDoMes);
                }
            }

            // Filtrar fundos com cotistas, com o CNPJ especificado e entre as datas especificadas
            var fundosFiltrados = fundos.Where(f => f.NR_COTST >= 1 && f.CNPJ_FUNDO == cnpj && f.DT_COMPTC >= dataInicial && f.DT_COMPTC <= dataFinal).ToList();

            // Calcular a rentabilidade
            var valorInicial = fundosFiltrados.First().VL_QUOTA;
            var valorFinal = fundosFiltrados.Last().VL_QUOTA;
            var rentabilidade = ((valorFinal / valorInicial) - 1) * 100;

            // Calcular Captação no periodo
            var captacaoTotal = fundosFiltrados.Sum(f => f.CAPTC_DIA);

            // Calcular Resgates no periodo
            var resgatesTotal = fundosFiltrados.Sum(f => f.RESG_DIA);

            // Calcular Resultado Captação Líquida Aplicações - Resgates
            var captacaoLiquida = captacaoTotal - resgatesTotal;

            // Gerar nome de arquivo com data e hora atual
            string nomeArquivo = $"C:\\Users\\ojeff\\Downloads\\FundosFiltrados_{DateTime.Now:yyyyMMddHHmmss}.xlsx";

            // Escrever os dados filtrados em um arquivo Excel
            using (var package = new ExcelPackage(new FileInfo(nomeArquivo)))
            {
                var worksheet = package.Workbook.Worksheets.Add("Fundos");
                worksheet.Cells["A1"].LoadFromCollection(fundosFiltrados, PrintHeaders: true);
                worksheet.Cells["B:B"].Style.Numberformat.Format = "dd-mm-yyyy"; // Formato de data para coluna DT_COMPTC
                worksheet.Cells["E:E"].Style.Numberformat.Format = "#,##0.00"; // Formato de moeda para coluna VL_PATRIM_LIQ
                worksheet.Cells["G:G"].Style.Numberformat.Format = "#,##0.00"; // Formato de moeda para coluna CAPTC_DIA
                worksheet.Cells["H:H"].Style.Numberformat.Format = "#,##0.00"; // Formato de moeda para coluna RESG_DIA

                // Escrever Rentabilidade no Período
                worksheet.Cells["I1"].Value = "Rentabilidade no Período:";
                worksheet.Cells["I2"].Value = Math.Round(rentabilidade, 2); // Arredondar a rentabilidade para 2 casas decimais

                // Escrever Captação no Período
                worksheet.Cells["J1"].Value = "Captação no Período:";
                worksheet.Cells["J2"].Value = captacaoTotal;

                // Escrever Resgates no Período
                worksheet.Cells["K1"].Value = "Resgates no Período:";
                worksheet.Cells["K2"].Value = resgatesTotal;

                // Escrever Captação Líquida
                worksheet.Cells["L1"].Value = "Captação Líquida:";
                worksheet.Cells["L2"].Value = captacaoLiquida;

                package.Save();
            }

            Console.WriteLine($"Arquivo gerado com sucesso: {nomeArquivo}");

            // Perguntar ao usuário se deseja gerar uma nova solicitação ou encerrar o aplicativo
            Console.Write("Deseja gerar nova solicitação? (S/N): ");
        } while (Console.ReadLine().ToUpper() == "S");
    }

    public static async Task<List<FundoInvestimento>> BaixarDadosFundos(string ano, string mes)
    {
        var url = $"https://dados.cvm.gov.br/dados/FI/DOC/INF_DIARIO/DADOS/inf_diario_fi_{ano}{mes}.zip";
        var response = await client.GetAsync(url);

        if (!response.IsSuccessStatusCode)
        {
            // Se a resposta não for bem-sucedida (por exemplo, se o arquivo para o mês especificado não existir), retorne uma lista vazia
            return new List<FundoInvestimento>();
        }

        var zipPath = $"inf_diario_fi_{ano}{mes}.zip";
        await using (var stream = await response.Content.ReadAsStreamAsync())
        using (var fileStream = File.Create(zipPath))
        {
            await stream.CopyToAsync(fileStream);
        }

        using (var archive = ZipFile.OpenRead(zipPath))
        {
            var entry = archive.Entries[0];
            using (var reader = new StreamReader(entry.Open()))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture) { Delimiter = ";" }))
            {
                var fundos = csv.GetRecords<FundoInvestimento>().ToList();
                return fundos;
            }
        }
    }

    public static bool CnpjValido(string cnpj)
    {
        // Remover caracteres não numéricos
        cnpj = Regex.Replace(cnpj, "[^0-9]", "");

        // CNPJ deve ter 14 caracteres numéricos
        if (cnpj.Length != 14)
            return false;

        // Calcular os dois dígitos verificadores
        int[] multiplier1 = { 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2 };
        int[] multiplier2 = { 6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2 };

        string tempCnpj = cnpj.Substring(0, 12);

        int sum = 0;
        for (int i = 0; i < 12; i++)
        {
            sum += int.Parse(tempCnpj[i].ToString()) * multiplier1[i];
        }

        int remainder = sum % 11;
        remainder = remainder < 2 ? 0 : 11 - remainder;

        string digit1 = remainder.ToString();

        tempCnpj += digit1;

        sum = 0;
        for (int i = 0; i < 13; i++)
        {
            sum += int.Parse(tempCnpj[i].ToString()) * multiplier2[i];
        }

        remainder = sum % 11;
        remainder = remainder < 2 ? 0 : 11 - remainder;

        string digit2 = remainder.ToString();

        tempCnpj += digit2;

        // Verificar se o CNPJ original é igual ao CNPJ com os dígitos verificadores calculados
        return cnpj == tempCnpj;
    }
}
