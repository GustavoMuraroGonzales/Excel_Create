using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

internal class Program
{
    
    private static void Main(string[] args)
    {
        // Diretório onde o arquivo será salvo
        string diretorio = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), $"TesteDeExcel_{DateTime.Now:dd.MM.yyyy}_{DateTime.Now:HH.mm.ss}.xlsx");

        SpreadsheetDocument spreadsheetDocument;
        WorkbookPart workbookPart;
        CriaçãoExcel(diretorio, out spreadsheetDocument, out workbookPart);

        workbookPart.Workbook.Save();
        spreadsheetDocument.Dispose();


    }

    public static void CriaçãoExcel(string diretorio, out SpreadsheetDocument spreadsheetDocument, out WorkbookPart workbookPart)
    {
        // Criando um novo arquivo excel
        spreadsheetDocument = SpreadsheetDocument.Create(diretorio, SpreadsheetDocumentType.Workbook);

        // Adicione um WorkbookPart ao documento.
        workbookPart = spreadsheetDocument.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        // Adicione um WorksheetPart ao WorkbookPart.
        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        // Adicione planilhas à pasta de trabalho.
        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

        // Anexe uma nova planilha e associe-a à pasta de trabalho.
        Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
        sheets.Append(sheet);

        // Adicionar alguns dados à planilha
        SheetData? sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();


        // Criando Cabeçalhos
        var headers = new[]
        {
                    "ID",
                    "Cliente",
                    "ID Cliente",
                    "Produto",
                    "IdProduto",
                    "Valor do Produto",
                    "Quantidade",
                    "Data de compra",
                    "Ativo",
                    "Elegivel Devolucao",
                    "Entregue",
                    "Data Recebimento",
                    "Valor Total"
                };

        // Criar linha de cabeçalho
        var headerRow = new Row();
        foreach (var header in headers)
        {
            headerRow.AppendChild(new Cell { CellValue = new CellValue(header), DataType = CellValues.String });
        }

        sheetData?.AppendChild(headerRow);

        // Criar dados de exemplo
        var data = new[]
        {
                    new Dados()
                    {
                        Id = Guid.NewGuid(),
                        IdCliente = Guid.NewGuid(),
                        NomeCliente = "John Doe",
                        Produto = "Product 1",
                        IdProduto = Guid.NewGuid(),
                        ValorProduto = 10.99m,
                        Quantidade = 2,
                        DataCompra = DateTime.Now,
                        Ativo = true,
                        ElegivelDevolucao = false,
                        Entregue = true,
                        DataRecebimento = DateTime.Now.AddDays(3),
                    },
                    new Dados()
                    {
                        Id = Guid.NewGuid(),    
                        IdCliente = Guid.NewGuid(),
                        NomeCliente = "Alinain Pick",
                        Produto = "Product 2",
                        IdProduto = Guid.NewGuid(),
                        ValorProduto = 25.33m,
                        Quantidade = 5,
                        DataCompra = DateTime.Now,
                        Ativo = true,
                        ElegivelDevolucao = false,
                        Entregue = true,
                        DataRecebimento = DateTime.Now.AddDays(3),
                    },
                    new Dados()
                    {
                        Id = Guid.NewGuid(),
                        IdCliente = Guid.NewGuid(),
                        NomeCliente = "Maria",
                        Produto = "Product 3",
                        IdProduto = Guid.NewGuid(),
                        ValorProduto = 15.99m,
                        Quantidade = 3,
                        DataCompra = DateTime.Now,
                        Ativo = true,
                        ElegivelDevolucao = false,
                        Entregue = true,
                        DataRecebimento = DateTime.Now.AddDays(3),
                        
                    }
                };

        // Criar linhas com os dados
        foreach (var item in data)
        {
            var row = new Row();
            row.AppendChild(new Cell { CellValue = new CellValue(item.Id.ToString()), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue(item.NomeCliente), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue(item.IdCliente.ToString()), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue(item.Produto), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue(item.IdProduto.ToString()), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue(item.ValorProduto.ToString("F2")), DataType = CellValues.Number });
            row.AppendChild(new Cell { CellValue = new CellValue(item.Quantidade.ToString()), DataType = CellValues.Number });
            row.AppendChild(new Cell { CellValue = new CellValue(item.DataCompra.ToString("dd.MM.yyyy HH:mm:ss")), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue(item.Ativo ? "Sim" : "Não"), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue(item.ElegivelDevolucao ? "Sim" : "Não"), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue(item.Entregue ? "Sim" : "Não"), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue(item.DataRecebimento.ToString("dd.MM.yyyy HH:mm:ss")), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue((item.Quantidade * item.ValorProduto).ToString("F2")), DataType = CellValues.Number });
            sheetData?.AppendChild(row);
        }
    }

    public class Dados()
    {
        public Guid Id { get; set; }    
        public Guid IdCliente { get; set; }
        public string? NomeCliente { get; set; }
        public string? Produto { get; set; }
        public Guid IdProduto { get; set; }
        public decimal ValorProduto { get; set; }
        public int Quantidade { get; set; }
        public DateTime DataCompra { get; set; }
        public bool Ativo { get; set; }
        public bool ElegivelDevolucao { get; set; }
        public bool Entregue { get; set; }
        public DateTime DataRecebimento { get; set; }
        public decimal ValorTotal { get; set; }

    }
}

