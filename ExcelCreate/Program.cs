using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

internal class Program
{
    private static void C(string[] args)
    {
        // Criando um novo arquivo excel
        using var spreadsheetDocument = SpreadsheetDocument.Create("example .xlsx", SpreadsheetDocumentType.Workbook);
        // Adicione um WorkbookPart ao documento.
        WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        // Adicione um WorksheetPart ao WorkbookPart.
        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        // Adicione planilhas à pasta de trabalho.
        Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

        // Anexe uma nova planilha e associe-a à pasta de trabalho.
        Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
        sheets.Append(sheet);

        workbookPart.Workbook.Save();

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
        worksheetPart.Worksheet.AppendChild(headerRow);

        // Criar dados de exemplo
        var data = new[]
        {
                    new
                    {
                        Id = Guid.NewGuid(),
                        Cliente = "John Doe",
                        IdCliente = Guid.NewGuid(),
                        Produto = "Product 1",
                        IdProduto = Guid.NewGuid(),
                        ValorDoProduto = 10.99m,
                        Quantidade = 2,
                        DataDeCompra = DateTime.Now,
                        Ativo = true,
                        ElegivelDevolucao = false,
                        Entregue = true,
                        DataRecebimento = DateTime.Now.AddDays(3),
                    },
                    new
                    {
                        Id = Guid.NewGuid(),
                        Cliente = "Jane Doe",
                        IdCliente = Guid.NewGuid(),
                        Produto = "Product 2",
                        IdProduto = Guid.NewGuid(),
                        ValorDoProduto = 5.99m,
                        Quantidade = 3,
                        DataDeCompra = DateTime.Now.AddDays(-1),
                        Ativo = false,
                        ElegivelDevolucao = true,
                        Entregue = false,
                        DataRecebimento = DateTime.Now.AddDays(5),
                    }
                };

        // Criar linhas com os dados
        foreach (var item in data)
        {
            var row = new Row();
            row.AppendChild(new Cell { CellValue = new CellValue(item.Id.ToString()), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue(item.Cliente), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue(item.IdCliente.ToString()), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue(item.Produto), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue(item.IdProduto.ToString()), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue(item.ValorDoProduto.ToString("F2")), DataType = CellValues.Number });
            row.AppendChild(new Cell { CellValue = new CellValue(item.Quantidade.ToString()), DataType = CellValues.Number });
            row.AppendChild(new Cell { CellValue = new CellValue(item.DataDeCompra.ToString("dd.MM.yyyy HH:mm:ss")), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue(item.Ativo ? "Sim" : "Não"), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue(item.ElegivelDevolucao ? "Sim" : "Não"), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue(item.Entregue ? "Sim" : "Não"), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue(item.DataRecebimento.ToString("dd.MM.yyyy HH:mm:ss")), DataType = CellValues.String });
            row.AppendChild(new Cell { CellValue = new CellValue((item.Quantidade * item.ValorDoProduto).ToString("F2")), DataType = CellValues.Number });
            worksheetPart.Worksheet.AppendChild(row);
        }
    }
}