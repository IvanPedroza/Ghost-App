module ReQcHelpers

open System
open OfficeOpenXml
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Wordprocessing
open System.IO
open System.Linq


//finds ID values of cs builds in excel doc and assigns them to their respective variables
let codesetIdentifiers (param : string) (sheetName : ExcelWorksheet) =
    let list = List.init sheetName.Dimension.End.Row (fun i -> (i+1,1))
    let coordinates = List.find (fun (row,col) -> param.Equals ((string sheetName.Cells.[row,col].Value).Trim(), StringComparison.InvariantCultureIgnoreCase)) list
    let row, _colnum = coordinates

    let lot = sheetName.Cells.[row, 1].Value |> string
    let csName = sheetName.Cells.[row, 2].Value |> string
    let geneNumber =  sheetName.Cells.[row, 5].Value |> string
    let scale =  sheetName.Cells.[row, 7].Value |> string

    lot, csName, geneNumber, scale


    //writes cs identifying info at top of documents
let gelsCsInfoHeader (body : Body) paragraphIndex runIndex =
    let paragraph = body.Elements<Paragraph>().ElementAt(paragraphIndex)
    let run = paragraph.Elements<Run>().ElementAt(runIndex)
    let runProperties = run.Elements<RunProperties>().First()
    let underline = runProperties.AppendChild<Underline>(new Underline(Val = EnumValue<UnderlineValues>DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single))
    let position = runProperties.AppendChild<Position>(new Position(Val = StringValue("4")))
    let fontSize = runProperties.AppendChild<FontSize>(new FontSize(Val = StringValue("20")))
    run.Elements<RunProperties>().Equals(fontSize) |> ignore
    run.Elements<RunProperties>().Equals(underline) |>ignore
    run.Elements<RunProperties>().Equals(position) |> ignore
    let text = run.AppendChild(new Text())
    text

    //writes on body table cells of documents
let gelsTableFiller (body : Body) tableIndex rowIndex cellIndex paragraphIndex =
    let lot = body.Elements<Table>().ElementAt(tableIndex)
    let tableRow = lot.Elements<TableRow>().ElementAt(rowIndex)
    let tableCell = tableRow.Elements<TableCell>().ElementAt(cellIndex)
    let paragraph = tableCell.Elements<Paragraph>().ElementAt(paragraphIndex)
    let run = paragraph.AppendChild(new Run())
    let runProperties = run.AppendChild(new RunProperties())
    let fontSize = runProperties.AppendChild<FontSize>(new FontSize(Val = StringValue("20")))
    run.Elements<RunProperties>().Equals(fontSize) |> ignore
    let text = run.AppendChild(new Text())
    text

//Creates empty list - finds the Excel cell location for input and deconstructs the tuple into row and column numbers
let gelsListFunction (item : string) (sheetName : ExcelWorksheet) columnIndex =
    let list = List.init sheetName.Dimension.End.Row (fun i -> (i+1,1))
    let coordinates = List.find (fun (row,col) -> item.Equals ((string sheetName.Cells.[row,col].Value).Trim(), StringComparison.InvariantCultureIgnoreCase)) list
    let row, _colnum = coordinates
    let value = sheetName.Cells.[row,columnIndex].Value |> string
    value

