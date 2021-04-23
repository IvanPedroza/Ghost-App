module ZagHelpers

open System
open OfficeOpenXml
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Wordprocessing
open System.Linq


//finds ID values of cs builds in excel doc and assigns them to their respective variables
let codesetIdentifiers (param : string) (sheetName : ExcelWorksheet) =
    let list = List.init 100 (fun i -> (i+1,1))
    let coordinates = List.find (fun (row,col) -> param.Equals ((string sheetName.Cells.[row,col].Value).Trim(), StringComparison.InvariantCultureIgnoreCase)) list
    let row, _colnum = coordinates

    let lot = sheetName.Cells.[row, 1].Value |> string
    let csName = sheetName.Cells.[row, 2].Value |> string
    let species = sheetName.Cells.[row, 3].Value |> string
    let customer =  sheetName.Cells.[row, 4].Value |> string
    let geneNumber =  sheetName.Cells.[row, 5].Value |> string
    let scale =  sheetName.Cells.[row, 7].Value |> string
    let formulation = sheetName.Cells.[row,  9].Value|> string
    let shipDate =  sheetName.Cells.[row, 10].Value |> string
    lot, csName, species, customer, geneNumber, scale, formulation, shipDate

    //writes cs identifying info at top of documents
let zagCsInfoHeader (body : Body) paragraphIndex runIndex =
    let paragraph = body.Elements<Paragraph>().ElementAt(paragraphIndex)
    let run = paragraph.Elements<Run>().ElementAt(runIndex)
    let text = run.AppendChild(new Text())
    text

    //Creates empty list - finds the Excel cell location for input and deconstructs the tuple into row and column numbers
let zagReagentsList (item : string) (sheetName : ExcelWorksheet) columnIndex =
    let list = List.init 100 (fun i -> (i+1,1))
    let coordinates = List.find (fun (row,col) -> item.Equals ((string sheetName.Cells.[row,col].Value).Trim(), StringComparison.InvariantCultureIgnoreCase)) list
    let row, _colnum = coordinates
    let value = sheetName.Cells.[row,columnIndex].Value |> string
    value

    //writes on reagents table of zag documents
let zagLotNumberFiller (body : Body) tableIndex rowIndex cellIndex paragraphIndex =
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

    //writes on body table cells of zag documents
let zagCalculationsWriter (body : Body) tableIndex rowIndex cellIndex paragraphIndex runIndex =
    let calc = body.Elements<Table>().ElementAt(tableIndex)
    let calcTableRow = calc.Elements<TableRow>().ElementAt(rowIndex)
    let calcTableCell = calcTableRow.Elements<TableCell>().ElementAt(cellIndex)
    let calcParagraph = calcTableCell.Elements<Paragraph>().ElementAt(paragraphIndex)
    let calcRun = calcParagraph.Elements<Run>().ElementAt(runIndex)
    let calcText = calcRun.AppendChild(new Text())
    calcText

    //formats the text size of footnote symbol
let writeFootNote (body : Body) tableIndex rowIndex cellIndex paragraphIndex runIndex =
    let table = body.Elements<Table>().ElementAt(tableIndex)
    let row = table.Elements<TableRow>().ElementAt(rowIndex)
    let cell = row.Elements<TableCell>().ElementAt(cellIndex)
    let paragraph = cell.Elements<Paragraph>().ElementAt(paragraphIndex)
    let run = paragraph.Elements<Run>().ElementAt(runIndex)
    let runProperties = run.Elements<RunProperties>().First()
    let fontSize = runProperties.AppendChild<FontSize>(new FontSize(Val = StringValue("10")))
    run.Elements<RunProperties>().Equals(fontSize) |> ignore
    let text = run.AppendChild(new Text())
    text

    //adds footnote symbol to the same location in all docs when conditionals are met
let zagNote (body : Body) (inputParams : string list) (param : string) : unit =
    let firstLot = inputParams.[0]

    let table = body.Elements<Table>().ElementAt(1)
    let row = table.Elements<TableRow>().ElementAt(23)
    let cell = row.Elements<TableCell>().ElementAt(0)
    let paragraph = cell.Elements<Paragraph>().ElementAt(4)
    let run = paragraph.AppendChild (new Run())
    let runProperties = run.AppendChild(new RunProperties())
    let fontSize = runProperties.AppendChild<FontSize>(new FontSize(Val = StringValue("18")))
    run.Elements<RunProperties>().Equals(fontSize) |> ignore
    let text = run.AppendChild (new Text())

    if param = firstLot then
        let restOfList = inputParams.[1..inputParams.Length-1] |> String.concat ", "
        let theFirstNote = "① Calculations include " + restOfList + "."
        text.Text <- theFirstNote
    else
        let otherNote = "① Calculations are on " + firstLot.ToUpper() + "."
        text.Text <- otherNote


