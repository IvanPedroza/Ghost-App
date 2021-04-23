module PurificationHelpers

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
let purificationCsInfoHeader (body : Body) paragraphIndex runIndex =
    let paragraph = body.Elements<Paragraph>().ElementAt(paragraphIndex)
    let run = paragraph.Elements<Run>().ElementAt(runIndex)
    let runProperties = run.Elements<RunProperties>().First()
    let underline = runProperties.AppendChild<Underline>(new Underline(Val = EnumValue<UnderlineValues>DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single))
    let position = runProperties.AppendChild<Position>(new Position(Val = StringValue("2")))
    run.Elements<RunProperties>().Equals(underline) |>ignore
    run.Elements<RunProperties>().Equals(position) |> ignore
    let text = run.AppendChild(new Text())
    text

    //writes on body table cells of documents
let writingCalculations (body : Body) tableIndex tableRowIndex tableCellIndex paragraphIndex runIndex = 
    let table = body.Elements<Table>().ElementAt(tableIndex)
    let row = table.Elements<TableRow>().ElementAt(tableRowIndex)
    let cell = row.Elements<TableCell>().ElementAt(tableCellIndex)
    let paragraph = cell.Elements<Paragraph>().ElementAt(paragraphIndex)
    let run = paragraph.Elements<Run>().ElementAt(runIndex)
    let runProperties = run.AppendChild(new RunProperties())
    let underline = runProperties.AppendChild<Underline>(new Underline(Val = EnumValue<UnderlineValues>DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single))
    run.Elements<RunProperties>().Equals(underline) |>ignore 
    let fontSize = runProperties.AppendChild<FontSize>(new FontSize(Val = StringValue("18")))
    run.Elements<RunProperties>().Equals(fontSize) |> ignore
    let text = run.AppendChild(new Text())
    text

    //writes on reagents table of zag documents
let fillingPurificationLots (body : Body) tableIndex tableRowIndex tableCellIndex paragraphIndex runIndex = 
    let table = body.Elements<Table>().ElementAt(tableIndex)
    let row = table.Elements<TableRow>().ElementAt(tableRowIndex)
    let cell = row.Elements<TableCell>().ElementAt(tableCellIndex)
    let paragraph = cell.Elements<Paragraph>().ElementAt(paragraphIndex)
    if not (paragraph.Elements<Run>().Any()) then
        let run = paragraph.AppendChild<Run>(new Run())
        let runProperties = run.AppendChild(new RunProperties())
        let fontSize = runProperties.AppendChild<FontSize>(new FontSize(Val = StringValue("18")))
        run.Elements<RunProperties>().Equals(fontSize) |> ignore
        let text = run.AppendChild<Text>(new Text())
        text
    else
        let run = paragraph.Elements<Run>().ElementAt(runIndex)
        let runProperties = run.AppendChild(new RunProperties())
        let fontSize = runProperties.AppendChild<FontSize>(new FontSize(Val = StringValue("18")))
        run.Elements<RunProperties>().Equals(fontSize) |> ignore
        let text = run.AppendChild(new Text())
        text

//Creates empty list and finds thje Excel cell location for input and deconstructs the tuple into row and column numbers
let purificationReagentsList (item : string) (sheetName : ExcelWorksheet) columnIndex =
    let list = List.init 1000000 (fun i -> (i+1,1)) 
    let coordinates = List.find (fun (row,col) -> item.Equals((string sheetName.Cells.[row,col].Value).Trim(), StringComparison.InvariantCultureIgnoreCase)) list
    let row, _colnum = coordinates
    let value = sheetName.Cells.[row,columnIndex].Value |> string
    value

    //calculates theoretical volumes of cs specified in the SOP
let theoreticalVolume (scale : float) (geneNumeber : float) : float = 
    match scale with 
        | 6.0 -> geneNumeber * 17.0
        | 7.5 -> geneNumeber * 22.0
        | 9.0 -> geneNumeber * 27.0
        | 10.5 -> geneNumeber * 32.0
        | 12.0 -> geneNumeber * 37.0
        | 13.5 -> geneNumeber * 42.0
        | 15.0 -> geneNumeber * 47.0
        | 18.0 -> geneNumeber * 57.0
        | 21.0 -> geneNumeber * 67.0
        | 60.0 -> geneNumeber * 197.0
        | 180.0 -> geneNumeber * 591.0
        | _ -> failwith "Error..."