module HelperFunctions

open System
open OfficeOpenXml
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Wordprocessing
open System.IO
open System.Linq

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





let rqstFormDropdowns (body : Body) paragraphIndex sdtRunIndex sdtContentRunIndex runIndex textIndex =
    let paragraph = body.Elements<Paragraph>().ElementAt(paragraphIndex)
    let sdtRun = paragraph.Elements<SdtRun>().ElementAt(sdtRunIndex)
    let sdtContentRun =sdtRun.Elements<SdtContentRun>().First()
    let run = sdtContentRun.Elements<Run>().First()
    let text = run.Elements<Text>().First()
    text

//assings form parameter to read cs formulation and populate cs type
let formToCodeSetType (form : string) : string =
    match form with
    | "XT" -> "RNA"
    | "TBD" -> "TBD"
    | "DX" -> "RNA"
    | "STD" | "miRNA" -> "Panel/CodeSet Plus (RNA)"
    | _ -> failwith "Error ..."

//defining a function with two parameters "csname" and "form"
//function looks at string of csname for startsWith "string"
let determineFormulation (csname : string) (form : string) : string =
    match csname with 
        | csname when csname.StartsWith("CNV", StringComparison.CurrentCultureIgnoreCase)  -> "CNV (DNA)"
        | csname when csname.StartsWith("PLS", StringComparison.CurrentCultureIgnoreCase)  -> "Panel/CodeSet Plus (RNA)"
        | csname when csname.StartsWith("PLS_CNV", StringComparison.CurrentCultureIgnoreCase)  -> "Panel/CodeSet Plus (DNA)"
        | csname when csname.StartsWith("miR", StringComparison.CurrentCultureIgnoreCase)  -> "miRNA"
        | csname when csname.StartsWith("DNA", StringComparison.CurrentCultureIgnoreCase)  -> "DNA"
        | csname when csname.StartsWith("miX", StringComparison.CurrentCultureIgnoreCase)  -> "miRGE/miXED"
        | csname when csname.StartsWith("CHIP", StringComparison.CurrentCultureIgnoreCase)  -> "CHIP"
        | _ -> formToCodeSetType form

let requestForms (body : Body) tableIndex rowIndex cellIndex paragraphIndex=
    let lot = body.Elements<Table>().ElementAt(tableIndex)
    let tableRow = lot.Elements<TableRow>().ElementAt(rowIndex)
    let tableCell = tableRow.Elements<TableCell>().ElementAt(cellIndex)
    let paragraph = tableCell.Elements<Paragraph>().ElementAt(paragraphIndex)
    let run = paragraph.AppendChild(new Run())
    let runProperties = run.AppendChild(new RunProperties())
    let bold = runProperties.AppendChild<Bold>(new Bold())
    let font = runProperties.AppendChild<RunFonts>(new RunFonts(Ascii = StringValue("Cambria (Headings)")))
    let fontSize = runProperties.AppendChild<FontSize>(new FontSize(Val = StringValue("22")))
    run.Elements<RunProperties>().Equals(bold) |> ignore
    run.Elements<RunProperties>().Equals(font) |> ignore
    run.Elements<RunProperties>().Equals(fontSize) |> ignore
    let text = run.AppendChild(new Text())
    text

//Calling system time to extract year for formatting later
let year = (DateTime.Now.Year.ToString())






////Start of ligation BR functions
let ligationsCsInfoHeader (body : Body) paragraphIndex runIndex =
    let paragraph = body.Elements<Paragraph>().ElementAt(paragraphIndex)
    let run = paragraph.Elements<Run>().ElementAt(runIndex)
    let runProperties = run.Elements<RunProperties>().First()
    let underline = runProperties.AppendChild<Underline>(new Underline(Val = EnumValue<UnderlineValues>DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single))
    let position = runProperties.AppendChild<Position>(new Position(Val = StringValue("4")))
    run.Elements<RunProperties>().Equals(underline) |>ignore
    run.Elements<RunProperties>().Equals(position) |> ignore
    let text = run.AppendChild(new Text())
    text

let ligationsTableFiller (body : Body) tableIndex tableRowIndex tableCellIndex paragraphIndex runIndex= 
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
            

let footNoteSize (body : Body) tableIndex tableRowIndex tableCellIndex paragraphIndex runIndex = 
    let table = body.Elements<Table>().ElementAt(tableIndex)
    let row = table.Elements<TableRow>().ElementAt(tableRowIndex)
    let cell = row.Elements<TableCell>().ElementAt(tableCellIndex)
    let paragraph = cell.Elements<Paragraph>().ElementAt(paragraphIndex)
    let run = paragraph.Elements<Run>().ElementAt(runIndex)
    let runProperties = run.AppendChild(new RunProperties())
    let underline = runProperties.AppendChild<Underline>(new Underline(Val = EnumValue<UnderlineValues>DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single))
    run.Elements<RunProperties>().Equals(underline) |>ignore
    let fontSize = runProperties.AppendChild<FontSize>(new FontSize(Val = StringValue("8")))
    run.Elements<RunProperties>().Equals(fontSize) |> ignore
    let text = run.AppendChild(new Text())
    text

let oligoStamp (scale : float) : (string * float * float * float * float * float * float * string) = 
    match scale with
        |6.0 -> ("2.4", 2.4, 2.0, 6.0, 1.0, 5.7, 0.5,  "17.6")
        |7.5 -> ("3", 3.0, 2.5, 7.5, 1.25, 7.125, 0.625, "22") 
        |9.0 -> ("3.6", 3.6, 3.0, 9.0, 1.5, 8.55, 0.75, "26.4")  
        |10.5 -> ("4.2", 4.2, 3.5, 10.5, 1.75, 9.975, 0.875, "30.8") 
        |12.0 -> ("4.8", 4.8, 4.0, 12.0, 2.0, 11.4, 1.0, "35.2")
        |13.5 -> ("5.4", 5.4, 4.5, 13.5, 2.25, 12.825, 1.125, "39.6")
        |15.0 -> ("6", 6.0, 5.0, 15.0, 2.5, 14.25, 1.25, "44")
        |18.0 -> ("7.2", 7.2, 6.0, 18.0, 3.0, 17.1, 1.5, "52.8")
        |21.0 -> ("8.4", 8.4, 7.0, 21.0, 3.5, 19.95, 1.75, "61.6")
        |60.0 -> ("24", 24.0, 20.0, 60.0, 10.0, 57.0, 5.0, "176")
        |_ -> failwith "Error..."

let footnotes (body : Body) (inputParams : string list) (param : string) : unit =
    if inputParams.Length > 1 then 
        (footNoteSize body 0 30 1 7 2).Text <- "①"
        (footNoteSize body 0 30 1 7 10).Text <- "①"
        (footNoteSize body 0 30 1 10 2).Text <- "①"
        (footNoteSize body 0 30 1 10 6).Text <- "①"
        (footNoteSize body 0 30 1 10 10).Text <- "①"
        (footNoteSize body 0 31 1 0 4).Text <- "①"
        (footNoteSize body 0 31 1 1 4).Text <- "①"
        (footNoteSize body 0 31 1 2 4).Text <- "①"
        (footNoteSize body 0 31 1 3 5).Text <- "①"
        (footNoteSize body 0 31 1 4 3).Text <- "①"
        (footNoteSize body 0 31 1 5 3).Text <- "①"
        (footNoteSize body 0 31 1 6 3).Text <- "①"
        (footNoteSize body 0 31 1 8 5).Text <- "①"
        let lastLot = inputParams.Last()
        if param = lastLot then
            let restOfList = inputParams.[0..inputParams.Length - 2]  |> String.concat ", "
            let note = "① Calculations include " + restOfList + "."
            (ligationsTableFiller body 0 38 0 2 0).Text <- note
        else
            let note = "① Calculations are on " + lastLot + "."
            (ligationsTableFiller body 0 38 0 2 0).Text <- note

//Creates empty list - finds the Excel cell location for input and deconstructs the tuple into row and column numbers
let ligationsListFunction (item : string) (sheetName : ExcelWorksheet) columnIndex =
    let list = List.init 100 (fun i -> (i+1,1))
    let coordinates = List.find (fun (row,col) -> item.Equals ((string sheetName.Cells.[row,col].Value).Trim(), StringComparison.InvariantCultureIgnoreCase)) list
    let row, _colnum = coordinates
    let value = sheetName.Cells.[row,columnIndex].Value |> string
    value

let roundupbyfive(i) : float = 
    (System.Math.Ceiling(i / 5.0) * 5.0)







///gel qc specific functions
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

let gelsListFunction (item : string) (sheetName : ExcelWorksheet) columnIndex =
    let list = List.init 100 (fun i -> (i+1,1))
    let coordinates = List.find (fun (row,col) -> item.Equals ((string sheetName.Cells.[row,col].Value).Trim(), StringComparison.InvariantCultureIgnoreCase)) list
    let row, _colnum = coordinates
    let value = sheetName.Cells.[row,columnIndex].Value |> string
    value





/////GP zag specific forms

let zagCsInfoHeader (body : Body) paragraphIndex runIndex =
    let paragraph = body.Elements<Paragraph>().ElementAt(paragraphIndex)
    let run = paragraph.Elements<Run>().ElementAt(runIndex)
    let text = run.AppendChild(new Text())
    text

let zagReagentsList (item : string) (sheetName : ExcelWorksheet) columnIndex =
    let list = List.init 100 (fun i -> (i+1,1))
    let coordinates = List.find (fun (row,col) -> item.Equals ((string sheetName.Cells.[row,col].Value).Trim(), StringComparison.InvariantCultureIgnoreCase)) list
    let row, _colnum = coordinates
    let value = sheetName.Cells.[row,columnIndex].Value |> string
    value

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


let zagCalculationsWriter (body : Body) tableIndex rowIndex cellIndex paragraphIndex runIndex =
    let calc = body.Elements<Table>().ElementAt(tableIndex)
    let calcTableRow = calc.Elements<TableRow>().ElementAt(rowIndex)
    let calcTableCell = calcTableRow.Elements<TableCell>().ElementAt(cellIndex)
    let calcParagraph = calcTableCell.Elements<Paragraph>().ElementAt(paragraphIndex)
    let calcRun = calcParagraph.Elements<Run>().ElementAt(runIndex)
    let calcText = calcRun.AppendChild(new Text())
    calcText

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





//////purification form specific functions 

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



