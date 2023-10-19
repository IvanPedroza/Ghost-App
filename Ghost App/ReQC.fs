module ReQC

open System
open OfficeOpenXml
open DocumentFormat.OpenXml.Packaging
open System.IO
open HelperFunctions

// Used to fill out Batch Records of CS lots that need a second QC run
let reQcStart(inputParams : string list) (reQcForm : string) (ghost : ExcelWorksheet)(myTools : ExcelWorksheet) =
    // Gets current user for use in path and error logging
    let user =  Environment.UserName 

    // Cycles through specified CS lots
    for param in inputParams do 

        // User interface
        Console.WriteLine ("How many probes are being reQC-ed for " + param)
        let reQcGeneNumber = Console.ReadLine()

        // Gets reagent lots from LIMS
        let negative = gelsListFunction "gelqc" myTools 2

        // Reads in Batch Record template
        let docArray = File.ReadAllBytes(reQcForm)
        use _copyDoc = new MemoryStream(docArray)
        use reQcDocument = WordprocessingDocument.Open(_copyDoc, true)
        let reQCBody = reQcDocument.MainDocumentPart.Document.Body

        // Reads in CS indentifying information from LIMS
        let lot, csName, species, customer, geneNumber, scale, formulation, shipDate = (codesetIdentifiers param ghost)

        // Fills out CS indentifying information on Batch Record template
        (gelsCsInfoHeader reQCBody 2 3).Text <- lot + " " + csName
        (gelsCsInfoHeader reQCBody 2 8).Text <- reQcGeneNumber
        (gelsCsInfoHeader reQCBody 2 12).Text <- geneNumber.ToString()
        (gelsCsInfoHeader reQCBody 2 18).Text <- scale.ToString() + " pmol"
        (gelsTableFiller reQCBody 0 2 4 0).Text <- negative

        // Saves filled out Batch Record to temp folder for pringting and subsequent deletion
        let reQcBatchRecordPath = "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " reQC Batch Record" + ".docx"
        reQcDocument.SaveAs(reQcBatchRecordPath).Close() |> ignore
