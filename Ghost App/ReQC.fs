module ReQC

open System
open OfficeOpenXml
open DocumentFormat.OpenXml.Packaging
open System.IO
open HelperFunctions


let reQcStart(inputParams : string list) (reQcForm : string) (ghost : ExcelWorksheet)(myTools : ExcelWorksheet) =
    let user =  Environment.UserName 
    for param in inputParams do 

        Console.WriteLine ("How many probes are being reQC-ed for " + param)
        let reQcGeneNumber = Console.ReadLine()

        let negative = gelsListFunction "gelqc" myTools 2

        let docArray = File.ReadAllBytes(reQcForm)
        use _copyDoc = new MemoryStream(docArray)
        use reQcDocument = WordprocessingDocument.Open(_copyDoc, true)
        let reQCBody = reQcDocument.MainDocumentPart.Document.Body

        let lot, csName, species, customer, geneNumber, scale, formulation, shipDate = (codesetIdentifiers param ghost)

        (gelsCsInfoHeader reQCBody 2 3).Text <- lot + " " + csName
        (gelsCsInfoHeader reQCBody 2 8).Text <- reQcGeneNumber
        (gelsCsInfoHeader reQCBody 2 12).Text <- geneNumber.ToString()
        (gelsCsInfoHeader reQCBody 2 18).Text <- scale.ToString() + " pmol"
        (gelsTableFiller reQCBody 0 2 4 0).Text <- negative

        let reQcBatchRecordPath = "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " reQC Batch Record" + ".docx"
        reQcDocument.SaveAs(reQcBatchRecordPath).Close() |> ignore