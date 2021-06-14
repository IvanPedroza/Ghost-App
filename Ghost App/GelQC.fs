module GelQC

open System
open OfficeOpenXml
open DocumentFormat.OpenXml.Packaging
open System.IO
open GelHelpers


let gelQcStart(inputParams : string list) (gelForm : string) (ghost : ExcelWorksheet)(myTools : ExcelWorksheet) =
    let user =  Environment.UserName 
    for param in inputParams do 

        let negative = gelsListFunction "gelqc" myTools 2

        let docArray = File.ReadAllBytes(gelForm)
        use _copyDoc = new MemoryStream(docArray)
        use gelDocument = WordprocessingDocument.Open(_copyDoc, true)
        let gelBody = gelDocument.MainDocumentPart.Document.Body

        let lot, csName, species, customer, geneNumber, scale, formulation, shipDate = (codesetIdentifiers param ghost)


        let plateCount = System.Math.Floor((geneNumber|>float) / 96.0)

        let totalGenesToGel =
            if plateCount > 1.0 then 
                let unGeledGenes = 96.0 * plateCount
                let genesToGel = (geneNumber|>float) - unGeledGenes
                genesToGel.ToString() + "/" + geneNumber
            else 
                if param.EndsWith("RW", StringComparison.InvariantCultureIgnoreCase) then 
                    Console.WriteLine ("How many probes are you geling for " + param + "?")
                    let genesToGel = Console.ReadLine ()
                    genesToGel + "/" + geneNumber
                else
                    geneNumber

        (gelsCsInfoHeader gelBody 2 5).Text <- lot + " " + csName
        (gelsCsInfoHeader gelBody 2 12).Text <- totalGenesToGel
        (gelsCsInfoHeader gelBody 2 17).Text <- scale
        (gelsTableFiller gelBody 0 1 2 0).Text <- negative

        let gelBatchRecordPath = "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Gel Batch Record" + ".docx"
        gelDocument.SaveAs(gelBatchRecordPath).Close() |> ignore
