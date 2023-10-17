module GelQC

// Import necessary modules and namespaces
open System
open OfficeOpenXml
open DocumentFormat.OpenXml.Packaging
open System.IO
open HelperFunctions

// Define a function named gelQcStart with four parameters
let gelQcStart(inputParams : string list) (gelForm : string) (ghost : ExcelWorksheet)(myTools : ExcelWorksheet) =

    // Capture the username of the current environment
    let user =  Environment.UserName 

    // Loop over each parameter in the inputParams list
    for param in inputParams do 

        // Call a function to retrieve data from LIMS
        let negative = gelsListFunction "gelqc" myTools 2

        // Read the content of a file specified by gelForm path
        let docArray = File.ReadAllBytes(gelForm)

        // Create a memory stream and initialize it with the content of docArray
        use _copyDoc = new MemoryStream(docArray)

        // Open the WordprocessingDocument
        use gelDocument = WordprocessingDocument.Open(_copyDoc, true)

        // Get the body of the Word document
        let gelBody = gelDocument.MainDocumentPart.Document.Body

        // Extract lot, csName, species, customer, geneNumber, scale, formulation, and shipDate from LIMS
        let lot, csName, species, customer, geneNumber, scale, formulation, shipDate = (codesetIdentifiers param ghost)

        // Update specific sections of the Word document with extracted information
        (gelsCsInfoHeader gelBody 2 5).Text <- lot + " " + csName
        (gelsCsInfoHeader gelBody 2 12).Text <- geneNumber.ToString()
        (gelsCsInfoHeader gelBody 2 17).Text <- scale.ToString()
        (gelsTableFiller gelBody 0 1 2 0).Text <- negative

        // Define the path for the Gel Batch Record document PLACEHOLDER
        let gelBatchRecordPath = "C:/Users/" + user + "/AppData/Local/Temp/ " + param + " Gel Batch Record" + ".docx"

        // Save and close the modified Word document
        gelDocument.SaveAs(gelBatchRecordPath).Close() |> ignore
