module ZagQC

open System
open OfficeOpenXml
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open System.IO
open ZagHelpers




let zagStart (inputParams : string list) (zagForm : string) (ghost : ExcelWorksheet)(myTools : ExcelWorksheet) =
    let user =  Environment.UserName
    Console.WriteLine "How many plates are you running?"
    let plateInput = Console.ReadLine() |> float


    Console.WriteLine "Which reagents will you use?"
    let reagentsInput = Console.ReadLine ()
    
    let gelVolume = (((plateInput - 1.0) * 5.0) + 20.0).ToString()
    let ureaVolume = (plateInput * 2.625).ToString()
    let diVolume = (plateInput * 315.0).ToString()
    let tenMerVolume = (plateInput * 60.0).ToString()

    for param in inputParams do
         
        let lot, csName, species, customer, geneNumber, scale, formulation, shipDate = (codesetIdentifiers param ghost)

        let gelLot = zagReagentsList reagentsInput myTools 2 
        let sybrLot = zagReagentsList reagentsInput myTools 3
        let ibLot = zagReagentsList reagentsInput myTools 4
        let ccLot = zagReagentsList reagentsInput myTools 5
        let ureaLot = zagReagentsList reagentsInput myTools 6
        let tenmerLot = zagReagentsList reagentsInput myTools 7
        let mpLot = zagReagentsList reagentsInput myTools 8 

        let docArray = File.ReadAllBytes(zagForm)
        use docCopy = new MemoryStream(docArray)
        use zagDocument = WordprocessingDocument.Open (docCopy, true)
        let zagBody = zagDocument.MainDocumentPart.Document.Body

        let numberOfPlates = (geneNumber |> float) / 96.0 
        let totalPlates = Math.Ceiling(numberOfPlates)
        let lotPlates = 
            if totalPlates = 1.0 then 
                (zagCsInfoHeader zagBody 2 11).Text <- geneNumber.ToString()
                "1"

            else
                Console.WriteLine ("For " + param + " are you running the last plate?")
                let answer = Console.ReadLine()
                if answer = "yes" then
                    (zagCsInfoHeader zagBody 2 11).Text <- geneNumber.ToString()
                    "1 - " + totalPlates.ToString()
                else       
                    let qcPlateNumber = System.Math.Round((geneNumber |> float) / 96.0)
                    let genesForQc = (96.0 * qcPlateNumber)
                    (zagCsInfoHeader zagBody 2 11).Text <- genesForQc.ToString()
                    "1 - " + (totalPlates - 1.0).ToString()
         

        //CS lot identifier info and number of plates being run
        (zagCsInfoHeader zagBody 1 5).Text <- lot + " " + csName
        (zagCsInfoHeader zagBody 1 13).Text <- lotPlates
        (zagCsInfoHeader zagBody 1 20).Text <- totalPlates |> string
        (zagCsInfoHeader zagBody 2 17).Text <- geneNumber |> string
        (zagCsInfoHeader zagBody 2 21).Text <- scale |> string

        //finds cell of each reagent
        (zagLotNumberFiller zagBody 0 1 4 0).Text <- gelLot
        (zagLotNumberFiller zagBody 0 2 4 0).Text <- sybrLot
        (zagLotNumberFiller zagBody 0 3 4 0).Text <- ibLot
        (zagLotNumberFiller zagBody 0 4 4 0).Text <- ccLot
        (zagLotNumberFiller zagBody 0 5 4 0).Text <- ureaLot
        (zagLotNumberFiller zagBody 0 7 4 0).Text <- tenmerLot
        (zagLotNumberFiller zagBody 0 8 4 0).Text <- mpLot
        (zagLotNumberFiller zagBody 0 9 4 0).Text <- "N/A"

        //Calculations text and footnotes
        let gelCalculations = (zagCalculationsWriter zagBody 1 2 1 2 8)
        let gelFootNote = (writeFootNote zagBody 1 2 1 2 9)
        let sybrCalculations = (zagCalculationsWriter zagBody 1 2 1 4 6)
        let sybrFootNote = (writeFootNote zagBody 1 2 1 4 7)
        let ureaCalculations = (zagCalculationsWriter zagBody 1 7 1 2 2)
        let ureaFootNote = (writeFootNote zagBody 1 7 1 2 3)
        let diCalculations = (zagCalculationsWriter zagBody 1 7 1 4 5)
        let diFootNote = (writeFootNote zagBody 1 7 1 4 6)
        let tenMerCalculations = (zagCalculationsWriter zagBody 1 7 1 6 11)
        let tenMerFootNote = (writeFootNote zagBody 1 7 1 6 12)
         

        //Adds footnote to calculations section and comment section
        let firstLot = inputParams.[0]
        let restOfList = inputParams.[1..inputParams.Length-1]
        if inputParams.Length > 1 then       
            if param = firstLot then
                gelCalculations.Text <- gelVolume
                gelFootNote.Text <- "①"
                sybrCalculations.Text <- gelVolume
                sybrFootNote.Text <- "①"
                ureaCalculations.Text <- ureaVolume
                ureaFootNote.Text <- "①"
                diCalculations.Text <- diVolume
                diFootNote.Text <- "①"
                tenMerCalculations.Text <- tenMerVolume
                tenMerFootNote.Text <- "①"
            else
                gelFootNote.Text <- "①"
                sybrFootNote.Text <- "①"
                ureaFootNote.Text <- "①"
                diFootNote.Text <- "①"
                tenMerFootNote.Text <- "①"
            zagNote zagBody inputParams param
        else
            gelCalculations.Text <- gelVolume
            sybrCalculations.Text <- gelVolume
            ureaCalculations.Text <- ureaVolume
            diCalculations.Text <- diVolume
            tenMerCalculations.Text <- tenMerVolume
            gelFootNote.Text <- ""
            sybrFootNote.Text <- ""
            ureaFootNote.Text <- ""
            diFootNote.Text <- ""
            tenMerFootNote.Text <- ""

        let zagBatchRecordPath = "C:/Users/" + user + "/AppData/Local/Temp/ "+param + " Zag Batch Record" + ".docx"
        zagDocument.SaveAs(zagBatchRecordPath).Close()